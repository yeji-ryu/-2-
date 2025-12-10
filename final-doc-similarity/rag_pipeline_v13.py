# -*- coding: utf-8 -*-
import argparse
import json
import re
import subprocess
import sys
import os
import numpy as np
from pathlib import Path
from qdrant_client import QdrantClient, models
from typing import List, Dict, Tuple, Optional

# --- Import functions from user-provided scripts ---
try:
    import pdf_parser
    import hwpx_parser
    import docx_parser
    import embedding
    import qdrant_dense
    import cosine
    import word
except ImportError as e:
    print(f"[오류] 필수 모듈 임포트 실패: {e}")
    sys.exit(1)

# --- Configuration ---
QDRANT_PATH = "./qdrant_storage"
DENSE_COLLECTION = "my_document_store"
SPARSE_COLLECTION = "bm25_chunks"
SPARSE_NAME = "bm25"
BM25_BUILD_DIR = Path("./bm25_index")
VECTOR_DIM = 1024
BM25_TOKENIZER = "okt"
BM25_PREPARE_DELIM = "***"

# 하이브리드 가중치
BM25_WEIGHT = 0.6
COSINE_WEIGHT = 0.4

# BM25 최대값 JSON
BM25_MAX_JSON = r"C:\Users\LG\HS\bm25_file_scores.json"

# 문서 메타 정보 파일
DOC_META_JSON = BM25_BUILD_DIR / "doc_meta.json"


###################################################
# Utility
###################################################

def truncate_text(text: str, max_chars: int = 1500) -> str:
    """LLM에 넘기는 텍스트 길이를 제한."""
    if not text:
        return ""
    text = text.strip()
    if len(text) <= max_chars:
        return text
    return text[:max_chars] + "\n...[이하 생략]..."


def parse_document(file_path: Path) -> str:
    ext = file_path.suffix.lower()
    print(f"[Parse] {file_path.name} (type: {ext})")
    try:
        if ext == ".pdf":
            return pdf_parser.parse_pdf_to_text(file_path)
        elif ext == ".hwpx":
            return hwpx_parser.parse_hwpx_to_text(file_path)
        elif ext == ".docx":
            return docx_parser.parse_docx_to_text(file_path)
        else:
            print(f"[경고] 지원하지 않는 형식: {ext}")
            return ""
    except Exception as e:
        print(f"[오류] {file_path.name} 파싱 실패: {e}")
        return ""


SENT_SPLIT_RE = re.compile(r"(?<=[.!?])\s+|[\r\n]+")

def chunk_text(text: str) -> List[str]:
    return [s.strip() for s in SENT_SPLIT_RE.split(text) if s.strip()]


def load_bm25_max_table() -> Dict[str, float]:
    if not os.path.exists(BM25_MAX_JSON):
        print(f"[오류] BM25 최대값 JSON 없음: {BM25_MAX_JSON}")
        return {}

    try:
        data = json.load(open(BM25_MAX_JSON, "r", encoding="utf-8"))
    except Exception as e:
        print(f"[오류] BM25 JSON 로드 실패: {e}")
        return {}

    table = {}
    for entry in data:
        doc_id = entry.get("best_doc_id")
        best_score = entry.get("best_score", 0.0)
        if doc_id:
            table[doc_id] = best_score
    return table


def load_doc_meta(build_dir: Path = BM25_BUILD_DIR) -> Dict[str, Dict[str, str]]:
    path = build_dir / "doc_meta.json"
    if not path.exists():
        print(f"[경고] doc_meta.json 없음")
        return {}

    try:
        data = json.load(open(path, "r", encoding="utf-8"))
    except:
        return {}

    out = {}
    for e in data:
        doc_id = e.get("doc_id")
        if doc_id:
            out[doc_id] = {
                "filename": e.get("filename", ""),
                "path": e.get("path", ""),
            }
    return out


###################################################
# LLM — Ollama 호출 (HTTP + CLI 폴백)
###################################################

def generate_with_ollama(model: str, prompt: str) -> str:
    """
    1) HTTP API 먼저 시도
    2) 실패 시 CLI 폴백
    """
    # --- 1) HTTP ---
    try:
        import requests
        url = "http://localhost:11434/api/generate"
        payload = {"model": model, "prompt": prompt, "stream": False}
        r = requests.post(url, json=payload, timeout=600)
        if r.status_code == 200:
            return r.json().get("response", "").strip()
        else:
            print(f"[정보] Ollama HTTP 응답: {r.status_code}, CLI로 폴백")
    except Exception as e:
        print(f"[정보] Ollama HTTP 실패: {e}")

    # --- 2) CLI 폴백 ---
    try:
        result = subprocess.run(
            ["ollama", "run", model],
            input=prompt,
            capture_output=True,
            text=True,
            check=True
        )
        return result.stdout.strip()
    except Exception as e:
        print(f"[오류] Ollama CLI 실패: {e}")
        return ""


###################################################
# Indexing 단계
###################################################

def run_indexing(corpus_dir: Path, bm25_build_dir: Path):
    print("--- 1. 인덱싱 시작 ---")
    corpus_dir.mkdir(exist_ok=True)
    bm25_build_dir.mkdir(parents=True, exist_ok=True)

    doc_paths = [
        p for p in corpus_dir.iterdir()
        if p.suffix.lower() in [".pdf", ".docx", ".hwpx"]
    ]

    client = QdrantClient(path=QDRANT_PATH)

    client.recreate_collection(
        DENSE_COLLECTION,
        vectors_config=models.VectorParams(size=VECTOR_DIM, distance=models.Distance.COSINE),
    )

    all_lines = []
    doc_meta = []

    for fp in doc_paths:
        doc_id = fp.stem
        text = parse_document(fp)
        if not text:
            continue
        chunks = chunk_text(text)

        print(f"[Index-Dense] {doc_id} — {len(chunks)} chunks")
        try:
            vecs = embedding.embed_texts(chunks)
            temp = bm25_build_dir / f"{doc_id}_tmp.npy"
            np.save(temp, vecs)
            qdrant_dense.append_npy_only(client, str(temp), file_id=doc_id)
            os.remove(temp)
        except Exception as e:
            print(f"[오류] Dense 인덱싱 실패: {e}")

        all_lines.append(doc_id)
        all_lines.extend(chunks)
        all_lines.append(BM25_PREPARE_DELIM)

        doc_meta.append({"doc_id": doc_id, "filename": fp.name, "path": str(fp.resolve())})

    json.dump(doc_meta, open(DOC_META_JSON, "w", encoding="utf-8"), ensure_ascii=False, indent=2)

    # BM25 prepare & build
    all_lines_json = bm25_build_dir / "all_lines.json"
    json.dump(all_lines, open(all_lines_json, "w", encoding="utf-8"), ensure_ascii=False)

    subprocess.run([sys.executable, "bm25.py", "prepare-delimited",
                    "--json", str(all_lines_json),
                    "--out_dir", str(bm25_build_dir),
                    "--delim", BM25_PREPARE_DELIM,
                    "--use_title_as_id"], check=True)

    subprocess.run([sys.executable, "bm25.py", "build-json",
                    "--json", str(bm25_build_dir/"corpus_chunks.json"),
                    "--save_dir", str(bm25_build_dir),
                    "--tokenizer", BM25_TOKENIZER], check=True)

    subprocess.run([sys.executable, "qdrant_sparse.py",
                    "--json", str(bm25_build_dir/"bm25_docs.json"),
                    "--bm25", str(bm25_build_dir/"bm25_matrix.npz"),
                    "--chunk2doc", str(bm25_build_dir/"bm25_chunk2doc.json"),
                    "--qdrant-path", QDRANT_PATH,
                    "--collection", SPARSE_COLLECTION,
                    "--sparse-name", SPARSE_NAME,
                    "--recreate"], check=True)

    print("--- 인덱싱 완료 ---")
###################################################
# Query 단계 (하이브리드 스코어 + LLM 리포트)
###################################################

def run_query(
    query_file: Path,
    bm25_build_dir: Path,
) -> Optional[Dict]:
    """외부 문서 1개로 전체 코퍼스를 쿼리하고, 하이브리드 점수를 계산한다."""
    print(f"--- 2. 쿼리 시작: {query_file.name} ---")

    if not query_file.exists():
        print(f"[오류] 쿼리 파일 없음: {query_file}")
        return None

    bm25_vocab_path = bm25_build_dir / "bm25_vocab.json"
    bm25_chunk2doc_path = bm25_build_dir / "bm25_chunk2doc.json"
    if not (bm25_vocab_path.exists() and bm25_chunk2doc_path.exists()):
        print("[오류] BM25 인덱스 파일이 없습니다. 먼저 index를 실행하세요.")
        return None

    # BM25 문서별 최대값 테이블
    bm25_max_table = load_bm25_max_table()

    # 1) 외부 문서 파싱 & 청킹
    query_text = parse_document(query_file)
    if not query_text:
        print("[오류] 쿼리 문서 텍스트 없음")
        return None

    chunks = chunk_text(query_text)
    if not chunks:
        print("[오류] 쿼리 문서에서 청크 추출 실패")
        return None

    print(f"[Query] {len(chunks)}개 청크 생성.")

    # 2) Dense 검색 (코사인)
    print("[Query-Dense] 코사인 유사도 검색 중...")
    dense_detailed_results: Dict[str, Dict] = {}
    try:
        query_vectors = embedding.embed_texts(chunks)
        _, all_dense_ranks = cosine.cosine_rank_with_qdrant(query_vectors)
        dense_detailed_results = {r["file_id"]: r for r in all_dense_ranks}
        print(f"[Query-Dense] {len(dense_detailed_results)}개 문서 결과.")
    except Exception as e:
        print(f"[오류] Dense 검색 실패: {e}")

    # 3) Sparse 검색 (BM25)
    print("[Query-Sparse] BM25 유사도 검색 중...")
    sparse_ranks_list: List[Tuple[str, float]] = []
    try:
        sparse_ranks_list = word.bm25_search_qdrant(
            chunks=chunks,
            qdrant_path=QDRANT_PATH,
            collection=SPARSE_COLLECTION,
            sparse_name=SPARSE_NAME,
            vocab_path=bm25_vocab_path,
            chunk2doc_path=bm25_chunk2doc_path,
            tokenizer=BM25_TOKENIZER,
            topk=300,  # 너무 크지 않게
        )
        print(f"[Query-Sparse] {len(sparse_ranks_list)}개 문서 결과.")
    except Exception as e:
        print(f"[오류] Sparse 검색 실패: {e}")

    # 4) 하이브리드 스코어 계산
    bm25_scores = {doc_id: score for doc_id, score in sparse_ranks_list}
    cosine_scores = {doc_id: r.get("DocSim", 0.0) for doc_id, r in dense_detailed_results.items()}

    all_doc_ids = set(bm25_scores.keys()) | set(cosine_scores.keys())
    if not all_doc_ids:
        print("[오류] 어떤 문서도 검색 결과에 포함되지 않았습니다.")
        return None

    results = []
    for doc_id in sorted(all_doc_ids):
        bm25_raw = bm25_scores.get(doc_id, 0.0)
        cos_raw = cosine_scores.get(doc_id, 0.0)

        # BM25 정규화 (문서별 최대값으로 나누기)
        bm25_doc_max = bm25_max_table.get(doc_id, 0.0)
        if bm25_doc_max > 0.0:
            bm25_rel = bm25_raw / bm25_doc_max
        else:
            bm25_rel = 0.0
        bm25_pct = bm25_rel * 100.0

        # 코사인 0~1 → 0~100
        if cos_raw < 0:
            cos_raw_clamped = 0.0
        elif cos_raw > 1:
            cos_raw_clamped = 1.0
        else:
            cos_raw_clamped = cos_raw
        cos_pct = cos_raw_clamped * 100.0

        hybrid_pct = BM25_WEIGHT * bm25_pct + COSINE_WEIGHT * cos_pct
        hybrid_score = hybrid_pct / 100.0

        results.append(
            {
                "doc_id": doc_id,
                "bm25_raw": bm25_raw,
                "bm25_max": bm25_doc_max,
                "bm25_rel": bm25_rel,
                "bm25_pct": bm25_pct,
                "cosine_DocSim": cos_raw,
                "cosine_pct": cos_pct,
                "hybrid_score": hybrid_score,
                "hybrid_pct": hybrid_pct,
                "cosine_metrics": dense_detailed_results.get(doc_id),
            }
        )

    results.sort(key=lambda r: r["hybrid_score"], reverse=True)

    out = {
        "query_document": query_file.stem,
        "query_text": query_text,
        "bm25_weight": BM25_WEIGHT,
        "cosine_weight": COSINE_WEIGHT,
        "results": results,
    }

    print("--- ✅ 쿼리 및 하이브리드 스코어 계산 완료. ---")
    return out


###################################################
# 메인 (CLI)
###################################################

def main():
    parser = argparse.ArgumentParser(description="RAG Pipeline v13 (Hybrid + LLM 리포트)")
    subparsers = parser.add_subparsers(dest="command", required=True)

    # index
    p_index = subparsers.add_parser("index", help="코퍼스 인덱싱")
    p_index.add_argument(
        "--corpus_dir",
        type=str,
        default="./corpus_docs",
        help="인덱싱할 문서 디렉터리",
    )

    # query
    p_query = subparsers.add_parser("query", help="외부 문서로 코퍼스 쿼리")
    p_query.add_argument(
        "--query_file",
        type=str,
        required=True,
        help="외부 문서 파일 경로",
    )
    p_query.add_argument(
        "--corpus_dir",
        type=str,
        default="./corpus_docs",
        help="내부 문서(코퍼스) 디렉터리",
    )
    p_query.add_argument(
        "--out_json",
        type=str,
        default="",
        help="전체 유사도 결과 JSON 저장 경로",
    )
    p_query.add_argument(
        "--out_summary",
        type=str,
        default="",
        help="LLM 요약 마크다운 저장 경로",
    )
    p_query.add_argument(
        "--ollama_model",
        type=str,
        default="llama3.1",
        help="Ollama 모델 이름",
    )

    args = parser.parse_args()

    if args.command == "index":
        run_indexing(Path(args.corpus_dir), BM25_BUILD_DIR)
        return

    if args.command == "query":
        query_path = Path(args.query_file)
        result = run_query(query_path, BM25_BUILD_DIR)
        if not result:
            return

        results = result["results"]
        query_doc_id = result["query_document"]
        query_text = result["query_text"]

        # 상위 3개만 사용
        top3 = results[:3]
        print("\n--- 하이브리드 상위 3개 문서 ---")
        for i, r in enumerate(top3, 1):
            print(f"{i}. {r['doc_id']} (Hybrid={r['hybrid_pct']:.2f}%)")

        # JSON 저장 (LLM 이전 버전)
        if args.out_json:
            try:
                with open(args.out_json, "w", encoding="utf-8") as f:
                    json.dump(result, f, ensure_ascii=False, indent=2)
                print(f"\n[저장] 유사도 전체 결과 JSON: {args.out_json}")
            except Exception as e:
                print(f"[경고] JSON 저장 실패: {e}")

        # LLM 리포트 생성
        if not top3:
            print("[경고] 상위 문서가 없어 LLM 리포트를 생성하지 않습니다.")
            return

        # A/B/C 라벨 매핑
        labels = ["A", "B", "C"]
        labeled_docs = []
        for idx, r in enumerate(top3):
            label = labels[idx]
            labeled_docs.append((label, r))

        # 내부 문서 원문 로딩
        doc_meta = load_doc_meta(BM25_BUILD_DIR)
        corpus_dir = Path(args.corpus_dir)
        internal_text_map: Dict[str, str] = {}

        for label, rec in labeled_docs:
            doc_id = rec["doc_id"]
            meta = doc_meta.get(doc_id)
            candidate_path: Optional[Path] = None

            if meta:
                p = Path(meta.get("path", ""))
                if p.exists():
                    candidate_path = p
                else:
                    fn = meta.get("filename", "")
                    if fn:
                        p2 = corpus_dir / fn
                        if p2.exists():
                            candidate_path = p2
            else:
                # 메타가 없으면 이름으로 추정
                for ext in [".pdf", ".docx", ".hwpx"]:
                    p3 = corpus_dir / f"{doc_id}{ext}"
                    if p3.exists():
                        candidate_path = p3
                        break

            if not candidate_path:
                print(f"[경고] 내부 문서 원문을 찾을 수 없음: {doc_id}")
                continue

            text = parse_document(candidate_path)
            if text:
                internal_text_map[doc_id] = text

        if not internal_text_map:
            print("[경고] 내부 문서 텍스트가 없어 LLM 리포트를 건너뜁니다.")
            return

        # LLM 프롬프트 구성
        # 외부/내부 텍스트는 길이 제한 후 삽입
        ext_text_trunc = truncate_text(query_text)

        prompt_parts = []

        # 설명 + 출력 포맷 강제
        prompt_header_template = """
당신은 "외부 문서와 내부 문서 TOP3 간 유사도 및 기밀 유출 위험"을 분석하는 한국어 보안 분석가입니다.

아래 규칙을 반드시 지키세요:
- 비교는 외부 문서 vs 내부 TOP3 문서만 수행하세요.
- 내부 문서끼리 비교하거나 유사점/차이점을 말하는 것은 절대 금지입니다.
- 요약은 반드시 본문 내용 기반 3~4줄의 자연스러운 서술문이어야 하며, 목차 제목만 나열하는 행위는 금지입니다.
- 유사 문장 예시는 반드시 제공된 텍스트에서 실제 발췌한 연속된 문장/구절만 사용하세요.
- 텍스트에 없는 정보(숫자, 리스크, 사업 구조 등)를 절대 추측하거나 만들어내지 마세요.

### 출력 형식 (마크다운)

# 외부 문서 요약
- (3~5줄로, 이 문서가 어떤 회사/사업/투자 내용을 다루는지, 목적은 무엇인지, 본문에서 드러난 핵심 논지를 요약 작성)

# 내부 문서 요약
## {top1_id}
- (3~4줄로, 해당 문서의 내용·회사·사업·핵심 포인트·리스크를 실제 본문 내용에 기반하여 요약)

## {top2_id}
- (동일하게 3~4줄로 본문 내용 기반 요약)

## {top3_id}
- (동일하게 3~4줄로 본문 내용 기반 요약)

# 유사도 분석

## 외부 문서 vs {top1_id}
### 유사한 내용 (2~4줄)
- 외부 문서와 {top1_id}가 어떤 측면에서 비슷한지 설명

### 유사한 문장 예시 (실제 텍스트 기반, 최대 5쌍)
- 형식:
  - 외부: "..."
    내부({top1_id}): "..."

### 기밀 유출 위험 평가 (2~4줄)
- 위 유사 문장 쌍을 근거로 위험도를 "낮음/중간/높음"으로 평가하고 이유를 설명

---

## 외부 문서 vs {top2_id}
### 유사한 내용 (2~4줄)
### 유사한 문장 예시 (최대 5쌍)
### 기밀 유출 위험 평가 (2~4줄)

---

## 외부 문서 vs {top3_id}
### 유사한 내용 (2~4줄)
### 유사한 문장 예시 (최대 5쌍)
### 기밀 유출 위험 평가 (2~4줄)

주의:
- 내부 문서끼리 비교 금지.
- 회사 이름·수치 등 텍스트에 없는 내용은 언급 금지.
- 모든 예시는 반드시 실제 텍스트 기반으로 작성할 것.
"""



        prompt_parts.append(prompt_header_template)

        # 외부 문서 정보
        prompt_parts.append(
            f"""
[외부 문서 ID]
{query_doc_id}

[외부 문서 내용 일부]
{ext_text_trunc}
"""
        )

        # 내부 문서 정보 (A/B/C)
        for label, rec in labeled_docs:
            doc_id = rec["doc_id"]
            if doc_id not in internal_text_map:
                continue
            bm25_pct = rec["bm25_pct"]
            cos_pct = rec["cosine_pct"]
            hybrid_pct = rec["hybrid_pct"]
            int_text_trunc = truncate_text(internal_text_map[doc_id])

            part = f"""
[내부 문서 {label}]
ID: {doc_id}
점수: BM25%={bm25_pct:.1f}, Cosine%={cos_pct:.1f}, Hybrid%={hybrid_pct:.1f}

[내부 문서 {label} 내용 일부]
{int_text_trunc}
"""
            prompt_parts.append(part)

        full_prompt = "\n\n".join(prompt_parts)

        print("\n[LLM] Ollama를 이용해 유사도/기밀 유출 분석 리포트를 생성합니다...")
        llm_output = generate_with_ollama(args.ollama_model, full_prompt)

        if not llm_output:
            print("[경고] LLM에서 결과를 받지 못했습니다.")
            return

        # 결과 딕셔너리에 LLM 마크다운 포함
        result["llm_summary_markdown"] = llm_output

        # JSON 재저장 (LLM 포함)
        if args.out_json:
            try:
                with open(args.out_json, "w", encoding="utf-8") as f:
                    json.dump(result, f, ensure_ascii=False, indent=2)
                print(f"[저장] LLM 요약 포함 전체 결과 JSON 업데이트: {args.out_json}")
            except Exception as e:
                print(f"[경고] JSON 재저장 실패: {e}")

        # 요약 마크다운 파일 저장
        if args.out_summary:
            try:
                with open(args.out_summary, "w", encoding="utf-8") as f:
                    f.write(llm_output)
                print(f"[저장] LLM 마크다운 요약 파일: {args.out_summary}")
            except Exception as e:
                print(f"[경고] 요약 파일 저장 실패: {e}")

        # 콘솔에 최종 리포트 찍기
        print("\n======================= LLM 유사도/기밀 분석 리포트 =======================\n")
        print(llm_output)
        print("\n===========================================================================")


if __name__ == "__main__":
    main()
