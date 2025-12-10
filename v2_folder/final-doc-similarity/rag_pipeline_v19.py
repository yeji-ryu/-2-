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
    import cosine1  # <--- 문서 단위 코사인 (수정된 버전)
    import word    # <--- word.py (BM25)
except ImportError as e:
    print(f"--- [오류] 필수 모듈 임포트 실패 ---")
    print(f"오류 원인: {e}")
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

# BM25 최대값 JSON (파일별 best_score 사용)
BM25_MAX_JSON = r"C:\Users\LG\HS\bm25_file_scores.json"


# ============================
# 공통 유틸 (코사인, 문장 필터)
# ============================
def l2norm(X: np.ndarray) -> np.ndarray:
    return X / (np.linalg.norm(X, axis=1, keepdims=True) + 1e-12)


def cosine_matrix(A: np.ndarray, B: np.ndarray) -> np.ndarray:
    A = l2norm(A)
    B = l2norm(B)
    return A @ B.T


def is_meaningful_sentence(text: str) -> bool:
    """
    너무 짧거나, 숫자/기호만 있는 문장은 매칭에서 제외하기 위한 필터.
    예: "1.", "4." 같은 것들은 False.
    """
    s = text.strip()
    if len(s) < 2:
        return False

    # 한글/영문/숫자만 남긴 뒤 길이 체크
    meaningful = re.sub(r"[^0-9A-Za-z가-힣]", "", s)
    if len(meaningful) < 2:
        return False

    return True


# ============================
# BM25 최대값 로딩
# ============================
def load_bm25_max_table() -> Dict[str, float]:
    """
    bm25_file_scores.json 형식:
    [
      {
        "file": "그린퓨처_투자보고서.docx",
        "best_score": 4692.6849,
        "best_doc_id": "그린퓨처_투자보고서",
        "top_matches": [... ]
      },
      ...
    ]

    → { "그린퓨처_투자보고서": 4692.6849, ... } 형태로 변환
    """
    if not os.path.exists(BM25_MAX_JSON):
        print(f"[오류] BM25 최대값 JSON 파일을 찾을 수 없습니다: {BM25_MAX_JSON}")
        return {}
    try:
        with open(BM25_MAX_JSON, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        print(f"[오류] BM25 최대값 JSON 로드 실패: {e}")
        return {}

    table: Dict[str, float] = {}
    if isinstance(data, list):
        for entry in data:
            try:
                doc_id = entry.get("best_doc_id")
                best_score = float(entry.get("best_score", 0.0))
                if doc_id:
                    table[doc_id] = best_score
            except Exception:
                continue
    else:
        print("[경고] bm25_file_scores.json 형식이 리스트가 아닙니다. 빈 테이블 사용.")
    return table


# ============================
# 문서 파싱/청킹
# ============================
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
            print(f"경고: 지원하지 않는 파일 형식입니다 ({ext}). 건너뜁니다.")
            return ""
    except Exception as e:
        print(f"오류: {file_path.name} 파싱 중 오류 발생: {e}")
        return ""


SENT_SPLIT_RE = re.compile(r"(?<=[.!?])\s+|[\r\n]+")
def chunk_text(text: str) -> List[str]:
    return [s.strip() for s in SENT_SPLIT_RE.split(text) if s.strip()]


def run_cli_command(cmd_list: List[str], error_msg: str):
    print(f"\n[Run] {' '.join(cmd_list)}")
    try:
        subprocess.run(
            cmd_list,
            check=True,
            encoding="utf-8",
            stdout=sys.stdout,
            stderr=sys.stderr,
        )
    except subprocess.CalledProcessError as e:
        print(f"{error_msg}: {e}")
        sys.exit(1)
    except FileNotFoundError:
        print(f"오류: 'python' 또는 스크립트(예: bm25.py)를 찾을 수 없습니다. 경로를 확인하세요.")
        sys.exit(1)


# ============================
# all_lines.json → {doc_id: [문장들]} 로 변환
# ============================
def load_doc_chunks_table(bm25_build_dir: Path) -> Dict[str, List[str]]:
    all_lines_path = bm25_build_dir / "all_lines.json"
    if not all_lines_path.exists():
        print(f"[오류] all_lines.json 없음: {all_lines_path}")
        return {}

    try:
        with open(all_lines_path, "r", encoding="utf-8") as f:
            lines = json.load(f)  # list[str]
    except Exception as e:
        print(f"[오류] all_lines.json 로드 실패: {e}")
        return {}

    table: Dict[str, List[str]] = {}
    current_doc = None
    current_chunks: List[str] = []

    for line in lines:
        if line == BM25_PREPARE_DELIM:
            if current_doc is not None:
                table[current_doc] = current_chunks
            current_doc = None
            current_chunks = []
            continue

        if current_doc is None:
            current_doc = line
            current_chunks = []
            continue

        current_chunks.append(line)

    # 마지막 문서 처리 (마지막에 delim 없이 끝났을 경우 대비)
    if current_doc is not None and current_doc not in table:
        table[current_doc] = current_chunks

    return table


# ============================
# 상위 문서들에 대해 문장 매칭 계산
# ============================
def compute_sentence_matches_for_top_docs(
    query_chunks: List[str],
    query_vectors: np.ndarray,
    top_results: List[Dict],
    bm25_build_dir: Path,
    max_docs: int = 3,
    min_score: float = 0.4,
) -> Dict[str, List[Dict]]:
    """
    - all_lines.json 기반으로 상위 문서들의 문장 리스트를 가져온 뒤
    - 쿼리 문장 vs 해당 문서 문장들 간 코사인 유사도 계산
    - 각 문서별로 (query_sentence, matched_sentence, score) 리스트를 반환
    """
    if query_vectors is None or query_vectors.size == 0:
        return {}

    # 쿼리 문장에서 의미 없는 문장은 미리 제거
    meaningful_indices = [
        i for i, s in enumerate(query_chunks) if is_meaningful_sentence(s)
    ]
    if not meaningful_indices:
        return {}

    Q = query_vectors[meaningful_indices]
    filtered_query_sentences = [query_chunks[i] for i in meaningful_indices]

    # 문서별 문장 테이블 로드
    doc_table = load_doc_chunks_table(bm25_build_dir)
    if not doc_table:
        return {}

    result_map: Dict[str, List[Dict]] = {}

    for r in top_results[:max_docs]:
        doc_id = r["doc_id"]
        corpus_chunks = doc_table.get(doc_id, [])
        if not corpus_chunks:
            continue

        # 해당 문서 문장 임베딩
        corpus_vectors = embedding.embed_texts(corpus_chunks)
        if corpus_vectors.size == 0:
            continue

        S = cosine_matrix(Q, corpus_vectors)  # (Q_filtered x C)
        best_idx = np.argmax(S, axis=1)
        best_vals = np.max(S, axis=1)

        matches: List[Dict] = []
        for q_sent, idx_c, val in zip(filtered_query_sentences, best_idx, best_vals):
            if val < min_score:
                continue  # 너무 낮은 유사도는 버림
            matched_sent = corpus_chunks[idx_c] if idx_c < len(corpus_chunks) else ""
            matches.append(
                {
                    "query_sentence": q_sent,
                    "matched_sentence": matched_sent,
                    "score": float(val),
                }
            )

        # 점수 순으로 정렬
        matches.sort(key=lambda x: x["score"], reverse=True)
        result_map[doc_id] = matches

    return result_map


# ============================
# Phase 1: Indexing
# ============================
def run_indexing(corpus_dir: Path, bm25_build_dir: Path):
    print("--- 1. 인덱싱 시작 ---")
    corpus_dir.mkdir(exist_ok=True)
    bm25_build_dir.mkdir(parents=True, exist_ok=True)

    doc_paths = [
        p for p in corpus_dir.iterdir()
        if p.is_file() and p.suffix.lower() in [".pdf", ".docx", ".hwpx"]
    ]
    if not doc_paths:
        print(f"경고: '{corpus_dir}'에 처리할 문서가 없습니다.")
        return

    client = QdrantClient(path=QDRANT_PATH)

    print(f"[Index-Dense] 기존 {DENSE_COLLECTION} 컬렉션을 삭제하고 새로 생성합니다.")
    client.recreate_collection(
        collection_name=DENSE_COLLECTION,
        vectors_config=models.VectorParams(
            size=VECTOR_DIM,
            distance=models.Distance.COSINE,
        ),
    )

    all_lines_for_bm25_prepare = []

    for file_path in doc_paths:
        file_id = file_path.stem
        print(f"\n[Index] {file_id}{file_path.suffix}")
        text = parse_document(file_path)
        if not text:
            continue

        chunks = chunk_text(text)
        if not chunks:
            print(f"경고: {file_id}에서 청크를 찾을 수 없습니다.")
            continue

        print(f"[Index-Dense] {len(chunks)}개 청크 임베딩 및 Qdrant 업로드...")

        try:
            vectors = embedding.embed_texts(chunks)
            temp_npy_path = bm25_build_dir / f"{file_id}_temp_vecs.npy"
            np.save(temp_npy_path, vectors)
            qdrant_dense.append_npy_only(client, str(temp_npy_path), file_id=file_id)
            os.remove(temp_npy_path)
            print(f"[Index-Dense] {file_id} 완료.")
        except Exception as e:
            print(f"오류: {file_id} Dense 인덱싱 실패: {e}")

        # BM25용 라인들 준비
        all_lines_for_bm25_prepare.append(file_id)
        all_lines_for_bm25_prepare.extend(chunks)
        all_lines_for_bm25_prepare.append(BM25_PREPARE_DELIM)

    client.close()

    if not all_lines_for_bm25_prepare:
        print("오류: BM25 인덱싱을 건너뜁니다.")
        return

    print("\n[Index-Sparse] BM25 인덱스 생성 및 Qdrant 업로드...")

    all_lines_json = bm25_build_dir / "all_lines.json"
    chunk_json = bm25_build_dir / "corpus_chunks.json"
    chunk2doc = bm25_build_dir / "bm25_chunk2doc.json"
    vocab_json = bm25_build_dir / "bm25_vocab.json"
    matrix_npz = bm25_build_dir / "bm25_matrix.npz"
    docs_json = bm25_build_dir / "bm25_docs.json"

    with open(all_lines_json, "w", encoding="utf-8") as f:
        json.dump(all_lines_for_bm25_prepare, f, ensure_ascii=False, indent=2)

    cmd_prepare = [
        sys.executable, "bm25.py", "prepare-delimited",
        "--json", str(all_lines_json),
        "--out_dir", str(bm25_build_dir),
        "--delim", BM25_PREPARE_DELIM,
        "--use_title_as_id",
    ]
    run_cli_command(cmd_prepare, "BM25 prepare-delimited 실패")

    cmd_build = [
        sys.executable, "bm25.py", "build-json",
        "--json", str(chunk_json),
        "--save_dir", str(bm25_build_dir),
        "--tokenizer", BM25_TOKENIZER,
    ]
    run_cli_command(cmd_build, "BM25 build-json 실패")

    cmd_upsert = [
        sys.executable, "qdrant_sparse.py",
        "--json", str(docs_json),
        "--bm25", str(matrix_npz),
        "--chunk2doc", str(chunk2doc),
        "--qdrant-path", QDRANT_PATH,
        "--collection", SPARSE_COLLECTION,
        "--sparse-name", SPARSE_NAME,
        "--recreate",
    ]
    run_cli_command(cmd_upsert, "Qdrant Sparse 업로드 실패")

    print("\n--- ✅ 모든 인덱싱 완료. ---")


# ============================
# Phase 2: Querying + Hybrid Score
# ============================
def run_query(
    query_file: Path,
    bm25_build_dir: Path,
) -> Optional[Dict]:
    """외부 문서 1개로 전체 코퍼스를 쿼리하고, 하이브리드 점수를 계산한다."""
    print(f"--- 쿼리 시작: {query_file.name} ---")

    if not query_file.exists():
        print(f"오류: 쿼리 파일을 찾을 수 없습니다: {query_file}")
        return None

    bm25_vocab_path = bm25_build_dir / "bm25_vocab.json"
    bm25_chunk2doc_path = bm25_build_dir / "bm25_chunk2doc.json"

    if not (bm25_vocab_path.exists() and bm25_chunk2doc_path.exists()):
        print(f"오류: BM25 인덱스 파일이 없습니다 ('{bm25_build_dir}' 경로).")
        print("먼저 'index' 명령을 실행하세요.")
        return None

    # BM25 문서별 최대값 로드
    bm25_max_table = load_bm25_max_table()

    # 1) 쿼리 문서 파싱 및 청킹
    text = parse_document(query_file)
    if not text:
        return None

    chunks = chunk_text(text)
    if not chunks:
        print("오류: 쿼리 문서에서 청크를 추출할 수 없습니다.")
        return None

    print(f"[Query] {len(chunks)}개 청크 생성.")

    # 2) Dense 검색 (cosine1에서 문서 단위 DocSim만 받음)
    print("[Query-Dense] 코사인 유사도 검색 중...")
    dense_detailed_results: Dict[str, Dict] = {}
    query_vectors = None

    try:
        query_vectors = embedding.embed_texts(chunks)
        _, all_dense_ranks = cosine1.cosine_rank_with_qdrant(query_vectors, chunks)
        dense_detailed_results = {r["file_id"]: r for r in all_dense_ranks}
        print(f"[Query-Dense] {len(dense_detailed_results)}개 문서 결과.")
    except Exception as e:
        print(f"오류: Dense 검색 실패: {e}")
        query_vectors = None

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
            topk=10000,
        )
        print(f"[Query-Sparse] {len(sparse_ranks_list)}개 문서 결과.")
    except Exception as e:
        print(f"오류: Sparse 검색 실패: {e}")

    # 4) 하이브리드 스코어 계산
    bm25_scores = {doc_id: score for doc_id, score in sparse_ranks_list}
    cosine_scores = {doc_id: r.get("DocSim", 0.0) for doc_id, r in dense_detailed_results.items()}

    # 문서 id 전체 집합
    all_doc_ids = set(bm25_scores.keys()) | set(cosine_scores.keys())
    if not all_doc_ids:
        print("오류: 어떤 문서도 검색 결과에 포함되지 않았습니다.")
        return None

    results = []
    for doc_id in sorted(all_doc_ids):
        bm25_raw = bm25_scores.get(doc_id, 0.0)
        cos_raw = cosine_scores.get(doc_id, 0.0)

        # --- BM25 정규화: 문서별 최대값으로 나누기 ---
        bm25_doc_max = bm25_max_table.get(doc_id, 0.0)
        if bm25_doc_max > 0.0:
            bm25_rel = bm25_raw / bm25_doc_max   # 0~1
        else:
            bm25_rel = 0.0

        bm25_pct = bm25_rel * 100.0

        # --- Cosine: 0~1 값을 0~100 스케일로 ---
        if cos_raw < 0.0:
            cos_raw_clamped = 0.0
        elif cos_raw > 1.0:
            cos_raw_clamped = 1.0
        else:
            cos_raw_clamped = cos_raw

        cos_pct = cos_raw_clamped * 100.0

        # --- 하이브리드 ---
        hybrid_pct = BM25_WEIGHT * bm25_pct + COSINE_WEIGHT * cos_pct
        hybrid_score = hybrid_pct / 100.0  # 0~1

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
                "cosine_metrics": None,  # 나중에 sentence_matches로 채움
            }
        )

    # 하이브리드 기준 정렬
    results.sort(key=lambda r: r["hybrid_score"], reverse=True)

    # 5) 상위 문서들에 대해 문장 매칭 재계산
    sentence_matches_map = compute_sentence_matches_for_top_docs(
        query_chunks=chunks,
        query_vectors=query_vectors,
        top_results=results,
        bm25_build_dir=bm25_build_dir,
        max_docs=3,       # 상위 3개 문서
        min_score=0.5,    # 최소 코사인 점수 (조절 가능)
    )

    for r in results:
        doc_id = r["doc_id"]
        matches = sentence_matches_map.get(doc_id)
        if matches is not None:
            r["cosine_metrics"] = {"matches": matches}

    query_doc_id = query_file.stem

    out = {
        "query_document": query_doc_id,
        "query_sentences": chunks,
        "bm25_weight": BM25_WEIGHT,
        "cosine_weight": COSINE_WEIGHT,
        "results": results,
    }

    print("--- ✅ 쿼리 및 하이브리드 스코어 계산 완료. ---")
    return out


# ============================
# LLM PAYLOAD 생성 함수
# ============================
def build_llm_payload(result: Dict, topn: int = 3) -> Dict:
    query_doc_id = result.get("query_document")
    results = result.get("results", [])

    payload = {
        "query_document": query_doc_id,
        "top_matches": []
    }

    for r in results[:topn]:
        matches = (r.get("cosine_metrics") or {}).get("matches", []) or []
        # 이미 score 기준으로 정렬되어 있으므로 그대로 사용 가능
        payload["top_matches"].append({
            "doc_id": r["doc_id"],
            "hybrid_pct": r.get("hybrid_pct"),
            "bm25_pct": r.get("bm25_pct"),
            "cosine_pct": r.get("cosine_pct"),
            "sentence_matches": [
                {
                    "query_sentence": m["query_sentence"],
                    "matched_sentence": m["matched_sentence"],
                    "cosine_score": m["score"],
                }
                for m in matches[:5]  # 상위 5개만
            ]
        })

    return payload


def save_llm_payload(payload: Dict, out_dir: Path = BM25_BUILD_DIR) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"llm_input_{payload['query_document']}.json"

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    print(f"[LLM] 입력 JSON 저장 완료 → {out_path}")
    return out_path


# ============================
# 로컬 LLM(Ollama) 요약 함수
# ============================
def summarize_with_llm_local(payload: Dict) -> str:
    prompt = (
        "다음 JSON은 쿼리 문서와 상위 3개 문서의 문장 단위 매칭 결과입니다.\n"
        "이 정보를 이용해서, 아래 2단계 형식으로 한국어로만 출력하세요.\n\n"
        "1단계: 문서 요약 (짧게)\n"
        "- 먼저 쿼리 문서와 상위 3개 문서(총 4개)에 대해 각각 1~2문장으로 짧게 요약하세요.\n"
        "- 형식 예시:\n"
        "  [요약]\n"
        "  - 쿼리 문서(파일명: XXX): 이 문서는 ~~ 회사의 투자보고서이며, 투자 최종 권고사항은 reject이다.\n"
        "  - 문서1(파일명: YYY): ...\n"
        "  - 문서2(파일명: ZZZ): ...\n"
        "  - 문서3(파일명: WWW): ...\n"
        "- 각 요약에는 문서의 성격(예: 투자보고서, 계약서, 제안서 등)과 핵심 의사결정(예: 투자 여부, 승인/반려 등)이 있으면 꼭 포함하세요.\n\n"
        "2단계: 유사 문장 쌍\n"
        "- 그 다음, 각 상위 문서별로 쿼리 문장과의 유사도가 높은 문장 쌍을 점수 순서대로 최대 5개까지 보여주세요.\n"
        "- 형식 예시:\n"
        "  [유사 문장]\n"
        "  문서명: YYY\n"
        "  1) 쿼리 문장: \"...\"\n"
        "     상대 문장: \"...\"\n"
        "     유사도 점수: 0.xx\n"
        "  2) ...\n"
        "  문서명: ZZZ\n"
        "  1) ...\n\n"
        "추가 지침:\n"
        "- 유사도 점수는 cosine_score를 사용하고, 내림차순(큰 값 → 작은 값)으로 정렬한 것처럼 표현하세요.\n"
        "- 완전히 동일한 문장은 유사도 1.0로 간주해도 됩니다.\n"
        "- 의미적으로 많이 비슷해 보이는 문장(0.7 이상)은 꼭 포함하세요.\n"
        "- 전체 출력은 한국어로만 작성하세요.\n\n"
        "JSON 데이터:\n"
        f"{json.dumps(payload, ensure_ascii=False, indent=2)}"
    )

    try:
        result = subprocess.run(
            ["ollama", "run", "qwen2.5:7b"],
            input=prompt,
            text=True,
            capture_output=True,
            encoding="utf-8"
        )
        if result.returncode != 0:
            return f"[LLM 오류] 로컬 LLM 실행 실패 (코드 {result.returncode})\nSTDERR:\n{result.stderr}"
        return result.stdout.strip()
    except Exception as e:
        return f"[LLM 오류] 로컬 LLM 호출 실패: {e}"


# ============================
# main
# ============================
def main():
    parser = argparse.ArgumentParser(description="RAG Pipeline (Indexing and Hybrid Querying)")
    subparsers = parser.add_subparsers(dest="command", required=True)

    parser_index = subparsers.add_parser("index", help="문서 코퍼스를 인덱싱합니다.")
    parser_index.add_argument(
        "--corpus_dir",
        type=str,
        default="./corpus_docs",
        help="인덱싱할 문서가 포함된 디렉터리",
    )

    parser_query = subparsers.add_parser(
        "query",
        help="새 문서로 코퍼스를 쿼리하고, 하이브리드 유사도 점수를 계산합니다.",
    )
    parser_query.add_argument(
        "--query_file",
        type=str,
        required=True,
        help="유사도를 비교할 새 문서 파일",
    )
    parser_query.add_argument(
        "--topk",
        type=int,
        default=3,
        help="콘솔에 요약해서 보여줄 상위 문서 개수 (전체 결과는 JSON으로 모두 출력)",
    )

    args = parser.parse_args()

    if args.command == "index":
        run_indexing(Path(args.corpus_dir), BM25_BUILD_DIR)

    elif args.command == "query":
        result = run_query(Path(args.query_file), BM25_BUILD_DIR)
        if not result:
            return

        results = result["results"]

        # 1. 사람 눈용 요약 (하이브리드 상위 N개)
        topk = max(1, args.topk)
        print(f"\n--- 유사도 상위 {topk}개 문서 ---")
        for i, r in enumerate(results[:topk], 1):
            doc_id = r["doc_id"]
            bm25_raw = r["bm25_raw"]
            bm25_pct = r["bm25_pct"]
            cos_raw = r["cosine_DocSim"]
            cos_pct = r["cosine_pct"]
            hybrid_pct = r.get("hybrid_pct", 0.0)
            print(
                f"{i}. {doc_id}  "
                f"Hybrid={hybrid_pct:6.2f}%  "
                f"(BM25={bm25_raw:.1f}, BM25%={bm25_pct:5.1f}%, "
                f"Cosine={cos_raw:.4f} / {cos_pct:5.1f}%)"
            )

        # 2. LLM용 JSON 생성 및 저장
        payload = build_llm_payload(result, topn=3)
        payload_path = save_llm_payload(payload, BM25_BUILD_DIR)

        # 3. 로컬 LLM(qwen2.5:7b) 요약 실행
        print("\n[LLM] 상위 3개 문서 유사 구간 및 요약 생성 중...")
        summary = summarize_with_llm_local(payload)

        print("\n====== LLM 유사도 요약 ======")
        print(summary)
        print("====== LLM 유사도 요약 끝 ======")


if __name__ == "__main__":
    main()
