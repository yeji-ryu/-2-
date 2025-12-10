# -*- coding: utf-8 -*-
import argparse
import json
import re
import subprocess
import sys
import os
import time
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
    import cosine  # <--- cosine.py (필수)
    import word    # <--- word.py (필수)
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

# --- Helper Functions ---

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

# --- Phase 1: Indexing ---

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

# --- Phase 2: Querying + LLM 판결 ---

def run_query(query_file: Path, bm25_build_dir: Path) -> Optional[Dict]:
    """외부 문서 1개로 전체 코퍼스를 쿼리하고, LLM 판결에 쓸 데이터(judge_data)를 만든다."""
    print(f"--- 2. 쿼리 시작: {query_file.name} ---")

    if not query_file.exists():
        print(f"오류: 쿼리 파일을 찾을 수 없습니다: {query_file}")
        return None

    bm25_vocab_path = bm25_build_dir / "bm25_vocab.json"
    bm25_chunk2doc_path = bm25_build_dir / "bm25_chunk2doc.json"

    if not (bm25_vocab_path.exists() and bm25_chunk2doc_path.exists()):
        print(f"오류: BM25 인덱스 파일이 없습니다 ('{bm25_build_dir}' 경로).")
        print("먼저 'index' 명령을 실행하세요.")
        return None

    # 1) 쿼리 문서 파싱 및 청킹
    text = parse_document(query_file)
    if not text:
        return None

    chunks = chunk_text(text)
    if not chunks:
        print("오류: 쿼리 문서에서 청크를 추출할 수 없습니다.")
        return None

    print(f"[Query] {len(chunks)}개 청크 생성.")

    # 2) Dense 검색 (cosine.py에서 모든 상세 지표 받기)
    print("[Query-Dense] 코사인 유사도 검색 중...")
    dense_detailed_results: Dict[str, Dict] = {}

    try:
        query_vectors = embedding.embed_texts(chunks)
        _, all_dense_ranks = cosine.cosine_rank_with_qdrant(query_vectors)
        dense_detailed_results = {r["file_id"]: r for r in all_dense_ranks}
        print(f"[Query-Dense] {len(dense_detailed_results)}개 문서 결과.")
    except Exception as e:
        print(f"오류: Dense 검색 실패: {e}")

    # 3) Sparse 검색 (BM25) — topk 줄여서 속도 절약
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
            topk=30,  # 100 -> 30으로 줄여서 BM25 쿼리량 축소
        )
        print(f"[Query-Sparse] {len(sparse_ranks_list)}개 문서 결과.")
    except Exception as e:
        print(f"오류: Sparse 검색 실패: {e}")

    # 4) Dense/BM25 결과를 LLM용 judge_data로 정리

    # 4a. BM25 상위 3개 (doc_id, score)
    sparse_top_3 = sparse_ranks_list[:3]

    # 4b. 코사인 상위 3개 (상세 지표)
    sorted_dense_results = sorted(
        dense_detailed_results.values(),
        key=lambda x: x.get("DocSim", 0.0),
        reverse=True,
    )

    cosine_top_3 = []
    for r in sorted_dense_results[:3]:
        metrics = {
            "doc_id": r.get("file_id"),
            "DocSim": r.get("DocSim"),
            "PctMean_Q_C": r.get("PctMean_Q2C"),
            "PctMean_C_Q": r.get("PctMean_C2Q"),
            "TrimMean_Q_C": r.get("TrimMean_Q2C"),
            "TrimMean_C_Q": r.get("TrimMean_C2Q"),
            "Cov_85_Q_C_pct": r.get("Cov85_Q2C_pct"),
            "Cov_85_C_Q_pct": r.get("Cov85_C2Q_pct"),
            "Cov_90_Q_C_pct": r.get("Cov90_Q2C_pct"),
            "Cov_90_C_Q_pct": r.get("Cov90_C2Q_pct"),
            "Q_chunks": r.get("Q_chunks"),
            "C_chunks": r.get("C_chunks"),
        }
        cosine_top_3.append(metrics)

    # 4c. BM25 전체 map (doc_id -> score)
    bm25_map: Dict[str, float] = {doc_id: score for doc_id, score in sparse_ranks_list}

    # 4d. 후보 doc_id 집합 = 코사인 top3 + BM25 top3 (중복 제거)
    cosine_ids = [m["doc_id"] for m in cosine_top_3 if m.get("doc_id")]
    bm25_ids = [doc_id for doc_id, _ in sparse_top_3]

    candidate_ids: List[str] = []
    for did in cosine_ids + bm25_ids:
        if did not in candidate_ids:
            candidate_ids.append(did)

    # 4e. candidates 리스트 구성 (LLM에 넘길 최소 정보)
    candidates = []
    for did in candidate_ids:
        dense = dense_detailed_results.get(did, {})
        cand = {
            "doc_id": did,
            "DocSim": dense.get("DocSim"),
            "PctMean_Q_C": dense.get("PctMean_Q2C"),
            "PctMean_C_Q": dense.get("PctMean_C2Q"),
            "TrimMean_Q_C": dense.get("TrimMean_Q2C"),
            "TrimMean_C_Q": dense.get("TrimMean_C2Q"),
            "Cov_85_Q_C_pct": dense.get("Cov85_Q2C_pct"),
            "Cov_85_C_Q_pct": dense.get("Cov85_C2Q_pct"),
            "Cov_90_Q_C_pct": dense.get("Cov90_Q2C_pct"),
            "Cov_90_C_Q_pct": dense.get("Cov90_C2Q_pct"),
            "Q_chunks": dense.get("Q_chunks"),
            "C_chunks": dense.get("C_chunks"),
            "bm25_score": bm25_map.get(did),
        }
        candidates.append(cand)

    query_doc_id = query_file.stem

    judge_data = {
        "query_document": query_doc_id,
        "top_3_cosine_similarity_metrics": cosine_top_3,
        "top_3_bm25_score": [
            {"doc_id": doc_id, "bm25_score": score}
            for doc_id, score in sparse_top_3
        ],
        "candidates": candidates,
    }

    print("--- ✅ 쿼리 완료. ---")
    return judge_data

# --- Main CLI ---

def main():
    parser = argparse.ArgumentParser(description="RAG Pipeline (Indexing and Querying)")
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
        help="새 문서로 코퍼스를 쿼리하고, LLM이 최종 유사 문서를 판결합니다.",
    )
    parser_query.add_argument(
        "--query_file",
        type=str,
        required=True,
        help="유사도를 비교할 새 문서 파일",
    )
    parser_query.add_argument(
        "--ollama_model",
        type=str,
        default=None,
        help="Ollama에서 사용할 모델 이름 (예: 'qwen2.5:7b'). 미지정 시 OLLAMA_MODEL env 또는 기본값 사용.",
    )

    args = parser.parse_args()

    if args.command == "index":
        run_indexing(Path(args.corpus_dir), BM25_BUILD_DIR)

    elif args.command == "query":
        judge_data = run_query(Path(args.query_file), BM25_BUILD_DIR)
        if not judge_data:
            return

        # 1. 최종 결과 (JSON) 출력
        print("--- 최종 결과 (JSON) ---")
        json_output = json.dumps(judge_data, indent=2, ensure_ascii=False)
        print(json_output)

        # 1-1. 사람 눈용 요약
        print("\n--- 요약: Top-3 코사인 유사도 ---")
        for i, r in enumerate(judge_data["top_3_cosine_similarity_metrics"], 1):
            print(
                f"{i}. {r['doc_id']}  "
                f"DocSim={r['DocSim']:.4f}  "
                f"PctMean_Q_C={r['PctMean_Q_C']:.4f}  "
                f"TrimMean_Q_C={r['TrimMean_Q_C']:.4f}  "
                f"Cov85_Q_C={r['Cov_85_Q_C_pct']:.1f}%"
            )

        print("\n--- 요약: Top-3 BM25 ---")
        for i, r in enumerate(judge_data["top_3_bm25_score"], 1):
            print(f"{i}. {r['doc_id']}  BM25={r['bm25_score']:.1f}")

        # 2. Ollama용 payload (최대한 압축: query_document + candidates만)
        judge_payload = {
            "query_document": judge_data["query_document"],
            "candidates": judge_data["candidates"],
        }
        compact_json = json.dumps(judge_payload, ensure_ascii=False)

        # 2-1. Ollama 모델 이름
        model_name = (
            args.ollama_model
            or os.environ.get("OLLAMA_MODEL")
            or "qwen2.5:7b"
        )
        print(f"\n--- Ollama 모델: {model_name} ---")

        # 2-2. 프롬프트 (짧게, 규칙만)
        prompt = f"""
당신은 투자 보고서·기밀 문서 간 유사도를 판결하는 전문 판사입니다.
아래 JSON에 주어진 후보 문서들(candidates)만 보고, 쿼리 문서와 가장 유사한 **단 하나의 문서(doc_id)**를 선택해야 합니다.

[쿼리 문서]
- query_document: "{judge_payload['query_document']}"

[분석 데이터]
아래 JSON에는 판결 대상 후보 문서들의 지표만 포함되어 있습니다.

---[분석 데이터 시작]---
{compact_json}
---[분석 데이터 끝]---

[데이터 구조]
- candidates: 판결 대상 문서 리스트
  - doc_id
  - DocSim
  - PctMean_Q_C, TrimMean_Q_C  (쿼리->문서 의미 유사도 대표값)
  - bm25_score                  (키워드 기반 BM25 점수, 없으면 null)
  - Cov_85_Q_C_pct, Cov_90_Q_C_pct (강한 매칭 비율, %)

[규칙]
1. **candidates 배열 안의 값만 사용**해야 합니다. JSON에 없는 수치나 순위를 새로 만들어 내지 마십시오.
2. PctMean_Q_C와 TrimMean_Q_C, bm25_score를 모두 중요한 지표로 보고,
   의미 유사도 50%, BM25 50% 비중으로 균형 있게 해석하십시오.
3. bm25_score가 가장 큰 문서는 "BM25 기준 1위"라고 명확히 표현해야 합니다.
4. Cov_85_Q_C_pct / Cov_90_Q_C_pct가 0%에 가까우면
   "강한 문장 재사용은 거의 없고, 섹터/서술 수준에서의 유사도" 정도로 해석하십시오.
5. candidates에 없는 doc_id는 언급하지 마십시오.

[reason 작성 가이드]
- 선택한 문서와 나머지 후보 문서들의 PctMean_Q_C, TrimMean_Q_C, DocSim, bm25_score를
  소수점 셋째 자리까지 비교하면서, 왜 그 문서를 선택했는지 설명하십시오.
- 의미 유사도(PctMean_Q_C, TrimMean_Q_C)와 BM25 점수를 각각 어떻게 평가했는지,
  두 축을 50:50 비중으로 본다는 점을 설명에 녹여 주십시오.
- 마지막 문장에서 당신의 판단 기준에 따른 유출 위험도를 다음 중 하나로 명시하십시오.
  - "결론적으로 이번 케이스의 유출 위험은 '낮음'으로 판단됩니다."
  - "결론적으로 이번 케이스의 유출 위험은 '중간'으로 판단됩니다."
  - "결론적으로 이번 케이스의 유출 위험은 '높음'으로 판단됩니다."

[출력 형식]
반드시 아래 JSON 형식 **하나만** 출력하십시오. 추가 설명, 마크다운, 코드 블록을 포함하지 마십시오.

{{
  "final_doc_id": "<가장 유사하다고 판단한 문서의 doc_id>",
  "reason": "<위 기준에 따라 수치 비교·해석·유출 위험도까지 포함한 설명을 한국어로 3~6문장으로 작성>"
}}

지금부터 위 JSON만 출력하십시오.
"""

        # 3. Ollama 실행 (스피너 + 타임아웃 10초 + num_predict 줄이기)
                # 3. Ollama 실행 (스피너 + 타임아웃 10초 + num_predict 줄이기)
        MAX_OLLAMA_SECONDS = 10  # 최대 10초까지만 기다림

        print("\n--- Ollama 호출 ---")
        start_time = time.time()
        try:
            # num_predict를 줄여서 답변 길이(=시간) 줄이기
            proc = subprocess.Popen(
                ["ollama", "run", model_name, "--num-predict", "96"],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
        except FileNotFoundError:
            print("오류: 'ollama' 명령을 찾을 수 없습니다.")
            print("→ Ollama가 설치되어 있고 PATH에 잡혀 있는지 확인하세요.")
            print("→ 또는 PowerShell에서 직접 다음과 같이 실행할 수 있습니다:")
            print("   ollama run qwen2.5:7b")
            return

        # 프롬프트 전달 후 stdin 닫기
        try:
            proc.stdin.write(prompt)
            proc.stdin.close()
        except Exception as e:
            print(f"오류: Ollama stdin 전송 중 문제 발생: {e}")

        spinner = ["|", "/", "-", "\\"]
        spin_idx = 0

        # 진행 상황 표시 (실제 퍼센트는 모름, 경과 시간 + 스피너만)
        while True:
            retcode = proc.poll()
            elapsed = time.time() - start_time
            sys.stdout.write(
                f"\r[Ollama] 실행 중... {elapsed:5.1f}s 경과 {spinner[spin_idx]}"
            )
            sys.stdout.flush()
            spin_idx = (spin_idx + 1) % len(spinner)

            # 타임아웃 초과 시 강제 종료
            if elapsed > MAX_OLLAMA_SECONDS and retcode is None:
                proc.kill()
                sys.stdout.write("\r[Ollama] 타임아웃으로 프로세스를 종료했습니다.\n")
                sys.stdout.flush()
                break

            if retcode is not None:
                break
            time.sleep(0.3)

        # 줄 정리
        sys.stdout.write("\r[Ollama] 실행 완료 또는 중단.                 \n")
        sys.stdout.flush()

        stdout, stderr = proc.communicate()
        total_elapsed = time.time() - start_time
        print(f"[Ollama] 총 소요 시간: {total_elapsed:.1f}초")

        if stderr:
            print("\n[Ollama stderr]")
            print(stderr.strip())

        print("\n[Ollama raw output]")
        print(stdout.strip())


        # 4. LLM 출력에서 JSON 파싱
        final_doc_id = None
        reason = None

        try:
            out_str = stdout.strip()
            first = out_str.find("{")
            last = out_str.rfind("}")
            if first != -1 and last != -1 and last > first:
                json_str = out_str[first : last + 1]
            else:
                json_str = out_str

            if json_str:
                data = json.loads(json_str)
                final_doc_id = data.get("final_doc_id") or data.get("doc_id")
                reason = data.get("reason") or data.get("explanation")
        except Exception as e:
            print(f"\n[경고] LLM 출력 JSON 파싱 실패: {e}")

        if final_doc_id:
            print("\n--- LLM 판결 결과 ---")
            print(f"가장 유사한 문서: {final_doc_id}")
            if reason:
                print(f"근거: {reason}")
        else:
            print("\n[경고] LLM 판결 결과를 추출하지 못했습니다. 위 Ollama 출력 내용을 직접 확인하세요.")

        # 5. 시스템 기준 유출 위험도 (코사인 50% + BM25 50%) 계산
        risk_label = None
        combined_score = None

        try:
            candidates = judge_data.get("candidates", [])
            if final_doc_id and candidates:
                target = None
                for c in candidates:
                    if c.get("doc_id") == final_doc_id:
                        target = c
                        break

                if target is not None:
                    P = float(target.get("PctMean_Q_C") or 0.0)
                    B = float(target.get("bm25_score") or 0.0)

                    max_P = max(float(c.get("PctMean_Q_C") or 0.0) for c in candidates) or 0.0
                    max_B = max(float(c.get("bm25_score") or 0.0) for c in candidates) or 0.0

                    P_norm = (P / max_P) if max_P > 0 else 0.0
                    B_norm = (B / max_B) if max_B > 0 else 0.0

                    combined_score = 0.5 * P_norm + 0.5 * B_norm

                    if combined_score >= 0.7:
                        risk_label = "높음"
                    elif combined_score >= 0.4:
                        risk_label = "중간"
                    else:
                        risk_label = "낮음"

                    print(
                        f"\n--- 시스템 계산 유출 위험도 (코사인 50% + BM25 50%) ---\n"
                        f"선택 문서: {final_doc_id}\n"
                        f"- PctMean_Q_C: {P:.4f} (max={max_P:.4f}, 정규화={P_norm:.3f})\n"
                        f"- BM25: {B:.1f} (max={max_B:.1f}, 정규화={B_norm:.3f})\n"
                        f"- 결합 점수: {combined_score:.3f}\n"
                        f"=> 시스템 판단 유출 위험도: '{risk_label}'"
                    )
                else:
                    print("\n[알림] candidates 안에서 선택된 문서를 찾을 수 없어 시스템 위험도 계산을 건너뜁니다.")
        except Exception as e:
            print(f"\n[경고] 시스템 위험도 계산 중 오류 발생: {e}")

if __name__ == "__main__":
    main()
