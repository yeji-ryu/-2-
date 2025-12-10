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

# --- Import functions ---
try:
    import pdf_parser
    import hwpx_parser
    import docx_parser
    import embedding
    import qdrant_dense
    import cosine
    import word
    import hwpx_chunking  # <--- 수정된 청킹 모듈
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

def chunk_text(text: str) -> List[str]:
    """
    수정된 hwpx_chunking을 사용하여 문장 중간의 불릿까지 감지해 청킹함.
    """
    if not text.strip():
        return []
    try:
        # 여기서 hwpx_chunking.py 의 강화된 로직 실행
        return hwpx_chunking.chunk_by_bullets(text)
    except Exception as e:
        print(f"[Warning] 청킹 중 오류 발생, 기본 분할로 대체: {e}")
        SENT_SPLIT_RE = re.compile(r"(?<=[.!?])\s+|[\r\n]+")
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
        print(f"오류: 명령어를 찾을 수 없습니다. 경로를 확인하세요.")
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

    print(f"[Index-Dense] 기존 {DENSE_COLLECTION} 컬렉션 재생성")
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
        if not text: continue

        chunks = chunk_text(text)
        if not chunks:
            print(f"경고: {file_id}에서 청크를 찾을 수 없습니다.")
            continue

        print(f"[Index-Dense] {len(chunks)}개 청크 처리 중...")

        try:
            vectors = embedding.embed_texts(chunks)
            temp_npy_path = bm25_build_dir / f"{file_id}_temp_vecs.npy"
            np.save(temp_npy_path, vectors)
            qdrant_dense.append_npy_only(client, str(temp_npy_path), file_id=file_id)
            os.remove(temp_npy_path)
        except Exception as e:
            print(f"오류: {file_id} Dense 인덱싱 실패: {e}")

        all_lines_for_bm25_prepare.append(file_id)
        all_lines_for_bm25_prepare.extend(chunks)
        all_lines_for_bm25_prepare.append(BM25_PREPARE_DELIM)

    client.close()

    if not all_lines_for_bm25_prepare:
        return

    print("\n[Index-Sparse] BM25 인덱스 생성...")
    all_lines_json = bm25_build_dir / "all_lines.json"
    chunk_json = bm25_build_dir / "corpus_chunks.json"
    chunk2doc = bm25_build_dir / "bm25_chunk2doc.json"
    vocab_json = bm25_build_dir / "bm25_vocab.json"
    matrix_npz = bm25_build_dir / "bm25_matrix.npz"
    docs_json = bm25_build_dir / "bm25_docs.json"

    with open(all_lines_json, "w", encoding="utf-8") as f:
        json.dump(all_lines_for_bm25_prepare, f, ensure_ascii=False, indent=2)

    # BM25 파이프라인 실행
    run_cli_command([sys.executable, "bm25.py", "prepare-delimited", "--json", str(all_lines_json), "--out_dir", str(bm25_build_dir), "--delim", BM25_PREPARE_DELIM, "--use_title_as_id"], "BM25 prepare 실패")
    run_cli_command([sys.executable, "bm25.py", "build-json", "--json", str(chunk_json), "--save_dir", str(bm25_build_dir), "--tokenizer", BM25_TOKENIZER], "BM25 build 실패")
    run_cli_command([sys.executable, "qdrant_sparse.py", "--json", str(docs_json), "--bm25", str(matrix_npz), "--chunk2doc", str(chunk2doc), "--qdrant-path", QDRANT_PATH, "--collection", SPARSE_COLLECTION, "--sparse-name", SPARSE_NAME, "--recreate"], "Sparse 업로드 실패")

    print("\n--- ✅ 인덱싱 완료 ---")

# --- Phase 2: Querying ---

def run_query(query_file: Path, bm25_build_dir: Path) -> Optional[Dict]:
    print(f"--- 2. 쿼리 시작: {query_file.name} ---")
    if not query_file.exists():
        print("오류: 파일 없음")
        return None

    text = parse_document(query_file)
    if not text: return None

    # 디버깅용 출력 (텍스트가 잘렸는지 확인)
    print(f"[DEBUG] 텍스트 앞부분:\n{text[:200]}...")

    chunks = chunk_text(text)
    print(f"[Query] {len(chunks)}개 청크 생성.")
    if not chunks: return None

    # Dense 검색
    print("[Query-Dense] 코사인 유사도 검색...")
    query_vectors = embedding.embed_texts(chunks)
    _, all_dense_ranks = cosine.cosine_rank_with_qdrant(query_vectors)
    dense_results = {r["file_id"]: r for r in all_dense_ranks}

    # Sparse 검색
    print("[Query-Sparse] BM25 검색...")
    sparse_ranks = word.bm25_search_qdrant(
        chunks=chunks,
        qdrant_path=QDRANT_PATH,
        collection=SPARSE_COLLECTION,
        sparse_name=SPARSE_NAME,
        vocab_path=bm25_build_dir / "bm25_vocab.json",
        chunk2doc_path=bm25_build_dir / "bm25_chunk2doc.json",
        tokenizer=BM25_TOKENIZER,
        topk=100,
    )

    # 결과 정리
    top3_sparse = sparse_ranks[:3]
    top3_dense = sorted(dense_results.values(), key=lambda x: x.get("DocSim", 0.0), reverse=True)[:3]

    cosine_top_3 = []
    for r in top3_dense:
        cosine_top_3.append({
            "doc_id": r.get("file_id"),
            "DocSim": r.get("DocSim"),
            "PctMean_Q_C": r.get("PctMean_Q2C"),
            "TrimMean_Q_C": r.get("TrimMean_Q2C"),
            "Cov_85_Q_C_pct": r.get("Cov85_Q2C_pct"),
        })

    judge_data = {
        "query_document": query_file.stem,
        "top_3_cosine_similarity_metrics": cosine_top_3,
        "top_3_bm25_score": [{"doc_id": d, "bm25_score": s} for d, s in top3_sparse],
    }
    return judge_data

def main():
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers(dest="command", required=True)

    idx_p = subparsers.add_parser("index")
    idx_p.add_argument("--corpus_dir", default="./corpus_docs")

    qry_p = subparsers.add_parser("query")
    qry_p.add_argument("--query_file", required=True)
    qry_p.add_argument("--ollama_model", default="qwen2.5:7b")

    args = parser.parse_args()

    if args.command == "index":
        run_indexing(Path(args.corpus_dir), BM25_BUILD_DIR)
    elif args.command == "query":
        res = run_query(Path(args.query_file), BM25_BUILD_DIR)
        if not res: return
        
        json_out = json.dumps(res, indent=2, ensure_ascii=False)
        print("--- 최종 분석 데이터 (JSON) ---")
        print(json_out)

        # Ollama 프롬프트
        prompt = f"""
당신은 문서 유사도 판독 전문가입니다. 아래 데이터를 분석하여 가장 유사한 문서 하나를 선정하십시오.

[쿼리 문서]: {res['query_document']}

[데이터]:
{json_out}

[판단 기준]:
1. PctMean_Q_C와 TrimMean_Q_C가 가장 높은 문서를 우선 선택합니다. (DocSim은 참고용)
2. BM25 점수는 키워드 일치 여부를 확인하는 보조 지표입니다.
3. PctMean > 0.6 이면 유출 위험 '높음', > 0.45 면 '중간', 그 외는 '낮음'입니다.

[출력 형식(JSON Only)]:
{{
  "final_doc_id": "문서명",
  "reason": "한국어로 상세 근거 작성 (수치 비교 포함)"
}}
"""
        print(f"\n--- Ollama ({args.ollama_model}) 실행 중 ---")
        subprocess.run(["ollama", "run", args.ollama_model], input=prompt, text=True)

if __name__ == "__main__":
    main()