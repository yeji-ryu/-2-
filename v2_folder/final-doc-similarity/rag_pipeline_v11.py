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

# 하이브리드 가중치
BM25_WEIGHT = 0.6
COSINE_WEIGHT = 0.4

# BM25 최대값 JSON (파일별 best_score 사용)
BM25_MAX_JSON = r"C:\Users\LG\HS\bm25_file_scores.json"

# --- Helper Functions ---

def load_bm25_max_table() -> Dict[str, float]:
    """
    bm25_file_scores.json 형식:
    [
      {
        "file": "그린퓨처_투자보고서.docx",
        "best_score": 4692.6849,
        "best_doc_id": "그린퓨처_투자보고서",
        "top_matches": [...]
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

# --- Phase 2: Querying + Hybrid Score ---

def run_query(
    query_file: Path,
    bm25_build_dir: Path,
) -> Optional[Dict]:
    """외부 문서 1개로 전체 코퍼스를 쿼리하고, 하이브리드 점수를 계산한다."""
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

    # 3) Sparse 검색 (BM25) - topk 제한 없이 충분히 크게
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
            bm25_rel = bm25_raw / bm25_doc_max   # 0~1, 동일 문서면 1에 가까워짐
        else:
            bm25_rel = 0.0

        bm25_pct = bm25_rel * 100.0

        # --- Cosine: 0~1 값을 0~100 스케일로 ---
        # DocSim 자체가 0~1 범위라 가정, 혹시 벗어나면 클램핑
        if cos_raw < 0.0:
            cos_raw_clamped = 0.0
        elif cos_raw > 1.0:
            cos_raw_clamped = 1.0
        else:
            cos_raw_clamped = cos_raw

        cos_pct = cos_raw_clamped * 100.0

        # --- 하이브리드: (BM25%, Cosine%)를 6:4로 가중합 ---
        # 예: BM25%=89, Cos%=26 → Hybrid%=0.6*89 + 0.4*26
        hybrid_pct = BM25_WEIGHT * bm25_pct + COSINE_WEIGHT * cos_pct
        hybrid_score = hybrid_pct / 100.0  # 0~1

        results.append(
            {
                "doc_id": doc_id,
                "bm25_raw": bm25_raw,
                "bm25_max": bm25_doc_max,
                "bm25_rel": bm25_rel,          # 0~1
                "bm25_pct": bm25_pct,          # 0~100
                "cosine_DocSim": cos_raw,
                "cosine_pct": cos_pct,         # 0~100
                "hybrid_score": hybrid_score,  # 0~1
                "hybrid_pct": hybrid_pct,      # 0~100
                "cosine_metrics": dense_detailed_results.get(doc_id),
            }
        )

    # 하이브리드 기준 정렬
    results.sort(key=lambda r: r["hybrid_score"], reverse=True)

    query_doc_id = query_file.stem

    out = {
        "query_document": query_doc_id,
        "bm25_weight": BM25_WEIGHT,
        "cosine_weight": COSINE_WEIGHT,
        "results": results,
    }

    print("--- ✅ 쿼리 및 하이브리드 스코어 계산 완료. ---")
    return out

# --- Main CLI ---

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
        default=10,
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

        # 2. 사람 눈용 요약 (하이브리드 상위 N개)
        topk = max(1, args.topk)
        print(f"\n--- 요약: 하이브리드 상위 {topk}개 문서 ---")
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

if __name__ == "__main__":
    main()
