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
            topk=100,
        )
        print(f"[Query-Sparse] {len(sparse_ranks_list)}개 문서 결과.")
    except Exception as e:
        print(f"오류: Sparse 검색 실패: {e}")

    # 4) LLM에 줄 데이터로 정리
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

    query_doc_id = query_file.stem

    judge_data = {
        "query_document": query_doc_id,
        "top_3_cosine_similarity_metrics": cosine_top_3,
        "top_3_bm25_score": [
            {"doc_id": doc_id, "bm25_score": score}
            for doc_id, score in sparse_top_3
        ],
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

        # 2. Ollama용 프롬프트 구성 (LLM이 실제로 판결하도록)
        model_name = (
            args.ollama_model
            or os.environ.get("OLLAMA_MODEL")
            or "qwen2.5:7b"
        )

        print(f"\n--- Ollama 모델: {model_name} ---")

        prompt = f"""
당신은 투자 보고서·기밀 문서 간 유사도를 판결하는 전문 판사입니다.
아래에 주어진 분석 지표를 사용해, 쿼리 문서와 가장 유사한 **단 하나의 문서(doc_id)**를 선택해야 합니다.

[쿼리 문서]
- query_document: "{judge_data['query_document']}"

[분석 데이터]
아래 JSON에는 코사인 유사도 기반 지표와 BM25 기반 키워드 유사도 상위 3개 문서가 정리되어 있습니다.

---[분석 데이터 시작]---
{json_output}
---[분석 데이터 끝]---

[지표 설명]
- top_3_cosine_similarity_metrics:
  - DocSim: 전체적인 문맥/의미 유사도 (1에 가까울수록 유사).
  - PctMean_Q_C / PctMean_C_Q: 상위 퍼센타일 구간의 최대 유사도 평균 (쿼리→코퍼스 / 코퍼스→쿼리).
  - TrimMean_Q_C / TrimMean_C_Q: 상하위 일부를 제거한 트림 평균 유사도.
  - Cov_85_Q_C_pct / Cov_90_Q_C_pct: 쿼리 청크 중 0.85 / 0.90 이상으로 강하게 매칭되는 비율(%).
- top_3_bm25_score:
  - bm25_score: 키워드 기반 BM25 점수 (높을수록 텍스트 키워드가 잘 맞음).

[지표 사용 우선순위]
1. **가장 중요한 기준**은 PctMean_Q_C와 TrimMean_Q_C입니다. 이 두 값이 높은 문서를 우선적으로 고려하십시오.
2. DocSim은 보조적인 문맥 유사도 지표로 사용합니다. PctMean/TrimMean과 함께 해석하십시오.
3. BM25 점수는 키워드 관점에서 유사한지 확인하는 보조 지표입니다. BM25 1위/2위/3위의 상대적 차이를 정확하게 반영해야 합니다.
4. Cov_85_Q_C_pct / Cov_90_Q_C_pct는 "카피 수준의 강한 문장 재사용"이 있는지 확인하는 **참고 지표**일 뿐, PctMean/TrimMean보다 우선하지 않습니다.

[일관성 규칙 (반드시 지켜야 함)]
1. JSON 안의 수치를 그대로 사용하고, **존재하지 않는 값이나 순위를 만들어 내지 마십시오.**
2. BM25에 대해 설명할 때는, 반드시 **bm25_score가 가장 높은 문서를 1위로 언급**해야 합니다.
   - 예: bm25_score가 779, 621, 572이면, 779를 가진 문서가 1위임을 분명히 말해야 합니다.
3. 어떤 문서에 대해 "BM25 점수가 가장 높다"고 말할 때는, 실제로 그 문서의 bm25_score가 다른 후보보다 크다는 사실이 JSON에 의해 확인되어야 합니다.
4. 상반되는 서술(예: "DocSim은 낮다"고 하면서 0.45 이상인 경우)을 하지 마십시오.

[유출 위험도 판단 규칙]
아래 규칙을 사용하여 기밀 문서 유출 위험도를 '낮음', '중간', '높음' 중 하나로 정하십시오.  
(이 규칙을 reason 안에서 그대로 설명할 필요는 없지만, 결과는 반드시 이 규칙을 만족해야 합니다.)

- 우선, 선택된 문서의 주요 값:
  - D = DocSim
  - P = PctMean_Q_C
  - T = TrimMean_Q_C
- 규칙:
  1) 만약 D ≥ 0.55 또는 P ≥ 0.60 이거나, Cov_85_Q_C_pct ≥ 10%인 경우  
     → 유출 위험 = "높음"
  2) 그렇지 않고, D ≥ 0.40 또는 P ≥ 0.45 인 경우  
     → 유출 위험 = "중간"
  3) 위 두 조건에 모두 해당하지 않으면  
     → 유출 위험 = "낮음"

[reason에 반드시 포함해야 할 내용]
- 선택한 문서와 나머지 후보 문서들의 PctMean_Q_C, TrimMean_Q_C, DocSim 값을 **소수점 셋째 자리까지** 직접 비교하십시오.
  - 예: "라움토목  PctMean 0.485 / Trim 0.439 / DocSim 0.417 vs 누누팩토리 0.445 / 0.400 / 0.386 …"
- BM25에 대해서는, **실제 값과 순위에 맞게** 비교하십시오.
  - 예: "BM25 기준 1위는 레오나건설(779), 2위는 라움토목(621) …"
- Cov_85_Q_C_pct / Cov_90_Q_C_pct가 0%에 가깝다면 
  "강한 문장 단위 재사용은 거의 없고, 섹터/서술 수준의 유사도"라는 식으로 해석하십시오.
- 마지막 문장에서는 **유출 위험을 한 단어로 요약하여** 다음 중 하나를 그대로 사용하십시오:
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



        # 3. Ollama 실행
        print("\n--- Ollama 호출 ---")
        try:
            result = subprocess.run(
                ["ollama", "run", model_name],
                input=prompt,
                text=True,
                capture_output=True,
                check=False,
            )
        except FileNotFoundError:
            print("오류: 'ollama' 명령을 찾을 수 없습니다.")
            print("→ Ollama가 설치되어 있고 PATH에 잡혀 있는지 확인하세요.")
            print("→ 또는 PowerShell에서 직접 다음과 같이 실행할 수 있습니다:")
            print("   ollama run qwen2.5:7b")
            return

        stdout = (result.stdout or "").strip()
        stderr = (result.stderr or "").strip()

        if stderr:
            print("\n[Ollama stderr]")
            print(stderr)

        print("\n[Ollama raw output]")
        print(stdout)

        # 4. LLM 출력에서 JSON 파싱
        final_doc_id = None
        reason = None

        try:
            first = stdout.find("{")
            last = stdout.rfind("}")
            if first != -1 and last != -1 and last > first:
                json_str = stdout[first : last + 1]
            else:
                json_str = stdout

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

if __name__ == "__main__":
    main()
