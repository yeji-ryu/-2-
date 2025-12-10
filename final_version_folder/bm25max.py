# -*- coding: utf-8 -*-
"""
corpus_docs 폴더 안의 .hwpx / .docx / .pdf 파일들을 모두 파싱해서
각 외부 문서에 대해 BM25 기준으로 가장 유사한 내부 문서와 점수를 계산하는 스크립트.

전제:
- bm25 인덱스 폴더 안에 이미 다음 파일들이 있어야 한다.
    - bm25_matrix.npz
    - bm25_vocab.json
    - bm25_docs.json
    - bm25_chunk2doc.json
- 토크나이저는 simple 또는 mecab 사용 (기본: mecab)
"""

import argparse
from pathlib import Path
import json

import hwpx_parser
import docx_parser
import pdf_parser

from bm25 import bm25_doc_similarity_from_chunks


def write_pretty_json(path, obj):
    path = Path(path)
    path.write_text(
        json.dumps(obj, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def simple_chunk(text: str):
    """아주 단순한 청킹: 줄 단위로 나누고 빈 줄 제거."""
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    return lines if lines else [text.strip()]


def parse_file(path: Path) -> str:
    suf = path.suffix.lower()
    try:
        if suf == ".hwpx":
            return hwpx_parser.parse_hwpx_to_text(str(path))
        elif suf == ".docx":
            return docx_parser.parse_docx_to_text(str(path))
        elif suf == ".pdf":
            return pdf_parser.parse_pdf_to_text(str(path))
        else:
            print(f"[SKIP] 지원하지 않는 확장자: {path.name}")
            return ""
    except Exception as e:
        print(f"[ERROR] {path.name} 파싱 실패: {e}")
        return ""


def main():
    ap = argparse.ArgumentParser(
        description="corpus_docs(.hwpx/.docx/.pdf) 전체에 대해 BM25 최대값 계산"
    )
    ap.add_argument(
        "--corpus_dir",
        required=True,
        help="외부 문서(.hwpx/.docx/.pdf)가 들어있는 폴더",
    )
    ap.add_argument(
        "--save_dir",
        required=True,
        help="BM25 인덱스 폴더",
    )
    ap.add_argument(
        "--chunk2doc",
        required=True,
        help="bm25_chunk2doc.json 경로",
    )
    ap.add_argument(
        "--tokenizer",
        choices=["simple", "mecab"],
        default="mecab",
        help="BM25 인덱스에서 사용할 토크나이저",
    )
    ap.add_argument(
        "--mode",
        choices=["sum", "max", "topMsum"],
        default="topMsum",
        help="문서 단위 집계 방식 (M은 bm25.py 내부 기본값 사용)",
    )
    ap.add_argument(
        "--topk",
        type=int,
        default=10,
        help="각 문서에 대해 상위 몇 개 내부 문서를 기록할지",
    )
    ap.add_argument(
        "--out_json",
        required=True,
        help="최종 결과 JSON 저장 경로",
    )
    args = ap.parse_args()

    corpus_dir = Path(args.corpus_dir)
    if not corpus_dir.is_dir():
        raise SystemExit(f"❌ corpus_dir 가 폴더가 아님: {corpus_dir}")

    files = sorted(list(corpus_dir.glob("*.hwpx"))
                   + list(corpus_dir.glob("*.docx"))
                   + list(corpus_dir.glob("*.pdf")))

    if not files:
        raise SystemExit(f"❌ corpus_dir 안에 hwpx/docx/pdf 파일이 없습니다: {corpus_dir}")

    print(f"[Batch] 총 파일 수: {len(files)}개\n")

    results = []

    for f in files:
        print(f"=== {f.name} ===")
        text = parse_file(f)

        if not text.strip():
            print("  → 파싱 실패 또는 내용 없음")
            results.append(
                {
                    "file": f.name,
                    "best_score": 0.0,
                    "best_doc_id": None,
                    "top_matches": [],
                }
            )
            continue

        chunks = simple_chunk(text)

        ranked = bm25_doc_similarity_from_chunks(
            chunks=chunks,
            save_dir=args.save_dir,
            chunk2doc_path=args.chunk2doc,
            tokenizer=args.tokenizer,
            mode=args.mode,
            topk=args.topk,
        )

        if not ranked:
            print("  → 코퍼스 vocab과 겹치는 토큰 없음")
            results.append(
                {
                    "file": f.name,
                    "best_score": 0.0,
                    "best_doc_id": None,
                    "top_matches": [],
                }
            )
            continue

        best_doc_id, best_score = ranked[0]
        print(f"  → best: {best_score:.6f} ({best_doc_id})")

        results.append(
            {
                "file": f.name,
                "best_score": float(best_score),
                "best_doc_id": best_doc_id,
                "top_matches": [
                    {"doc_id": doc_id, "score": float(score)}
                    for doc_id, score in ranked
                ],
            }
        )

    print("\n===== 요약: BM25 최대값 =====")
    for r in results:
        print(f"{r['file']:40s}  {r['best_score']:12.6f}  ({r['best_doc_id']})")

    write_pretty_json(args.out_json, results)
    print(f"\n[저장 완료] → {args.out_json}")


if __name__ == "__main__":
    main()

# uv run python bm25max.py --corpus_dir "C:\Users\gkseh\hansung\HS\corpus_docs" --save_dir "C:\Users\gkseh\hansung\HS\bm25_index" --chunk2doc "C:\Users\gkseh\hansung\HS\bm25_index\bm25_chunk2doc.json" --tokenizer mecab --out_json "C:\Users\gkseh\hansung\HS\bm25_file_scores.json"