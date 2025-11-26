import json, argparse, re
from pathlib import Path
from typing import List
from unittest.mock import sentinel

parser = argparse.ArgumentParser(description="Sentence-only splitter (.?! + newlines)")
parser.add_argument("--in", dest="input_path", required=True, help="입력 JSON 경로")
parser.add_argument("--out", dest="output_path", required=True, help="출력 JSON 경로")
args = parser.parse_args()

INPUT_PATH = Path(args.input_path)
OUTPUT_PATH = Path(args.output_path)

# 1) 입력 로드
with open(INPUT_PATH, "r", encoding="utf-8") as f:
    data = json.loads(f.read())

text = (data.get("text") or data.get("content") or "").strip()
if not text:
    raise ValueError("입력 JSON에 'text' 또는 'content'가 비어 있습니다.")

# 2) 문장 분리 (. ? ! + 개행)
SENT_SPLIT_RE = re.compile(r"(?<=[.!?])\s+|[\r\n]+")

def chunk_text(text: str) -> List[str]:
    """
    긴 본문 문자열을 문장 단위로 청킹해서 리스트로 반환.
    - '.', '?', '!' 뒤 공백
    - 줄바꿈 기준 분리
    """
    return [s.strip() for s in SENT_SPLIT_RE.split(text) if s.strip()]

# 3) 저장 — 리스트 형태로
with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
    json.dump(sentinel, f, ensure_ascii=False, indent=2)
def main():
    parser = argparse.ArgumentParser(description="Sentence-only splitter (.?! + newlines)")
    parser.add_argument("--in", dest="input_path", required=True, help="입력 JSON 경로")
    parser.add_argument("--out", dest="output_path", required=True, help="출력 JSON 경로")
    args = parser.parse_args()

    INPUT_PATH = Path(args.input_path)
    OUTPUT_PATH = Path(args.output_path)

    # 1) 입력 로드
    with open(INPUT_PATH, "r", encoding="utf-8") as f:
        data = json.loads(f.read())

    text = (data.get("text") or data.get("content") or "").strip()
    if not text:
        raise ValueError("입력 JSON에 'text' 또는 'content'가 비어 있습니다.")

    # 2) 문장 분리 (공통 함수 사용)
    sentences = chunk_text(text)

    # 3) 저장 — 리스트 형태로
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(sentences, f, ensure_ascii=False, indent=2)

    print(f"분리된 문장 수: {len(sentences)}")
    print(f"저장: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()


#uv run python docx_chunking.py --in "C:/Users/gkseh/hansung/pz1023/docxparser.json" --out "C:/Users/gkseh/hansung/pz1023/docxchunk.json"
