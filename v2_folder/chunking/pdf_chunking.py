import json, argparse, re
from pathlib import Path

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
    raise ValueError("❌ 입력 JSON에 'text' 또는 'content'가 비어 있습니다.")

# 2) 문장 분리 (. ? ! + 개행)  — 윈도우 \r\n 포함
SENT_SPLIT_RE = re.compile(r'(?<=(?<!\d)\.(?!\d))\s+|(?<=[!?])\s+|[\r\n]+')
sentences = [s.strip() for s in SENT_SPLIT_RE.split(text) if s.strip()]

# 3) 저장
out = [{"sent_id": i, "text": s} for i, s in enumerate(sentences)]
with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=2)

print(f"🧩 분리된 문장 수: {len(sentences)}")
print(f"💾 저장: {OUTPUT_PATH}")

#uv run python pdfchunk.py --in "C:/Users/gkseh/hansung/pz1023/pdfparser.json" --out "C:/Users/gkseh/hansung/pz1023/pdfchunk.json"
