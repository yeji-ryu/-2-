# filename: bullet_chunker_extended.py
# usage:
#   python bullet_chunker_extended.py input.json --output chunks.json
#
# 기능:
# - 불릿/점 기호(❍, •, ‥)를 기준으로 문장 단위 청킹
# - 불릿 표식 제거, 앞뒤 공백 정리
# - 중복 제거(순서 유지)
# - 표(| 포함)은 한 줄로 병합

import argparse
import io
import json
import re
import sys
from collections import OrderedDict

# 청킹 기준 기호
BULLETS = r"[❍•‥]"

def load_text(path: str) -> str:
    if path == "-":
        data = sys.stdin.read()
    else:
        with io.open(path, "r", encoding="utf-8") as f:
            data = f.read()
    obj = json.loads(data)
    return obj["text"] if isinstance(obj, dict) and "text" in obj else str(obj)

def normalize_line(s: str) -> str:
    s = s.replace("\u00A0", " ").replace("\u200b", "").replace("\u202f", " ")
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"\s*\|\s*", "|", s)  # 표 공백 정리
    s = re.sub(rf"^\s*{BULLETS}\s*", "", s)  # 선두 불릿 제거
    return s.strip()

def dedupe_keep_order(items):
    seen = OrderedDict()
    out = []
    for it in items:
        key = re.sub(r"\s+", " ", it.lower()).strip()
        key = re.sub(r"\s*\|\s*", "|", key)
        if key and key not in seen:
            seen[key] = True
            out.append(it)
    return out

def chunk_by_bullets(text: str):
    # 표 단위는 한 줄로 합치기
    lines = text.replace("||||", "\n").splitlines()
    merged_lines = []
    buffer = []
    for line in lines:
        if "|" in line:
            buffer.append(line.strip())
        else:
            if buffer:
                merged_lines.append(" ".join(buffer))
                buffer = []
            merged_lines.append(line)
    if buffer:
        merged_lines.append(" ".join(buffer))

    text = "\n".join(merged_lines)
    # 불릿 기준으로 줄바꿈 삽입
    text = re.sub(rf"\s*{BULLETS}\s*", r"\n\g<0> ", text)

    parts = []
    cur = None
    for line in text.splitlines():
        if re.match(rf"^\s*{BULLETS}", line):
            # 새 항목 시작
            seg = re.sub(rf"^\s*{BULLETS}\s*", "", line)
            if cur:
                parts.append(cur)
            cur = seg
        else:
            # 기존 항목 계속 이어쓰기
            if cur is not None:
                cur += " " + line.strip()
    if cur:
        parts.append(cur)

    parts = [normalize_line(p) for p in parts if p.strip()]
    parts = dedupe_keep_order(parts)
    return parts

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input", help="입력 JSON 파일 경로 (또는 '-' 로 STDIN)")
    ap.add_argument("--output", "-o", help="출력 JSON 파일 경로", default=None)
    args = ap.parse_args()

    text = load_text(args.input)
    chunks = chunk_by_bullets(text)
    out_json = json.dumps(chunks, ensure_ascii=False, indent=2)

    if args.output:
        with io.open(args.output, "w", encoding="utf-8") as f:
            f.write(out_json + "\n")
    else:
        sys.stdout.write(out_json + "\n")

if __name__ == "__main__":
    main()


#python sentence_chunk.py input.json --output chunks.json
