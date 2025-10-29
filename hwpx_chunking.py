# filename: bullet_chunker.py
# usage:
#   python bullet_chunker.py input.json --output chunks.json
#   python bullet_chunker.py -   # STDIN으로 입력
#
# 동작:
# - 불릿 기호 '❍'만 기준으로 청킹
# - 불릿 표식/앞뒤 공백 제거
# - 중복 항목 제거(순서 보존)
# - 표식/파이프 공백 최소 정규화(보수적)

import argparse
import io
import json
import re
import sys
from collections import OrderedDict

def load_text(path: str) -> str:
    if path == "-":
        raw = sys.stdin.read()
    else:
        with io.open(path, "r", encoding="utf-8") as f:
            raw = f.read()
    obj = json.loads(raw)
    return obj["text"] if isinstance(obj, dict) and "text" in obj else str(obj)

def normalize_line(s: str) -> str:
    # 공백/특수공백 정리
    s = s.replace("\u00A0", " ").replace("\u200b", "").replace("\u202f", " ")
    s = re.sub(r"\s+", " ", s)
    # 파이프 주변 공백 정리(표 라인일 때도 한 줄 유지)
    s = re.sub(r"\s*\|\s*", "|", s)
    # 선두 불릿/대시류 제거 (혹시 끼어 있으면)
    s = re.sub(r"^\s*[❍•\-–—]\s*", "", s)
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

def chunk_by_bullet(text: str):
    # 편의를 위해 라인 사이 불릿이 섞여 있어도 분할되도록 전처리
    t = text.replace("||||", "\n")
    # '❍' 앞에 줄바꿈을 삽입해 세그먼트 경계를 확실히
    t = re.sub(r"\s*❍\s*", r"\n❍ ", t)
    parts = []
    cur = None
    for line in t.splitlines():
        if "❍" in line:
            # 새 불릿 시작
            seg = line.split("❍", 1)[1]  # '❍' 뒤만 취함
            if cur is not None:
                parts.append(cur)
            cur = seg
        else:
            # 현재 불릿이 열려 있으면 계속 붙임
            if cur is not None:
                # 표/파이프 등은 한 줄로 이어붙이기
                cur += " " + line
            else:
                # 불릿 외 본문은 버림(요청: ❍만으로 청킹)
                pass
    if cur is not None:
        parts.append(cur)

    # 정리
    parts = [normalize_line(p) for p in parts]
    parts = [p for p in parts if p]  # 빈 것 제거
    parts = dedupe_keep_order(parts)
    return parts

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input", help="입력 JSON 파일 경로 (또는 '-' 로 STDIN)")
    ap.add_argument("--output", "-o", help="출력 JSON 파일 경로(미지정 시 STDOUT)", default=None)
    args = ap.parse_args()

    text = load_text(args.input)
    chunks = chunk_by_bullet(text)
    out = json.dumps(chunks, ensure_ascii=False, indent=2)

    if args.output:
        with io.open(args.output, "w", encoding="utf-8") as f:
            f.write(out + "\n")
    else:
        sys.stdout.write(out + "\n")

if __name__ == "__main__":
    main()
