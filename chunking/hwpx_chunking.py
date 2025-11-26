# filename: hwpx_chunking.py
import argparse
import io
import json
import re
import sys
from collections import OrderedDict

# 청킹 기준 기호 (HWP에서 자주 쓰는 특수문자 포함)
# ❍, •, ‥, ․(작은 점), -(하이픈) 등
BULLETS_REGEX = r"([❍•‥\uF06C\uF06E\uF0D8\u2022])"

def load_text(path: str) -> str:
    if path == "-":
        data = sys.stdin.read()
    else:
        with io.open(path, "r", encoding="utf-8") as f:
            data = f.read()
    if isinstance(data, str):
        try:
            obj = json.loads(data)
            return obj["text"] if isinstance(obj, dict) and "text" in obj else str(obj)
        except json.JSONDecodeError:
            return data
    return str(data)

def normalize_line(s: str) -> str:
    s = s.replace("\u00A0", " ").replace("\u200b", "").replace("\u202f", " ")
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"\s*\|\s*", "|", s)
    # 청킹 후 맨 앞의 불릿 기호는 제거하여 깔끔하게 만듦
    s = re.sub(rf"^\s*{BULLETS_REGEX}\s*", "", s)
    return s.strip()

def dedupe_keep_order(items):
    seen = OrderedDict()
    out = []
    for it in items:
        key = re.sub(r"\s+", " ", it.lower()).strip()
        if key and key not in seen:
            seen[key] = True
            out.append(it)
    return out

def chunk_by_bullets(text: str):
    # 1. 전처리: 줄바꿈 문자 통일
    text = text.replace("||||", "\n").replace("\r\n", "\n")

    # ★ 핵심 수정: 문장 중간에 숨은 불릿 앞에도 강제로 줄바꿈(\n) 삽입 ★
    # 예: "제목 ❍ 내용" -> "제목 \n❍ 내용"
    text = re.sub(BULLETS_REGEX, r"\n\1", text)

    lines = text.splitlines()
    
    # 2. 표(| 포함된 라인) 병합 로직
    merged_lines = []
    buffer = []
    
    for line in lines:
        stripped = line.strip()
        if not stripped: continue 

        if "|" in stripped:
            buffer.append(stripped)
        else:
            if buffer:
                merged_lines.append(" ".join(buffer))
                buffer = []
            merged_lines.append(stripped)
    
    if buffer:
        merged_lines.append(" ".join(buffer))

    # 3. 불릿 기준으로 그룹핑 (청킹)
    final_text = "\n".join(merged_lines)
    parts = []
    cur_lines = []
    
    # 줄의 시작이 불릿인지 확인하는 패턴
    start_bullet_pattern = re.compile(rf"^\s*{BULLETS_REGEX}")

    for line in final_text.splitlines():
        # 불릿으로 시작하면 새로운 청크 시작
        if start_bullet_pattern.match(line):
            if cur_lines:
                parts.append(" ".join(cur_lines))
                cur_lines = []
            
            # 불릿 기호 자체는 제거하고 내용만 담음 (선택사항)
            clean_line = re.sub(rf"^\s*{BULLETS_REGEX}\s*", "", line).strip()
            cur_lines.append(clean_line)
        else:
            # 불릿이 아니면 앞 내용에 이어 붙임
            cur_lines.append(line.strip())

    if cur_lines:
        parts.append(" ".join(cur_lines))

    # 4. 정제 및 중복 제거
    final_parts = [normalize_line(p) for p in parts if p.strip()]
    final_parts = dedupe_keep_order(final_parts)
    
    return final_parts

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input", help="입력 JSON 파일 경로")
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
