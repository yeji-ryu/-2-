# filename: hwpx_chunking.py
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
    # 입력이 JSON 객체라면 text 필드 추출, 아니면 그대로 사용
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
    s = re.sub(r"\s*\|\s*", "|", s)  # 표 공백 정리
    # 여기서 선두 불릿을 지우면 나중에 구분이 안되므로, 청킹 로직 안에서 처리하거나 유지
    # 사용자 의도상 내용 정제 때는 지우는 게 맞으므로 유지하되, 청킹 로직 확인 필요
    s = re.sub(rf"^\s*{BULLETS}\s*", "", s)  
    return s.strip()

def dedupe_keep_order(items):
    seen = OrderedDict()
    out = []
    for it in items:
        # 비교를 위한 키 생성 (소문자, 공백 압축)
        key = re.sub(r"\s+", " ", it.lower()).strip()
        key = re.sub(r"\s*\|\s*", "|", key)
        if key and key not in seen:
            seen[key] = True
            out.append(it)
    return out

def chunk_by_bullets(text: str):
    # 1. 전처리: 줄바꿈 문자 통일
    text = text.replace("||||", "\n")
    lines = text.splitlines()

    # 2. 표(| 포함된 라인) 병합 로직 개선
    merged_lines = []
    buffer = []
    
    for line in lines:
        stripped = line.strip()
        if not stripped: continue # 빈 줄 건너뛰기

        # '|'가 포함된 줄은 표로 간주하여 버퍼에 모음
        if "|" in stripped:
            buffer.append(stripped)
        else:
            # 버퍼에 내용이 있으면 털어내고 병합 (표 끝남)
            if buffer:
                merged_lines.append(" ".join(buffer))
                buffer = []
            merged_lines.append(stripped)
    
    # 남은 버퍼 처리
    if buffer:
        merged_lines.append(" ".join(buffer))

    # 재조립된 텍스트
    text = "\n".join(merged_lines)

    # 3. 청킹 로직 (핵심 수정 부분)
    parts = []
    cur_lines = []
    
    # 정규식 컴파일
    bullet_pattern = re.compile(rf"^\s*{BULLETS}")

    for line in text.splitlines():
        # 불릿으로 시작하는 줄을 만나면
        if bullet_pattern.match(line):
            # 기존에 모아둔 내용이 있다면 하나의 청크로 저장 (Flush)
            if cur_lines:
                parts.append(" ".join(cur_lines))
                cur_lines = []
            
            # 불릿 기호 제거 후 현재 줄 시작 (새 청크 시작)
            clean_line = re.sub(rf"^\s*{BULLETS}\s*", "", line).strip()
            cur_lines.append(clean_line)
        else:
            # 불릿이 아니면 현재 청크에 계속 이어 붙임
            # (수정 전에는 여기서 cur가 None이면 버렸음 -> 이제는 그냥 cur_lines에 추가)
            cur_lines.append(line.strip())

    # 마지막 남은 청크 저장
    if cur_lines:
        parts.append(" ".join(cur_lines))

    # 4. 정제 및 중복 제거
    final_parts = [normalize_line(p) for p in parts if p.strip()]
    final_parts = dedupe_keep_order(final_parts)
    
    return final_parts

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input", help="입력 JSON 파일 경로 (또는 '-' 로 STDIN)")
    ap.add_argument("--output", "-o", help="출력 JSON 파일 경로", default=None)
    args = ap.parse_args()

    text = load_text(args.input)
    chunks = chunk_by_bullets(text)
    
    # 리스트 형태로 바로 덤프
    out_json = json.dumps(chunks, ensure_ascii=False, indent=2)

    if args.output:
        with io.open(args.output, "w", encoding="utf-8") as f:
            f.write(out_json + "\n")
    else:
        sys.stdout.write(out_json + "\n")

if __name__ == "__main__":
    main()