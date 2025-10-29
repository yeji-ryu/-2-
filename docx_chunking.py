# sentence_chunker.py
import sys, json, re, unicodedata, argparse
from typing import Any, Iterable, List

# --- stdout을 UTF-8로 강제 (PowerShell cp949 오류 방지) ---
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# ===== 분리 기준 =====
SENT_END = r"\.!\?。！？…"  # 종결부호
BULLETS = "❍●○■□▪▫◼◻◾◽•·∙‧▶►▸–—"  # 불릿/도트 기호
BULLET_PATTERN = "[" + re.escape(BULLETS) + "]"

# 종결부호 기준 분리 (look-behind 미사용)
SENT_SPLIT_REGEX = re.compile(
    r"([" + SENT_END + r"]+)(?=\s+(?=[\uAC00-\uD7A3A-Za-z0-9\"'“”‘’(【\[-–—]))"
)
# 불릿 기준 분리
BULLET_SPLIT_REGEX = re.compile(r"\s*" + BULLET_PATTERN + r"\s*")

# 약어 과분할 완화
ABBREV_TAIL = re.compile(
    r"(?:\b(?:Mr|Mrs|Ms|Dr|Prof|Inc|Ltd|Co|No|vs|etc|e\.g|i\.e)\.)$",
    re.IGNORECASE,
)
NOISE_REGEX = re.compile(r"[<>]{3,}")
DIGIT_TITLE_FIX = re.compile(r"(\d)([^\d\s])")
HEADING_FIX = re.compile(r"([IVXⅰ-ⅳⅠ-Ⅻ]+)([^\s])", re.IGNORECASE)

def normalize_text(s: str) -> str:
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{2,}", "\n", s)
    s = DIGIT_TITLE_FIX.sub(r"\1 \2", s)
    s = HEADING_FIX.sub(r"\1 \2", s)
    return s.strip()

def collect_strings(obj: Any) -> Iterable[str]:
    if isinstance(obj, dict):
        if "text" in obj and isinstance(obj["text"], str):
            yield obj["text"]
        for k, v in obj.items():
            if k != "text":
                yield from collect_strings(v)
    elif isinstance(obj, list):
        for v in obj:
            yield from collect_strings(v)
    elif isinstance(obj, str):
        yield obj

# ===== 표(테이블) 감지 =====
TABLE_HINT_WORDS = ["구분", "항목", "합계", "총계"]
TABLE_PUNCT_WEAK = re.compile(r"^[^\.!\?。！？…]*$")
MANY_NUMBERS = re.compile(r"(?:\d[\d,\.%원억]*){3,}")
PIPE_SEP = re.compile(r"\s*\|\s*")
MANY_SPACES = re.compile(r"\s{2,}")

def is_table_like(line: str) -> bool:
    l = line.strip()
    if not l:
        return False
    if "|" in l and PIPE_SEP.search(l):
        return True
    if any(w in l for w in TABLE_HINT_WORDS):
        return True
    if TABLE_PUNCT_WEAK.match(l) and (MANY_NUMBERS.search(l) or MANY_SPACES.search(l)):
        return True
    return False

def merge_table_block(lines: List[str], start_idx: int) -> (str, int):
    buf = []
    i = start_idx
    while i < len(lines):
        line = lines[i].strip()
        if not line or not is_table_like(line):
            break
        # 파이프(|) 구분자 표 형태
        if "|" in line and PIPE_SEP.search(line):
            cols = [c.strip() for c in PIPE_SEP.split(line) if c.strip()]
            buf.append(" | ".join(cols))
        else:
            # 다중 공백 압축
            line = MANY_SPACES.sub(" ", line)
            buf.append(line)
        i += 1
    merged = " ".join(buf)
    merged = re.sub(r"\s{2,}", " ", merged).strip()
    return merged, i

# ===== 문장 분리 =====
def split_sentences_in_line(text: str) -> List[str]:
    text = text.strip()
    if not text:
        return []

    # 불릿 기준 1차 분리
    parts_by_bullet = [p for p in BULLET_SPLIT_REGEX.split(text) if p and p.strip()]
    sentences: List[str] = []

    for chunk in parts_by_bullet:
        chunk = chunk.strip()
        if not chunk:
            continue
        pieces = SENT_SPLIT_REGEX.split(chunk)
        buf = ""
        for p in pieces:
            if not p.strip():
                continue
            if re.fullmatch("[" + SENT_END + r"]+", p):
                buf += p
            else:
                if buf:
                    if ABBREV_TAIL.search(buf):
                        buf = f"{buf} {p}"
                        continue
                    sentences.append(buf.strip())
                buf = p
        if buf:
            sentences.append(buf.strip())

    # 너무 긴 문장은 공백 경계로 나누기
    final = []
    for s in sentences:
        if len(s) <= 300:
            final.append(s)
            continue
        start = 0
        while start < len(s):
            end = min(len(s), start + 300)
            if end < len(s):
                m = re.search(r"\s", s[end:end + 50])
                if m:
                    end = end + m.start()
            final.append(s[start:end].strip())
            start = end
    return final

def dedup_consecutive(items: List[str]) -> List[str]:
    out = []
    prev = None
    for x in items:
        if not x:
            continue
        if prev and normalize_text(x) == normalize_text(prev):
            continue
        out.append(x)
        prev = normalize_text(x)
    return out

def read_text_any_encoding(path: str, encs: List[str]) -> str:
    with open(path, "rb") as f:
        raw = f.read()
    for e in encs:
        try:
            return raw.decode(e)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")

def load_json_any_encoding(path: str, forced: str = None):
    order = ["utf-8", "utf-8-sig", "cp949", "euc-kr"]
    if forced:
        order = [forced] + [e for e in order if e != forced]
    text = read_text_any_encoding(path, order)
    return json.loads(text)

# ===== main =====
def main():
    ap = argparse.ArgumentParser(description="문장/불릿/표 단위 청킹")
    ap.add_argument("input", help="입력 JSON 경로")
    ap.add_argument("--output", help="출력 파일 경로 (UTF-8)")
    ap.add_argument("--format", choices=["json", "jsonl"], default="json")
    ap.add_argument("--input-encoding", help="입력 인코딩 강제 (utf-8, cp949 등)")
    args = ap.parse_args()

    data = load_json_any_encoding(args.input, forced=args.input_encoding)
    strings = list(collect_strings(data))
    big_text = "\n".join(normalize_text(s) for s in strings if s.strip())

    # 불릿 앞에 줄바꿈 추가
    prepped = re.sub(BULLET_PATTERN, lambda m: f"\n{m.group(0)} ", big_text)
    lines = [l.strip() for l in prepped.split("\n")]

    chunks = []
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue
        line = NOISE_REGEX.sub("", line).strip()
        if not line:
            i += 1
            continue

        # 표(테이블) 블록 병합
        if is_table_like(line):
            merged, next_i = merge_table_block(lines, i)
            if merged:
                chunks.append(merged)
            i = next_i
            continue

        # 일반 문장 분리
        for s in split_sentences_in_line(line):
            chunks.append(s)
        i += 1

    chunks = [re.sub(r"\s{2,}", " ", c).strip() for c in chunks if c.strip()]
    chunks = dedup_consecutive(chunks)

    if args.format == "jsonl":
        out_text = "\n".join(json.dumps({"sentence": c}, ensure_ascii=False) for c in chunks)
    else:
        out_text = json.dumps({"sentences": chunks}, ensure_ascii=False, indent=2)

    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(out_text)
    else:
        print(out_text)

if __name__ == "__main__":
    main()
