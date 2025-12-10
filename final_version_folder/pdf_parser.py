# -*- coding: utf-8 -*-
import argparse
import json
import re
import unicodedata
from pathlib import Path
import pdfplumber

# 글머리 기호
BULLET_CHARS = ''.join([
    '•', '◦', '●', '○', '‣', '∙',
    '', '', '', '', '', '', '', '', '', '-', '–'
])
BULLET_START_RE = re.compile(rf'^[\s]*[{re.escape(BULLET_CHARS)}]+[\s·\t]*')

# 문장 종결/경계 문자(이 뒤는 보통 줄이 바뀌어도 OK)
ENDERS = set(list('.!?…?!。！？：:;)]」』”’>'))

# 항목/헤더 패턴(시작 줄이면 새 줄로 유지)
HEADING_RE = re.compile(
    r'^(기업명|대표자|주요\s*사업장|주요\s*사업|주요\s*고객사\s*/\s*계약처|개요|요약|목차)\s*:'
)

# 번호/장/절 등 목록/헤더 같은 것
LIST_OR_SECTION_RE = re.compile(r'^(\(?[0-9０-９IVXivx]+[\).]|\d+\s|제\s*\d+\s*(장|절)|[A-Za-z]\))\s')


def clean_line(line: str) -> str:
    """글머리 제거 + 라인 내부 공백 정리(줄바꿈 판단은 여기서 안 함)"""
    s = unicodedata.normalize('NFKC', line)
    s = BULLET_START_RE.sub('', s)
    s = re.sub(r'\s+', ' ', s.strip())
    return s

JOIN_DASHES = "-\u2010\u2011\u2013\u2014\u2212"

_HANGUL_ONLY = re.compile(r'^[\uAC00-\uD7A3]')

def is_heading_like(s: str) -> bool:
    return bool(HEADING_RE.match(s) or LIST_OR_SECTION_RE.match(s))

def _is_hangul(ch: str) -> bool:
    return bool(ch) and bool(_HANGUL_ONLY.match(ch))

def should_join(prev: str, curr: str) -> bool:
    """
    강제 줄바꿈처럼 보이면 True.
    단, **한글-한글**로 이어지는 경우에만 붙임.
    """
    if not prev or not curr:
        return False

    # 하이픈 줄바꿈: 다음 글자가 한글일 때만 붙임 (예: "에너지저장시스- \n 템")
    if prev[-1] in JOIN_DASHES and _is_hangul(curr[0]):
        return True

    # 이전 줄이 종결/경계 문자로 끝나면 새 줄 유지
    if prev[-1] in ENDERS:
        return False
    
    # 현재 줄이 헤더/새 항목/글머리로 시작하면 새 줄 유지
    if is_heading_like(curr) or BULLET_START_RE.match(curr):
        return False

    # ★ 한글-한글 이어짐일 때만 붙임
    return _is_hangul(prev[-1]) and _is_hangul(curr[0])

def pdf_to_text(pdf_path: Path, out_json: Path):
    """PDF → 텍스트 (의도치 않은 줄바꿈은 제거, 의미 있는 줄바꿈은 유지)"""
    merged_lines = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            for raw in text.split("\n"):
                cur = clean_line(raw)

                if not cur:
                    continue

                if merged_lines and should_join(merged_lines[-1], cur):
                    # 하이픈 줄바꿈이었다면 하이픈 제거 후 붙이기
                    if merged_lines[-1].endswith('-'):
                        merged_lines[-1] = merged_lines[-1][:-1] + cur
                    else:
                        # 한국어/영문이 연속될 땐 공백 없이 붙이는 편이 자연스러움
                        merged_lines[-1] = merged_lines[-1] + cur
                else:
                    merged_lines.append(cur)

            # 페이지 구분 빈 줄을 넣고 싶으면 주석 해제
            # merged_lines.append("")

    text_with_newlines = "\n".join(ln for ln in merged_lines if ln).strip()

    output_data = {
        "filename": pdf_path.name,
        "text": text_with_newlines
    }

    out_json.parent.mkdir(parents=True, exist_ok=True)
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(output_data, f, ensure_ascii=False, indent=2)

    print(f"✅ 텍스트 추출 완료 → {out_json}")
    return output_data

def parse_pdf_to_text(path: str | Path) -> str:
    """
    server에서 사용할 PDF → 전체 텍스트 문자열 변환 헬퍼.
    파일을 쓰지 않고 텍스트만 반환.
    """
    p = Path(path)
    merged_lines = []

    with pdfplumber.open(p) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for raw in text.split("\n"):
                cur = clean_line(raw)
                if not cur:
                    continue
                if merged_lines and should_join(merged_lines[-1], cur):
                    if merged_lines[-1].endswith('-'):
                        merged_lines[-1] = merged_lines[-1][:-1] + cur
                    else:
                        merged_lines[-1] = merged_lines[-1] + cur
                else:
                    merged_lines.append(cur)

    return "\n".join(ln for ln in merged_lines if ln).strip()



if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="PDF 파일 파서 (워드랩 줄바꿈 제거, 의미 줄바꿈 유지)")
    parser.add_argument("--in", dest="input_path", help="입력 PDF 파일 경로 (기본: ./input.pdf)")
    parser.add_argument("--out", dest="output_path", help="출력 JSON 파일 경로 (기본: ./parsed_pdf.json)")
    args = parser.parse_args()

    cwd = Path.cwd()
    input_path = Path(args.input_path) if args.input_path else cwd / "input.pdf"
    output_path = Path(args.output_path) if args.output_path else cwd / "parsed_pdf.json"

    if not input_path.exists():
        raise FileNotFoundError(f"❌ 파일을 찾을 수 없습니다: {input_path}")

    print(f"📂 PDF 파싱 시작: {input_path.name}")
    pdf_to_text(input_path, output_path)

#uv run python pdf_parser.py --in "C:/Users/gkseh/hansung/pz1023/그린퓨처 투자보고서.pdf" --out "C:/Users/gkseh/hansung/pz1023/pdfparser.json"