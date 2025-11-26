# -*- coding: utf-8 -*-
import argparse
import json
import re
from pathlib import Path
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl


def _normalize_keep_newlines(s: str) -> str:
    """줄바꿈은 유지하되, 각 줄의 불필요한 공백 제거"""
    lines = s.splitlines()
    cleaned = []
    for ln in lines:
        ln = re.sub(r"\s+", " ", ln.strip())
        if ln:
            cleaned.append(ln)
    return "\n".join(cleaned)


def parse_docx(file_path: Path):
    """DOCX 문서에서 텍스트 및 표 데이터를 추출"""
    doc = Document(file_path)
    content = []

    for block in doc.element.body:
        # 일반 문단
        if isinstance(block, CT_P):
            paragraph = Paragraph(block, doc)
            text = _normalize_keep_newlines(paragraph.text)
            if text:
                content.append(text)

        # 표
        elif isinstance(block, CT_Tbl):
            table = Table(block, doc)
            rows = []
            for row in table.rows:
                row_data = [cell.text.strip() for cell in row.cells]
                if any(row_data):
                    rows.append(" | ".join(re.sub(r"\s+", " ", c) for c in row_data if c))
            if rows:
                content.append("\n".join(rows))
    return content

def parse_docx_to_blocks(path: str | Path):
    """
    서버/다른 스크립트에서 재사용하기 위한 래퍼.
    DOCX → [블록 문자열, ...] 형태로 그대로 반환.
    """
    return parse_docx(Path(path))


def parse_docx_to_text(path: str | Path) -> str:
    """
    server에서 사용할 DOCX → 전체 텍스트 문자열 변환 헬퍼.
    (표는 ' | '로 인라인, 줄바꿈은 \n 유지)
    """
    p = Path(path)
    blocks = parse_docx(p)
    return "\n".join(blocks).strip()



if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="DOCX 파일 파서 (본문 + 표, 줄바꿈 유지)")
    parser.add_argument("--in", dest="input_path", help="입력 DOCX 파일 경로")
    parser.add_argument("--out", dest="output_path", help="출력 JSON 파일 경로 (선택사항)")
    args = parser.parse_args()

    cwd = Path.cwd()
    input_path = Path(args.input_path) if args.input_path else cwd / "input.docx"

    # 입력 파일 이름 기반으로 자동 출력 파일명 생성
    base_name = input_path.stem  # 확장자 제외 파일명
    output_path = Path(args.output_path) if args.output_path else cwd / f"{base_name}_parsed.json"

    if not input_path.exists():
        raise FileNotFoundError(f"❌ 파일을 찾을 수 없습니다: {input_path}")

    print(f"📂 DOCX 파싱 시작: {input_path.name}")
    blocks = parse_docx(input_path)

    text_with_newlines = "\n".join(blocks).strip()

    output_data = {
        "filename": input_path.name,
        "text": text_with_newlines
    }

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(output_data, f, ensure_ascii=False, indent=2)

    print(f"✅ 파싱 완료 → {output_path}")
  
#uv run python docx_parser.py --in "C:/Users/gkseh/hansung/pz1023/그린퓨처 투자보고서.docx"
