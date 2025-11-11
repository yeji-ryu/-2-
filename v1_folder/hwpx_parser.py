# -*- coding: utf-8 -*-
"""
HWPX → JSON parser with EasyOCR (images → text in-place, one line per image)

Updates (2025-10-10)
- 이미지 OCR 결과: 한 이미지당 한 줄로 합쳐서 출력(join)
- TOC 제거 기본 적용(목차/목 차/Contents 등)
- 표 중복 제거: 테이블을 원자 블록으로 처리
- 이미지 참조 불명 시 아카이브 내 남은 이미지 순차 매핑 시도

출력 규칙
- 문단: 한 문단 = 한 줄
- 표: 각 행을 `col1 | col2 | ...` 로 직렬화(한 행 = 한 줄)
- 이미지: 같은 위치에서 OCR 결과 **한 줄**로 삽입(없으면 자리표시)

CLI
    python hwpx_parser_with_easyocr.py INPUT.hwpx --ocr --ocr-lang ko,en --pretty --out out.json
"""
from __future__ import annotations
import sys, os, re, io, zipfile, argparse, json, tempfile, shutil
import xml.etree.ElementTree as ET
from typing import List, Dict, Tuple, Optional, Iterable

# ------- Optional EasyOCR -------
try:
    import easyocr  # type: ignore
    _HAS_EASYOCR = True
except Exception:
    _HAS_EASYOCR = False

# ------- Namespaces (best-effort) -------
NS = {
    'w': 'http://www.hancom.co.kr/hwpml/2011/wordprocessor',
    'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
    'hc': 'http://www.hancom.co.kr/hwpml/2011/common',
    'r':  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
for k, v in NS.items():
    ET.register_namespace(k, v)

# Localnames we care about
T_P   = 'p'
T_RUN = 'run'
T_T   = 't'
T_TBL = 'tbl'
T_TR  = 'tr'
T_TC  = 'tc'
T_IMG = 'img'

IMG_EXTS = {'.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.tiff', '.webp'}

# ----------------- helpers -----------------
def _local(tag: str) -> str:
    return tag.rsplit('}', 1)[1] if '}' in tag else tag

def _iter_sections(zf: zipfile.ZipFile) -> List[str]:
    names = [n for n in zf.namelist() if n.startswith('Contents/') and n.endswith('.xml')]
    sections = sorted([n for n in names if re.search(r'/section\d+\.xml$', n)],
                      key=lambda s: int(re.search(r'(\d+)', s).group(1)) if re.search(r'(\d+)', s) else 0)
    others = sorted(set(names) - set(sections))
    return sections + others

def _text_from_para(p: ET.Element) -> str:
    parts: List[str] = []
    for run in p.iter():
        if _local(run.tag) == T_T and run.text:
            parts.append(run.text)
    s = re.sub(r'\s+', ' ', ''.join(parts)).strip()
    return s

def _rows_from_table(tbl: ET.Element) -> List[List[str]]:
    rows: List[List[str]] = []
    for tr in list(tbl):
        if _local(tr.tag) != T_TR:
            continue
        row: List[str] = []
        for tc in list(tr):
            if _local(tc.tag) != T_TC:
                continue
            cell_parts: List[str] = []
            for node in tc.iter():
                if _local(node.tag) == T_T and node.text:
                    cell_parts.append(node.text)
            row.append(re.sub(r'\s+', ' ', ''.join(cell_parts)).strip())
        if any(x for x in row):
            rows.append(row)
    return rows

def _image_index(zf: zipfile.ZipFile) -> Dict[str, str]:
    idx: Dict[str, str] = {}
    for name in zf.namelist():
        ext = os.path.splitext(name)[1].lower()
        if ext in IMG_EXTS and name.startswith('Contents/'):
            base = os.path.basename(name)
            idx[base] = name
            idx[os.path.splitext(base)[0]] = name
    return idx

# --- OCR ---
def _ocr_image(reader, img_bytes: bytes) -> List[str]:
    try:
        import numpy as np, cv2
        nparr = np.frombuffer(img_bytes, np.uint8)
        img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
        if img is not None:
            return [s.strip() for s in reader.readtext(img, detail=0) if s.strip()]
    except Exception:
        pass
    # Fallback temp file path
    with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
        tmp.write(img_bytes)
        tmp.flush()
        p = tmp.name
    try:
        return [s.strip() for s in reader.readtext(p, detail=0) if s.strip()]
    finally:
        try: os.unlink(p)
        except Exception: pass

# --- TOC filter ---
TOC_PATTERNS = [
    re.compile(r'\b목\s*차\b'),
    re.compile(r'\b목차\b'),
    re.compile(r'\bTable of Contents\b', re.I),
    re.compile(r'^(?:[IVXLC]+|[Ⅰ-Ⅹ]+)[\.)]\s'),  # roman numerals
]
DOT_LEADER = re.compile(r'\.{3,}')

def _is_toc_line(s: str) -> bool:
    if not s:
        return False
    if any(p.search(s) for p in TOC_PATTERNS):
        return True
    # 점선 리더+페이지번호
    if DOT_LEADER.search(s) and re.search(r'\b\d{1,3}\b', s):
        return True
    # 섹션 인덱스 스타일(숫자+항목) 과도 나열
    if re.match(r'^(?:\d+[\.|)]\s+){1,4}.+?$', s):
        return True
    return False

# --- top-level block traversal (table is atomic) ---
def _yield_top_blocks(root: ET.Element) -> Iterable[ET.Element]:
    """Depth-first traversal but treat <tbl> as atomic: don't yield its children separately."""
    stack = [root]
    while stack:
        el = stack.pop()
        children = list(el)
        lname = _local(el.tag)
        if lname == T_TBL:
            yield el
            continue  # atomic
        if lname in (T_P, T_IMG):
            yield el
        for ch in reversed(children):
            stack.append(ch)

# ----------------- main parse -----------------
def parse_hwpx_to_lines(path: str, use_ocr: bool = False, ocr_lang: str = 'ko,en', drop_toc: bool = True) -> List[str]:
    if not zipfile.is_zipfile(path):
        raise ValueError('Not a valid HWPX (zip) file: %s' % path)

    lines: List[str] = []

    reader = None
    if use_ocr and _HAS_EASYOCR:
        langs = [x.strip() for x in ocr_lang.split(',') if x.strip()]
        reader = easyocr.Reader(langs)
    elif use_ocr and not _HAS_EASYOCR:
        print('[warn] --ocr requested but easyocr is not available. Proceeding without OCR.', file=sys.stderr)

    with zipfile.ZipFile(path, 'r') as zf:
        img_idx = _image_index(zf)
        remaining_imgs = [n for n in zf.namelist() if os.path.splitext(n)[1].lower() in IMG_EXTS]
        rem_i = 0

        for sec in _iter_sections(zf):
            try:
                root = ET.fromstring(zf.read(sec))
            except Exception:
                continue

            last_table_rows: List[str] = []

            for el in _yield_top_blocks(root):
                lname = _local(el.tag)

                # Paragraph
                if lname == T_P:
                    s = _text_from_para(el)
                    if not s:
                        continue
                    if drop_toc and _is_toc_line(s):
                        continue
                    if drop_toc and s.strip() in ('목차', '목 차', 'Contents', 'Table of Contents'):
                        continue
                    lines.append(s)

                # Table (atomic block)
                elif lname == T_TBL:
                    rows = _rows_from_table(el)
                    ser = [' | '.join(r) for r in rows if any(c for c in r)]
                    if ser and ser != last_table_rows:
                        lines.extend(ser)
                        last_table_rows = ser

                # Image
                elif lname == T_IMG:
                    if use_ocr and reader is not None:
                        ref = None
                        for attr in ('src', 'ref', '{%s}link' % NS['r'], 'link', 'r:id'):
                            if attr in el.attrib:
                                ref = el.attrib.get(attr)
                                break
                        img_name = None
                        if ref:
                            base = os.path.basename(ref)
                            img_name = img_idx.get(base) or img_idx.get(os.path.splitext(base)[0])
                        if not img_name and rem_i < len(remaining_imgs):
                            img_name = remaining_imgs[rem_i]
                            rem_i += 1
                        if img_name and img_name in zf.namelist():
                            try:
                                ocr_txts = _ocr_image(reader, zf.read(img_name))
                                if ocr_txts:
                                    # --- one line per image ---
                                    one_line = ' '.join(t.strip() for t in ocr_txts if t.strip()).strip()
                                    if one_line:
                                        lines.append(one_line)
                                    else:
                                        lines.append('그림입니다. (OCR 결과 없음)')
                                else:
                                    lines.append('그림입니다. (OCR 결과 없음)')
                            except Exception:
                                lines.append('그림입니다. (OCR 오류)')
                        else:
                            lines.append('그림입니다. (이미지 참조 불명)')
                    else:
                        lines.append('그림입니다.')

        # Optional: if no "그림입니다"가 없고 남은 이미지가 있으면 후처리로 OCR
        if use_ocr and reader is not None and not any('그림입니다' in s for s in lines):
            for ip in remaining_imgs:
                try:
                    ocr_txts = _ocr_image(reader, zf.read(ip))
                    if ocr_txts:
                        # --- one line per image (post pass) ---
                        one_line = ' '.join(t.strip() for t in ocr_txts if t.strip()).strip()
                        if one_line:
                            lines.append(one_line)
                except Exception:
                    pass

    return lines

# ----------------- CLI -----------------
def main():
    ap = argparse.ArgumentParser(description='HWPX → JSON lines with EasyOCR. TOC filtered, tables deduped.')
    ap.add_argument('input')
    ap.add_argument('--ocr', action='store_true', help='Enable EasyOCR for images')
    ap.add_argument('--ocr-lang', default='ko,en', help='Languages for EasyOCR (comma sep)')
    ap.add_argument('--keep-toc', action='store_true', help='Do NOT filter out table of contents')
    ap.add_argument('--out', default='-', help='Output .json path or - for stdout')
    ap.add_argument('--pretty', action='store_true', help='Pretty-print JSON')
    args = ap.parse_args()

    lines = parse_hwpx_to_lines(
        args.input,
        use_ocr=args.ocr,
        ocr_lang=args.ocr_lang,
        drop_toc=(not args.keep_toc)
    )

    js = json.dumps(lines, ensure_ascii=False, indent=2 if args.pretty else None)
    if args.out == '-' or args.out.lower() == 'stdout':
        print(js)
    else:
        with open(args.out, 'w', encoding='utf-8') as f:
            f.write(js)
        print(f'[ok] wrote {args.out}')

if __name__ == '__main__':
    main()
