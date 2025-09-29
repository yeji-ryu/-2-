# -*- coding: utf-8 -*-
# hwp_flat_lines_json_v2.py
# - .hwp(HWP5, OLE) / .hwpx(ZIP+XML) 자동 감지
# - 문서와 유사하게 "세로로 한 줄씩" JSON 배열에 담아서 출력
# - HWPX: p/para/paragraph/br/tbl/tr/tc/title 등에서 줄바꿈
# - HWP : 문단 텍스트 레코드(66)마다 줄바꿈
# - 옵션: --hard-wrap N  (긴 줄 강제 줄나눔)

from __future__ import annotations
import sys, struct, zlib, json, re
from pathlib import Path
import zipfile, xml.etree.ElementTree as ET

# 콘솔 UTF-8
if hasattr(sys.stdout, "reconfigure"):
    try: sys.stdout.reconfigure(encoding="utf-8")
    except Exception: pass

# hwp(OLE) 선택 의존성
try:
    import olefile
    _HAS_OLE = True
except Exception:
    _HAS_OLE = False

OLE_SIG = b"\xD0\xCF\x11\xE0"
ZIP_SIG = b"PK\x03\x04"
HWPTAG_PARA_TEXT = 66

def detect_container(p: Path) -> str:
    if not p.is_file():
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {p}")
    head = p.read_bytes()[:8]
    if head.startswith(OLE_SIG): return "HWP5"
    if head.startswith(ZIP_SIG): return "HWPX"
    ext = p.suffix.lower()
    if ext == ".hwp": return "HWP5"
    if ext == ".hwpx": return "HWPX"
    return "UNKNOWN"

# ---------------- HWP (.hwp) ----------------
def _read_fileheader(ole: "olefile.OleFileIO") -> dict:
    with ole.openstream("FileHeader") as fp:
        raw = fp.read()
    if len(raw) < 48: raise ValueError("FileHeader 길이 오류")
    ver,   = struct.unpack("<I", raw[32:36])
    attr1, = struct.unpack("<I", raw[36:40])  # bit0=compressed, bit1=encrypted
    return {"compressed": bool(attr1 & 1), "encrypted": bool(attr1 & 2), "ver": ver}

def _iter_records(buf: bytes):
    off, n = 0, len(buf)
    while off + 4 <= n:
        (hdr,) = struct.unpack("<I", buf[off:off+4]); off += 4
        tag   =  hdr        & 0x3FF
        size  = (hdr >> 20) & 0xFFF
        if size == 0xFFF:
            if off + 4 > n: break
            (size,) = struct.unpack("<I", buf[off:off+4]); off += 4
        if off + size > n:
            payload = buf[off:n]; off = n
        else:
            payload = buf[off:off+size]; off += size
        yield tag, payload

def _zlib_if_needed(raw: bytes, expect: bool) -> bytes:
    if not expect: return raw
    try: return zlib.decompress(raw)
    except Exception: return raw

def extract_hwp_lines(path: Path) -> list[str]:
    if not _HAS_OLE:
        raise RuntimeError("olefile 모듈이 필요합니다. (pip install olefile)")
    out: list[str] = []
    with olefile.OleFileIO(str(path)) as ole:
        hdr = _read_fileheader(ole)
        if hdr["encrypted"]:
            raise RuntimeError("암호화된 HWP는 지원하지 않습니다.")
        entries = [e for e in ole.listdir(streams=True, storages=True)
                   if len(e)==2 and e[0]=="BodyText" and e[1].startswith("Section")]
        entries.sort(key=lambda e:int("".join(ch for ch in e[1] if ch.isdigit()) or "0"))
        for e in entries:
            with ole.openstream("/".join(e)) as fp:
                raw = fp.read()
            data = _zlib_if_needed(raw, hdr["compressed"])
            # 레코드 순회: 문단 텍스트(66)마다 줄 하나
            for tag, payload in _iter_records(data):
                if tag == HWPTAG_PARA_TEXT:
                    s = payload.decode("utf-16le", errors="ignore")
                    s = s.replace("\r","\n")
                    # 빈 줄 제거, 제어문자 제거
                    for ln in s.split("\n"):
                        ln = "".join(ch if (ch >= " " or ch in "\t") else " " for ch in ln).strip()
                        if ln:
                            out.append(ln)
    return out

# ---------------- HWPX (.hwpx) ----------------
def _lname(tag: str) -> str:
    """{ns}local → local"""
    if "}" in tag: return tag.rsplit("}",1)[1]
    return tag

# 문단/줄 경계로 취급할 요소 로컬명(소문자)
_PARA_END = {
    "p","para","paragraph","line","li","item","title","subtitle","caption",
}
_BR = {"br","linebreak"}
# 표 관련 요소에서 최소 줄바꿈: 행/셀 시작 전후로 잘라주면 세로 정렬에 유리
_TABLE_BREAK = {"tbl","table","tr","row","tc","cell","th","thead","tbody","tfoot"}

def extract_hwpx_lines(path: Path) -> list[str]:
    out: list[str] = []
    with zipfile.ZipFile(str(path)) as z:
        sections = sorted([n for n in z.namelist()
                           if n.startswith("Contents/section") and n.endswith(".xml")])
        if not sections:
            sections = sorted([n for n in z.namelist() if n.lower().endswith(".xml")])

        for name in sections:
            try:
                root = ET.fromstring(z.read(name))
            except Exception:
                continue

            buf: list[str] = []  # 현재 문단/셀 버퍼

            def flush():
                # 버퍼 → 한 줄
                nonlocal buf
                line = "".join(buf)
                # 연속 공백 정리
                line = re.sub(r"[ \t\u00A0]{2,}", " ", line)
                line = line.strip()
                if line:
                    out.append(line)
                buf = []

            # 깊이우선 순회
            stack = [root]
            while stack:
                node = stack.pop()
                # 노드의 텍스트
                if node.text:
                    buf.append(node.text)

                lname = _lname(node.tag).lower()
                if lname in _BR:
                    flush()
                # 표 요소에서 살짝 잘라 넣기 (셀/행/표 경계)
                if lname in _TABLE_BREAK:
                    flush()

                # 자식 push (역순으로 push하여 원래 순서대로 pop)
                children = list(node)
                for child in reversed(children):
                    stack.append(child)

                # 노드 종료 텍스트 tail
                if node.tail:
                    buf.append(node.tail)

                # 문단/목록 항목/제목 계열 끝나면 줄바꿈
                if lname in _PARA_END:
                    flush()

            # 섹션 끝나고 남은 버퍼
            flush()

    return out

# ---------------- 유틸 ----------------
def hard_wrap_lines(lines: list[str], width: int) -> list[str]:
    if width <= 0: return lines
    wrapped: list[str] = []
    for ln in lines:
        s = ln
        while len(s) > width:
            # 문장부호/공백 기준으로 자연스레 자르기
            cut = s.rfind(" ", 0, width)
            if cut < width * 0.6:
                # 공백이 없으면 문장부호 시도
                for p in (". ", ") ", "] ", "· ", "• ", ", "):
                    pos = s.rfind(p, 0, width)
                    if pos > 0: cut = pos + len(p)-1; break
            if cut <= 0: cut = width
            wrapped.append(s[:cut].rstrip())
            s = s[cut:].lstrip()
        if s: wrapped.append(s)
    return wrapped

# ---------------- 메인 ----------------
def extract_any_lines(path: Path) -> tuple[str, list[str]]:
    kind = detect_container(path)
    if kind == "HWP5": return "HWP5", extract_hwp_lines(path)
    if kind == "HWPX": return "HWPX", extract_hwpx_lines(path)
    raise ValueError("알 수 없는 형식입니다(.hwp/.hwpx)")

def main(argv=None):
    import argparse
    ap = argparse.ArgumentParser(description="HWP/HWPX → JSON(lines) (문단/표 경계 줄바꿈 반영)")
    ap.add_argument("input", help="입력 파일(.hwp/.hwpx)")
    ap.add_argument("--out", help="출력 JSON 경로(UTF-8)")
    ap.add_argument("--pretty", action="store_true", help="들여쓰기 출력")
    ap.add_argument("--hard-wrap", type=int, default=0, help="긴 줄 강제 줄나눔 폭(0=해제)")
    args = ap.parse_args(argv)

    p = Path(args.input)
    fmt, lines = extract_any_lines(p)

    if args.hard_wrap and args.hard_wrap > 0:
        lines = hard_wrap_lines(lines, args.hard_wrap)

    obj = {"file": str(p), "format": fmt, "lines": lines}
    text = json.dumps(obj, ensure_ascii=False, indent=2 if args.pretty else None)

    if args.out:
        Path(args.out).write_text(text, encoding="utf-8")
    else:
        sys.stdout.write(text)

if __name__ == "__main__":
    main()
