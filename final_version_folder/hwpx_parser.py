#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
HWPX → {"text": "전체 본문..."}
수정사항:
1. 문단/표 블록 간 연결을 공백이 아닌 개행(\n)으로 변경 (청킹 호환성)
2. root.iter() 대신 재귀 탐색을 사용하여 표 내부 문단의 중복 추출 방지
3. 표의 행(row) 구분도 개행(\n)으로 처리하여 가독성 확보
"""

import argparse
import json
import os
import sys
import xml.etree.ElementTree as ET
import zipfile
from typing import List


def _local(tag: str) -> str:
    return tag.split("}", 1)[1] if "}" in tag else tag


def _collect_text(elem: ET.Element) -> str:
    """요소 내부의 모든 텍스트를 재귀적으로 수집"""
    parts = []
    if elem.text:
        parts.append(elem.text)
    for ch in elem:
        parts.append(_collect_text(ch))
        if ch.tail:
            parts.append(ch.tail)
    txt = "".join(parts)
    # 텍스트 내의 과도한 공백/개행 정리
    return " ".join(line.strip() for line in txt.splitlines() if line.strip())


def _flatten_single_tbl(tbl_elem: ET.Element) -> str:
    """<tbl> → 행 단위로 개행, 셀 단위는 '|'로 구분"""
    rows = []
    # 재귀적으로 tr(row) 탐색
    for tr in tbl_elem.iter():
        if _local(tr.tag).lower() in ("tr", "row"):
            cells = []
            for tc in tr.iter():
                if _local(tc.tag).lower() in ("tc", "cell"):
                    cell_text = _collect_text(tc)
                    cells.append(cell_text)
            if cells:
                rows.append("|".join(cells))
    # 행과 행 사이는 줄바꿈으로 연결 (기존 || 방식보다 청킹에 유리)
    return "\n".join(rows)


def _parse_element_recursive(elem: ET.Element) -> List[str]:
    """
    요소를 재귀적으로 순회하며 p(문단)와 tbl(표)을 추출.
    *중요*: tbl을 만나면 내부를 평탄화하고, 그 내부의 p는 별도로 추출하지 않음(중복 방지).
    """
    blocks = []
    tag = _local(elem.tag).lower()

    if tag == "tbl":
        # 표를 만나면 평탄화해서 추가하고, 자식 요소로 더 깊이 들어가지 않음
        t = _flatten_single_tbl(elem)
        if t:
            blocks.append(t)
        return blocks
    
    elif tag == "p":
        # 문단을 만나면 텍스트 추출
        t = _collect_text(elem)
        if t:
            blocks.append(t)
        # 문단 안에 또다른 문단이나 표가 있는 경우는 드물지만, 
        # HWPX 구조상 p는 텍스트 컨테이너이므로 여기서 멈춤 (혹은 run 탐색)
        return blocks

    # 그 외(section, box 등) 컨테이너면 자식들을 계속 탐색
    for child in elem:
        blocks.extend(_parse_element_recursive(child))
    
    return blocks


def _parse_hwpx_blocks(hwpx_path: str) -> List[str]:
    if not zipfile.is_zipfile(hwpx_path):
        raise ValueError("유효한 HWPX(Zip) 파일이 아닙니다: " + hwpx_path)

    zf = zipfile.ZipFile(hwpx_path, "r")
    try:
        # Contents/ 경로의 xml 파일들을 우선 탐색 (본문)
        names = sorted(
            [n for n in zf.namelist() if n.startswith("Contents/") and n.lower().endswith(".xml")]
        )
        # 만약 Contents가 없으면 루트의 xml 탐색 (BodyText 등)
        if not names:
            names = sorted([n for n in zf.namelist() if n.lower().endswith(".xml")])

        blocks: List[str] = []
        for name in names:
            try:
                xml_data = zf.read(name)
                root = ET.fromstring(xml_data)
                # 재귀 파싱 호출
                blocks.extend(_parse_element_recursive(root))
            except Exception:
                continue

        return blocks

    finally:
        zf.close()


def parse_hwpx_to_text(path: str) -> str:
    """
    외부 호출용 함수
    블록들을 개행문자(\n)로 연결하여 반환해야 청킹 스크립트가 줄 단위로 인식함.
    """
    blocks = _parse_hwpx_blocks(path)
    return "\n".join(blocks)


def export_text_only(
    hwpx_path: str,
    out_path: str | None = None,
    to_stdout: bool = False,
    preserve_paragraphs: bool = False,
) -> None:
    blocks = _parse_hwpx_blocks(hwpx_path)

    # preserve_paragraphs 옵션과 무관하게 기본적으로 \n으로 합치는 것이 안전함
    # 단, JSON 출력 시 형식을 맞추기 위해 로직 유지
    text = "\n".join(blocks)

    oneline = json.dumps({"text": text}, ensure_ascii=False, separators=(",", ":"))

    if to_stdout or not out_path:
        print(oneline, end="")
    else:
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(oneline)


def _parse_args():
    ap = argparse.ArgumentParser(description="HWPX → text only JSON")
    ap.add_argument("input", help="입력 .hwpx 경로")
    ap.add_argument("output", nargs="?", help="출력 .json 경로 (미지정 시 --stdout 필요)")
    ap.add_argument("--stdout", action="store_true", help="파일 대신 표준출력으로 한 줄 JSON 출력")
    ap.add_argument(
        "--preserve-paragraphs",
        action="store_true",
        help="문단 사이에 \\n 유지(여전히 text 단일 필드)",
    )
    return ap.parse_args()


if __name__ == "__main__":
    args = _parse_args()

    if not os.path.exists(args.input):
        sys.exit(f"입력 파일이 없습니다: {args.input}")

    if not args.stdout and not args.output:
        sys.exit("출력 파일을 지정하거나 --stdout를 사용하세요.")

    export_text_only(
        args.input,
        args.output,
        to_stdout=args.stdout,
        preserve_paragraphs=args.preserve_paragraphs,
    )