#!/usr/bin/env python3



# -*- coding: utf-8 -*-



"""



HWPX → {"text": "<전체 본문 한 줄>"} (표는 | 규칙으로 인라인)



- 출력은 무조건 단일 키: text



- 표 인라인 규칙: 셀'|', 행'||', 표'|||'



- 기본: 완전 한 줄 JSON (ensure_ascii=False)







사용법:



  python hwpx_text_tables_only.py "아이번농장 투자 보고서.hwpx" "dkdlqjs.json"



  # 표준출력:



  python hwpx_to_text_only.py "입력.hwpx" --stdout > out.json







옵션:



  --preserve-paragraphs : 문단 사이에 \n 유지(여전히 {"text":"..."} 하나에 담김)



"""







import os



import sys



import json



import zipfile



import argparse



import xml.etree.ElementTree as ET



from typing import List







def _local(tag: str) -> str:



    return tag.split("}", 1)[1] if "}" in tag else tag







def _collect_text(elem: ET.Element) -> str:



    parts = []



    if elem.text:



        parts.append(elem.text)



    for ch in elem:



        parts.append(_collect_text(ch))



        if ch.tail:



            parts.append(ch.tail)



    # 줄 정돈



    txt = "".join(parts)



    txt = "\n".join(line.strip() for line in txt.splitlines() if line.strip())



    return txt







def _clean_inline(s: str) -> str:



    return " ".join(line.strip() for line in s.splitlines() if line.strip())







def _flatten_single_tbl(tbl_elem: ET.Element) -> str:



    """<tbl> → '|' 규칙 문자열(셀'|', 행'||')"""



    rows = []



    for tr in tbl_elem.iter():



        if _local(tr.tag).lower() in ("tr", "row"):



            cells = []



            for tc in tr.iter():



                if _local(tc.tag).lower() in ("tc", "cell"):



                    cells.append(_clean_inline(_collect_text(tc)))



            if cells:



                rows.append("|".join(cells))



    return "||".join(rows)







def _parse_hwpx_blocks(hwpx_path: str) -> List[str]:



    """문단(<p>)과 표(<tbl>)를 등장 순서대로 문자열 블록 리스트로"""



    if not zipfile.is_zipfile(hwpx_path):



        raise ValueError("유효한 HWPX(Zip) 파일이 아닙니다: " + hwpx_path)







    zf = zipfile.ZipFile(hwpx_path, "r")



    try:



        names = sorted([n for n in zf.namelist() if n.startswith("Contents/") and n.lower().endswith(".xml")])



        if not names:



            names = sorted([n for n in zf.namelist() if n.lower().endswith(".xml")])







        blocks: List[str] = []



        for name in names:



            try:



                root = ET.fromstring(zf.read(name))



            except Exception:



                continue



            for el in root.iter():



                tag = _local(el.tag).lower()



                if tag == "p":



                    t = _clean_inline(_collect_text(el))



                    if t:



                        blocks.append(t)



                elif tag == "tbl":



                    t = _flatten_single_tbl(el)



                    if t:



                        blocks.append(t)







        if blocks:



            return blocks







        # fallback: 문단/표 태그를 못 찾았을 때 전체 텍스트 긁기



        merged = []



        for name in names:



            try:



                root = ET.fromstring(zf.read(name))



            except Exception:



                continue



            merged.append(_collect_text(root))



        total = "\n".join(t for t in merged if t)



        return [line for line in total.split("\n") if line.strip()] if total.strip() else []



    finally:



        zf.close()







def export_text_only(hwpx_path: str, out_path: str = None, to_stdout: bool = False, preserve_paragraphs: bool = False):



    blocks = _parse_hwpx_blocks(hwpx_path)



    if preserve_paragraphs:



        text = "\n".join(blocks)



    else:



        # 표 블록은 이미 '|' 규칙으로 들어있으니 사이엔 공백만



        text = " ".join(blocks)







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



    ap.add_argument("--preserve-paragraphs", action="store_true",



                    help="문단 사이에 \\n 유지(여전히 text 단일 필드)")



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
