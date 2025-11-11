import json
import re
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P

# 문서 내의 모든 블록(단락, 표)을 순회하는 함수입니다.
def iterBlockItems(parent):
    # parent.element.body.iterchildren()의 각 child에 대해 반복합니다.
    for child in parent.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

# 텍스트를 문장 단위로 분리하는 함수입니다.
def splitSentences(text):
    # . 또는 \n 을 기준으로 문장들을 분리합니다.
    sentences = re.split(r"[.\n]", text)

    # 공백과 빈 문자열을 제거하고, 각 문장 뒤에 마침표를 다시 붙여서 반환합니다.
    return [s.strip() + "." for s in sentences if s.strip()]

# docx 파일을 파싱하고 내용을 정리하는 함수입니다.
def parseDocxClean(filePath):
    doc = Document(filePath)
    content = []

    # 문서의 모든 블록을 순회합니다.
    for block in iterBlockItems(doc):
        # 블록이 단락인 경우입니다.
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if text:
                content.extend(splitSentences(text))
        # 블록이 표인 경우입니다.
        elif isinstance(block, Table):
            # 표를 직렬화(serialization)합니다.
            # 각 행(row)을 순회합니다.
            for row in block.rows:
                # 각 셀(cell)의 텍스트를 가져와서 rowData 리스트에 저장합니다.
                rowData = [cell.text.strip() for cell in row.cells]
                # rowData에 내용이 있다면, " | "로 join하여 content 리스트에 추가합니다.
                if any(rowData):
                    content.append(" | ".join(rowData))

    return content

# 이 코드가 메인으로 실행될 때만 아래 내용을 실행합니다.
if __name__ == "__main__":
    filePath = "테스트.docx"
    parsedData = parseDocxClean(filePath)

    # 파싱된 데이터를 JSON 파일로 저장합니다.
    with open("parsed_result1.json", "w", encoding="utf-8") as fileObject:
        json.dump(parsedData, fileObject, ensure_ascii=False, indent=2)

    # 완료 메시지를 출력합니다.
    print("✅ 문장/줄 단위 JSON 생성 완료 → parsed_result1.json")
