"""Small, original document starters; no desktop automation or third-party assets."""
from jakal_hwpx import HwpxDocument
from jakal_hwpx.namespaces import NS

TEMPLATES = [
    {"id": "blank", "name": "빈 문서", "description": "A4 · 여백 20mm · 자유롭게 작성", "title": "새 문서"},
    {"id": "report", "name": "보고서", "description": "개요 · 본문 · 결론 · 참고 자료", "title": "보고서"},
    {"id": "meeting", "name": "회의록", "description": "일시 · 참석자 · 논의 · 후속 작업", "title": "회의록"},
    {"id": "table", "name": "표 문서", "description": "제목 · 기본 표 · 메모", "title": "정리 문서"},
]


def make_template(kind: str, title: str, question_count: int = 3) -> bytes:
    # Retain earlier API starters for compatibility, without promoting them in UI.
    if kind not in {item["id"] for item in TEMPLATES} | {"exam", "worksheet"} or not 1 <= question_count <= 20:
        raise ValueError("지원하지 않는 양식입니다.")
    doc = HwpxDocument.blank()
    doc.section_settings().set_margins(left=5669, right=5669, top=5669, bottom=5669)
    body = doc.append_paragraph_style(line_spacing=165).element.get("id")
    centered = doc.append_paragraph_style(alignment_horizontal="CENTER", line_spacing=165).element.get("id")
    large = doc.append_character_style(height=2000).element.get("id")
    heading = doc.append_character_style(height=1300).element.get("id")
    first = doc.sections[0].root_element.xpath("./hp:p", namespaces=NS)[0]
    first.set("paraPrIDRef", centered if kind != "blank" else body)
    if kind != "blank":
        doc.set_paragraph_text(0, 0, title)
        for run in first.xpath("./hp:run[hp:t]", namespaces=NS):
            run.set("charPrIDRef", large)
    def text(value="", *, label=False):
        doc.append_paragraph(value, para_pr_id=body, char_pr_id=heading if label else "0")
    if kind == "report":
        text("작성자:                         작성일:")
        for label, content in [("1. 개요", "문서의 목적과 배경을 작성하세요."), ("2. 본문", "주요 내용과 근거를 정리하세요."), ("3. 결론", "결론과 후속 작업을 작성하세요."), ("4. 참고 자료", "자료의 제목과 출처를 기록하세요.")]:
            text(); text(label, label=True); text(content); text()
    elif kind == "meeting":
        text("일시:                         장소:"); text("참석자:")
        for label in ["안건", "논의 내용", "결정 사항", "후속 작업"]:
            text(); text(label, label=True); text(); text()
    elif kind == "table":
        text(); doc.append_table(rows=4, columns=3, cell_width=16000, row_height=2600, cell_texts=[["항목", "내용", "비고"], ["", "", ""], ["", "", ""], ["", "", ""]])
        text(); text("메모", label=True); text()
    elif kind == "exam":
        text("학년:            반:            번호:            이름:")
        text("각 문항을 읽고 가장 알맞은 답을 고르세요.")
        for n in range(1, question_count + 1):
            text(); text(f"{n}. 문항 내용을 입력하세요. [배점]", label=True)
            text(); text("① 선지 1    ② 선지 2    ③ 선지 3    ④ 선지 4    ⑤ 선지 5"); text()
    elif kind == "worksheet":
        text("이름:                         날짜:")
        text(); text("학습 목표", label=True); text("이번 활동에서 확인할 내용을 입력하세요.")
        text(); text("관찰 및 기록", label=True)
        doc.append_table(rows=4, columns=3, cell_width=16000, row_height=2600, cell_texts=[["구분", "관찰 내용", "해석"], ["1", "", ""], ["2", "", ""], ["3", "", ""]])
        text(); text("활동 정리", label=True); text("관찰한 사실과 생각을 정리하세요."); text()
    raw = doc.compile()
    reopened = HwpxDocument.from_bytes(raw)
    if reopened.validation_errors() or reopened.xml_validation_errors() or reopened.reference_validation_errors():
        raise ValueError("양식 문서 검증에 실패했습니다.")
    return raw
