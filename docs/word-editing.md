# Word 세부 편집 API

`WordDocument.edit_word_details()`는 Word 원본 XML 위에서 문단·표·그림·영역을 편집하는
트랜잭션입니다. 공통 모델을 새 Word 문서로 다시 작성하지 않으므로 다른 패키지 항목과
대상 외의 서식을 보존합니다. 기존 `edit_word()`는 python-docx/OOXML 직접 편집용으로
계속 사용할 수 있습니다.

```python
from jakal_hwpx.addons.docx import WordDocument

doc = WordDocument.blank()
with doc.edit_word_details() as edit:
    body = edit.body
    title = body.append_paragraph("검토 보고서", style="Title")
    title.format_text(font_size_pt=20, bold=True, color="000000")
    status = body.append_paragraph("상태: 초안")
    status.replace_text("초안", "검토 완료")
    status.format_text(0, 3, bold=True)

    table = body.insert_table(len(body.blocks), rows=2, cols=2,
                              values=[["항목", "값"], ["합계", "135"]])
    table.format(widths_pt=[240, 160], repeat_header=True)
    table.format_cell(0, 0, fill_color="E8EDF2", padding_pt={"left": 6})
    table.cell(0, 0).paragraph(0).format_text(bold=True)
    edit.header().paragraph(0).set_text("검토 보고서")

# 성공 시 공통 모델이 갱신됩니다. 편집 핸들은 이 시점부터 사용할 수 없습니다.
doc.save("edited.docx")
doc.save("edited.hwpx")
```

## 위치와 단위

- `body.blocks`는 문단·표가 문서에 배치된 순서입니다. 삽입·이동 인덱스는 이 목록 기준입니다.
- `paragraph(i)`, `table(i)`는 각각 문단 또는 표만 세어 선택합니다.
- `section(i)`는 지정한 섹션의 본문입니다. 기존 마지막 섹션 한정 추가 제한을 해제했습니다.
- 길이는 `*_pt`라는 이름의 **포인트 단위**입니다. 기존 공통 API의 HWP 단위와 구분합니다.
- `format_text(start, end)`는 0부터 시작하는 Python 문자열 범위이며 끝 위치는 제외합니다.
  링크의 글자를 포함한 편집 가능한 일반 텍스트를 기준으로 하며, 수식·그림·필드 명령은
  문자 수에 포함하지 않습니다. `paragraph.text`로 이 편집 대상 텍스트를 확인할 수 있습니다.
- `move(i)`는 이동 후 대상 목록에서의 위치입니다. 범위를 벗어나면 원본을 이동하지 않습니다.

## 문단과 글자 서식

| 작업 | API |
| --- | --- |
| 특정 위치 삽입 | `container.insert_paragraph(index, text, style=None)` |
| 끝에 추가 | `container.append_paragraph(text, style=None)` |
| 일반 텍스트 치환 | `paragraph.replace_text(old, new, count=-1)` |
| 일반 텍스트 전체 수정 | `paragraph.set_text(text)` |
| 문단 삭제·이동 | `paragraph.delete()`, `paragraph.move(index, destination=container)` |
| 문단 서식 | `paragraph.format(style=None, alignment="left", space_after_pt=6, ...)` |
| 글자 범위 서식 | `paragraph.format_text(start=0, end=None, bold=True, ...)` |
| 기존 이름 있는 스타일 수정 | `edit.format_style("Normal", font={...}, paragraph={...})` |

문단 속성은 `alignment`(`left`, `center`, `right`, `justify`, `distribute`),
`left_indent_pt`, `right_indent_pt`, `first_line_indent_pt`, `space_before_pt`,
`space_after_pt`, `line_spacing`(배수), `keep_with_next`, `keep_together`,
`page_break_before`, `widow_control`을 지원합니다. 상속 가능한 속성에 `None`을 주면
직접 지정을 해제합니다. `style=None`은 현재 스타일을 유지합니다.

글자 속성은 `font_name`, `font_size_pt`, `color`, `bold`, `italic`, `underline`,
`strike`, `superscript`, `subscript`입니다. 글꼴 크기는 양수이며 0.5pt 단위이고,
색상은 `224466` 또는 `#224466` 형태입니다. 밑줄은 bool 또는 Word 밑줄 이름을 소문자로
지정합니다. 예: `single`, `double`, `dotted_heavy`, `dash_long`, `wave`.
같은 범위에서 위첨자와 아래첨자를 동시에 켤 수 없습니다.

범위 서식은 필요한 run만 나누고 링크·책갈피·그림·수식과 다른 XML 자식은 유지합니다.
사용 중인 비목록 문단 스타일의 속성과 속성의 부재도 HWPX 왕복에 보존합니다.
따라서 다른 출력 템플릿의 제목 테두리·간격이 편집한 스타일에 다시 생기지 않습니다.
번호 스타일은 기존 목록 복원 경로에서 처리합니다.
필드가 있는 문단의 `set_text()`는 명령 유실을 막기 위해 거부합니다. 일반 글자의
`replace_text()`를 사용하거나 `edit_word()`에서 필드를 명시적으로 편집하세요.

## 표

```python
with doc.edit_word_details() as edit:
    table = edit.body.table(0)
    table.set_cell_text(1, 1, "150")
    table.add_row(["추가 항목", "20"], index=1)
    table.format(widths_pt=[240, 160], repeat_header=True)
    table.format_row(0, height_pt=24, exact=False)
    table.format_cell(0, 0, fill_color="E8EDF2", vertical_alignment="center",
                      padding_pt={"left": 6, "right": 6, "top": 4, "bottom": 4})
```

`delete_row(index)`, `merge(top, left, bottom, right)`, 표 자체의 `delete()`·`move()`도
지원합니다. `cell(row, col)`은 내부 문단·표 편집용 컨테이너를 반환합니다.
`set_cell_text()`는 기존 run 서식을 보존하며 여러 문단·중첩 표·필드가 있는 셀은
내부 문단을 직접 선택하도록 안내합니다.

병합 표의 행 추가·삭제와 마지막 행 삭제는 거부합니다. 병합 표의 열 너비는 실제
그리드에 맞춰 적용합니다. 열 추가·삭제나 복잡한 병합 재구성은 현재 `edit_word()`에서
직접 처리해야 합니다.

## 그림

```python
with doc.edit_word_details() as edit:
    picture = edit.body.pictures[0]
    picture.replace_image(image_bytes)
    picture.resize(width_pt=144)  # 한쪽 크기만 주면 현재 가로세로 비율 유지
    picture.crop(left=.1, right=.1)
    picture.rotate(15)
    picture.position(x_pt=12, y_pt=6, horizontal_from="margin",
                      vertical_from="paragraph", wrap="square")
```

그림 교체는 선택한 그림의 관계만 바꾸므로 같은 원본 이미지를 공유하는 다른 그림은
유지됩니다. `paragraph.add_picture(data, width_pt=72, height_pt=None, alt_text=None)`로
그림을 추가하고 `picture.delete()`로 삭제합니다. 자르기는 각 변에서 제거할 비율이며
가시 영역이 남아야 합니다. 배치 기준은 가로 `page`·`margin`·`column`, 세로
`page`·`margin`·`paragraph`·`line`이고 감싸기는 `square`, `top_and_bottom`,
`front`, `behind`를 지원합니다.

## 머리말·꼬리말·텍스트 상자·콘텐츠 컨트롤

`edit.header(section_index=0, variant="default", unlink=False)`와 `edit.footer(...)`는
해당 영역의 컨테이너를 반환합니다. `variant`는 `default`, `first`, `even`입니다.
이전 섹션의 영역을 상속하는 경우, 독립적으로 수정하려면 `unlink=True`를 명시해야 합니다.
`first`와 `even` 선택 시 필요한 Word 설정도 켭니다.

`edit.body.textbox(index)` 또는 `edit.header().textbox(index)`로 텍스트 상자 내부를
편집합니다. Word의 현대적 표현과 VML 대체 표현이 함께 있으면 변경 내용을 동기화하고
오래된 VML 보조 캐시를 제거합니다.

콘텐츠 컨트롤 안의 일반 텍스트를 바꾸면 연결된 custom XML 바인딩도 갱신합니다.
서식만 바꾸거나 아무것도 수정하지 않은 경우 연결 XML은 다시 쓰지 않습니다.
외부 컨테이너에서 컨트롤 내부 문단을 삭제·이동하거나 중간에 블록을 끼워 넣는 구조 변경은
거부하며, 컨트롤 앞에 삽입할 때는 새 문단을 컨트롤 밖에 둡니다.

## 트랜잭션과 남은 경계

편집 블록을 정상 종료하면 Word 백엔드와 공통 모델을 함께 갱신합니다. 예외가 밖으로
전파되면 편집 전 원본 바이트로 되돌립니다. 편집 핸들은 블록 종료 후 만료되며 중첩된
`edit_word()`/`edit_word_details()`는 거부합니다. 편집 중에는 공통 모델을 별도로
수정하지 말고 활성 편집기를 사용하세요. 저장은 편집 블록 밖에서 실행합니다.

다른 OPC 파트 사이의 이동, 섹션 구분 문단의 삭제·이동, 필드·책갈피·댓글 경계를
끊는 삭제·이동은 거부합니다. 같은 파트의 본문과 표 셀 사이 이동은 지원합니다.
그룹 도형, 고급 수식, IF·TOC의 한글 내 자동 계산 지원 범위는 이번 API와 별개입니다.

저장소 루트에서 실행하는 예제는 다음과 같습니다. `original.docx`, `edited.docx`, `edited.hwpx`,
`restored.docx`를 생성합니다.

```sh
python examples/docx_editing.py
python examples/docx_editing.py --image path/to/image.png --output-dir build/my-editing-demo
python -m pytest tests/test_word_editing.py
```

## 검증 범위

2026-11-02 기준 전용 편집 회귀 검사 16개와 전체 검사 441개가 통과했습니다.
선택적 로컬 문서 샘플이 필요한 검사 12개는 건너뛰었습니다. 실행 예제의 글자·문단·표·
머리말·꼬리말·그림을 편집한 뒤 DOCX→HWPX→DOCX로 왕복하고 Word에서 출력한
1페이지를 비교했습니다. 흰 배경을 제외한 영역의 변경률은 0%, 잉크 영역 IoU는 1.0이며
전체 픽셀도 일치했습니다. 이 결과는 해당 예제의 검증 결과이며 모든 Word 문서의
배치 동일성을 보증하지는 않습니다.
