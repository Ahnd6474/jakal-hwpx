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

추가 글자 속성은 `highlight`(예: `yellow`, `cyan`, `none`),
`character_spacing_pt`(음수 가능), `scale_percent`(정수 1~600),
`font_names`와 `languages`입니다. `font_names={"eastAsia": "맑은 고딕", "ascii": "Arial"}`처럼
`ascii`·`hAnsi`·`eastAsia`·`cs` 채널을 따로 지정할 수 있습니다. 언어는
`languages={"latin": "en-US", "eastAsia": "ko-KR", "bidi": "ar-SA"}` 형식입니다.
장평은 DOCX에서 1~600%를 허용하지만 HWPX의 네이티브 매핑은 50~200% 범위이며,
범위 밖에서는 경고하고 원래 Word 속성을 왕복 보존합니다.

고정값 줄 간격은 `paragraph.format(line_spacing_pt=18, line_spacing_rule="exact")`로
지정합니다. `at_least`도 지원하며 배수 `line_spacing`과 동시에 줄 수 없습니다.
`tab_stops=[{"position_pt": 120, "alignment": "right", "leader": "dot"}]`로 직접 탭 정의를
교체합니다. 정렬은 `left`·`center`·`right`·`decimal`·`bar`·`clear`, 채움선은
`none`·`dot`·`hyphen`·`underscore`·`heavy`·`middleDot`입니다. `None`은 직접 탭 정의를
제거하고 상속합니다. 언어·탭·세부 글자 속성도 기존 스타일에 적용할 수 있습니다.

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

병합 표에서도 행·열을 추가·삭제할 수 있습니다. `add_column(values=None, index=None,
width_pt=72)`, `delete_column(index)`로 열을 수정하고 `split_cell(row, col)`로 기존 병합을
원래 그리드의 개별 셀로 풉니다. 분할 시 기존 본문은 왼쪽 위 셀에 둡니다.
새 그리드를 만드는 임의 행×열 분할은 포함하지 않습니다.

병합 내부에 행·열을 삽입하면 병합 범위를 확장합니다. `values`는 그리드 위치 기준이며
이미 병합된 셀의 연속 위치에는 `None` 또는 빈 문자열을 사용해야 합니다. 병합 시작 행을
삭제하면 기존 내용과 서식을 첫 잔존 행으로 옮깁니다. 마지막 행·열 삭제, 생략된 가장자리
셀이 있는 불완전 그리드, 레거시 `hMerge` 구조는 거부합니다. 중첩 표의 병합은 바깥 표의
행·열 편집을 막지 않습니다.

셀 테두리는 `format_cell(row, col, borders={"bottom": {"style": "double",
"width_pt": 1.5, "color": "224466"}})`로 지정합니다. 각 변의 `None`은 직접 정의를
제거하고 `{"style": "nil"}`은 선을 숨깁니다. 병합 셀의 실제 연속 행에도 테두리를 적용하고
왕복 복원하므로 중간 행에서 세로 테두리가 끊기지 않습니다.

## 페이지·섹션·다단

`edit.section(i)`는 기존 본문 컨테이너 메서드와 다음 설정을 함께 제공합니다.

```python
with doc.edit_word_details() as edit:
    edit.section(0).format_page(
        width_pt=595, height_pt=842, orientation="portrait",
        margins_pt={"left": 48, "right": 48, "top": 45, "bottom": 45},
        header_distance_pt=20, footer_distance_pt=20, gutter_pt=0,
    )
    # 본문 blocks 위치 앞에서 나눕니다. 반환값은 나눈 뒤의 섹션입니다.
    next_section = edit.insert_section(2, break_type="next_page")
    next_section.format_page(orientation="landscape")
    next_section.format_columns(count=2, gap_pt=18, separator=True)
```

방향을 바꾸면 용지 너비·높이도 맞춰 교환합니다. 여백은 용지 가장자리에서 측정하며,
본문 영역이 남지 않는 설정은 거부합니다. 섹션 나누기 유형은 `continuous`·`next_page`·
`even_page`·`odd_page`·`next_column`입니다. 표나 콘텐츠 컨트롤 내부는 나눌 수 없습니다.
균등 다단과 `widths_pt=[...]`를 주는 비균등 다단을 지원하며, 너비 개수와 용지 영역을
검증합니다. 다단은 공통 `SectionSettings.columns`와 HWPX `colPr`에 매핑합니다.
Word 섹션 시작 유형은 왕복 보존합니다. 한글의 구역 배치 방식까지 Word의 연속 구역과
동일하게 계산하는 것은 보장하지 않습니다.
다단의 순수 Python HWP 바이너리 출력은 지원하지 않아 오류를 냅니다. HWPX로 저장하거나
WordDocument에 명시적인 외부 변환기를 지정해 HWP로 변환하세요.

## 자동 목록·필드

```python
with doc.edit_word_details() as edit:
    listing = edit.create_list([
        {"format": "decimal", "text": "%1.", "left_indent_pt": 36, "hanging_pt": 18},
        {"format": "lowerLetter", "text": "%1.%2)"},
    ])
    first = edit.body.paragraph(0).set_list(listing, start=3)
    edit.body.paragraph(1).set_list(first.listing)  # 같은 번호 인스턴스 이어 매기기
    edit.body.paragraph(2).set_list(listing, level=1)
    edit.body.paragraph(3).remove_list()           # 스타일에서 상속한 번호도 억제

    field = edit.footer().paragraph(0).add_field("PAGE", "1")
    field.set_instruction("NUMPAGES").set_result("3").lock(False)
    print(field.instruction, field.result)
```

목록은 0~8 수준을 사용하며 `decimal`·`decimalZero`·`lowerRoman`·`upperRoman`·
`lowerLetter`·`upperLetter`·`bullet`을 지원합니다. 각 수준은 `format`, `text`, `start`,
`left_indent_pt`, `hanging_pt`, `font_name`, `suffix`를 받습니다. `listing.restart(5, level=0)`는
기존 목록과 다른 번호 인스턴스를 만들어 다른 문단의 시작값을 바꾸지 않습니다.
`listing` 핸들도 편집 트랜잭션 안에서만 유효합니다.

컨테이너 또는 문단의 `fields`는 단순 필드와 완결된 복합 필드를 문서 순서로 반환합니다.
다른 텍스트 상자·각주 등 별도 이야기 영역의 필드는 섞지 않습니다. 명령 수정은 `dirty`
표시를 켜며, 표시값 수정은 계산 결과를 직접 기록합니다. 필드 계산 엔진을 실행하지는
않습니다. 중첩·그림·수식·표·여러 문단이 있는 결과를 `set_result()`로 평탄화하는 작업은
거부하므로 내부 객체를 직접 수정해야 합니다. 중첩된 명령도 내부 필드를 선택해 수정합니다.
필드만 있는 머리말·꼬리말도 변환 대상에서 누락하지 않습니다.

## 각주·미주

```python
with doc.edit_word_details() as edit:
    note = edit.body.paragraph(0).add_note("각주 본문", kind="footnote")
    note.paragraph(0).format_text(font_size_pt=9)
    note.format_reference(font_size_pt=8, superscript=True)
    note.move_reference(edit.body.paragraph(1))
    edit.footnote(0).append_paragraph("추가 설명")
    # edit.endnote(0) / add_note(..., kind="endnote")
    # note.delete()  # 본문 참조와 정의를 함께 삭제
```

각주·미주는 별도 Word 파트의 실제 객체이며, 문서 순서가 아닌 해당 파트의 일반 정의
목록 순서로 선택합니다. 구분선 정의는 인덱스에서 제외합니다. 내부 문단·표·그림은 같은
컨테이너 API를 사용하며, 참조는 같은 문서 본문 안에서 이동합니다. 원본의 불투명 각주
파트는 실제 수정 시에만 다시 직렬화하므로 조회만 한 경우 원본 바이트를 유지합니다.

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
상속을 끊을 때 현재 내용과 그림·링크 관계를 새 파트에 복사합니다.
`first`와 `even` 선택 시 필요한 Word 설정도 켭니다.

`edit.body.textbox(index)` 또는 `edit.header().textbox(index)`로 텍스트 상자 내부를
편집합니다. Word의 현대적 표현과 VML 대체 표현이 함께 있으면 변경 내용을 동기화하고
오래된 VML 보조 캐시를 제거합니다.

콘텐츠 컨트롤 안의 일반 텍스트를 바꾸면 연결된 custom XML 바인딩도 갱신합니다.
서식만 바꾸거나 아무것도 수정하지 않은 경우 연결 XML은 다시 쓰지 않습니다.
외부 컨테이너에서 컨트롤 내부 문단을 삭제·이동하거나 중간에 블록을 끼워 넣는 구조 변경은
거부하며, 컨트롤 앞에 삽입할 때는 새 문단을 컨트롤 밖에 둡니다.

## 범위·링크·책갈피

```python
with doc.edit_word_details() as edit:
    p = edit.body.paragraph(0)
    p.range(2, 6).replace("새 글")
    edit.body.range(0, 3, 1, 5).replace("문단 사이\n치환")
    matches = edit.body.find_text("검색어", case_sensitive=False)
    edit.body.replace_all("초안", "검토 완료")
    p.range(0, 2).copy_to(edit.header().paragraph(0), 0)
    p.range(0, 2).move_to(edit.body.paragraph(2), 3)
    link = p.insert_hyperlink(0, "https://example.com", "외부 링크 ")
    link.set_target("https://example.com/new").set_text("새 링크 ")
    mark = p.range(0, 3).add_bookmark("reference")
    mark.rename("new_reference")  # 내부 링크와 REF/PAGEREF 명령도 함께 변경
```

범위의 문단 번호와 문자 인덱스는 0부터 시작하며 끝 문자는 제외합니다. 문단 사이는
검색에서 `\n`입니다. `find_text(..., regex=True)`는 정규식 검색을 지원하며 0폭 결과는
거부합니다. `replace_all()`의 검색은 리터럴입니다. 수식·그림·필드 결과는 문자 인덱스에
포함하지 않습니다. 본문과 표 셀은 별도로 검색하며 표·섹션 경계를 가로지르는 치환은
거부합니다. 오래된 텍스트 범위를 다시 쓰려 하면 오류를 냅니다.

복사할 때 글자 서식·링크·그림·수식을 유지하고 파트가 달라지면 관계 ID를 새 파트에
연결합니다. 선택 범위의 책갈피는 새 이름과 번호로 복사합니다. 같은 문단 안의 복사와
범위 밖으로의 이동도 지원합니다. 여러 문단을 선택한 경우 목적지는 원본 문단 밖이어야
합니다. 각주 참조 복사는 새 각주 정의가 필요하므로 거부하며, 이동은 같은 본문 안에서
할 수 있습니다. 필드·책갈피 경계를 잘못 끊는 선택이나 콘텐츠 컨트롤 내부 분할은 거부합니다.

`container.hyperlinks`, `edit.bookmarks`로 기존 대상을 얻습니다. 링크의 `unlink()`는
표시 내용을 유지하고 링크만 해제합니다. 책갈피 `delete()`도 본문을 유지하고 표식만
삭제합니다. 책갈피 이름 변경은 본문·머리말·꼬리말·각주·미주의 내부 링크와 필드를 찾습니다.

## 스타일·테마와 표 배치

`edit.create_style(name, kind="paragraph", base=None, next_style=None, font=None,
paragraph=None)`로 스타일을 만듭니다. 종류는 `paragraph`·`character`·`table`이며,
`edit.configure_style(name, base=..., next_style=...)`로 관계를 수정합니다. 상속 순환과
서로 다른 종류의 상속을 거부합니다. `edit.delete_style(name, replacement="Normal")`는
본문·관련 스타일·번호 정의의 참조를 대체한 뒤 삭제합니다. 기본 스타일은 삭제할 수 없습니다.

`edit.format_theme(colors={"accent1": "224466"}, major_fonts={"latin": "Arial"},
minor_fonts={"eastAsia": "맑은 고딕"})`로 테마 색과 글꼴을 수정합니다. 색 슬롯은
`dark1`·`light1`·`dark2`·`light2`·`accent1`~`accent6`·`hyperlink`·`followed_hyperlink`,
글꼴 채널은 `latin`·`eastAsia`·`complexScript`입니다.

`table.format_layout(wrap="around", x_pt=12, y_pt=6, horizontal_from="margin",
vertical_from="text", distance_pt={"left": 4})`로 떠 있는 표와 본문 감싸기를 설정합니다.
기준은 `page`·`margin`·`text`, 감싸기는 `inline`·`around`입니다. `alignment`,
`indent_pt`, `autofit`도 지원합니다. 떠 있는 표의 기준·거리·오프셋은 HWPX에 매핑합니다.
`format_row(index, allow_split=False)`는 Word의 페이지 간 행 분할을 막습니다. HWPX는
모든 행이 분할 금지인 경우 셀 단위 분할로 매핑하고, 개별 행 속성은 Word 왕복에 보존합니다.

`following = table.split(row_index)`는 빈 문단으로 구분한 새 표를 반환합니다.
수직 병합을 자르는 분할은 거부하므로 먼저 `split_cell()`을 실행해야 합니다. 떠 있는 표도
먼저 `format_layout(wrap="inline")`로 바꿔야 합니다. `table.join(following)`은 사이에
내용이 없고 열 그리드·너비·표 속성이 같은 두 표를 합칩니다. `adopt_style=True`를
명시하면 첫 표의 속성을 사용하며, 뒤 표의 반복 머리글 표시를 제거합니다.

## 구조화 수식

```python
with doc.edit_word_details() as edit:
    equation = edit.body.append_paragraph().add_equation("matrix {a & b # c & d}")
    equation.set_script("sum _ {i=1} ^ {n} {{x} _ {i}}")
    print(equation.script)
    # equation.delete()  # 수식 삭제
```

분수·근호·위/아래첨자에 더해 직사각 행렬, 합·곱·적분, 단일 인수 괄호와 `lim`을
OMML과 한글 수식 스크립트 사이에서 변환합니다. 행렬의 열은 `&`, 행은 `#`로 구분합니다.
연산자는 `sum`·`prod`·`coprod`·`int`·`oint`·`iint`·`iiint`·`bigcup`·`bigcap`이며
`_ {하한}`·`^ {상한}` 또는 `from {하한}`·`to {상한}` 뒤에 `{본문}`을 둡니다.
예: `int _ {0} ^ {1} {{x} ^ {2} dx}`, `left [ {matrix {1 & 2 # 3 & 4}} right ]`.

추가로 `cases {x & x>0 # -x & x<0}`, `pile {a # b}`, `lpile {a # b}`,
`rpile {a # b}`, `eqalign {x & =1 # y & =2}`, `{n} choose {k}`,
`binom {n} {k}`, `{a} atop {b}`를 지원합니다. `pmatrix`·`bmatrix`·`dmatrix`는
행렬을 괄호로 감쌉니다. 열별 정렬은 `matrix {lcol {a # b} rcol {c # d}}`처럼
`col`·`lcol`·`rcol`로 지정합니다. 한글의 공식
[수식 명령어](https://help.hancom.com/hoffice130_assistant/ko-KR/Hwp/insert/equation/equation(explanation).htm)를 사용합니다.
왼쪽·오른쪽 쌓기는 열 정렬이 있는 단일 열 OMML 행렬로 작성합니다. OMML의 `baseJc`는
수직 기준이므로 가로 정렬 값을 넣지 않습니다.

원래 Word 수식의 속성만 별도 기록하므로 피연산자 사본을 저장하지 않습니다. 구조가
같으면 제한 위치·글자 속성과 독립 수식 문단을 복원하고, 구조가 바뀌면 오래된 속성을
다른 노드에 붙이지 않습니다. 일반 `limLow`·`limUpp`는 한글에서 세로 쌓기로 표시하고,
여러 인수 구분 괄호는 구분 문자를 사이에 둔 행렬로 표시합니다. 구조가 유지된 채
내용이 편집되면 현재 피연산자로 원래 OMML 구조를 복원합니다. 대체 표현의 구조를
바꾸면 원래 구조 복원을 강제하지 않습니다. 색·굵기 명령, 아직 매핑하지 않은 장식·
도형 수식과 비스듬한 분수 등은 명시적으로 오류를 냅니다.

## 복합 필드 결과와 콘텐츠 컨트롤

```python
with doc.edit_word_details() as edit:
    field = edit.body.append_paragraph().add_field(
        'IF 1 = 1 "Result" "Other"', "Result", complex_field=True,
    )
    with field.edit_result() as result:
        result.paragraph(0).format_text(bold=True)
        result.append_paragraph("두 번째 결과 문단")
        result.insert_table(1, rows=1, cols=2, values=[["항목", "값"]])

    control = edit.body.insert_control(
        len(edit.body.blocks), text="본문", tag="article", title="Article",
    )
    control.append_paragraph("컨트롤 내부에 삽입")
    control.paragraph(0).set_text("수정한 본문")
    control.configure(tag="updated", title="Updated article")
    control.move(0, destination=edit.header())
    # control.unwrap()  # 내용은 유지하고 컨트롤만 해제
    # control.delete()  # 컨트롤과 내용을 삭제
```

`field.edit_result()`는 분리된 결과 범위를 편집한 뒤 정상 종료 시 반영합니다. 바깥의
명령·시작/끝 표식·주변 본문을 유지하고 내부 문단·표·링크·중첩 필드를 편집합니다.
`field.result_fields`로 내부 필드를 찾을 수도 있습니다. 편집 중 원래 결과가 바뀌면
오래된 복사본 반영을 거부하며, 반환한 결과 컨테이너는 블록 종료 후 만료됩니다.
simple field의 결과는 한 문단의 인라인 내용으로 제한합니다. 여러 블록이 필요하면
처음부터 `complex_field=True`로 생성합니다. 복합 필드 경계는 형제 문단에 있어야 하며
섹션 경계를 재구성하는 편집은 거부합니다.
Word는 열기·재계산 시 필드 명령으로 결과를 다시 만들 수 있습니다. 직접 작성한 복합
결과를 유지하려면 `field.lock()` 또는 생성 시 `locked=True`를 사용하고, Word 자동 계산을
다시 사용할 때 `lock(False)`로 풉니다. 결과 편집 자체는 잠금 상태를 자동 변경하지 않습니다.

`container.wrap_control(start, end, tag=..., title=...)`은 블록 인덱스의 반열린 범위를
감쌉니다. `end`를 생략하면 한 블록을 감쌉니다. 컨트롤 내부 구조는 해당 핸들에서
편집하며, 바깥 컨테이너에서 내부 블록 사이에 삽입하는 것은 계속 거부합니다.
`container.content_controls`로 찾고, 단일 문단 또는 인라인 컨트롤은 `set_text()`도
사용할 수 있습니다. 생성 종류는 기본 `rich_text`, 선택적으로 `plain_text`입니다.
기존 컨트롤 속성과 데이터 바인딩은 유지하며 바뀐 내용으로 custom XML을 갱신합니다.
컨트롤 UUID도 HWPX 왕복에 보존합니다. 드롭다운·날짜 선택 UI를 새로 생성하는 기능은
포함하지 않습니다.

## 명시적 필드 갱신과 목차 재생성

```python
report = doc.update_fields(data={"status": "approved"}, now="2026-11-04T13:14:15")
for row in report:
    print(row["instruction"], row["status"], row.get("reason"))

with doc.edit_word_details() as edit:
    toc = next(f for f in edit.body.fields if f.instruction.lstrip().startswith("TOC"))
    print(toc.rebuild_toc())  # entries, requires_layout
```

`edit.update_fields()`도 같은 갱신을 실행합니다. 숫자 비교·대소문자를 구분하는 문자열
동등 비교·지원 범위의 `*`/`?` 패턴을 사용하는 리터럴/중첩 `IF`, 데이터가 주어진
`MERGEFIELD`, 기본 문서 속성, 지원 날짜·시간 서식, 기본 `REF`를 갱신합니다.
상태는 `updated`·`unchanged`·`locked`·`unsupported`이며 잠긴 필드는 유지합니다.
`strict=True`는 미지원 항목이 있으면 결과를 쓰기 전에 오류를 냅니다. 날짜는 명시적
서식만 지원하며 로캘별 기본 서식이나 문자열 대소 비교를 추측하지 않습니다.
구조화된 결과를 새 문자열로 덮어쓰는 자동 갱신은 거부하고 `edit_result()`를 사용합니다.

`TOC.rebuild_toc()`는 본문 제목 수준을 읽어 현재 제목·들여쓰기·이동 링크를 다시 만듭니다.
지원 스위치는 `\o "1-3"`·`\h`·`\z`·`\u`·인수 없는 `\n`입니다. 기존 명령은 유지하지만
쪽 번호는 생성하지 않습니다. 쪽 번호가 필요한 명령이면 `requires_layout=True`로
보고하고 Word 재계산 대상으로 표시합니다. `include_page_numbers=True`는 조판 엔진이
필요하다는 오류를 냅니다. 이 API는 한글 앱 안의 IF·TOC 자동 계산기를 설치하지 않습니다.

## 한영 문단 배치 설정

`paragraph.format(word_wrap=False, auto_space_east_asian_latin=True,
auto_space_east_asian_numbers=False, snap_to_grid=False, suppress_line_numbers=True)`로
줄바꿈·한영/한숫자 간격·격자·줄 번호 제외를 편집합니다. 값은 bool이며 `None`은
상속으로 되돌립니다. 스타일 상속을 해석해 HWPX의 `breakLatinWord`, `autoSpacing`,
격자·줄 번호 속성에 연결하고 역변환에도 반영합니다. 이 설정은 조판 차이의 일부를
줄이지만 글꼴 대체·표 자동 높이·두 앱의 모든 줄바꿈 동일성을 보장하지 않습니다.
단독 수식 문단에는 Word의 수식 정렬을 적용하고 인라인 수식이 포함된 본문의 정렬은
유지합니다. 빈 머리말의 숨김 ID 앵커와 run 유무, 변경하지 않은 스타일의 상속도 보존합니다.
문단 스타일의 원본 글자 속성도 보존하며, HWPX 기본 스타일이 같은 이름의 가져온
Word 스타일을 덮어쓰지 않게 처리합니다. 이 수정으로 확장 예제의 0.72pt 세로 차이를
해소했고 동일한 Word 렌더링에서 좌표·전체 픽셀 일치를 확인했습니다.

## 트랜잭션과 객체 ID

편집 블록을 정상 종료하면 Word 백엔드와 공통 모델을 함께 갱신합니다. 예외가 밖으로
전파되면 편집 전 원본 바이트로 되돌립니다. 편집 핸들은 블록 종료 후 만료되며 중첩된
`edit_word()`/`edit_word_details()`는 거부합니다. 편집 중에는 공통 모델을 별도로
수정하지 말고 활성 편집기를 사용하세요. 저장은 편집 블록 밖에서 실행합니다.

많은 문단을 공통 API로 추가할 때는 `doc.append_paragraphs(texts, section_index=0)`를
사용합니다. 모든 삽입을 묶고 공통 모델은 한 번만 다시 읽습니다. 변경 없는 편집 블록은
모델을 다시 읽지 않으며, 여러 세부 편집도 한 `edit_word_details()` 블록 안에서 묶습니다.

본문·머리말·꼬리말 등 다른 OPC 파트로 블록을 이동할 때 링크·그림 관계를 다시 연결하고
대상 파트에서 책갈피 숫자 ID가 겹치지 않게 조정합니다. 객체 UUID는 유지합니다.
다른 활성 편집기 사이 이동, 각주·미주 참조가 포함된 파트 간 이동, 섹션 구분 문단의
삭제·이동, 필드·책갈피·댓글 경계를 끊는 삭제·이동은 거부합니다.
그룹 도형과 IF·TOC의 한글 내 자동 계산은 별도 지원 범위입니다.

`WordEndpoint` 또는 `doc.endpoint()`를 사용하면 객체 UUID와 버전 기반 일괄 수정을
사용할 수 있습니다. Python·HTTP 실행, 충돌 응답, 숨김 앵커를 통한 ID 보존과 남은
경계는 [Word endpoint](./word-endpoint.md)에 설명합니다.

저장소 루트에서 실행하는 예제는 다음과 같습니다. `original.docx`, `edited.docx`, `edited.hwpx`,
`restored.docx`를 생성합니다.

```sh
python examples/docx_editing.py
python examples/docx_editing.py --image path/to/image.png --output-dir build/my-editing-demo
python examples/docx_advanced_editing.py --image path/to/image.png
python examples/docx_module_extensions.py
python -m pytest tests/test_word_module_extensions.py
python -m pytest tests/test_word_editing.py
python -m pytest tests/test_word_advanced_editing.py
```

## 검증 범위

2026-11-02 기준 전용 편집 회귀 검사 16개와 전체 검사 441개가 통과했습니다.
선택적 로컬 문서 샘플이 필요한 검사 12개는 건너뛰었습니다. 실행 예제의 글자·문단·표·
머리말·꼬리말·그림을 편집한 뒤 DOCX→HWPX→DOCX로 왕복하고 Word에서 출력한
1페이지를 비교했습니다. 흰 배경을 제외한 영역의 변경률은 0%, 잉크 영역 IoU는 1.0이며
전체 픽셀도 일치했습니다. 이 결과는 해당 예제의 검증 결과이며 모든 Word 문서의
배치 동일성을 보증하지는 않습니다.

2026-11-04에는 병합 표 편집·한영 혼합·각주·그림·필드 전용 꼬리말·가로 용지·다단을
포함한 다페이지 예제를 추가했습니다. Word 편집본과 HWPX 왕복 복원본은 96dpi 출력의
모든 페이지에서 흰 배경 제외 변경률 0%, 잉크 IoU 1.0을 확인했습니다. 한글에서도
HWPX 열기·PDF 출력·재저장을 확인했으며 균등·비균등 다단과 구분선이 유지됐습니다.
Word와 한글의 직접 렌더링 사이에는 글꼴 배치·행 높이·탭 채움선 등 차이가 남습니다.
Word 왕복 픽셀 일치와 두 앱 간 화면 일치는 서로 다른 검사 결과입니다.
