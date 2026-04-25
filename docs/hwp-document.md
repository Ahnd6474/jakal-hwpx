# `HwpDocument`

`HwpDocument`는 HWP binary 문서를 직접 다루는 고수준 객체 API입니다. 내부적으로는 `HwpBinaryDocument`를 사용하지만, 일반적인 HWP 편집은 stream/record를 직접 건드리지 않고 이 클래스로 처리합니다.

```python
from jakal_hwpx import HwpDocument
```

## 생성과 직렬화

| API | 설명 |
|---|---|
| `HwpDocument.blank(*, converter=None)` | bundled pure profile을 기반으로 빈 HWP 문서를 만듭니다. |
| `HwpDocument.open(path, *, converter=None)` | HWP 파일을 엽니다. |
| `HwpDocument.blank_from_profile(profile_root, *, converter=None)` | 특정 pure profile을 기반으로 빈 문서를 만듭니다. |
| `save(path=None)` | HWP 파일로 저장합니다. |
| `save_as_hwpx(path)` | HWPX 파일로 저장합니다. |
| `to_hwpx_document(*, force_refresh=False)` | `HwpxDocument`로 materialize합니다. |
| `to_hancom_document(*, force_refresh=False)` | `HancomDocument`로 materialize합니다. |
| `bridge()` | `HwpHwpxBridge`로 감쌉니다. |

## 문서 단위 접근

| API | 설명 |
|---|---|
| `source_path` | 원본 HWP 경로입니다. |
| `binary_document()` | 내부 `HwpBinaryDocument`를 반환합니다. |
| `file_header()` | HWP file header 정보를 반환합니다. |
| `document_properties()` | 문서 속성 record를 반환합니다. |
| `docinfo_model()` | `DocInfoModel`을 반환합니다. |
| `list_stream_paths()` | OLE stream path 목록을 반환합니다. |
| `bindata_stream_paths()` | BinData stream path 목록을 반환합니다. |
| `stream_capacity(path)` | 해당 stream의 capacity 정보를 반환합니다. |
| `preview_text()` | preview text를 반환합니다. |
| `set_preview_text(value)` | preview text를 설정합니다. |
| `get_document_text()` | 본문 텍스트를 추출합니다. |

## Sections

| API | 설명 |
|---|---|
| `sections()` | `HwpSection` 목록을 반환합니다. |
| `section(index)` | 특정 `HwpSection`을 반환합니다. |
| `section_model(section_index=0)` | 저수준 `SectionModel`을 반환합니다. |
| `ensure_section_count(section_count)` | 섹션 개수를 보장합니다. |
| `apply_section_settings(...)` | 용지, 여백, visibility, grid, numbering 등을 적용합니다. |
| `apply_section_page_border_fills(page_border_fills, *, section_index=0)` | page border fill을 적용합니다. |
| `apply_section_page_numbers(page_numbers, *, section_index=0)` | page number control을 적용합니다. |
| `apply_section_note_settings(...)` | 각주/미주 설정을 적용합니다. |

## Selectors

Selector는 HWP native wrapper list를 반환합니다. `section_index=None`이면 전체 섹션에서 찾습니다.

| API | 반환 |
|---|---|
| `paragraphs(section_index=0)` | `list[HwpParagraphObject]` |
| `controls(section_index=None)` | `list[HwpControlObject]` |
| `tables(section_index=None)` | `list[HwpTableObject]` |
| `pictures(section_index=None)` | `list[HwpPictureObject]` |
| `hyperlinks(section_index=None)` | `list[HwpHyperlinkObject]` |
| `bookmarks(section_index=None)` | `list[HwpBookmarkObject]` |
| `notes(section_index=None)` | `list[HwpNoteObject]` |
| `page_numbers(section_index=None)` | `list[HwpPageNumObject]` |
| `fields(section_index=None)` | `list[HwpFieldObject]` |
| `mail_merge_fields(section_index=None)` | `list[HwpFieldObject]` |
| `calculation_fields(section_index=None)` | `list[HwpFieldObject]` |
| `cross_references(section_index=None)` | `list[HwpFieldObject]` |
| `doc_property_fields(section_index=None)` | `list[HwpFieldObject]` |
| `date_fields(section_index=None)` | `list[HwpFieldObject]` |
| `shapes(section_index=None)` | `list[HwpShapeObject]` |
| `equations(section_index=None)` | `list[HwpEquationObject]` |
| `oles(section_index=None)` | `list[HwpOleObject]` |
| `forms(section_index=None)` | `list[HwpFormObject]` |
| `memos(section_index=None)` | `list[HwpMemoObject]` |
| `charts(section_index=None)` | `list[HwpChartObject]` |

## Shape subtype selector

| API | 반환 |
|---|---|
| `lines(section_index=None)` | `list[HwpLineShapeObject]` |
| `connect_lines(section_index=None)` | `list[HwpConnectLineShapeObject]` |
| `rectangles(section_index=None)` | `list[HwpRectangleShapeObject]` |
| `ellipses(section_index=None)` | `list[HwpEllipseShapeObject]` |
| `arcs(section_index=None)` | `list[HwpArcShapeObject]` |
| `polygons(section_index=None)` | `list[HwpPolygonShapeObject]` |
| `curves(section_index=None)` | `list[HwpCurveShapeObject]` |
| `containers(section_index=None)` | `list[HwpContainerShapeObject]` |
| `textarts(section_index=None)` | `list[HwpTextArtShapeObject]` |

## 편집 API

| API | 설명 |
|---|---|
| `replace_text_same_length(old, new, *, section_index=None, count=-1)` | 같은 길이 텍스트 치환을 수행합니다. |
| `append_paragraph(text, *, section_index=0, ...)` | 문단을 추가합니다. |
| `append_table(...)` | 표를 추가합니다. |
| `append_picture(...)` | 그림을 추가합니다. |
| `append_hyperlink(...)` | 하이퍼링크를 추가합니다. |
| `append_field(...)` | 필드를 추가합니다. |
| `append_auto_number(...)` | 자동 번호를 추가합니다. |
| `append_header(text, *, apply_page_type="BOTH", section_index=None)` | 머리말을 추가합니다. |
| `append_footer(text, *, apply_page_type="BOTH", section_index=None)` | 꼬리말을 추가합니다. |
| `append_bookmark(name, ...)` | 북마크를 추가합니다. |
| `append_note(text, *, kind="footNote", ...)` | 각주/미주를 추가합니다. |
| `append_footnote(text, ...)`, `append_endnote(text, ...)` | 각주/미주 convenience API입니다. |
| `append_form(...)`, `append_memo(...)`, `append_chart(...)` | form, memo, chart를 추가합니다. |
| `append_equation(script, ...)`, `append_shape(...)`, `append_ole(...)` | 수식, 도형, OLE를 추가합니다. |

## Validation

| API | 설명 |
|---|---|
| `strict_lint_errors()` | HWP strict lint issue 목록을 반환합니다. |
| `strict_validate()` | strict lint issue가 있으면 예외를 발생시킵니다. |
| `strict_lint_report(include_none=False)` | report list를 반환합니다. |
| `format_strict_lint_errors()` | issue를 문자열로 formatting합니다. |

## 예제

```python
from jakal_hwpx import HwpDocument

doc = HwpDocument.open("input.hwp")
doc.append_paragraph("검토 완료")
doc.append_hyperlink("https://example.com", text="참고 링크")

for table in doc.tables():
    table.set_cell_text(0, 0, "수정")

doc.strict_validate()
doc.save("build/output.hwp")
```

## 참고

- `replace_text_same_length()`는 HWP binary 구조를 안전하게 유지하기 위해 같은 길이 치환만 허용합니다. 길이가 바뀌는 일반 텍스트 편집은 wrapper API나 공통 모델 경로를 사용하세요.
- HWP stream과 record를 직접 조사해야 할 때는 [`HwpBinaryDocument`](./bridge-and-binary.md)를 사용합니다.
- HWP와 HWPX를 같은 코드로 처리하려면 [`HancomDocument`](./hancom-document.md)를 먼저 고려하세요.
