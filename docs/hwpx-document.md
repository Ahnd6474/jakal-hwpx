# `HwpxDocument`

`HwpxDocument`는 HWPX package를 직접 다루는 공개 API입니다. HWPX는 zip 안에 XML과 binary part를 담는 구조이므로, 이 클래스는 package part 접근, XML wrapper, HWPX 전용 validation을 제공합니다.

```python
from jakal_hwpx import HwpxDocument
```

## 생성과 직렬화

| API | 설명 |
|---|---|
| `HwpxDocument.blank()` | 빈 HWPX 문서를 만듭니다. |
| `HwpxDocument.open(path)` | HWPX 파일을 엽니다. |
| `HwpxDocument.from_bytes(raw_bytes)` | bytes에서 HWPX package를 엽니다. |
| `compile(validate=True)` | 현재 문서를 HWPX bytes로 컴파일합니다. |
| `save(path, validate=True)` | HWPX 파일로 저장합니다. |
| `to_hancom_document(*, converter=None)` | 공통 모델로 변환합니다. |
| `to_hwp_document(*, converter=None)` | HWP 문서로 materialize합니다. |
| `save_as_hwp(path, *, converter=None, force_refresh=True)` | HWP 파일로 저장합니다. |
| `bridge(*, converter=None)` | `HwpHwpxBridge`로 감쌉니다. |

## Package part

| API | 설명 |
|---|---|
| `list_part_paths()` | package 안의 part path 목록을 반환합니다. |
| `get_part(path, expected_type=None)` | part wrapper를 반환합니다. |
| `add_part(path, raw_bytes)` | 새 package part를 추가합니다. |
| `remove_part(path)` | package part를 제거합니다. |
| `add_or_replace_binary(name, data, *, media_type=None, manifest_id=None, is_embedded=True)` | binary part를 추가하거나 교체합니다. |
| `mimetype()` | `mimetype` part를 반환합니다. |
| `content_hpf()` | content metadata part를 반환합니다. |
| `header()` | header part를 반환합니다. |
| `container()` | container part를 반환합니다. |
| `preview_text()` | preview text part를 반환합니다. |

## XML 접근

| API | 설명 |
|---|---|
| `xml_part(path)` | XML part를 `HwpxXmlNode`로 반환합니다. |
| `section_xml(section_index=0)` | 섹션 XML을 반환합니다. |
| `header_xml()` | header XML을 반환합니다. |
| `content_hpf_xml()` | content metadata XML을 반환합니다. |
| `settings_xml()` | settings XML을 반환합니다. |
| `append_control_xml(xml, *, section_index=0, paragraph_index=None, char_pr_id=None)` | control XML을 본문에 삽입합니다. |
| `append_run_xml(xml, *, section_index=0, paragraph_index=None, char_pr_id=None)` | run XML을 본문에 삽입합니다. |

## Metadata와 text

| API | 설명 |
|---|---|
| `metadata()` | `DocumentMetadata`를 반환합니다. |
| `set_metadata(**values)` | metadata를 수정합니다. |
| `get_document_text(section_separator="\n\n")` | 본문 텍스트를 추출합니다. |
| `replace_text(old, new, count=-1, include_header=True)` | 텍스트를 치환하고 치환 횟수를 반환합니다. |
| `set_preview_text(text)` | preview text를 설정합니다. |

## Selectors

Selector는 XML wrapper list를 반환합니다. `section_index=None`이면 전체 섹션에서 찾습니다.

| API | 반환 |
|---|---|
| `headers(section_index=None)`, `footers(section_index=None)` | `list[HeaderFooterXml]` |
| `tables(section_index=None)` | `list[TableXml]` |
| `pictures(section_index=None)` | `list[PictureXml]` |
| `oles(section_index=None)` | `list[OleXml]` |
| `notes(section_index=None)` | `list[NoteXml]` |
| `memos(section_index=None)` | `list[MemoXml]` |
| `forms(section_index=None)` | `list[FormXml]` |
| `charts(section_index=None)` | `list[ChartXml]` |
| `bookmarks(section_index=None)` | `list[BookmarkXml]` |
| `fields(section_index=None)` | `list[FieldXml]` |
| `hyperlinks(section_index=None)` | `list[FieldXml]` |
| `auto_numbers(section_index=None)` | `list[AutoNumberXml]` |
| `equations(section_index=None)` | `list[EquationXml]` |
| `shapes(section_index=None)` | `list[ShapeXml]` |
| `styles()`, `paragraph_styles()`, `character_styles()` | style wrapper list |
| `memo_shapes()` | `list[MemoShapeXml]` |

## 문단 편집

| API | 설명 |
|---|---|
| `append_paragraph(text, *, section_index=0, template_index=None, para_pr_id=None, style_id=None, char_pr_id=None)` | 문단을 추가합니다. |
| `insert_paragraph(...)` | 지정 위치에 문단을 삽입합니다. |
| `set_paragraph_text(section_index, paragraph_index, text)` | 문단 텍스트를 교체합니다. |
| `delete_paragraph(section_index, paragraph_index)` | 문단을 삭제합니다. |
| `apply_style_to_paragraph(section_index, paragraph_index, *, style_id=None, para_pr_id=None, char_pr_id=None)` | 문단에 스타일을 적용합니다. |
| `apply_style_batch(*, section_index=None, text_contains=None, regex=None, style_id=None, para_pr_id=None, char_pr_id=None)` | 조건에 맞는 문단에 일괄 스타일을 적용하고 개수를 반환합니다. |

## Control 추가

| API | 반환 |
|---|---|
| `append_header(text="", *, apply_page_type="BOTH", hide_first=None, ...)` | `HeaderFooterXml` |
| `append_footer(text="", *, apply_page_type="BOTH", hide_first=None, ...)` | `HeaderFooterXml` |
| `append_note(text="", *, kind="footNote", number=None, ...)` | `NoteXml` |
| `append_footnote(text="", *, number=None, ...)` | `NoteXml` |
| `append_endnote(text="", *, number=None, ...)` | `NoteXml` |
| `append_form(label="", *, form_type="INPUT", ...)` | `FormXml` |
| `append_memo(text="", ...)` | `MemoXml` |
| `append_chart(title="", *, chart_type="BAR", ...)` | `ChartXml` |
| `append_auto_number(*, number=1, number_type="PAGE", kind="newNum", ...)` | `AutoNumberXml` |
| `append_new_number(*, number=1, number_type="PAGE", ...)` | `AutoNumberXml` |
| `append_equation(script, *, width=4800, height=2300, ...)` | `EquationXml` |
| `append_table(rows, columns, *, cell_texts=None, ...)` | `TableXml` |
| `append_picture(name, data, *, media_type=None, width=7200, height=7200, ...)` | `PictureXml` |
| `append_shape(*, kind="rect", text="", width=12000, height=3200, ...)` | `ShapeXml` |
| `append_ole(name, data, *, media_type="application/ole", ...)` | `OleXml` |
| `append_bookmark(name, ...)` | `BookmarkXml` |
| `append_field(*, field_type, display_text=None, name=None, parameters=None, ...)` | `FieldXml` |
| `append_hyperlink(target, *, display_text=None, ...)` | `FieldXml` |
| `append_mail_merge_field(field_name, *, display_text=None, ...)` | `FieldXml` |
| `append_calculation_field(expression, *, display_text=None, ...)` | `FieldXml` |
| `append_cross_reference(bookmark_name, *, display_text=None, ...)` | `FieldXml` |

## Styles

| API | 반환 |
|---|---|
| `append_style(name, *, english_name=None, style_id=None, style_type="PARA", ...)` | `StyleDefinitionXml` |
| `append_paragraph_style(*, style_id=None, template_id=None, alignment_horizontal=None, ...)` | `ParagraphStyleXml` |
| `append_character_style(*, style_id=None, template_id=None, text_color=None, ...)` | `CharacterStyleXml` |
| `append_memo_shape(...)` | `MemoShapeXml` |
| `get_style(style_id)` | `StyleDefinitionXml` |
| `get_paragraph_style(style_id)` | `ParagraphStyleXml` |
| `get_character_style(style_id)` | `CharacterStyleXml` |
| `get_memo_shape(memo_shape_id)` | `MemoShapeXml` |

## Validation

| API | 설명 |
|---|---|
| `validation_errors()` | 기본 validation issue 목록 |
| `validate()` | validation issue가 있으면 예외를 발생시킵니다. |
| `roundtrip_validate()` | compile/reopen 경로를 검증합니다. |
| `xml_validation_errors()` | XML 구조 issue |
| `schema_validation_errors(schema_map)` | XSD schema validation issue |
| `reference_validation_errors()` | part/reference issue |
| `save_reopen_validation_errors()` | 저장 후 다시 열기 issue |
| `strict_lint_errors()` | 엄격 lint issue |
| `strict_validate()` | strict lint issue가 있으면 예외를 발생시킵니다. |
| `strict_lint_report(include_none=False)` | 사람이 읽기 쉬운 report list |
| `format_strict_lint_errors()` | report를 문자열로 formatting |

## 예제

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("input.hwpx")
doc.set_metadata(title="최종본")
doc.replace_text("초안", "최종")
doc.append_footer("내부용", apply_page_type="BOTH")

errors = doc.format_strict_lint_errors()
if errors:
    print(errors)

doc.strict_validate()
doc.save("build/output.hwpx")
```

## 참고

- HWPX wrapper API는 XML 구조에 가까운 레이어입니다. 포맷 독립적인 문서 자동화에는 [`HancomDocument`](./hancom-document.md)를 먼저 고려하세요.
- `append_control_xml()`과 `append_run_xml()`은 강력하지만, XML 구조를 직접 책임져야 합니다.
- `save(validate=True)`는 저장 전 기본 validation을 실행합니다.
