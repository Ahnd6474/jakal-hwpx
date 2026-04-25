# `HancomDocument`

`HancomDocument`는 `jakal_hwpx`의 기본 문서 모델입니다. HWPX와 HWP를 서로 다른 내부 포맷으로 보지 않고, 같은 Python 객체 모델로 올려서 편집한 뒤 원하는 포맷으로 저장합니다.

```python
from jakal_hwpx import HancomDocument
```

## 사용 시점

| 상황 | 권장 |
|---|---|
| 새 문서를 만들고 HWP/HWPX 둘 다 저장 | `HancomDocument.blank()` |
| 입력 포맷이 HWP인지 HWPX인지 섞여 있음 | `read_hwp()`, `read_hwpx()` |
| 앱 코드에서 XML이나 binary record를 직접 다루고 싶지 않음 | `HancomDocument` |
| HWPX package 구조를 직접 고쳐야 함 | [`HwpxDocument`](./hwpx-document.md) |
| HWP native wrapper를 직접 써야 함 | [`HwpDocument`](./hwp-document.md) |

## 생성

| API | 설명 |
|---|---|
| `HancomDocument.blank(*, converter=None)` | 빈 공통 문서 모델을 만듭니다. |
| `HancomDocument.read_hwpx(path, *, converter=None)` | HWPX를 읽어 공통 모델로 변환합니다. |
| `HancomDocument.read_hwp(path, *, converter=None)` | HWP를 읽어 공통 모델로 변환합니다. |
| `HancomDocument.from_hwpx_document(document, *, converter=None)` | 이미 열린 `HwpxDocument`를 공통 모델로 올립니다. |
| `HancomDocument.from_hwp_document(document, *, converter=None)` | 이미 열린 `HwpDocument`를 공통 모델로 올립니다. |

## 저장과 변환

| API | 설명 |
|---|---|
| `to_hwpx_document()` | 현재 모델을 `HwpxDocument`로 materialize합니다. |
| `to_hwp_document(*, converter=None)` | 현재 모델을 `HwpDocument`로 materialize합니다. |
| `write_to_hwpx(path, *, validate=True)` | HWPX 파일로 저장합니다. |
| `write_to_hwp(path, *, converter=None)` | HWP 파일로 저장합니다. |
| `bridge(*, converter=None)` | `HwpHwpxBridge`로 감싸 materialization 경로를 명시합니다. |

## 최소 예제

```python
from jakal_hwpx import HancomDocument

doc = HancomDocument.blank()
doc.metadata.title = "분기 보고서"

doc.append_paragraph("요약")
doc.append_table(
    rows=2,
    cols=2,
    cell_texts=[["구분", "금액"], ["1분기", "120"]],
)

doc.write_to_hwpx("build/report.hwpx")
doc.write_to_hwp("build/report.hwp")
```

## 문서 속성

| 속성 | 타입 | 설명 |
|---|---|---|
| `metadata` | `HancomMetadata` | 문서 제목, 작성자, 날짜 등 |
| `sections` | `list[HancomSection]` | 섹션 목록 |
| `source_format` | `str \| None` | `read_hwp()` 또는 `read_hwpx()`로 읽은 원본 포맷 |
| `style_definitions` | `list[StyleDefinition]` | 문단/글자 스타일 정의 |
| `paragraph_styles` | `list[ParagraphStyle]` | 문단 속성 정의 |
| `character_styles` | `list[CharacterStyle]` | 글자 속성 정의 |
| `numbering_definitions` | `list[NumberingDefinition]` | 번호 정의 |
| `bullet_definitions` | `list[BulletDefinition]` | bullet 정의 |
| `memo_shape_definitions` | `list[MemoShapeDefinition]` | memo shape 정의 |

## Section model

`HancomDocument`는 섹션을 `HancomSection`으로 표현합니다. 섹션 안에는 설정, 머리말/꼬리말, 본문 block이 들어갑니다.

```python
section = doc.append_section()
section.settings.page_width = 59528
section.settings.page_height = 84188
```

| 타입 | 필드 |
|---|---|
| `HancomSection` | `settings`, `header_footer_blocks`, `blocks` |
| `SectionSettings` | `page_width`, `page_height`, `landscape`, `margins`, `page_border_fills`, `visibility`, `grid`, `start_numbers`, `page_numbers`, `footnote_pr`, `endnote_pr`, `line_number_shape`, `numbering_shape_id`, `memo_shape_id` |

## 추가 API

### Body blocks

| API | 반환 | 설명 |
|---|---|---|
| `append_paragraph(text, *, section_index=0)` | `Paragraph` | 문단을 추가합니다. |
| `append_table(*, rows, cols, cell_texts=None, ...)` | `Table` | 표를 추가합니다. |
| `append_picture(name, data, *, extension=None, width=7200, height=7200, section_index=0)` | `Picture` | 그림 binary를 문서에 추가합니다. |
| `append_shape(*, kind="rect", text="", width=12000, height=3200, ...)` | `Shape` | 도형을 추가합니다. |
| `append_equation(script, *, width=4800, height=2300, ...)` | `Equation` | 수식을 추가합니다. |
| `append_ole(name, data, *, width=42001, height=13501, ...)` | `Ole` | OLE 객체를 추가합니다. |
| `append_chart(title="", *, chart_type="BAR", categories=None, series=None, ...)` | `Chart` | 차트를 추가합니다. |

### References and fields

| API | 반환 | 설명 |
|---|---|---|
| `append_hyperlink(target, *, display_text=None, ...)` | `Hyperlink` | 하이퍼링크를 추가합니다. |
| `append_bookmark(name, *, section_index=0)` | `Bookmark` | 북마크를 추가합니다. |
| `append_field(*, field_type, display_text=None, name=None, parameters=None, ...)` | `Field` | 일반 필드를 추가합니다. |
| `append_mail_merge_field(field_name, *, display_text=None, ...)` | `Field` | 메일 머지 필드를 추가합니다. |
| `append_calculation_field(expression, *, display_text=None, ...)` | `Field` | 계산 필드를 추가합니다. |
| `append_cross_reference(bookmark_name, *, display_text=None, ...)` | `Field` | 북마크 참조 필드를 추가합니다. |
| `append_doc_property_field(property_name, *, display_text=None, ...)` | `Field` | 문서 속성 필드를 추가합니다. |
| `append_date_field(*, display_text=None, ...)` | `Field` | 날짜 필드를 추가합니다. |

### Section and annotation controls

| API | 반환 | 설명 |
|---|---|---|
| `append_header(text, *, apply_page_type="BOTH", section_index=0)` | `HeaderFooter` | 머리말을 추가합니다. |
| `append_footer(text, *, apply_page_type="BOTH", section_index=0)` | `HeaderFooter` | 꼬리말을 추가합니다. |
| `append_note(text, *, kind="footNote", number=None, section_index=0)` | `Note` | 각주/미주를 추가합니다. |
| `append_footnote(text, *, number=None, section_index=0)` | `Note` | 각주를 추가합니다. |
| `append_endnote(text, *, number=None, section_index=0)` | `Note` | 미주를 추가합니다. |
| `append_auto_number(*, number=1, number_type="PAGE", kind="newNum", section_index=0)` | `AutoNumber` | 자동 번호 control을 추가합니다. |
| `append_memo(text, *, author=None, memo_id=None, ...)` | `Memo` | memo/comment를 추가합니다. |
| `append_form(label="", *, form_type="INPUT", ...)` | `Form` | form object를 추가합니다. |

### Styles

| API | 반환 | 설명 |
|---|---|---|
| `append_style(name, *, style_id=None, style_type="PARA", ...)` | `StyleDefinition` | 스타일 정의를 추가합니다. |
| `append_paragraph_style(*, style_id=None, alignment_horizontal=None, ...)` | `ParagraphStyle` | 문단 속성 정의를 추가합니다. |
| `append_character_style(*, style_id=None, text_color=None, ...)` | `CharacterStyle` | 글자 속성 정의를 추가합니다. |
| `append_numbering_definition(...)` | `NumberingDefinition` | 번호 정의를 추가합니다. |
| `append_bullet_definition(...)` | `BulletDefinition` | bullet 정의를 추가합니다. |
| `append_memo_shape_definition(...)` | `MemoShapeDefinition` | memo shape 정의를 추가합니다. |

## 데이터 클래스

### Metadata

| 클래스 | 주요 필드 |
|---|---|
| `HancomMetadata` | `title`, `language`, `creator`, `subject`, `description`, `lastsaveby`, `created`, `modified`, `date`, `keyword`, `extra` |

### Body blocks

| 클래스 | 주요 필드 |
|---|---|
| `Paragraph` | `text`, `style_id`, `para_pr_id`, `char_pr_id`, `hwp_para_shape_id`, `hwp_style_id` |
| `Table` | `rows`, `cols`, `cell_texts`, `row_heights`, `col_widths`, `cell_spans`, `cell_border_fill_ids`, `table_border_fill_id`, `layout`, `out_margins` |
| `Picture` | `name`, `data`, `extension`, `width`, `height`, `shape_comment`, `layout`, `out_margins`, `rotation`, `crop`, `line_color`, `line_width` |
| `Shape` | `kind`, `text`, `width`, `height`, `fill_color`, `line_color`, `shape_comment`, `layout`, `rotation`, `specific_fields` |
| `Equation` | `script`, `width`, `height`, `shape_comment`, `text_color`, `base_unit`, `font`, `layout`, `rotation` |
| `Ole` | `name`, `data`, `width`, `height`, `object_type`, `draw_aspect`, `has_moniker`, `eq_baseline`, `extent` |
| `Chart` | `title`, `chart_type`, `categories`, `series`, `data_ref`, `legend_visible`, `width`, `height` |

### Controls

| 클래스 | 주요 필드 |
|---|---|
| `Hyperlink` | `target`, `display_text`, `metadata_fields` |
| `Bookmark` | `name` |
| `Field` | `field_type`, `display_text`, `name`, `parameters`, `editable`, `dirty`, `native_field_type` |
| `AutoNumber` | `kind`, `number`, `number_type` |
| `Note` | `kind`, `text`, `number` |
| `HeaderFooter` | `kind`, `text`, `apply_page_type` |
| `Memo` | `text`, `author`, `memo_id`, `anchor_id`, `order`, `visible` |
| `Form` | `label`, `form_type`, `name`, `value`, `checked`, `items`, `editable`, `locked`, `placeholder` |

## 참고

- `HancomDocument`는 앱 코드용 추상 모델입니다. HWPX XML wrapper가 필요하면 `to_hwpx_document()`를 호출하거나 `HwpxDocument`를 직접 사용합니다.
- HWP native wrapper가 필요하면 `to_hwp_document()`를 호출하거나 `HwpDocument`를 직접 사용합니다.
- `write_to_hwpx(validate=True)`는 HWPX 저장 시 validation을 실행합니다.
- 지원 범위와 안정성 기준은 [STABILITY_CONTRACT.md](../STABILITY_CONTRACT.md)를 따릅니다.
