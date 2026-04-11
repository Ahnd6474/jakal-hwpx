# `jakal_hwpx` Module Guide

`jakal_hwpx`의 가장 안정적인 공용 편집 모델은 `HancomDocument`입니다.

이 객체는 `.hwpx`와 `.hwp` 사이에 놓이는 공통 IR이며, 문서 구조를 Python dataclass 중심으로 다룰 수 있게 해 줍니다.  
실무 기준으로는 다음 흐름을 권장합니다.

- `.hwpx`를 읽어서 수정: `HancomDocument.read_hwpx()`
- `.hwp`를 읽어서 수정: `HancomDocument.read_hwp()`
- 새 문서 조립: `HancomDocument.blank()`
- 결과 저장:
  - `.hwpx`로 저장: `write_to_hwpx()`
  - `.hwp`로 저장: `write_to_hwp()`

즉, `HwpxDocument`와 `HwpDocument`는 중요하지만, 앱 코드의 중심 편집 객체는 `HancomDocument`로 두는 편이 가장 단순합니다.

## 권장 진입점

### 새 문서 만들기

```python
from jakal_hwpx import HancomDocument

doc = HancomDocument.blank()
doc.metadata.title = "API Sample"
doc.append_paragraph("Hello HancomDocument")
doc.write_to_hwpx("build/sample.hwpx")
```

### `.hwpx` 읽어서 수정하기

```python
from jakal_hwpx import HancomDocument

doc = HancomDocument.read_hwpx("input.hwpx")
doc.append_paragraph("added from IR", section_index=0)
doc.write_to_hwpx("build/edited.hwpx")
```

### `.hwp` 읽어서 수정하기

```python
from jakal_hwpx import HancomDocument

doc = HancomDocument.read_hwp("input.hwp")
doc.append_paragraph("added from IR", section_index=0)
doc.write_to_hwp("build/edited.hwp")
doc.write_to_hwpx("build/edited.hwpx")
```

## 왜 `HancomDocument` 중심인가

`HancomDocument`는 다음 장점이 있습니다.

- 입력 포맷이 `.hwpx`인지 `.hwp`인지 감출 수 있습니다.
- 문서를 XML 노드나 binary record 대신 Python 객체로 다룰 수 있습니다.
- section, block, style, metadata를 한 곳에서 조립할 수 있습니다.
- 같은 IR에서 `.hwpx`와 `.hwp`를 모두 생성할 수 있습니다.
- pure Python bridge가 기본이므로 Hancom 자동화에 의존하지 않는 경로를 기본으로 가져갈 수 있습니다.

## 핵심 객체 모델

### `HancomDocument`

문서 전체를 나타냅니다.

주요 필드:

- `metadata: HancomMetadata`
- `sections: list[HancomSection]`
- `style_definitions: list[StyleDefinition]`
- `paragraph_styles: list[ParagraphStyle]`
- `character_styles: list[CharacterStyle]`
- `source_format: str | None`

주요 생성 API:

- `HancomDocument.blank(*, converter=None) -> HancomDocument`
- `HancomDocument.read_hwpx(path, *, converter=None) -> HancomDocument`
- `HancomDocument.read_hwp(path, *, converter=None) -> HancomDocument`
- `HancomDocument.from_hwpx_document(document, *, converter=None) -> HancomDocument`
- `HancomDocument.from_hwp_document(document, *, converter=None) -> HancomDocument`

주요 변환/저장 API:

- `to_hwpx_document() -> HwpxDocument`
- `write_to_hwpx(path, *, validate=True) -> Path`
- `to_hwp_document(*, converter=None) -> HwpDocument`
- `write_to_hwp(path, *, converter=None) -> Path`

주요 append API:

- `append_section() -> HancomSection`
- `append_paragraph(text, *, section_index=0) -> Paragraph`
- `append_table(..., section_index=0) -> Table`
- `append_picture(name, data, ..., section_index=0) -> Picture`
- `append_hyperlink(target, ..., section_index=0) -> Hyperlink`
- `append_bookmark(name, *, section_index=0) -> Bookmark`
- `append_field(..., section_index=0) -> Field`
- `append_auto_number(..., section_index=0) -> AutoNumber`
- `append_note(..., section_index=0) -> Note`
- `append_footnote(..., section_index=0) -> Note`
- `append_endnote(..., section_index=0) -> Note`
- `append_equation(..., section_index=0) -> Equation`
- `append_shape(..., section_index=0) -> Shape`
- `append_ole(name, data, ..., section_index=0) -> Ole`
- `append_header(text, ..., section_index=0) -> HeaderFooter`
- `append_footer(text, ..., section_index=0) -> HeaderFooter`
- `append_style(...) -> StyleDefinition`
- `append_paragraph_style(...) -> ParagraphStyle`
- `append_character_style(...) -> CharacterStyle`

## IR dataclass 구조

### `HancomMetadata`

문서 메타데이터를 담습니다.

주요 필드:

- `title`
- `language`
- `creator`
- `subject`
- `description`
- `lastsaveby`
- `created`
- `modified`
- `date`
- `keyword`
- `extra`

예:

```python
from jakal_hwpx import HancomDocument

doc = HancomDocument.blank()
doc.metadata.title = "2026 Report"
doc.metadata.creator = "jakal-hwpx"
doc.metadata.subject = "Pure Python bridge"
doc.metadata.keyword = "hwp,hwpx,bridge"
```

### `HancomSection`

section 단위 컨테이너입니다.

주요 필드:

- `settings: SectionSettings`
- `header_footer_blocks: list[HeaderFooter]`
- `blocks: list[HancomBlock]`

`blocks` 안에는 문단, 표, 그림, 링크, 주석 같은 본문 block이 순서대로 들어갑니다.

### `HancomBlock`

본문에 들어갈 수 있는 block union입니다.

지원 타입:

- `Paragraph`
- `Table`
- `Picture`
- `Hyperlink`
- `Bookmark`
- `Field`
- `AutoNumber`
- `Note`
- `Equation`
- `Shape`
- `Ole`

## block 타입 상세

### `Paragraph`

주요 필드:

- `text: str`

```python
paragraph = doc.append_paragraph("body text")
paragraph.text = "updated body text"
```

### `Table`

주요 필드:

- `rows`
- `cols`
- `cell_texts`
- `row_heights`
- `col_widths`
- `cell_spans`
- `cell_border_fill_ids`
- `table_border_fill_id`

```python
table = doc.append_table(
    rows=2,
    cols=3,
    cell_texts=[
        ["A1", "B1", "C1"],
        ["A2", "B2", "C2"],
    ],
    col_widths=[3000, 3000, 3000],
)
table.cell_texts[1][1] = "UPDATED"
```

### `Picture`

주요 필드:

- `name`
- `data`
- `extension`
- `width`
- `height`

```python
from pathlib import Path

pic = doc.append_picture(
    "chart.png",
    Path("assets/chart.png").read_bytes(),
    extension="png",
    width=9000,
    height=6000,
)
```

### `Hyperlink`

주요 필드:

- `target`
- `display_text`

```python
link = doc.append_hyperlink("https://example.com", display_text="Example")
link.target = "https://openai.com"
link.display_text = "OpenAI"
```

### `Bookmark`

주요 필드:

- `name`

```python
bookmark = doc.append_bookmark("summary_anchor")
bookmark.name = "summary_anchor_v2"
```

### `Field`

주요 필드:

- `field_type`
- `display_text`
- `name`
- `parameters`
- `editable`
- `dirty`

```python
field = doc.append_field(
    field_type="FORMULA",
    display_text="42",
    parameters={"Expression": "40+2", "Command": "40+2"},
)
field.parameters["Expression"] = "50+8"
field.display_text = "58"
```

### `AutoNumber`

주요 필드:

- `kind`
- `number`
- `number_type`

```python
doc.append_auto_number(number=1, number_type="PAGE", kind="newNum")
```

### `Note`

주요 필드:

- `kind`
- `text`
- `number`

```python
footnote = doc.append_footnote("footnote text", number=1)
endnote = doc.append_endnote("endnote text", number=1)
footnote.text = "updated footnote"
```

### `Equation`

주요 필드:

- `script`
- `width`
- `height`

```python
eq = doc.append_equation("x^2+y^2=z^2", width=5200, height=2400)
eq.script = "\\frac{1}{2}"
```

### `Shape`

주요 필드:

- `kind`
- `text`
- `width`
- `height`
- `fill_color`
- `line_color`

```python
shape = doc.append_shape(
    kind="rect",
    text="Box label",
    width=12000,
    height=3200,
    fill_color="#F5F5F5",
    line_color="#222222",
)
```

### `Ole`

주요 필드:

- `name`
- `data`
- `width`
- `height`

```python
from pathlib import Path

ole = doc.append_ole(
    "embedded.bin",
    Path("assets/object.bin").read_bytes(),
    width=42001,
    height=13501,
)
```

### `HeaderFooter`

주요 필드:

- `kind`
- `text`
- `apply_page_type`

```python
header = doc.append_header("Document Header", apply_page_type="BOTH")
footer = doc.append_footer("Page Footer", apply_page_type="BOTH")
header.text = "Updated Header"
```

## style 타입 상세

### `StyleDefinition`

주요 필드:

- `name`
- `style_id`
- `english_name`
- `style_type`
- `para_pr_id`
- `char_pr_id`
- `next_style_id`
- `lang_id`
- `lock_form`

### `ParagraphStyle`

주요 필드:

- `style_id`
- `alignment_horizontal`
- `alignment_vertical`
- `line_spacing`

### `CharacterStyle`

주요 필드:

- `style_id`
- `text_color`
- `height`

예:

```python
style = doc.append_style(
    "Body Center",
    style_id="100",
    para_pr_id="100",
    char_pr_id="100",
)
para_style = doc.append_paragraph_style(
    style_id="100",
    alignment_horizontal="CENTER",
    line_spacing=160,
)
char_style = doc.append_character_style(
    style_id="100",
    text_color="#112233",
    height=1100,
)
```

## `SectionSettings` 상세

`HancomSection.settings`에 들어가는 객체입니다.

주요 필드:

- `page_width`
- `page_height`
- `landscape`
- `margins`
- `page_border_fills`
- `visibility`
- `grid`
- `start_numbers`
- `page_numbers`
- `footnote_pr`
- `endnote_pr`
- `line_number_shape`

설정 예:

```python
from jakal_hwpx import HancomDocument

doc = HancomDocument.blank()
section = doc.sections[0]

section.settings.page_width = 60000
section.settings.page_height = 85000
section.settings.landscape = "NARROWLY"
section.settings.margins = {
    "left": 7000,
    "right": 7000,
    "top": 5000,
    "bottom": 5000,
    "header": 3000,
    "footer": 3000,
    "gutter": 0,
}
section.settings.visibility = {
    "hideFirstHeader": "0",
    "hideFirstFooter": "0",
    "border": "SHOW_ALL",
    "fill": "SHOW_ALL",
    "showLineNumber": "1",
}
section.settings.grid = {
    "lineGrid": 0,
    "charGrid": 0,
    "wonggojiFormat": 0,
}
section.settings.start_numbers = {
    "pageStartsOn": "BOTH",
    "page": "1",
    "pic": "1",
    "tbl": "1",
    "equation": "1",
}
section.settings.page_numbers = [
    {
        "pos": "BOTTOM_CENTER",
        "formatType": "DIGIT",
        "sideChar": "NONE",
    }
]
section.settings.footnote_pr = {
    "numberShape": "DIGIT",
    "placement": "EACH_COLUMN",
}
section.settings.endnote_pr = {
    "numberShape": "DIGIT",
    "placement": "END_OF_DOCUMENT",
}
```

## 대표 사용 패턴

### 1. `.hwpx`를 읽어서 IR 기준으로 수정 후 다시 `.hwpx` 저장

```python
from jakal_hwpx import HancomDocument

doc = HancomDocument.read_hwpx("input.hwpx")

doc.metadata.title = "Updated Title"
doc.append_paragraph("added paragraph", section_index=0)
doc.append_hyperlink("https://example.com/final", display_text="Final Link", section_index=0)
doc.write_to_hwpx("build/output.hwpx")
```

### 2. `.hwp`를 읽어서 `.hwpx`와 `.hwp` 둘 다 내보내기

```python
from jakal_hwpx import HancomDocument

doc = HancomDocument.read_hwp("input.hwp")
doc.append_paragraph("bridge-added", section_index=0)
doc.append_bookmark("appendix_anchor", section_index=0)
doc.write_to_hwp("build/output.hwp")
doc.write_to_hwpx("build/output.hwpx")
```

### 3. 새 문서를 section 중심으로 조립

```python
from pathlib import Path
from jakal_hwpx import HancomDocument

doc = HancomDocument.blank()
doc.metadata.title = "Generated Document"

section = doc.sections[0]
section.settings.page_width = 60000
section.settings.page_height = 85000

doc.append_header("Generated Header", section_index=0)
doc.append_footer("Generated Footer", section_index=0)
doc.append_paragraph("First paragraph", section_index=0)
doc.append_table(
    rows=2,
    cols=2,
    cell_texts=[["A1", "B1"], ["A2", "B2"]],
    section_index=0,
)
doc.append_picture(
    "logo.png",
    Path("assets/logo.png").read_bytes(),
    extension="png",
    section_index=0,
)
doc.append_footnote("footnote from IR", number=1, section_index=0)

doc.write_to_hwpx("build/generated.hwpx")
doc.write_to_hwp("build/generated.hwp")
```

### 4. style과 본문을 함께 조립

```python
from jakal_hwpx import HancomDocument

doc = HancomDocument.blank()

doc.append_style(
    "Body Center",
    style_id="100",
    para_pr_id="100",
    char_pr_id="100",
)
doc.append_paragraph_style(
    style_id="100",
    alignment_horizontal="CENTER",
    line_spacing=160,
)
doc.append_character_style(
    style_id="100",
    text_color="#112233",
    height=1100,
)

doc.append_paragraph("Styled paragraph")
doc.write_to_hwpx("build/styled.hwpx")
```

## `HancomDocument`와 다른 객체의 관계

### `HwpxDocument`

`HwpxDocument`는 HWPX package를 직접 다루는 XML 중심 고수준 API입니다.

언제 쓰면 좋은가:

- 특정 XML wrapper 메서드를 직접 써야 할 때
- package part, manifest, preview text까지 직접 만져야 할 때
- 문단/표/그림 수정 후 즉시 validation API를 세밀하게 돌리고 싶을 때

`HancomDocument`와의 연결:

- `HancomDocument.from_hwpx_document(hwpx_doc)`
- `HancomDocument.to_hwpx_document()`

### `HwpDocument`

`HwpDocument`는 native HWP binary를 직접 다루는 high-level API입니다.

언제 쓰면 좋은가:

- HWP binary stream capacity나 preview text를 직접 봐야 할 때
- same-length text replacement 같은 native HWP 편집이 필요할 때
- HWP low-level behavior를 디버깅할 때

`HancomDocument`와의 연결:

- `HancomDocument.from_hwp_document(hwp_doc)`
- `HancomDocument.to_hwp_document()`

### `HwpBinaryDocument`

가장 저수준의 HWP binary reader/writer입니다.

앱 코드의 기본 편집 진입점으로는 권장하지 않습니다.  
inspection, reverse engineering, binary-level debugging 용도로 보는 편이 맞습니다.

## 검증과 저장 전략

`HancomDocument.write_to_hwpx()`는 내부적으로 `HwpxDocument.save(validate=True)`를 사용할 수 있으므로, 기본 `.hwpx` 저장에서는 validation을 같이 거는 편이 안전합니다.

예:

```python
from jakal_hwpx import HancomDocument

doc = HancomDocument.read_hwpx("input.hwpx")
doc.append_paragraph("validated append")
doc.write_to_hwpx("build/validated.hwpx", validate=True)
```

`.hwp` 쪽은 현재 기본 저장 경로가 pure Python writer이며, 필요하면 별도 smoke validation 스크립트로 Hancom 비교 검증을 돌릴 수 있습니다.

## 공개 top-level export

`jakal_hwpx` 루트 import에서 `HancomDocument` 중심으로 바로 쓸 수 있는 대표 타입은 다음입니다.

IR 중심:

- `HancomDocument`
- `HancomMetadata`
- `HancomSection`
- `Paragraph`
- `Table`
- `Picture`
- `Hyperlink`
- `Bookmark`
- `Field`
- `AutoNumber`
- `Note`
- `Equation`
- `Shape`
- `Ole`
- `HeaderFooter`
- `StyleDefinition`
- `ParagraphStyle`
- `CharacterStyle`
- `SectionSettings`

HWPX side:

- `HwpxDocument`
- `DocumentMetadata`
- `HeaderFooterXml`
- `TableXml`
- `TableCellXml`
- `PictureXml`
- `NoteXml`
- `BookmarkXml`
- `FieldXml`
- `AutoNumberXml`
- `EquationXml`
- `ShapeXml`
- `OleXml`
- `SectionSettingsXml`
- `StyleDefinitionXml`
- `ParagraphStyleXml`
- `CharacterStyleXml`

HWP side:

- `HwpDocument`
- `HwpBinaryDocument`
- `HwpBinaryFileHeader`
- `HwpDocumentProperties`
- `DocInfoModel`
- `HwpRecord`
- `RecordNode`
- `SectionModel`
- `SectionParagraphModel`

예외와 검증:

- `ValidationIssue`
- `HwpxError`
- `HwpxValidationError`
- `InvalidHwpxFileError`
- `InvalidHwpFileError`
- `HancomInteropError`
- `HwpBinaryEditError`

## 테스트와 검증 도구

로컬 검증:

```bash
python -m pip install -e .[dev]
python -m pytest -q
python scripts/run_stability_lab.py
```

선택적 Hancom smoke validation:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/setup_hancom_security_module.ps1 -DownloadIfMissing
powershell -ExecutionPolicy Bypass -File scripts/run_hancom_smoke_validation.ps1 -InputPath input.hwpx -OutputPath build\hancom-roundtrip.hwpx
```

현재 권장 운영 모델은 다음과 같습니다.

- 편집 중심 객체: `HancomDocument`
- HWPX package 조작: `HwpxDocument`
- native HWP 조작: `HwpDocument`
- binary inspection: `HwpBinaryDocument`
- 기본 변환/저장 경로: pure Python
- Hancom 경로: 선택 검증용
