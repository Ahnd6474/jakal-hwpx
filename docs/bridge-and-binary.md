# Bridge와 binary API

이 문서는 `HwpHwpxBridge`와 `HwpBinaryDocument`를 설명합니다. 둘 다 일반적인 문서 생성 코드의 첫 진입점은 아닙니다. 변환 경로를 명시하거나, HWP binary 구조를 조사할 때 사용합니다.

## `HwpHwpxBridge`

`HwpHwpxBridge`는 HWP, HWPX, `HancomDocument` 사이의 materialization 경로를 한 객체 안에서 관리합니다.

```python
from jakal_hwpx import HwpHwpxBridge
```

### Constructors

| API | 설명 |
|---|---|
| `HwpHwpxBridge.open(path, *, converter=None)` | 확장자와 내용 sniffing으로 HWP/HWPX를 열고 bridge를 만듭니다. |
| `HwpHwpxBridge.from_hwp(source, *, converter=None)` | path 또는 `HwpDocument`에서 bridge를 만듭니다. |
| `HwpHwpxBridge.from_hwpx(source, *, converter=None)` | path 또는 `HwpxDocument`에서 bridge를 만듭니다. |
| `HwpHwpxBridge.from_hancom(source, *, converter=None)` | `HancomDocument`에서 bridge를 만듭니다. |

### Materialization

| API | 설명 |
|---|---|
| `hancom_document(force_refresh=False)` | `HancomDocument`를 반환합니다. |
| `hwp_document(force_refresh=False)` | `HwpDocument`를 반환합니다. |
| `hwpx_document(force_refresh=False)` | `HwpxDocument`를 반환합니다. |
| `refresh_hancom()` | 공통 모델을 다시 materialize합니다. |
| `refresh_hwp()` | HWP 문서를 다시 materialize합니다. |
| `refresh_hwpx()` | HWPX 문서를 다시 materialize합니다. |

### Saving

| API | 설명 |
|---|---|
| `save_hwp(path, *, force_refresh=False)` | HWP로 저장합니다. |
| `save_hwpx(path, *, force_refresh=False)` | HWPX로 저장합니다. |
| `save(path, *, force_refresh=False)` | 확장자에 맞춰 저장합니다. |

### Example

```python
from jakal_hwpx import HwpHwpxBridge

bridge = HwpHwpxBridge.open("input.hwp")
doc = bridge.hancom_document()
doc.append_paragraph("추가 문단")

bridge.save_hwpx("build/output.hwpx", force_refresh=True)
bridge.save_hwp("build/output.hwp", force_refresh=True)
```

## `HwpBinaryDocument`

`HwpBinaryDocument`는 HWP OLE stream과 binary record를 직접 다룹니다. 문서 손상 원인을 찾거나, 아직 고수준 API로 매핑되지 않은 control을 분석할 때 사용합니다.

```python
from jakal_hwpx import HwpBinaryDocument
```

### 열기와 저장

| API | 설명 |
|---|---|
| `HwpBinaryDocument.open(path)` | HWP 파일을 엽니다. |
| `save(path=None, *, preserve_original_bytes=True)` | HWP를 저장합니다. |
| `save_copy(path, *, preserve_original_bytes=True)` | 다른 경로에 사본을 저장합니다. |

### Streams

| API | 설명 |
|---|---|
| `list_stream_paths()` | OLE stream path 목록을 반환합니다. |
| `section_stream_paths()` | body section stream path 목록을 반환합니다. |
| `bindata_stream_paths()` | BinData stream path 목록을 반환합니다. |
| `has_stream(path)` | stream 존재 여부를 반환합니다. |
| `stream_size(path)` | stream 크기를 반환합니다. |
| `read_stream(path, *, decompress=None)` | stream bytes를 읽습니다. |
| `write_stream(path, data, *, compress=None)` | stream bytes를 씁니다. |
| `add_stream(path, data)` | stream을 추가합니다. |
| `remove_stream(path)` | stream을 제거합니다. |
| `stream_capacity(path, data, *, compress=None)` | data를 썼을 때 capacity 정보를 계산합니다. |

### DocInfo와 section

| API | 설명 |
|---|---|
| `file_header()` | `HwpBinaryFileHeader`를 반환합니다. |
| `document_properties()` | `HwpDocumentProperties`를 반환합니다. |
| `docinfo_records()` | raw DocInfo record 목록을 반환합니다. |
| `docinfo_model()` | `DocInfoModel`을 반환합니다. |
| `replace_docinfo_model(model)` | DocInfo model을 교체합니다. |
| `section_records(section_index)` | section raw record 목록을 반환합니다. |
| `replace_section_records(section_index, records)` | section record를 교체합니다. |
| `section_model(section_index=0)` | `SectionModel`을 반환합니다. |
| `replace_section_model(section_index, model)` | section model을 교체합니다. |
| `ensure_section_count(section_count)` | section count를 보장합니다. |
| `reset_body_sections_to_blank()` | 본문 section을 빈 상태로 초기화합니다. |

### Section settings

| API | 설명 |
|---|---|
| `section_page_settings(section_index)` | 용지 설정을 읽습니다. |
| `set_section_page_settings(section_index, *, page_width=None, page_height=None, landscape=None, margins=None)` | 용지 설정을 씁니다. |
| `section_page_border_fills(section_index)` | page border fill을 읽습니다. |
| `set_section_page_border_fills(section_index, page_border_fills)` | page border fill을 씁니다. |
| `section_definition_settings(section_index)` | section definition 설정을 읽습니다. |
| `set_section_definition_settings(section_index, *, visibility=None, grid=None, start_numbers=None, numbering_shape_id=None, memo_shape_id=None)` | section definition 설정을 씁니다. |
| `section_page_numbers(section_index)` | page number control을 읽습니다. |
| `set_section_page_numbers(section_index, page_numbers)` | page number control을 씁니다. |
| `section_note_settings(section_index)` | 각주/미주 설정을 읽습니다. |
| `set_section_note_settings(section_index, *, footnote_pr=None, endnote_pr=None)` | 각주/미주 설정을 씁니다. |

### Text와 append helper

| API | 설명 |
|---|---|
| `paragraphs(section_index=0)` | binary paragraph wrapper 목록을 반환합니다. |
| `get_document_text()` | 본문 텍스트를 추출합니다. |
| `set_paragraph_text_same_length(section_index, paragraph_index, new_text)` | 같은 길이 문단 치환을 수행합니다. |
| `replace_paragraph_text_same_length(section_index, paragraph_index, old, new, *, count=-1)` | 특정 문단에서 같은 길이 치환을 수행합니다. |
| `replace_text_same_length(old, new, *, section_index=None, count=-1)` | 문서 또는 섹션 단위 같은 길이 치환을 수행합니다. |
| `append_paragraph(text, *, section_index=0, ...)` | binary section model에 문단을 추가합니다. |
| `append_table(...)`, `append_picture(...)`, `append_hyperlink(...)`, `append_field(...)` | control append helper입니다. |
| `append_header(...)`, `append_footer(...)`, `append_note(...)`, `append_bookmark(...)` | section/control append helper입니다. |
| `append_equation(...)`, `append_shape(...)`, `append_ole(...)` | graphics/control append helper입니다. |

### BinData

| API | 설명 |
|---|---|
| `add_embedded_bindata(data, *, extension, storage_id=None, flags=1)` | embedded BinData를 추가하고 `(storage_id, stream_path)`를 반환합니다. |
| `remove_embedded_bindata(storage_id)` | embedded BinData를 제거합니다. |

### 예제: 문서 record 조사

```python
from jakal_hwpx import HwpBinaryDocument

doc = HwpBinaryDocument.open("input.hwp")
print(doc.file_header().version)
print(doc.list_stream_paths())
print(doc.docinfo_model().id_mappings_record().named_counts())

section = doc.section_model(0)
for paragraph in section.paragraphs():
    print(paragraph.text())
```

### 예제: 안전한 no-op copy

```python
from jakal_hwpx import HwpBinaryDocument

doc = HwpBinaryDocument.open("input.hwp")
doc.save_copy("build/copy.hwp", preserve_original_bytes=True)
```

## Binary model class

| 클래스 | 설명 |
|---|---|
| `HwpRecord` | HWP binary record 단위 |
| `RecordNode` | record tree node |
| `TypedRecord` | tag별 typed record base |
| `DocInfoModel` | DocInfo record 집합 |
| `SectionModel` | section record 집합 |
| `SectionParagraphModel` | section paragraph model |
| `HwpBinaryFileHeader` | HWP file header |
| `HwpDocumentProperties` | 문서 속성 |
| `HwpStreamCapacity` | stream 재인코드 capacity 정보 |

## 참고

- `HwpBinaryDocument`는 강력하지만, 잘못된 record 조합은 한컴에서 열리지 않는 파일을 만들 수 있습니다.
- 일반 HWP 편집은 [`HwpDocument`](./hwp-document.md)를 먼저 사용하세요.
- HWP/HWPX를 같은 코드로 처리하는 앱은 [`HancomDocument`](./hancom-document.md)를 먼저 사용하세요.
