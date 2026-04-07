# `jakal_hwpx` 모듈 설명

이 문서는 `src/jakal_hwpx/` 아래의 Python 패키지를 설명합니다.

## 시작점

### `HwpxDocument`

`HwpxDocument`는 이 패키지의 기본 진입점입니다.

다음 작업을 맡습니다.

- 기존 `.hwpx` 열기
- 빈 문서 만들기
- 문서 메타데이터 읽기와 수정
- 문단 텍스트 수정
- 문단 추가
- 패키지 파트 추가와 교체
- 구조 및 참조 검증
- 저장과 round-trip 확인

예제:

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("example.hwpx")
doc.replace_text("before", "after")
doc.save("example-edited.hwpx")
```

### `HwpxDocument` API 요약

| 메서드 | 반환값 | 용도 |
| --- | --- | --- |
| `open(path)` | `HwpxDocument` | 기존 HWPX 패키지 열기 |
| `blank()` | `HwpxDocument` | 기본 파트가 들어 있는 새 문서 만들기 |
| `metadata()` | `DocumentMetadata` | 제목, 작성자, 주제, 키워드 같은 메타데이터 읽기 |
| `set_metadata(**values)` | `None` | 메타데이터 수정 |
| `get_document_text()` | `str` | 섹션 전체 본문 텍스트 추출 |
| `set_paragraph_text(section_index, paragraph_index, text)` | 문단 래퍼 | 특정 문단 텍스트 교체 |
| `append_paragraph(text, section_index=0)` | 문단 래퍼 | 섹션 끝에 문단 추가 |
| `replace_text(old, new, count=-1)` | `int` | 문서 전체에서 문자열 치환 |
| `add_section(text=...)` | `SectionPart` | 새 섹션 추가 |
| `section_settings(index)` | `SectionSettings` | 페이지 크기와 여백 설정 접근 |
| `tables()`, `pictures()`, `notes()`, `bookmarks()`, `fields()` | 래퍼 리스트 | 표, 그림, 주석, 북마크, 필드 순회 |
| `styles()`, `paragraph_styles()`, `character_styles()` | 래퍼 리스트 | 스타일 정의 조회 및 수정 |
| `apply_style_to_paragraph(...)` | `None` | 한 문단에 스타일 적용 |
| `apply_style_batch(...)` | `int` | 조건에 맞는 여러 문단에 스타일 적용 |
| `append_bookmark(...)`, `append_hyperlink(...)`, `append_mail_merge_field(...)`, `append_calculation_field(...)`, `append_cross_reference(...)` | 래퍼 객체 | 북마크와 필드 계열 요소 생성 |
| `add_or_replace_binary(...)` | `BinaryDataPart` | 패키지 내부 바이너리 추가 또는 교체 |
| `compile(validate=True)` | `bytes` | 메모리에서 패키지 직렬화 |
| `save(path, validate=True)` | `Path` | 디스크에 저장 |
| `validation_errors()` | `list[str]` | 패키지 구조 검증 |
| `xml_validation_errors()` | `list[str]` | XML 루트 및 구조 점검 |
| `reference_validation_errors()` | `list[str]` | 스타일, 필드, 북마크, manifest 참조 점검 |
| `save_reopen_validation_errors()` | `list[str]` | 저장 후 재오픈 기준의 실용적 검증 |

## 모듈 개요

이 패키지는 크기가 아주 크지 않아서, 대부분의 작업은 `document.py`에서 시작한 뒤 필요할 때 `elements.py`의 래퍼 타입으로 내려가면 됩니다.

| 모듈 | 역할 | 언제 쓰는가 |
| --- | --- | --- |
| `document.py` | 고수준 문서 컨테이너와 편집 API | 거의 모든 작업의 시작점 |
| `elements.py` | 표, 그림, 주석, 필드, 도형, 스타일 같은 요소 래퍼 | 특정 요소를 더 세밀하게 수정할 때 |
| `parts.py` | XML, 텍스트, 바이너리, preview, manifest 파트 모델 | 패키지 내부 파트에 직접 접근할 때 |
| `xmlnode.py` | XML 편의를 위한 작은 래퍼 | 확장 작업이나 내부 동작 이해가 필요할 때 |
| `namespaces.py` | namespace 맵, QName helper, section 매칭 상수 | XPath나 XML 조작을 직접 쓸 때 |
| `exceptions.py` | 패키지 전용 예외 타입 | 잘못된 HWPX 입력을 다룰 때 |

## 패키지 구성

### `document.py`

문서 컨테이너와 편집 흐름을 담당하는 핵심 모듈입니다.

주요 책임:

- HWPX zip 패키지 열기, 생성, 컴파일, 저장
- 고수준 편집 API 제공
- 내부 파트 보존과 갱신
- 구조 검증 수행

자주 쓰는 진입점:

- `HwpxDocument.open()` / `HwpxDocument.blank()`
- `set_metadata()`, `metadata()`, `get_document_text()`
- `set_paragraph_text()`, `append_paragraph()`, `replace_text()`
- `section_settings()`, `tables()`, `pictures()`, `notes()`, `fields()`
- `validation_errors()`, `reference_validation_errors()`, `save_reopen_validation_errors()`
- `compile()` / `save()`

### `elements.py`

문서 안의 고급 요소를 다루는 래퍼가 들어 있습니다.

예:

- 문단 스타일과 글자 스타일
- 표와 셀
- 그림
- 각주와 미주
- 수식
- 북마크와 필드
- 머리말과 꼬리말

주로 보게 되는 메서드:

- `Table.set_cell_text()` / `Table.append_row()`
- `HeaderFooterBlock.set_text()`
- `Field.set_display_text()` / `Field.set_hyperlink_target()`
- `SectionSettings.set_page_size()` / `SectionSettings.set_margins()`
- `CharacterStyle.set_text_color()` / `ParagraphStyle.set_alignment()`

### 자주 보이는 래퍼 타입

| 타입 | 주요 속성/메서드 | 설명 |
| --- | --- | --- |
| `HeaderFooterBlock` | `text`, `set_text()`, `replace_text()` | `headers()`와 `footers()`가 반환 |
| `Table` | `row_count`, `column_count`, `cells()`, `cell()`, `set_cell_text()`, `append_row()`, `merge_cells()` | 표 편집의 시작점 |
| `TableCell` | `row`, `column`, `text`, `row_span`, `col_span`, `set_text()` | 표 셀 래퍼 |
| `Picture` | `binary_item_id`, `shape_comment`, `binary_data()`, `replace_binary()` | 포함된 이미지 데이터 조회 및 교체 |
| `SectionSettings` | `page_width`, `page_height`, `landscape`, `margins()`, `set_page_size()`, `set_margins()` | 섹션별 페이지 레이아웃 |
| `StyleDefinition` | `style_id`, `name`, `set_name()`, `bind_refs()` | 상위 스타일 객체 |
| `ParagraphStyle` | `alignment_horizontal`, `line_spacing`, `set_alignment()`, `set_line_spacing()` | 문단 레벨 서식 |
| `CharacterStyle` | `text_color`, `height`, `set_text_color()`, `set_height()` | 문자 레벨 서식 |
| `Note` | `kind`, `number`, `text`, `set_text()` | 각주와 미주 래퍼 |
| `Bookmark` | `name`, `rename()` | 북마크 수정 |
| `Field` | `field_type`, `field_id`, `parameter_map()`, `set_parameter()`, `set_display_text()` | 하이퍼링크, 메일 머지, 계산식, 상호 참조를 포함 |
| `Equation` | `script`, `shape_comment` | 수식 스크립트 접근 |
| `ShapeObject` | `kind`, `shape_comment`, `text`, `set_text()` | 텍스트가 있는 도형 수정 |

### `parts.py`

패키지 내부 파트를 다루는 저수준 모델이 들어 있습니다.

주요 책임:

- `header.xml`, `section*.xml`, `content.hpf` 같은 파트 표현
- XML, 텍스트, 바이너리, preview 파트 구분
- `content.hpf` 메타데이터 helper 제공

### `xmlnode.py`

다른 모듈에서 공통으로 쓰는 작은 XML 래퍼입니다.

### `namespaces.py`

namespace 상수, QName helper, section 경로 매칭 로직이 들어 있습니다.

### `exceptions.py`

잘못된 패키지나 검증 실패를 위한 커스텀 예외 타입입니다.

## 검증 계층

패키지는 서로 다른 성격의 검증 함수를 제공합니다.

- `validation_errors()`
- `xml_validation_errors()`
- `reference_validation_errors()`
- `save_reopen_validation_errors()`
- `roundtrip_validate()`

실무적으로는 이렇게 보면 됩니다.

- `validation_errors()`는 패키지 구조 문제를 잡습니다.
- `xml_validation_errors()`는 XML 루트와 기본 구조 문제를 잡습니다.
- `reference_validation_errors()`는 스타일, 필드, 북마크, manifest 참조 문제를 잡습니다.
- `save_reopen_validation_errors()`는 저장 후 다시 열리는지 확인하는 현실적인 smoke check입니다.

## 공개 타입

`jakal_hwpx` 루트에서 다시 export하는 타입은 아래와 같습니다.

- `HwpxDocument`
- `DocumentMetadata`
- `HwpxPart`
- `XmlPart`
- `SectionPart`
- `HeaderPart`
- `ContentHpfPart`
- `SettingsPart`
- `VersionPart`
- `MimetypePart`
- `ContainerPart`
- `ContainerRdfPart`
- `ManifestPart`
- `BinaryDataPart`
- `GenericBinaryPart`
- `GenericTextPart`
- `GenericXmlPart`
- `PreviewImagePart`
- `PreviewTextPart`
- `ScriptPart`
- `Picture`
- `Table`
- `TableCell`
- `Note`
- `Equation`
- `Bookmark`
- `Field`
- `AutoNumber`
- `HeaderFooterBlock`
- `SectionSettings`
- `StyleDefinition`
- `ParagraphStyle`
- `CharacterStyle`
- `ShapeObject`
- `HwpxError`
- `HwpxValidationError`
- `InvalidHwpxFileError`

### 예외 타입

| 타입 | 언제 보게 되는가 |
| --- | --- |
| `InvalidHwpxFileError` | 입력 파일이 zip 기반 HWPX가 아닐 때 |
| `HwpxValidationError` | 열기, 컴파일, 저장 중 검증에 실패했을 때 |
| `HwpxError` | 패키지 전용 오류의 상위 타입 |

## round-trip과 편집 모델

### 안전한 편집 모델

이 패키지는 가능한 한 원래 패키지 구조를 보존하는 방향으로 설계되어 있습니다.

- 기존 파트를 메모리에 그대로 올립니다.
- 알 수 없는 파트도 타입 추론을 통해 보존합니다.
- 가능하면 zip entry 메타데이터도 유지합니다.
- 저장 후 다시 열어서 검증할 수 있습니다.

### 배포 보호(distribution protected) 문서

`content.hpf`가 배포 보호 상태를 나타내면 편집 가능한 section XML이 없을 수 있습니다.

이 경우:

- 문서는 열 수 있습니다.
- 구조 검증은 동작합니다.
- 고수준 편집 API는 사용할 수 없을 수 있습니다.

### 왜 element wrapper를 쓰는가

HWPX 편집은 보통 여러 층을 함께 만집니다.

- 문서 메타데이터
- 패키지 manifest
- section XML
- 스타일 테이블
- 포함된 바이너리

`elements.py`의 래퍼는 이런 세부 XML 처리와 XPath 반복을 호출부 밖으로 밀어내고, 문서 도메인에 맞는 API를 제공합니다.

## 자주 쓰는 패턴

### 문서 열기와 검증

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("example.hwpx")

print(doc.metadata())
print(doc.validation_errors())
print(doc.reference_validation_errors())
```

### 텍스트 수정과 저장

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("example.hwpx")
doc.replace_text("draft", "final")
doc.save("example-final.hwpx")
```

### 빈 문서 생성

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.blank()
doc.set_metadata(title="Generated")
doc.set_paragraph_text(0, 0, "Hello")
doc.save("build/blank.hwpx")
```

### 페이지 설정과 스타일 수정

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("example.hwpx")

settings = doc.section_settings(0)
settings.set_page_size(width=60000, height=85000)
settings.set_margins(left=7000, right=7000)

style = doc.styles()[0]
para_style = doc.paragraph_styles()[0]
char_style = doc.character_styles()[0]

style.set_name("Body Center")
para_style.set_alignment(horizontal="CENTER")
char_style.set_text_color("#112233")

doc.apply_style_to_paragraph(
    0,
    0,
    style_id=style.style_id,
    para_pr_id=para_style.style_id,
    char_pr_id=char_style.style_id,
)
doc.save("build/styled.hwpx")
```

### 표와 머리말/꼬리말 수정

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("example.hwpx")

if doc.headers():
    doc.headers()[0].set_text("Edited header")
if doc.footers():
    doc.footers()[0].set_text("Edited footer")

table = doc.tables()[0]
table.set_cell_text(0, 0, "Updated value")
table.append_row()[0].set_text("Appended row")

doc.save("build/structured-edit.hwpx")
```

### 북마크와 동적 필드 생성

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.blank()

bookmark = doc.append_bookmark("summary_anchor")
doc.append_hyperlink("https://example.com", display_text="Open Example")
doc.append_mail_merge_field("customer_name", display_text="CUSTOMER_NAME")
doc.append_calculation_field("40+2", display_text="42")
doc.append_cross_reference(bookmark.name or "summary_anchor", display_text="Go to summary")

assert doc.reference_validation_errors() == []
doc.save("build/fields.hwpx")
```

### 바이너리 추가

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.blank()
doc.add_or_replace_binary(
    "assets/custom.bin",
    b"abc123",
    media_type="application/octet-stream",
    manifest_id="custom_asset",
)

assert doc.validation_errors() == []
doc.save("build/with-binary.hwpx")
```

## 어떤 수준의 API를 쓸지 고르기

가능하면 가장 높은 수준의 API부터 쓰는 편이 낫습니다.

- 문서를 열고, 수정하고, 검증하고, 저장하는 작업은 `HwpxDocument`에서 시작하세요.
- 표, 필드, 주석, 그림, 스타일 같은 특정 요소를 더 세밀하게 수정해야 하면 `elements.py` 래퍼를 쓰세요.
- preview 파트, raw binary, custom package inspection처럼 패키지 내부 구조에 직접 접근해야 할 때만 `parts.py`로 내려가면 됩니다.

## 저장소 내 관련 자산

이 문서는 Python 패키지 자체를 다룹니다. 저장소에는 이 외에도 아래 항목이 있습니다.

- `examples/samples/`
- `examples/output_smoke/`
- `examples/output/`
- `tools/`

`tools/` 아래의 Java 기반 `.hwp -> .hwpx` 보조 도구는 importable Python 패키지의 일부가 아닙니다.
