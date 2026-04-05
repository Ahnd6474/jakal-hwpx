# jakal-hwpx

Python에서 HWPX 문서를 읽고, 수정하고, 검증하고, 다시 저장하기 위한 라이브러리입니다.

English version: [README.md](./README.md)

`jakal_hwpx` 패키지는 `src/jakal_hwpx` 아래에 있습니다. 이 문서는 실제로 import 해서 쓰는 라이브러리 기준으로 설명합니다. 저장소 안에는 예제와 검증 코드도 있지만, 그중 일부는 일반적인 체크아웃에 포함되지 않는 로컬 HWPX 코퍼스나 Windows 전용 도구를 전제로 하므로 여기서는 선택적인 유지보수 워크플로로만 다룹니다.

## 목차

- [설치](#설치)
- [빠른 시작](#빠른-시작)
- [jakal-hwpx는 무엇인가](#jakal-hwpx는-무엇인가)
- [왜 쓰는가](#왜-쓰는가)
- [API](#api)
  - [`HwpxDocument`](#hwpxdocument)
  - [`DocumentMetadata`](#documentmetadata)
  - [요소 래퍼](#요소-래퍼)
  - [파트 클래스](#파트-클래스)
  - [`HwpxXmlNode`](#hwpxxmlnode)
  - [예외](#예외)
- [예제](#예제)
- [유지보수용 워크플로](#유지보수용-워크플로)
- [추가 문서](#추가-문서)
- [라이선스](#라이선스)

## 설치

Python 3.11 이상이 필요합니다.

로컬 체크아웃에서 설치:

```bash
python -m pip install --upgrade pip
python -m pip install -e .
```

import 경로는 `jakal_hwpx`이고, `pyproject.toml`의 프로젝트 이름은 `jakal-hwpx`입니다.

테스트까지 실행하려면:

```bash
python -m pip install pytest
```

## 빠른 시작

```python
from pathlib import Path

from jakal_hwpx import HwpxDocument

source = Path("example.hwpx")
target = Path("example-edited.hwpx")

doc = HwpxDocument.open(source)
doc.set_metadata(title="Edited title", creator="jakal-hwpx")
doc.replace_text("old text", "new text", count=1)
doc.append_paragraph("Added from Python.", section_index=0)

if doc.headers():
    doc.headers()[0].set_text("Edited header")

doc.save(target)
print(doc.validation_errors())
# []
```

## jakal-hwpx는 무엇인가

`jakal_hwpx`는 zip 기반 HWPX 패키지를 다루는 Python 라이브러리입니다. 문서를 열고, 자주 수정하는 구조를 타입이 있는 래퍼로 노출하고, 변경하지 않은 파트는 원래 바이트를 최대한 유지한 채 다시 저장합니다.

메타데이터, 문단, 스타일, 페이지 설정, 표, 이미지, 머리글/바닥글, 주석, 북마크, 필드, 수식, 도형, 바이너리 자산, 저수준 XML 접근까지 다룹니다.

## 왜 쓰는가

HWPX는 XML 파일 하나가 아닙니다. 여러 XML 파트, manifest 항목, 바이너리 자산, 참조 정보, 레이아웃 데이터가 묶인 패키지입니다. 단순 unzip-edit-rezip 방식은 이런 연결을 쉽게 깨뜨립니다.

`jakal_hwpx`는 그 부분을 덜 위험하게 만들려는 라이브러리입니다.

- 손대지 않은 파트는 가능한 한 그대로 유지합니다.
- 문단 텍스트, 머리글/바닥글, 표, 필드, 스타일 같은 자주 쓰는 수정 경로를 타입 기반 API로 제공합니다.
- 저장 전에 패키지 구조와 파트 간 참조를 검증합니다.
- 고수준 래퍼가 아직 다루지 않는 경우에는 `HwpxXmlNode`로 직접 XML을 만질 수 있습니다.
- 배포 보호된 문서는 고수준 편집이 제한되더라도 보존과 재저장은 가능합니다.

그냥 XML만 순회하면 된다면 `lxml`만으로도 충분할 수 있습니다. 기존 HWPX 문서를 코드로 안전하게 수정해야 한다면 이 라이브러리가 더 나은 출발점입니다.

## API

### `HwpxDocument`

메인 진입점입니다.

#### 생성

| 시그니처 | 반환값 | 설명 |
|----------|--------|------|
| `HwpxDocument.open(path)` | `HwpxDocument` | 디스크의 `.hwpx` 파일을 엽니다. |
| `HwpxDocument.from_bytes(raw_bytes)` | `HwpxDocument` | 메모리의 바이트로부터 패키지를 엽니다. |

경로가 없으면 `HwpxDocument.open()`은 `FileNotFoundError`를 발생시킵니다.

#### 기본 프로퍼티와 파트 접근

| 멤버 | 반환값 | 설명 |
|------|--------|------|
| `mimetype` | `MimetypePart` | 원본 `mimetype` 엔트리 |
| `content_hpf` | `ContentHpfPart` | 메타데이터와 manifest 헬퍼를 포함한 `Contents/content.hpf` |
| `header` | `HeaderPart` | `Contents/header.xml` |
| `container` | `ContainerPart` | `META-INF/container.xml` |
| `preview_text` | `PreviewTextPart \| None` | `Preview/PrvText.txt`가 있으면 반환 |
| `is_distribution_protected` | `bool` | 보호된 암호화 파트를 쓰는 문서인지 여부 |
| `sections` | `list[SectionPart]` | 정렬된 section XML 파트 목록 |
| `get_part(path, expected_type=None)` | `HwpxPart` | 경로로 임의 파트 조회 |
| `list_part_paths()` | `list[str]` | 저장 시 사용하는 현재 파트 순서 |
| `add_part(path, raw_bytes)` | `HwpxPart` | 파트를 추가하거나 교체 |
| `remove_part(path)` | `None` | 파트를 제거 |

#### 기능 접근자

| 메서드 | 반환값 | 설명 |
|--------|--------|------|
| `metadata()` | `DocumentMetadata` | 문서 메타데이터 조회 |
| `headers(section_index=None)` | `list[HeaderFooterBlock]` | 문서 전체 또는 특정 섹션의 머리글 |
| `footers(section_index=None)` | `list[HeaderFooterBlock]` | 문서 전체 또는 특정 섹션의 바닥글 |
| `tables(section_index=None)` | `list[Table]` | 표 목록 |
| `pictures(section_index=None)` | `list[Picture]` | 그림 객체 목록 |
| `section_settings(section_index=0)` | `SectionSettings` | 페이지 크기와 여백 설정 |
| `notes(section_index=None)` | `list[Note]` | 각주와 미주 |
| `bookmarks(section_index=None)` | `list[Bookmark]` | 북마크 컨트롤 |
| `fields(section_index=None)` | `list[Field]` | 모든 필드 컨트롤 |
| `hyperlinks(section_index=None)` | `list[Field]` | 하이퍼링크 필드만 조회 |
| `mail_merge_fields(section_index=None)` | `list[Field]` | 메일 머지 필드만 조회 |
| `calculation_fields(section_index=None)` | `list[Field]` | 계산/수식 필드만 조회 |
| `cross_references(section_index=None)` | `list[Field]` | 상호 참조 필드만 조회 |
| `auto_numbers(section_index=None)` | `list[AutoNumber]` | 자동 번호 컨트롤 |
| `equations(section_index=None)` | `list[Equation]` | 수식 객체 |
| `shapes(section_index=None)` | `list[ShapeObject]` | 텍스트아트를 포함한 도형 객체 |
| `styles()` | `list[StyleDefinition]` | `header.xml`의 스타일 목록 |
| `paragraph_styles()` | `list[ParagraphStyle]` | 문단 스타일 레코드 |
| `character_styles()` | `list[CharacterStyle]` | 글자 스타일 레코드 |
| `get_style(style_id)` | `StyleDefinition` | style id로 스타일 조회 |
| `get_paragraph_style(style_id)` | `ParagraphStyle` | paragraph style id 조회 |
| `get_character_style(style_id)` | `CharacterStyle` | character style id 조회 |

#### 편집과 저장

| 메서드 | 반환값 | 설명 |
|--------|--------|------|
| `set_metadata(**values)` | `None` | 제목, 언어, 작성자, 설명, 날짜, 키워드 등 메타데이터 수정 |
| `get_document_text(section_separator="\n\n")` | `str` | 섹션 텍스트를 펼쳐서 반환 |
| `replace_text(old, new, count=-1, include_header=True)` | `int` | 편집 가능한 XML 파트 전반에서 텍스트 치환 |
| `append_paragraph(text, section_index=0, ...)` | `HwpxXmlNode` | 문단 추가 |
| `insert_paragraph(section_index, paragraph_index, text, ...)` | `HwpxXmlNode` | 특정 위치에 문단 삽입 |
| `set_paragraph_text(section_index, paragraph_index, text, ...)` | `HwpxXmlNode` | 문단 텍스트 교체 |
| `delete_paragraph(section_index, paragraph_index)` | `None` | 문단 삭제 |
| `apply_style_to_paragraph(section_index, paragraph_index, ...)` | `None` | 문단 하나에 스타일 id 적용 |
| `apply_style_batch(section_index=..., text_contains=..., regex=..., ...)` | `int` | 조건에 맞는 문단들에 스타일 id 일괄 적용 |
| `add_section(clone_from=0, text=None)` | `SectionPart` | 새 섹션 추가 |
| `remove_section(section_index)` | `None` | 섹션 삭제 후 메타데이터 갱신 |
| `set_preview_text(text)` | `PreviewTextPart` | 미리보기 텍스트 생성 또는 교체 |
| `add_or_replace_binary(name, data, media_type=None, manifest_id=None)` | `BinaryDataPart` | 바이너리 저장 및 manifest 갱신 |
| `append_bookmark(name, ...)` | `Bookmark` | 북마크 추가 |
| `append_field(field_type, ...)` | `Field` | 일반 필드 추가 |
| `append_hyperlink(target, display_text, ...)` | `Field` | 하이퍼링크 필드 추가 |
| `append_mail_merge_field(name, display_text, ...)` | `Field` | 메일 머지 필드 추가 |
| `append_calculation_field(expression, display_text, ...)` | `Field` | 계산 필드 추가 |
| `append_cross_reference(bookmark_name, display_text, ...)` | `Field` | 상호 참조 필드 추가 |
| `compile(validate=True)` | `bytes` | 메모리의 문서를 `.hwpx` 바이트로 컴파일 |
| `save(path, validate=True)` | `Path` | 디스크에 저장 |

#### 검증

| 메서드 | 반환값 | 설명 |
|--------|--------|------|
| `validation_errors()` | `list[str]` | 필수 파트, manifest, 중복 zip 경로, section count 등 구조 검사 |
| `validate()` | `None` | 오류가 있으면 `HwpxValidationError` 발생 |
| `roundtrip_validate()` | `None` | 임시 저장 후 다시 열어 검증 |
| `xml_validation_errors()` | `list[str]` | XML 루트 이름, section/table 기본 무결성 검사 |
| `schema_validation_errors(schema_map)` | `list[str]` | 특정 XML 파트에 대한 선택적 XSD 검증 |
| `reference_validation_errors()` | `list[str]` | 스타일, manifest, 북마크, 필드 참조 검사 |
| `save_reopen_validation_errors()` | `list[str]` | 저장 후 재오픈했을 때의 구조 오류 반환 |
`is_distribution_protected`가 `True`인 문서는 암호화된 section 데이터를 고수준 편집 API로 수정할 수 없도록 막아 둡니다.

### `DocumentMetadata`

`HwpxDocument.metadata()`가 반환하는 간단한 dataclass입니다.

```python
DocumentMetadata(
    title: str | None = None,
    language: str | None = None,
    creator: str | None = None,
    subject: str | None = None,
    description: str | None = None,
    lastsaveby: str | None = None,
    created: str | None = None,
    modified: str | None = None,
    date: str | None = None,
    keyword: str | None = None,
    extra: dict[str, str] = {},
)
```

### 요소 래퍼

자주 수정하는 구조는 아래 래퍼들로 다룹니다.

| 클래스 | 용도 | 주요 멤버 |
|--------|------|-----------|
| `HeaderFooterBlock` | 머리글/바닥글 텍스트 | `kind`, `text`, `replace_text()`, `set_text()` |
| `TableCell` | 개별 셀 | `row`, `column`, `text`, `row_span`, `col_span`, `set_text()` |
| `Table` | 표 편집 | `row_count`, `column_count`, `cells()`, `rows()`, `cell()`, `set_cell_text()`, `append_row()`, `merge_cells()` |
| `Picture` | 이미지 메타데이터와 바이너리 연결 | `binary_item_id`, `shape_comment`, `binary_part_path()`, `binary_data()`, `replace_binary()`, `bind_binary_item()` |
| `StyleDefinition` | 문서 스타일 | `style_id`, `name`, `english_name`, `para_pr_id`, `char_pr_id`, `set_name()`, `set_english_name()`, `bind_refs()` |
| `ParagraphStyle` | 문단 서식 | `style_id`, `alignment_horizontal`, `line_spacing`, `set_alignment()`, `set_line_spacing()` |
| `CharacterStyle` | 글자 서식 | `style_id`, `text_color`, `height`, `set_text_color()`, `set_height()` |
| `SectionSettings` | 페이지 설정 | `page_width`, `page_height`, `landscape`, `margins()`, `set_page_size()`, `set_margins()` |
| `Note` | 각주/미주 | `kind`, `number`, `text`, `set_text()` |
| `Bookmark` | 북마크 | `name`, `rename()` |
| `Field` | 하이퍼링크, 메일 머지, 수식, 상호 참조 | `field_type`, `field_id`, `control_id`, `name`, `parameter_map()`, `get_parameter()`, `set_parameter()`, `display_text`, `set_display_text()`, `set_hyperlink_target()`, `configure_mail_merge()`, `configure_calculation()`, `configure_cross_reference()` |
| `AutoNumber` | 자동 번호 | `kind`, `number`, `number_type`, `set_number()`, `set_number_type()` |
| `Equation` | 수식 편집 | `script`, `shape_comment` |
| `ShapeObject` | 도형 주석과 textart 텍스트 | `kind`, `shape_comment`, `text`, `set_text()` |

### 파트 클래스

더 낮은 수준의 파트 모델도 export 됩니다.

| 클래스 | 의미 |
|--------|------|
| `HwpxPart` | 모든 패키지 파트의 베이스 클래스 |
| `GenericBinaryPart` | 임의 바이너리 파트 |
| `MimetypePart` | `mimetype` |
| `GenericTextPart` | 임의 UTF-8 텍스트 파트 |
| `PreviewTextPart` | `Preview/PrvText.txt` |
| `PreviewImagePart` | 미리보기 이미지 |
| `BinaryDataPart` | `BinData/*` |
| `ScriptPart` | 스크립트 텍스트 파트 |
| `XmlPart` | XML 기반 파트의 베이스 클래스 |
| `GenericXmlPart` | 임의 XML 파트 |
| `VersionPart` | `version.xml` |
| `ContainerPart` | `META-INF/container.xml` |
| `ManifestPart` | `META-INF/manifest.xml` |
| `ContainerRdfPart` | `META-INF/container.rdf` |
| `ContentHpfPart` | `Contents/content.hpf` |
| `HeaderPart` | `Contents/header.xml` |
| `SettingsPart` | `settings.xml` |
| `SectionPart` | `Contents/sectionN.xml` |

### `HwpxXmlNode`

`HwpxXmlNode`는 네임스페이스를 인식하는 저수준 XML 편집용 탈출구입니다.

고수준 래퍼가 아직 다루지 않는 구조를 손봐야 할 때 사용하면 됩니다.

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("example.hwpx")
section = doc.sections[0]
paragraph = section.paragraphs()[0]

for node in paragraph.xpath(".//hp:t"):
    if node.text:
        node.text = node.text.replace("before", "after")

doc.save("example-edited.hwpx")
```

### 예외

| 예외 | 발생 조건 |
|------|-----------|
| `HwpxError` | 패키지 기본 예외 |
| `InvalidHwpxFileError` | 입력이 존재하지만 유효한 zip 기반 HWPX 패키지가 아닐 때 |
| `HwpxValidationError` | `validate()`에서 구조 오류가 발견될 때 |

## 예제

### 기능 편집 예제

```python
doc = HwpxDocument.open("example.hwpx")

doc.headers()[0].set_text("Edited header")
doc.footers()[0].set_text("Edited footer")

section = doc.section_settings(0)
section.set_page_size(width=60000, height=84000)
section.set_margins(left=8500, right=8500)

table = doc.tables()[0]
table.set_cell_text(0, 0, "Edited cell")
table.append_row()

doc.append_bookmark("python_anchor")
doc.append_hyperlink("https://example.com", display_text="Example")
doc.append_mail_merge_field("customer_name", display_text="CUSTOMER_NAME")
doc.append_calculation_field("40+2", display_text="42")
doc.append_cross_reference("python_anchor", display_text="Anchor Ref")

doc.save("example-edited.hwpx")
```

### 검증 예제

```python
doc = HwpxDocument.open("example.hwpx")

print(doc.xml_validation_errors())
print(doc.reference_validation_errors())
print(doc.save_reopen_validation_errors())
```

## 유지보수용 워크플로

저장소에는 더 넓은 검증과 showcase 스크립트도 들어 있지만, 라이브러리 사용 자체에 꼭 필요하지는 않습니다.

현재 제약은 한 가지입니다.

- 대부분의 저장소 테스트는 패키지 소스와 함께 커밋되지 않는 로컬 HWPX 코퍼스를 전제로 합니다.

이 환경을 이미 갖춘 유지보수자라면 showcase 스크립트에 경로를 직접 넘겨서 사용할 수 있습니다.

```bash
python examples/build_showcase_bundle.py --corpus-dir <path-to-hwpx-corpus> --output-dir <path-to-output>
```

테스트 실행 전에는 `tests/conftest.py`를 먼저 확인하세요. 현재 테스트 스위트는 유지보수자용 로컬 코퍼스 배치를 전제로 합니다.

## 추가 문서

- [`HWPX_MODULE.md`](./HWPX_MODULE.md)
- [`examples/SHOWCASE.md`](./examples/SHOWCASE.md)

## 라이선스

이 저장소에는 아직 최상위 라이선스 파일이 없습니다. 배포하거나 재사용하려면 먼저 라이선스를 추가하세요.
