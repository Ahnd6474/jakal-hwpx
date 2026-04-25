# `jakal_hwpx` 모듈 문서

`jakal_hwpx`는 HWP/HWPX 문서를 Python 객체로 읽고, 편집하고, 다시 저장하기 위한 공개 모듈입니다. 이 문서는 `jakal_hwpx`의 public API 표면을 정리하는 인덱스입니다. 클래스별 세부 사용법은 `docs/` 아래 문서에서 관리합니다.

![jakal_hwpx의 문서 레이어](./docs/images/module-layers.svg)

## 문서 구성

| 문서 | 내용 |
|---|---|
| [docs/README.md](./docs/README.md) | `docs/` 폴더의 문서 목록과 읽는 순서 |
| [docs/hancom-document.md](./docs/hancom-document.md) | `HancomDocument`와 공통 문서 모델 |
| [docs/hwpx-document.md](./docs/hwpx-document.md) | `HwpxDocument`, HWPX package/XML API |
| [docs/hwp-document.md](./docs/hwp-document.md) | `HwpDocument`, HWP 전용 객체 API |
| [docs/bridge-and-binary.md](./docs/bridge-and-binary.md) | `HwpHwpxBridge`, `HwpBinaryDocument`, 저수준 조사 API |
| [STABILITY_CONTRACT.md](./STABILITY_CONTRACT.md) | 지원 범위, release gate, 안정성 기준 |

## 공개 import 규칙

공개 API는 루트 패키지에서 import하는 형태를 기준으로 문서화합니다.

```python
from jakal_hwpx import HancomDocument, HwpxDocument, HwpDocument
```

아래처럼 내부 모듈 경로를 직접 import하는 코드는 구현 세부사항에 묶일 수 있습니다.

```python
from jakal_hwpx.hancom_document import HancomDocument
from jakal_hwpx.document import HwpxDocument
```

이 문서에 이름이 있더라도, 안정적인 진입점은 `jakal_hwpx` 루트 import입니다.

## API 레이어

| 레이어 | 공개 클래스 | 역할 |
|---|---|---|
| 공통 문서 모델 | `HancomDocument` | HWP/HWPX를 같은 객체 모델로 읽고 편집하고 저장 |
| HWPX 포맷 모델 | `HwpxDocument` | HWPX package, XML part, HWPX 검증과 직접 편집 |
| HWP 포맷 모델 | `HwpDocument` | HWP binary 문서의 객체 API와 직접 저장 |
| 포맷 전환 도우미 | `HwpHwpxBridge` | HWP/HWPX/HancomDocument materialization 경로 관리 |
| HWP binary 모델 | `HwpBinaryDocument` | stream, record tree, DocInfo, SectionModel 조사와 재인코드 |

일반 애플리케이션 코드는 `HancomDocument`부터 시작하는 것이 기본입니다. 포맷별 내부 구조를 직접 제어해야 할 때만 `HwpxDocument` 또는 `HwpDocument`로 내려갑니다.

## 공개 API 맵

### 주요 편집 클래스

| 이름 | 설명 | 세부 문서 |
|---|---|---|
| `HancomDocument` | 공통 문서 편집 모델 | [HancomDocument](./docs/hancom-document.md) |
| `HwpxDocument` | HWPX 직접 편집 모델 | [HwpxDocument](./docs/hwpx-document.md) |
| `HwpDocument` | HWP 직접 편집 모델 | [HwpDocument](./docs/hwp-document.md) |
| `HwpHwpxBridge` | 포맷 전환 도우미 | [Bridge and binary](./docs/bridge-and-binary.md) |
| `HwpBinaryDocument` | HWP binary 저수준 모델 | [Bridge and binary](./docs/bridge-and-binary.md) |

### 공통 문서 모델

`HancomDocument`는 문서를 dataclass 기반 block으로 표현합니다.

| 이름 | 역할 |
|---|---|
| `HancomMetadata` | 제목, 작성자, 주제, 날짜 같은 문서 metadata |
| `HancomSection` | 섹션 설정, 머리말/꼬리말, 본문 block을 묶는 단위 |
| `SectionSettings` | 용지, 여백, visibility, numbering, note 설정 |
| `Paragraph`, `Table`, `Picture`, `Shape`, `Equation`, `Ole` | 본문 block |
| `Field`, `Hyperlink`, `Bookmark`, `AutoNumber`, `Note` | 문서 참조와 control block |
| `HeaderFooter` | 머리말/꼬리말 block |
| `StyleDefinition`, `ParagraphStyle`, `CharacterStyle` | 스타일 정의 |
| `NumberingDefinition`, `BulletDefinition`, `MemoShapeDefinition` | 번호, bullet, memo shape 정의 |
| `Memo`, `Form`, `Chart` | memo/comment, form object, chart 표현 |

필드 목록과 예제는 [docs/hancom-document.md](./docs/hancom-document.md)에 있습니다.

### HWPX XML wrapper

`HwpxDocument`의 selector와 append API는 XML wrapper를 반환합니다.

| 이름 | 대표 용도 |
|---|---|
| `ParagraphStyleXml`, `CharacterStyleXml`, `StyleDefinitionXml` | HWPX 스타일 XML |
| `HeaderFooterXml`, `SectionSettingsXml` | 섹션 관련 XML |
| `TableXml`, `TableCellXml`, `PictureXml`, `ShapeXml`, `EquationXml`, `OleXml` | 본문 control XML |
| `FieldXml`, `BookmarkXml`, `AutoNumberXml`, `NoteXml`, `MemoXml`, `FormXml`, `ChartXml` | 문서 control XML |
| `HwpxXmlNode` | 임의 XML 노드 wrapper |
| `HwpxPart`, `XmlPart`, `GenericXmlPart`, `GenericBinaryPart` 등 | package part wrapper |

### HWP object wrapper

`HwpDocument`는 HWP native control을 객체 wrapper로 노출합니다.

| 이름 | 대표 용도 |
|---|---|
| `HwpParagraphObject`, `HwpSection` | 본문 문단과 섹션 |
| `HwpTableObject`, `HwpTableCellObject` | HWP 표와 셀 |
| `HwpPictureObject`, `HwpShapeObject`, `HwpEquationObject`, `HwpOleObject` | 그래픽/control object |
| `HwpFieldObject`, `HwpHyperlinkObject`, `HwpBookmarkObject` | 필드와 참조 |
| `HwpHeaderFooterObject`, `HwpNoteObject`, `HwpPageNumObject` | 섹션 보조 control |
| `HwpFormObject`, `HwpMemoObject`, `HwpChartObject` | form, memo, chart |
| `HwpLineShapeObject`, `HwpRectangleShapeObject`, `HwpEllipseShapeObject`, `HwpConnectLineShapeObject` 등 | shape subtype wrapper |

### HWP binary record와 utility

저수준 조사 API는 HWP stream과 record를 직접 다룹니다.

| 이름 | 대표 용도 |
|---|---|
| `HwpRecord`, `RecordNode`, `TypedRecord` | binary record와 record tree |
| `DocInfoModel`, `SectionModel`, `SectionParagraphModel` | DocInfo와 section model |
| `HwpBinaryFileHeader`, `HwpDocumentProperties`, `HwpStreamCapacity` | 파일 헤더, 문서 속성, stream capacity |
| `DocumentPropertiesRecord`, `IdMappingsRecord`, `BinDataRecord`, `FaceNameRecord`, `StyleRecord` 등 | 주요 DocInfo record |
| `ControlHeaderRecord`, `TableControlRecord`, `PictureControlRecord`, `FieldControlRecord` 등 | HWP control record |
| `build_record_tree()`, `flatten_record_tree()`, `hwp_tag_name()` | record tree 보조 함수 |

### 예외

| 이름 | 발생 상황 |
|---|---|
| `HwpxError` | HWPX 처리 중 기본 예외 |
| `InvalidHwpxFileError` | HWPX package가 아니거나 필수 구조가 없을 때 |
| `HwpxValidationError` | HWPX validation 실패 |
| `InvalidHwpFileError` | HWP 파일 구조가 잘못됐을 때 |
| `HwpBinaryEditError` | HWP binary edit가 안전하지 않을 때 |
| `HancomInteropError` | 외부 Hancom converter/smoke 경로 실패 |
| `ValidationIssue` | strict lint와 validation report의 단위 항목 |

## 자주 쓰는 흐름

### 한 번 만들고 두 포맷으로 저장

```python
from jakal_hwpx import HancomDocument

doc = HancomDocument.blank()
doc.metadata.title = "보고서"
doc.append_paragraph("본문")
doc.append_table(rows=2, cols=2, cell_texts=[["A", "B"], ["1", "2"]])

doc.write_to_hwpx("build/report.hwpx")
doc.write_to_hwp("build/report.hwp")
```

### HWP를 읽어 HWPX로 저장

```python
from jakal_hwpx import HancomDocument

doc = HancomDocument.read_hwp("input.hwp")
doc.write_to_hwpx("build/output.hwpx")
```

### HWPX 직접 편집

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("input.hwpx")
doc.replace_text("초안", "최종")
doc.strict_validate()
doc.save("build/output.hwpx")
```

### HWP 직접 편집

```python
from jakal_hwpx import HwpDocument

doc = HwpDocument.open("input.hwp")
doc.append_paragraph("추가 문단")
doc.strict_validate()
doc.save("build/output.hwp")
```

## 검증과 release check

문서 저장 API는 내부 구조를 만들지만, 운영 코드에서는 저장 전후 검증을 함께 두는 것이 좋습니다.

```bash
python -m pytest tests/test_hancom_document.py tests/test_bridge.py -q
python scripts/check_release.py
```

실제 한컴 프로그램에서 열림 여부를 검증하려면 Windows에서 smoke validation 스크립트를 사용합니다.

```powershell
powershell -ExecutionPolicy Bypass -File scripts/run_hancom_smoke_validation.ps1 -InputPath input.hwpx -OutputPath build\roundtrip.hwpx
```

지원 범위의 최종 기준은 [STABILITY_CONTRACT.md](./STABILITY_CONTRACT.md)를 따릅니다.
