# `jakal_hwpx` 모듈 선택 가이드

이 문서는 `jakal_hwpx`를 처음 쓰는 사람이 "어떤 객체부터 잡아야 하는지" 빠르게 결정할 수 있게 정리한 가이드입니다.

핵심은 간단합니다.

- HWPX만 다루면 `HwpxDocument`
- HWP를 직접 다루면 `HwpDocument`
- HWP와 HWPX를 같은 코드로 다루고 싶으면 `HancomDocument`
- 포맷 전환을 묶고 싶으면 `HwpHwpxBridge`
- binary를 직접 조사해야 하면 `HwpBinaryDocument`

## 빠른 선택

| 하고 싶은 일 | 추천 객체 | 이유 |
|---|---|---|
| 새 문서를 만들고 HWPX/HWP 둘 다 저장 | `HancomDocument` | 공통 편집 모델이라 코드가 가장 단순함 |
| HWPX를 직접 편집 | `HwpxDocument` | 가장 넓은 직접 편집 범위 |
| HWP를 직접 편집 | `HwpDocument` | HWP 전용 객체 API가 바로 붙어 있음 |
| HWP와 HWPX를 오가며 저장 | `HwpHwpxBridge` | 전환 경로를 한 객체로 관리 가능 |
| record tree, stream, docinfo 조사 | `HwpBinaryDocument` | 가장 저수준 API |

## 권장 사용 순서

처음 시작할 때는 아래 순서로 생각하면 됩니다.

1. 포맷이 하나로 고정이면 `HwpxDocument` 또는 `HwpDocument`
2. 포맷이 섞이거나 나중에 바뀔 수 있으면 `HancomDocument`
3. 저장 결과를 HWP와 HWPX 양쪽으로 내보내야 하면 `HancomDocument` 또는 `HwpHwpxBridge`
4. 고수준 API로 안 보이는 payload를 조사해야 하면 `HwpBinaryDocument`

## 주요 객체

### `HancomDocument`

가장 권장되는 공통 편집 모델입니다.

이 객체를 쓰면:

- `.hwpx`를 읽어도 같은 구조
- `.hwp`를 읽어도 같은 구조
- 새 문서를 만들어도 같은 구조

즉, 앱 코드가 입력 포맷에 덜 묶입니다.

대표 진입점:

- `HancomDocument.blank()`
- `HancomDocument.read_hwpx(path)`
- `HancomDocument.read_hwp(path)`
- `write_to_hwpx(path)`
- `write_to_hwp(path)`

대표 append API:

- `append_paragraph()`
- `append_table()`
- `append_picture()`
- `append_shape()`
- `append_equation()`
- `append_ole()`
- `append_field()`
- `append_hyperlink()`
- `append_bookmark()`
- `append_note()`
- `append_header()`, `append_footer()`

짧은 예:

```python
from jakal_hwpx import HancomDocument

doc = HancomDocument.blank()
doc.metadata.title = "Sample"
doc.append_paragraph("Hello")
doc.write_to_hwpx("build/sample.hwpx")
doc.write_to_hwp("build/sample.hwp")
```

### `HwpxDocument`

HWPX를 직접 편집하는 주력 객체입니다.

이 객체를 쓰면 좋은 경우:

- 입력과 출력이 둘 다 HWPX일 때
- XML wrapper 메서드를 직접 쓰고 싶을 때
- package part, manifest, preview text 같은 HWPX package 요소를 직접 볼 때

대표 진입점:

- `HwpxDocument.blank()`
- `HwpxDocument.open(path)`
- `save(path)`
- `strict_lint_errors()`, `strict_validate()`

대표 append API:

- `append_paragraph()`
- `append_table()`
- `append_picture()`
- `append_shape()`
- `append_equation()`
- `append_ole()`
- `append_field()`
- `append_hyperlink()`
- `append_bookmark()`
- `append_note()`
- `append_header()`, `append_footer()`

짧은 예:

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("input.hwpx")
doc.replace_text("초안", "최종")
doc.strict_validate()
doc.save("build/output.hwpx")
```

### `HwpDocument`

HWP를 직접 편집하는 주력 객체입니다.

이 객체를 쓰면 좋은 경우:

- 기존 `.hwp`를 직접 수정하고 다시 저장할 때
- HWP control wrapper가 필요한 경우
- HWP 기준으로 tables, fields, notes, shapes를 다루고 싶을 때

대표 진입점:

- `HwpDocument.blank()`
- `HwpDocument.open(path)`
- `save(path)`
- `strict_lint_errors()`, `strict_validate()`

대표 append API:

- `append_paragraph()`
- `append_table()`
- `append_picture()`
- `append_shape()`
- `append_equation()`
- `append_ole()`
- `append_field()`
- `append_hyperlink()`
- `append_bookmark()`
- `append_note()`
- `append_header()`, `append_footer()`
- `append_auto_number()`

짧은 예:

```python
from jakal_hwpx import HwpDocument

doc = HwpDocument.open("input.hwp")
doc.append_paragraph("추가 문단")
doc.strict_validate()
doc.save("build/output.hwp")
```

### `HwpHwpxBridge`

문서를 한쪽에서 열고 다른 포맷으로 내보내거나, 중간에 기준 문서를 바꾸며 작업할 때 쓰는 도우미 객체입니다.

이 객체를 쓰면 좋은 경우:

- `input.hwp`를 받아 HWPX도 같이 만들고 싶을 때
- `input.hwpx`를 받아 HWP도 같이 만들고 싶을 때
- HWP/HWPX/HancomDocument를 반복해서 materialize하고 저장할 때

대표 진입점:

- `HwpHwpxBridge.open(path)`
- `HwpHwpxBridge.from_hwp(source)`
- `HwpHwpxBridge.from_hwpx(source)`
- `HwpHwpxBridge.from_hancom(source)`
- `save_hwp(path)`
- `save_hwpx(path)`

짧은 예:

```python
from jakal_hwpx import HwpHwpxBridge

bridge = HwpHwpxBridge.open("input.hwp")
bridge.save_hwpx("build/output.hwpx")
bridge.save_hwp("build/output_copy.hwp")
```

### `HwpBinaryDocument`

가장 저수준의 HWP API입니다.

이 객체는 일반적인 앱 코드의 기본 편집 진입점으로 권장하지는 않습니다. 대신 이런 상황에서 씁니다.

- record tree를 직접 보고 싶을 때
- `DocInfoModel`, `SectionModel`을 직접 다뤄야 할 때
- binary payload parse/build 도구를 확장할 때
- stream reencode 안정성을 검사할 때

짧은 예:

```python
from jakal_hwpx import HwpBinaryDocument

doc = HwpBinaryDocument.open("input.hwp")
print(doc.file_header().version)
print(doc.docinfo_model().id_mappings_record().named_counts())
print(doc.section_model(0).controls())
```

## 지원 범위

### 문서 흐름별

| 흐름 | 현재 상태 | 설명 |
|---|---|---|
| `HWPX -> HWPX` | 매우 좋음 | 가장 넓게 지원되는 직접 편집 경로 |
| `HWP -> HWP` | 강함 | HWP 저장과 재인코드 안정성이 높음 |
| `HWP -> HWPX` | 강함 | 의미 있는 구조를 잘 끌어올리는 편 |
| `HWPX -> HWP` | 좋음 | 지원된 편집 범위 기준으로 실사용 가능 |

### 컨트롤 패밀리별

| 항목 | 현재 수준 | 메모 |
|---|---|---|
| 문단, 스타일 | 높음 | 가장 안정적인 축 |
| 머리말, 꼬리말 | 높음 | `apply_page_type`까지 포함 |
| 필드, 북마크, 하이퍼링크 | 높음 | HWP native field subtype도 보존 |
| 각주, 미주, 자동번호, 페이지번호 | 높음 | setter와 bridge parity가 있음 |
| 섹션 설정, page border fill, visibility | 높음 | 읽기/쓰기와 bridge 테스트 포함 |
| 표 | 중상 | 실무 자동화에 필요한 surface는 넓음 |
| 그림 | 중상 | crop, rotation, line style 등 지원 |
| 도형, `connectLine` | 중상 | subtype wrapper와 주요 setter 포함 |
| 수식 | 중상 | layout, outMargins, rotation surface 포함 |
| OLE | 중상 | metadata, extent, line style 등 지원 |
| 차트 | 부분 | wrapper/typed access 중심 |
| form object | 부분 | 의미 있는 수정 API가 더 필요한 단계 |
| memo/comment | 부분 | typed access와 wrapper 중심 |

## 어떤 block/type을 얻게 되는가

`HancomDocument` 기준으로 본문은 다음 타입들로 다룹니다.

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

섹션은 `HancomSection`, 설정은 `SectionSettings`, 머리말/꼬리말은 `HeaderFooter`, 스타일은 `StyleDefinition`, `ParagraphStyle`, `CharacterStyle`로 다룹니다.

즉, 일반적인 앱 코드에서는 XML 노드나 binary record 대신 dataclass 중심으로 문서를 조립한다고 생각하면 됩니다.

## 실제로 자주 쓰는 패턴

### 1. HWPX를 읽어서 수정하고 다시 HWPX 저장

```python
from jakal_hwpx import HancomDocument

doc = HancomDocument.read_hwpx("input.hwpx")
doc.append_paragraph("추가 문단")
doc.write_to_hwpx("build/output.hwpx")
```

### 2. HWP를 읽어서 HWP와 HWPX 둘 다 내보내기

```python
from jakal_hwpx import HancomDocument

doc = HancomDocument.read_hwp("input.hwp")
doc.append_bookmark("appendix_anchor")
doc.write_to_hwp("build/output.hwp")
doc.write_to_hwpx("build/output.hwpx")
```

### 3. 새 문서를 공통 IR로 조립

```python
from pathlib import Path
from jakal_hwpx import HancomDocument

doc = HancomDocument.blank()
doc.metadata.title = "생성 문서"
doc.append_header("보고서")
doc.append_paragraph("첫 문단")
doc.append_table(rows=2, cols=2, cell_texts=[["A1", "B1"], ["A2", "B2"]])
doc.append_picture("logo.png", Path("assets/logo.png").read_bytes(), extension="png")
doc.write_to_hwpx("build/generated.hwpx")
```

### 4. HWP control을 직접 수정

```python
from jakal_hwpx import HwpDocument

doc = HwpDocument.open("input.hwp")
field = doc.fields()[0]
field.set_display_text("변경됨")

table = doc.tables()[0]
table.set_cell_text(0, 0, "수정")

doc.save("build/edited.hwp")
```

## 공개 API와 내부 API

호환성 기준으로 믿을 수 있는 표면은 `jakal_hwpx` 루트 import입니다.

예:

```python
from jakal_hwpx import HancomDocument, HwpxDocument, HwpDocument
```

반대로 아래처럼 내부 모듈 경로를 직접 import하는 건 구현 세부사항으로 보는 편이 맞습니다.

```python
from jakal_hwpx.document import HwpxDocument
from jakal_hwpx.hancom_document import HancomDocument
```

이 경로들은 내부 구조 정리 때 바뀔 수 있습니다.

## 검증 전략

안전하게 쓰려면 저장 전에 검증을 같이 거는 편이 좋습니다.

HWPX:

- `strict_lint_errors()`
- `strict_validate()`

HWP:

- `strict_lint_errors()`
- `strict_validate()`

추가로, 실제 한컴에서 열리는지까지 보고 싶다면 Windows에서 smoke validation 스크립트를 돌리면 됩니다.

```powershell
powershell -ExecutionPolicy Bypass -File scripts/setup_hancom_security_module.ps1 -DownloadIfMissing
powershell -ExecutionPolicy Bypass -File scripts/run_hancom_smoke_validation.ps1 -InputPath input.hwpx -OutputPath build\roundtrip.hwpx
```

## 고급 도구

일반적인 앱 코드에서는 아래 도구를 바로 쓸 일이 많지 않지만, corpus 분석이나 template-backed HWP 작업에는 유용합니다.

- `build_hwp_pure_profile()`
- `append_feature_from_profile()`
- `run_template_lab()`
- `scan_hwp_collection()`
- donor selection 도구들

## 정리

실무 기준으로는 이렇게 기억하면 충분합니다.

- 기본 편집 객체: `HancomDocument`
- HWPX 직접 작업: `HwpxDocument`
- HWP 직접 작업: `HwpDocument`
- 포맷 전환 도우미: `HwpHwpxBridge`
- binary 조사: `HwpBinaryDocument`

입문자는 `HancomDocument`부터 시작하면 됩니다. 그다음에 정말 필요한 경우에만 `HwpxDocument`나 `HwpDocument`로 내려가면 됩니다.
