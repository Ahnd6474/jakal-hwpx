# jakal-hwpx

파이썬에서 HWPX와 HWP 문서를 읽고, 수정하고, 검증하고, 변환하는 라이브러리입니다.

이 프로젝트는 한컴 GUI를 그대로 흉내 내는 편집기가 아닙니다. 대신 코드로 문서를 만들고, 구조화된 내용을 바꾸고, 저장 전후를 검증하는 데 초점을 둡니다. 템플릿 생성, 대량 치환, 표/필드/머리말/꼬리말 수정 같은 자동화 작업에 잘 맞습니다.

## 목차

- [설치](#설치)
- [빠른 시작](#빠른-시작)
- [이 라이브러리로 할 수 있는 일](#이-라이브러리로-할-수-있는-일)
- [어떤 객체를 쓰면 되는가](#어떤-객체를-쓰면-되는가)
- [지원 범위](#지원-범위)
- [주요 API](#주요-api)
- [예제](#예제)
- [검증과 테스트](#검증과-테스트)
- [추가 문서](#추가-문서)
- [라이선스](#라이선스)

## 설치

요구 사항:

- Python 3.11 이상
- Git LFS 불필요

PyPI에서 설치:

```bash
python -m pip install --upgrade pip
python -m pip install jakal-hwpx
```

로컬 체크아웃에서 설치:

```bash
python -m pip install --upgrade pip
python -m pip install .
```

개발 의존성까지 설치:

```bash
python -m pip install -e .[dev]
```

패키지 이름은 `jakal-hwpx`, import 경로는 `jakal_hwpx`입니다.

## 빠른 시작

새 문서를 만들고 HWPX와 HWP 둘 다 저장하려면 `HancomDocument`가 가장 단순합니다.

```python
from jakal_hwpx import HancomDocument

doc = HancomDocument.blank()
doc.metadata.title = "분기 보고서"

doc.append_header("사내 배포용")
doc.append_paragraph("매출 요약")
doc.append_table(rows=2, cols=2, cell_texts=[["Q1", "Q2"], ["120", "135"]])
doc.append_equation("x+y", width=3200, height=1800)

doc.write_to_hwpx("build/report.hwpx")
doc.write_to_hwp("build/report.hwp")
```

HWPX만 직접 만질 거라면 `HwpxDocument`가 가장 짧습니다.

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.blank()
doc.set_metadata(title="Hello")
doc.append_paragraph("Hello HWPX")
doc.save("build/hello.hwpx")
```

기존 HWP를 바로 수정하려면 `HwpDocument`를 쓰면 됩니다.

```python
from jakal_hwpx import HwpDocument

doc = HwpDocument.open("input.hwp")
doc.append_paragraph("추가 문단")
doc.append_hyperlink("https://example.com", text="Example")
doc.strict_validate()
doc.save("build/edited.hwp")
```

## 이 라이브러리로 할 수 있는 일

- HWPX 문서를 직접 열고 수정하고 저장
- HWP 문서를 순수 Python으로 열고 수정하고 저장
- HWP와 HWPX를 `HancomDocument` 기준으로 오가며 같은 편집 모델로 다루기
- 문단, 스타일, 표, 그림, 필드, 각주/미주, 머리말/꼬리말, 섹션 설정 같은 구조화된 요소 수정
- 저장 전후에 엄격 검증과 왕복 테스트로 결과 확인
- 필요하면 저수준 API로 내려가 HWP binary record를 직접 조사

이 라이브러리는 특히 이런 작업에 잘 맞습니다.

- 문서 템플릿 생성
- 필드 채우기
- 대량 텍스트 치환
- 표/도형/그림 데이터 갱신
- 문서 포맷 변환 자동화
- 저장 후 깨지지 않는지 확인하는 배치 처리

## 어떤 객체를 쓰면 되는가

처음 고를 때는 아래 표만 보면 됩니다.

| 객체 | 언제 쓰면 좋은가 | 설명 |
|---|---|---|
| `HancomDocument` | HWP와 HWPX를 같은 코드로 다루고 싶을 때 | 가장 권장되는 공통 편집 모델 |
| `HwpxDocument` | 입력과 출력이 둘 다 HWPX일 때 | 가장 넓은 직접 편집 범위 |
| `HwpDocument` | 기존 HWP를 직접 수정하거나 HWP로 저장할 때 | HWP 전용 객체 API |
| `HwpHwpxBridge` | HWP/HWPX 사이를 오가며 한쪽 결과물을 재료로 다른 쪽을 만들 때 | 문서 전환을 묶는 도우미 객체 |
| `HwpBinaryDocument` | binary record, stream, docinfo를 직접 봐야 할 때 | 가장 저수준 API |

짧게 말하면:

- 처음 시작: `HancomDocument`
- HWPX만 편집: `HwpxDocument`
- HWP 직접 편집: `HwpDocument`
- reverse engineering이나 디버깅: `HwpBinaryDocument`

## 지원 범위

### 문서 흐름별 지원 수준

| 흐름 | 현재 수준 | 설명 |
|---|---|---|
| `HWPX -> HWPX` | 매우 좋음 | 가장 넓게 지원되는 직접 편집 경로 |
| `HWP -> HWP` | 강함 | HWP 편집과 재인코드 안정성이 높음 |
| `HWP -> HWPX` | 강함 | HWP를 더 풍부한 XML 구조로 올리는 경로 |
| `HWPX -> HWP` | 좋음 | 지원된 편집 범위 기준으로 안정적이지만 정규화가 가장 많음 |

### 컨트롤 패밀리별 지원 수준

표의 의미:

- `높음`: 고수준 수정 API가 있고 왕복 테스트가 붙어 있음
- `중상`: 잘 되지만 포맷 제약이나 변환 정규화가 더 큼
- `부분`: typed access나 wrapper는 있지만 수정 범위가 상대적으로 좁음

| 항목 | `HWPX -> HWPX` | `HWP -> HWP` | `HWP -> HWPX` | `HWPX -> HWP` |
|---|---:|---:|---:|---:|
| 문단, 텍스트, 스타일 | 높음 | 높음 | 높음 | 높음 |
| 머리말, 꼬리말 | 높음 | 높음 | 높음 | 중상 |
| 필드, 북마크, 하이퍼링크 | 높음 | 높음 | 높음 | 높음 |
| 각주, 미주, 자동번호, 페이지번호 | 높음 | 높음 | 높음 | 중상 |
| 섹션 설정, page border fill, visibility | 높음 | 높음 | 높음 | 중상 |
| 표 | 높음 | 중상 | 높음 | 중상 |
| 그림 | 높음 | 중상 | 높음 | 중상 |
| 도형, `connectLine` | 높음 | 중상 | 높음 | 중상 |
| 수식 | 높음 | 중상 | 높음 | 중상 |
| OLE | 높음 | 중상 | 높음 | 중상 |
| 차트 | 부분 | 부분 | 부분 | 부분 |
| form object | 부분 | 부분 | 부분 | 부분 |
| memo/comment | 부분 | 부분 | 부분 | 부분 |

### 안정성에 대해

이 프로젝트는 한컴 GUI보다 "자동화 자유도" 쪽이 강합니다. 즉, 사용자가 마우스로 아무 데나 들어가서 편집하는 범용 GUI 편집기 역할보다는, 코드로 반복 작업을 처리하고 저장 후 안정성을 확인하는 데 더 잘 맞습니다.

무변경 `HWP -> HWP` 재인코드 쪽은 lossless audit이 붙어 있고, 최근 기준으로 corpus 64건이 byte-identical입니다. 지원된 control 계열을 수정하고 다시 여는 경로도 회귀 테스트로 계속 잡고 있습니다.

## 주요 API

### `HancomDocument`

가장 권장되는 공통 편집 모델입니다.

주요 메서드:

- `HancomDocument.blank()`
- `HancomDocument.read_hwpx(path)`
- `HancomDocument.read_hwp(path)`
- `append_paragraph()`
- `append_table()`, `append_picture()`, `append_shape()`, `append_equation()`, `append_ole()`
- `append_field()`, `append_hyperlink()`, `append_bookmark()`, `append_note()`
- `append_header()`, `append_footer()`
- `write_to_hwpx(path)`
- `write_to_hwp(path)`

### `HwpxDocument`

HWPX를 직접 편집할 때 쓰는 주력 객체입니다.

주요 메서드:

- `HwpxDocument.open(path)`
- `HwpxDocument.blank()`
- `append_paragraph()`
- `append_header()`, `append_footer()`
- `append_table()`, `append_picture()`, `append_shape()`, `append_equation()`, `append_ole()`
- `append_field()`, `append_hyperlink()`, `append_bookmark()`, `append_note()`
- `section_settings()`
- `strict_lint_errors()`, `strict_validate()`
- `save(path)`

### `HwpDocument`

HWP 편집에 쓰는 주력 객체입니다.

주요 메서드:

- `HwpDocument.open(path)`
- `HwpDocument.blank()`
- `append_paragraph()`
- `append_table()`, `append_picture()`, `append_shape()`, `append_equation()`, `append_ole()`
- `append_field()`, `append_hyperlink()`, `append_bookmark()`, `append_note()`
- `append_header()`, `append_footer()`, `append_auto_number()`
- `tables()`, `pictures()`, `fields()`, `notes()`, `section(index)`
- `strict_lint_errors()`, `strict_validate()`
- `save(path)`

### `HwpHwpxBridge`

이미 가진 문서를 기준으로 HWP/HWPX/HancomDocument를 오가며 저장할 때 편리한 도우미 객체입니다.

주요 메서드:

- `HwpHwpxBridge.open(path)`
- `HwpHwpxBridge.from_hwp(source)`
- `HwpHwpxBridge.from_hwpx(source)`
- `hancom_document()`
- `hwp_document()`
- `hwpx_document()`
- `save_hwp(path)`
- `save_hwpx(path)`
- `save(path)`

### `HwpBinaryDocument`

가장 저수준의 HWP API입니다.

다음 상황에서 씁니다.

- binary stream을 직접 보고 싶을 때
- `DocInfoModel`, `SectionModel`을 다뤄야 할 때
- reencode나 probe 작업을 할 때
- 지원 매핑을 더 넓히기 위해 record tree를 조사할 때

## 예제

### HWPX 열고 수정 후 저장

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("input.hwpx")
doc.replace_text("초안", "최종")
doc.append_paragraph("승인 완료")
doc.strict_validate()
doc.save("build/edited.hwpx")
```

### HWP 열고 수정 후 저장

```python
from jakal_hwpx import HwpDocument

doc = HwpDocument.open("input.hwp")
doc.append_paragraph("추가 문단")
doc.append_hyperlink("https://example.com", text="Example")
doc.strict_validate()
doc.save("build/edited.hwp")
```

### HWP를 HWPX로 변환

```python
from jakal_hwpx import HancomDocument

doc = HancomDocument.read_hwp("input.hwp")
doc.write_to_hwpx("build/exported.hwpx")
```

### HWPX를 HWP로 변환

```python
from jakal_hwpx import HancomDocument

doc = HancomDocument.read_hwpx("input.hwpx")
doc.write_to_hwp("build/exported.hwp")
```

### bridge 도우미로 포맷 전환

```python
from jakal_hwpx import HwpHwpxBridge

bridge = HwpHwpxBridge.open("input.hwp")
bridge.save_hwpx("build/bridge.hwpx")
bridge.save_hwp("build/bridge-copy.hwp")
```

### HWP record 살펴보기

```python
from jakal_hwpx import HwpBinaryDocument

doc = HwpBinaryDocument.open("input.hwp")
print(doc.file_header().version)
print(doc.docinfo_model().id_mappings_record().named_counts())
print(doc.section_model(0).controls())
```

## 검증과 테스트

기본 테스트:

```bash
python -m pip install -e .[dev]
python -m pytest -q
```

release gate와 안정성 검사:

```bash
python scripts/check_release.py
```

선택적 Hancom smoke validation:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/setup_hancom_security_module.ps1 -DownloadIfMissing
powershell -ExecutionPolicy Bypass -File scripts/run_hancom_smoke_validation.ps1 -InputPath examples\samples\hwpx\AI와_특이점_보고서.hwpx -OutputPath .codex-temp\hancom-smoke\sample.roundtrip.hwpx
powershell -ExecutionPolicy Bypass -File scripts/run_hancom_corpus_smoke_validation.ps1
```

일반적인 사용에선 pure Python 경로가 기본입니다. Hancom 쪽은 "정말 한컴에서 열어도 괜찮은지"를 추가로 확인할 때 쓰면 됩니다.

## 추가 문서

- [HWPX_MODULE.md](./HWPX_MODULE.md): 어떤 객체를 언제 써야 하는지 정리한 모듈 선택 가이드
- [STABILITY_CONTRACT.md](./STABILITY_CONTRACT.md): 지원 범위와 release gate 기준
- [scripts/check_release.py](./scripts/check_release.py): release gate 스크립트
- [scripts/audit_hwp_lossless_roundtrip.py](./scripts/audit_hwp_lossless_roundtrip.py): HWP lossless reencode audit
- [scripts/run_bridge_stability_lab.py](./scripts/run_bridge_stability_lab.py): HWP/HWPX bridge stability matrix
- [examples/SHOWCASE.md](./examples/SHOWCASE.md): 생성 예시 모음
- [RELEASING.md](./RELEASING.md): 배포 절차
- [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md): 샘플 문서와 재배포 관련 고지

고급 도구도 루트 import에서 바로 쓸 수 있습니다.

- `build_hwp_pure_profile()`
- `append_feature_from_profile()`
- `run_template_lab()`
- `hwp_collection` donor scanning 도구

이 도구들은 corpus 분석이나 template-backed HWP 작업에 유용하지만, 일반적인 앱 코드의 기본 진입점은 아닙니다.

## 라이선스

프로젝트 소스 코드는 [MIT](./LICENSE) 라이선스를 따릅니다.

샘플 문서와 생성 결과물은 별도 권리를 가질 수 있습니다. 재배포 전에는 [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md)를 확인하세요.
