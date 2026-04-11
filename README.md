# jakal-hwpx

`jakal-hwpx`는 ZIP 기반 HWPX 문서를 열고, 수정하고, 검증하고, 다시 저장하기 위한 Python 라이브러리입니다.

## 저장소 구성

- `src/jakal_hwpx`: 패키지 소스 코드
- `tests`: 테스트 코드
- `examples/samples/hwpx`: 고정 테스트 fixture로 쓰는 HWPX 샘플
- `examples/output_smoke`: 소규모 smoke corpus
- `examples/output`: showcase 생성 결과물

## 요구 사항

- Python 3.11 이상
- Git LFS 불필요

## 설치

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

개발 모드 설치:

```bash
python -m pip install -e .[dev]
```

샘플 `.hwp`/`.hwpx` 문서는 저장소에 일반 Git blob으로 포함되어 있으므로 클론 후 별도 `git lfs install` 또는 `git lfs pull` 단계가 필요하지 않습니다.

- 패키지 이름: `jakal-hwpx`
- import 경로: `jakal_hwpx`

## 빠른 시작

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.blank()
doc.set_metadata(title="예제", creator="jakal-hwpx")
doc.set_paragraph_text(0, 0, "Hello HWPX")
doc.save("build/hello.hwpx")
```

이 패키지는 ZIP 기반 HWPX 패키지를 주 대상으로 동작합니다. `.hwp <-> .hwpx` pure-python 브리지도 포함하지만, Hancom과 byte-identical한 완전 호환 변환기를 목표로 하지는 않습니다.

실험적 저수준 `.hwp` 지원도 포함합니다. `HwpBinaryDocument`는 Compound File 기반 legacy HWP 문서에서 `FileHeader`, `DocInfo`, `BodyText/Section*`, `PrvText`를 읽고, 기존 스트림 크기를 보존하는 범위에서 미리보기 텍스트와 본문 텍스트의 same-length 치환을 저장할 수 있습니다.

객체 기반 실험 API인 `HwpDocument`도 포함합니다. 이 레이어는 `HwpSection`, `HwpParagraphObject`를 통해 `.hwp` 문서를 `HwpxDocument`와 비슷한 방식으로 다루도록 맞춘 래퍼입니다. 현재 `.hwp -> HancomDocument IR -> HwpxDocument`와 `HwpxDocument -> HancomDocument IR -> .hwp`를 pure Python으로 수행하므로 `tables()`, `pictures()`, `append_paragraph()`, `append_table()`, `save(...hwpx)` 같은 고수준 흐름을 Hancom 자동화 없이 사용할 수 있습니다.

`build_hwp_pure_profile()`는 `hwp_collection/`에서 `table + picture + hyperlink`가 함께 있는 donor를 골라 `base.hwp + feature templates`를 추출하고, `HwpDocument.blank_from_profile()`은 그 profile만으로 `append_table_pure()`, `append_picture_pure()`, `append_hyperlink_pure()`를 수행합니다.
기본 bundled profile도 패키지에 포함되므로, donor나 profile 경로를 따로 넘기지 않아도 `HwpDocument.blank()`와 `append_*_pure()`를 바로 사용할 수 있습니다.

```python
from jakal_hwpx import HwpBinaryDocument, HwpDocument

binary_doc = HwpBinaryDocument.open("input.hwp")
print(binary_doc.file_header().version)
print(binary_doc.preview_text())

doc = HwpDocument.open("input.hwp")
paragraph = next(p for p in doc.section(0).paragraphs() if "2027" in p.text)
paragraph.replace_text_same_length("2027", "2028", count=1)
doc.save("build/edited.hwp")

# Pure-python HWP <-> HWPX bridge:
doc = HwpDocument.open("input.hwp")
doc.append_paragraph("Bridge paragraph")
doc.append_hyperlink("https://example.com", text="Example")
doc.save("build/edited_via_bridge.hwp")
doc.save("build/edited_via_bridge.hwpx")

# Pure-python experimental path:
from jakal_hwpx import HwpDocument, build_hwp_pure_profile

profile = build_hwp_pure_profile("hwp_collection", "build/hwp_pure_profile")
doc = HwpDocument.blank_from_profile(profile.root)
doc.append_table_pure()
doc.append_picture_pure()
doc.append_hyperlink_pure()
doc.save("build/pure_profile_output.hwp")

# Bundled profile path:
doc = HwpDocument.blank()
doc.append_table_pure()
doc.append_picture_pure()
doc.append_hyperlink_pure()
doc.save("build/bundled_pure_output.hwp")
```

## 주요 API

대부분의 사용은 `HwpxDocument`에서 시작합니다.

- `HwpxDocument.open(path)`: 기존 HWPX 파일 열기
- `HwpxDocument.blank()`: 최소 구성의 새 문서 만들기
- `metadata()` / `set_metadata()`: 문서 메타데이터 조회 및 수정
- `get_document_text()`: 본문 텍스트 추출
  - 주의: 이 출력은 검색/미리보기용 평탄화 텍스트이며, 편집 대상 문단 인덱스와 1:1로 대응하지 않습니다.
- `set_paragraph_text()`: 특정 문단 텍스트 교체
- `append_paragraph()`: 문단 추가
- `replace_text()`: 문서 전체 텍스트 치환
- `section_settings()`: 페이지 설정 조회 및 수정
- `tables()`, `pictures()`, `notes()`, `fields()`: 구조화된 요소 접근
- `validation_errors()`: 문서 구조 검증
- `save(path)`: 파일 저장
  - `append_row()` now fails fast when the template row contains preserved controls such as bookmarks or fields.
  - `append_paragraph(..., template_index=...)` now fails fast when the selected template paragraph contains preserved controls.
  - `validate=True` 기본값은 제어문 보존 검사까지 포함합니다. 실패 시 예외 메시지에 섹션 경로와 누락된 제어문 종류가 함께 표시됩니다.

자주 쓰는 helper 타입:

- `SectionSettings`
- `Table`, `TableCell`
- `HeaderFooterBlock`
- `Bookmark`, `Field`, `Note`, `Equation`, `ShapeObject`

패키지 구조와 타입 설명은 [HWPX_MODULE.md](./HWPX_MODULE.md)를 참고하세요.

## 공개 API 계약

공식 지원하는 공개 API는 `src/jakal_hwpx/__init__.py`에서 다시 export하는 top-level `jakal_hwpx` import 표면입니다.

안정적으로 지원하는 진입점:

- `HwpxDocument`, `DocumentMetadata`
- `Table`, `TableCell`, `Picture`, `Note`, `Equation`, `Bookmark`, `Field`, `AutoNumber`
- `HeaderFooterBlock`, `SectionSettings`, `StyleDefinition`, `ParagraphStyle`, `CharacterStyle`, `ShapeObject`
- `HwpBinaryDocument`, `HwpBinaryFileHeader`, `HwpRecord`, `HwpParagraph`
- `HwpDocument`, `HwpSection`, `HwpParagraphObject`, `HwpDocumentProperties`
- `HwpxPart`, `XmlPart`, `SectionPart`, `HeaderPart`, `ContentHpfPart`, `SettingsPart`, `VersionPart`, `MimetypePart`, `ContainerPart`, `ContainerRdfPart`, `ManifestPart`, `BinaryDataPart`, `GenericBinaryPart`, `GenericTextPart`, `GenericXmlPart`, `PreviewImagePart`, `PreviewTextPart`, `ScriptPart`
- `HwpxError`, `HwpxValidationError`, `InvalidHwpxFileError`, `InvalidHwpFileError`, `HwpBinaryEditError`

반대로 `jakal_hwpx.document`, `jakal_hwpx.parts` 같은 내부 모듈 경로 import는 구현 세부사항이며, 호환성 계약 대상으로 보지 않습니다.

## 지원 범위

지원:

- Python `3.11`, `3.12`, `3.13`
- ZIP 기반 `HWPX` 패키지의 열기, 수정, 검증, 컴파일, 저장
- legacy binary `.hwp` 파일의 저수준 열기, 스트림 파싱, 미리보기/본문 same-length 수정, 저장
- Hancom automation 없이 `.hwp`의 pure-python high-level HWPX bridge editing / export
- `hwp_collection` 기반 pure-python profile build 및 template-backed `append_table_pure()/append_picture_pure()/append_hyperlink_pure()`
- `jakal_hwpx` top-level export를 통한 문서 편집 흐름

비지원:

- Hancom과 동일 fidelity의 `.hwp <-> .hwpx` 완전 변환기
- multi-section / header-footer / note / bookmark까지 포함한 full-fidelity HWP binary writer
- donor/profile 없이 모든 HWP control을 임의 구조로 생성하는 범용 pure-python authoring
- top-level `jakal_hwpx` export 밖의 내부 import 경로에 대한 호환성 보장

## 예제

문서 열기와 검증:

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("input.hwpx")

print(doc.metadata())
print(doc.get_document_text())
print(doc.validation_errors())
print(doc.reference_validation_errors())
```

메타데이터와 본문 수정:

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("input.hwpx")
doc.set_metadata(title="수정된 제목", creator="Docs Team", keyword="example")
doc.replace_text("초안", "최종")
doc.append_paragraph("추가 문단", section_index=0)
doc.save("build/edited.hwpx")
```

페이지 설정과 표 수정:

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("input.hwpx")

settings = doc.section_settings(0)
settings.set_page_size(width=60000, height=85000)
settings.set_margins(left=7000, right=7000, top=5000, bottom=5000)

table = doc.tables()[0]
table.set_cell_text(0, 0, "수정됨")
table.append_row()[0].set_text("새 행")

doc.save("build/layout-updated.hwpx")
```

북마크, 하이퍼링크, 계산 필드 추가:

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.blank()
bookmark = doc.append_bookmark("summary_anchor")
doc.append_hyperlink("https://example.com", display_text="Example")
doc.append_calculation_field("40+2", display_text="42")
doc.append_cross_reference(bookmark.name or "summary_anchor", display_text="요약으로 이동")
doc.save("build/fields.hwpx")
```

## 테스트

```bash
python -m pip install -e .[dev]
python -m pytest -q
python scripts/run_stability_lab.py
```

Windows Hancom smoke validation:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/setup_hancom_security_module.ps1 -DownloadIfMissing
powershell -ExecutionPolicy Bypass -File scripts/run_hancom_smoke_validation.ps1 -InputPath examples\samples\hwpx\AI와_특이점_보고서.hwpx -OutputPath .codex-temp\hancom-smoke\sample.roundtrip.hwpx
powershell -ExecutionPolicy Bypass -File scripts/run_hancom_corpus_smoke_validation.ps1
```

Pure-python bridge가 기본 경로이고, Hancom smoke validator는 선택적 호환성 검증용입니다. 이 스크립트는 공식 security module을 등록하고 `RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")`를 호출한 뒤 구조화된 `.run.json` 로그를 남깁니다. 또한 기존 `Hwp.exe` 프로세스 때문에 COM 검증이 흔들릴 수 있는 경우 즉시 실패합니다.

테스트는 `examples/samples/hwpx/` 아래의 고정 fixture corpus를 사용합니다. 로컬과 CI가 같은 경로를 사용하도록 고정해 두었기 때문에 테스트 커버리지가 실행 환경에 따라 흔들리지 않습니다.

## 추가 문서

- [HWPX_MODULE.md](./HWPX_MODULE.md): 패키지 구조와 API 설명
- [scripts/run_stability_lab.py](./scripts/run_stability_lab.py): synthetic paragraph/container matrix round-trip harness
- [scripts/setup_hancom_security_module.ps1](./scripts/setup_hancom_security_module.ps1): install and register the official Hancom automation security module
- [scripts/run_hancom_smoke_validation.ps1](./scripts/run_hancom_smoke_validation.ps1): single-file Hancom `Open()` / `SaveAs()` smoke validation with fast-fail diagnostics
- [scripts/run_hancom_corpus_smoke_validation.ps1](./scripts/run_hancom_corpus_smoke_validation.ps1): corpus-wide Hancom smoke validation runner
- [examples/SHOWCASE.md](./examples/SHOWCASE.md): showcase 생성 흐름
- [RELEASING.md](./RELEASING.md): 배포 체크리스트
- [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md): 샘플 문서와 재배포 관련 고지

## 라이선스

프로젝트가 직접 작성한 소스 코드는 [MIT License](./LICENSE)를 따릅니다.

샘플 문서와 커밋된 생성 결과물은 별도 권리를 가질 수 있습니다. 재배포 전에는 [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md)를 확인하세요.
