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

이 패키지는 ZIP 기반 HWPX 패키지를 대상으로 동작합니다. `.hwp -> .hwpx` 변환기는 번들하지 않습니다.

## 주요 API

대부분의 사용은 `HwpxDocument`에서 시작합니다.

- `HwpxDocument.open(path)`: 기존 HWPX 파일 열기
- `HwpxDocument.blank()`: 최소 구성의 새 문서 만들기
- `metadata()` / `set_metadata()`: 문서 메타데이터 조회 및 수정
- `get_document_text()`: 본문 텍스트 추출
- `set_paragraph_text()`: 특정 문단 텍스트 교체
- `append_paragraph()`: 문단 추가
- `replace_text()`: 문서 전체 텍스트 치환
- `section_settings()`: 페이지 설정 조회 및 수정
- `tables()`, `pictures()`, `notes()`, `fields()`: 구조화된 요소 접근
- `validation_errors()`: 문서 구조 검증
- `save(path)`: 파일 저장

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
- `HwpxPart`, `XmlPart`, `SectionPart`, `HeaderPart`, `ContentHpfPart`, `SettingsPart`, `VersionPart`, `MimetypePart`, `ContainerPart`, `ContainerRdfPart`, `ManifestPart`, `BinaryDataPart`, `GenericBinaryPart`, `GenericTextPart`, `GenericXmlPart`, `PreviewImagePart`, `PreviewTextPart`, `ScriptPart`
- `HwpxError`, `HwpxValidationError`, `InvalidHwpxFileError`

반대로 `jakal_hwpx.document`, `jakal_hwpx.parts` 같은 내부 모듈 경로 import는 구현 세부사항이며, 호환성 계약 대상으로 보지 않습니다.

## 지원 범위

지원:

- Python `3.11`, `3.12`, `3.13`
- ZIP 기반 `HWPX` 패키지의 열기, 수정, 검증, 컴파일, 저장
- `jakal_hwpx` top-level export를 통한 문서 편집 흐름

비지원:

- legacy binary `.hwp` 입력 파일
- `.hwp -> .hwpx` 변환기 번들 제공
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
```

테스트는 `examples/samples/hwpx/` 아래의 고정 fixture corpus를 사용합니다. 로컬과 CI가 같은 경로를 사용하도록 고정해 두었기 때문에 테스트 커버리지가 실행 환경에 따라 흔들리지 않습니다.

## 추가 문서

- [HWPX_MODULE.md](./HWPX_MODULE.md): 패키지 구조와 API 설명
- [examples/SHOWCASE.md](./examples/SHOWCASE.md): showcase 생성 흐름
- [RELEASING.md](./RELEASING.md): 배포 체크리스트
- [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md): 샘플 문서와 재배포 관련 고지

## 라이선스

프로젝트가 직접 작성한 소스 코드는 [MIT License](./LICENSE)를 따릅니다.

샘플 문서와 커밋된 생성 결과물은 별도 권리를 가질 수 있습니다. 재배포 전에는 [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md)를 확인하세요.
