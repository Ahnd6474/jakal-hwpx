# jakal-hwpx

`jakal-hwpx`는 `HWPX` 문서를 읽고, 수정하고, 검증하고, 다시 저장하기 위한 Python 패키지입니다.

## 저장소 구성

- `src/jakal_hwpx`: 실제 Python 패키지
- `examples/samples`: 샘플 `hwpx`, `hwp` 문서
- `examples/output_smoke`: 테스트용 smoke corpus
- `examples/output`: showcase 산출물
- `tools`: `.hwp -> .hwpx` 변환에 쓰는 유지보수용 Java 도구

## 설치

Python 3.11 이상이 필요합니다.

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

패키지 이름은 `jakal-hwpx`, import 경로는 `jakal_hwpx`입니다.

## 빠른 시작

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.blank()
doc.set_metadata(title="예제", creator="jakal-hwpx")
doc.set_paragraph_text(0, 0, "Hello HWPX")
doc.save("build/hello.hwpx")
```

모듈 구조와 주요 편집 API는 [HWPX_MODULE.md](./HWPX_MODULE.md)를 참고하세요.

## 핵심 API

대부분의 사용자는 `HwpxDocument`와 그 메서드가 반환하는 몇 가지 래퍼 타입만 알면 충분합니다.

- `HwpxDocument`: HWPX 열기, 생성, 수정, 검증, 컴파일, 저장
- `SectionSettings`: 섹션별 용지 크기와 여백 조회 및 수정
- `Table`, `TableCell`: 표 내용 조회와 편집
- `HeaderFooterBlock`: 머리말과 꼬리말 텍스트 조회 및 교체
- `Bookmark`, `Field`, `Note`, `Equation`, `ShapeObject`: 북마크, 필드, 각주, 수식, 도형 같은 고급 요소 편집
- `HwpxPart` 계열 클래스: 패키지 내부 파트에 직접 접근해야 할 때 쓰는 저수준 API

### `HwpxDocument` 요약

| 메서드 | 용도 |
| --- | --- |
| `open(path)` | 기존 HWPX 파일 열기 |
| `blank()` | 기본 파트가 들어 있는 새 문서 만들기 |
| `metadata()` / `set_metadata()` | 문서 메타데이터 조회 및 수정 |
| `get_document_text()` | 섹션 전체 본문 텍스트 추출 |
| `set_paragraph_text()` | 특정 문단 텍스트 교체 |
| `append_paragraph()` | 섹션 끝에 문단 추가 |
| `replace_text()` | 문서 전체에서 문자열 치환 |
| `section_settings()` | 페이지 크기와 여백 설정 접근 |
| `tables()`, `pictures()`, `notes()`, `fields()` | 고급 요소 래퍼 조회 |
| `validation_errors()` | 저장 전 패키지 유효성 점검 |
| `save(path)` | 파일로 저장 |

### 자주 보게 되는 래퍼 타입

| 타입 | 대표 사용처 |
| --- | --- |
| `SectionSettings` | 페이지 크기, 여백, 방향 수정 |
| `Table` / `TableCell` | 표 텍스트 수정, 행 추가, 셀 병합 |
| `HeaderFooterBlock` | 머리말과 꼬리말 텍스트 교체 |
| `Field` | 하이퍼링크, 메일 머지, 계산식, 상호 참조 필드 수정 |
| `Picture` | 포함된 이미지 바이너리 조회 및 교체 |
| `Note` | 각주와 미주 수정 |
| `ShapeObject` | 텍스트가 들어간 도형 수정 |

## 예제 코드

### 문서 열기와 검증

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("input.hwpx")

print(doc.metadata())
print(doc.get_document_text())
print(doc.validation_errors())
print(doc.reference_validation_errors())
```

### 메타데이터와 본문 수정

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("input.hwpx")
doc.set_metadata(title="수정된 제목", creator="Docs Team", keyword="example")
doc.replace_text("초안", "최종")
doc.append_paragraph("추가한 문단", section_index=0)
doc.save("build/edited.hwpx")
```

### 페이지 설정과 표 수정

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

### 하이퍼링크와 필드 추가

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

테스트는 기본적으로 아래 순서로 샘플을 찾습니다.

1. `JAKAL_HWPX_SAMPLE_DIR`
2. `all_hwpx_flat/`
3. `examples/output_smoke/`
4. `examples/output/`
5. `examples/samples/hwpx/`

## 추가 문서

- [HWPX_MODULE.md](./HWPX_MODULE.md): 패키지 구조와 API 설명
- [examples/SHOWCASE.md](./examples/SHOWCASE.md): showcase 생성 흐름
- [RELEASING.md](./RELEASING.md): 배포 체크리스트
- [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md): 샘플 문서와 번들 도구 관련 고지

## 라이선스

프로젝트가 직접 작성한 소스 코드는 [MIT License](./LICENSE)를 따릅니다.

샘플 문서, 커밋된 산출물, `tools/` 아래 번들 도구 자산은 별도 권리나 상위 라이선스를 따를 수 있습니다. 재배포 전에는 [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md)를 함께 확인하세요.
