# `WordDocument`

`Document`는 공통 문서 클래스이고 `WordDocument(Document)`는 Word 전용 구현입니다.
기존 `HancomDocument`는 `Document`의 호환 별칭입니다.

```bash
pip install -e ".[docx]"
```

```python
from jakal_hwpx import Document, WordDocument

doc = Document.open("input.docx")
assert isinstance(doc, WordDocument)

doc.append_paragraph("추가 문단")
doc.replace_text("이전 내용", "수정 내용")
doc.save("output.docx")
doc.save("output.hwpx")  # 파이썬으로 직접 생성
doc.save("output.hwp")   # 파이썬으로 직접 생성
```

## 보존 구조

Word 전용 XML, 스타일, 콘텐츠 컨트롤, 주석, 수식, 도형, 미디어와 관계 정보는 원본
DOCX 패키지에 유지됩니다. 공통 모델은 편집을 위한 투영이며 원본을 대체하지 않습니다.
수정하지 않은 DOCX 저장은 원본과 바이트까지 같습니다. 텍스트나 글자 서식을 수정하면
해당 XML만 반영하고 나머지 패키지 항목은 유지합니다.

공통 `append_paragraph()`, `append_table()`, `append_picture()`, `append_hyperlink()`,
`append_section()`, `append_header()`, `append_footer()`, `replace_text()`는 Word 구현으로 연결됩니다. 현재 본문 append는
마지막 섹션을 대상으로 합니다. 반영할 수 없는 공통 모델 수정은 출력 전에 오류를 내며
원본 서식을 조용히 버리지 않습니다.

## Word 세부 편집

```python
with doc.edit_word() as word:
    word.paragraphs[0].runs[0].bold = True
    word.sections[0].header.paragraphs[0].text = "머리말"
    word.tables[0].cell(0, 0).paragraphs[0].runs[0].italic = True
```

`word`는 python-docx 문서입니다. 이 경로에서 직접 다루지 못하는 Word 요소도 원본
패키지에 보존됩니다. 블록이 끝나면 공통 모델을 갱신합니다. `word_document`를 직접 수정한
경우 `refresh_from_word()`로 뷰를 갱신할 수 있습니다. 갱신 후에는 블록과 스타일 참조를
다시 얻으세요. 예외가 발생한 `edit_word()` 블록의 수정은 되돌립니다.

## HWPX와 HWP 변환

`WordDocument`는 기본적으로 DOCX를 공통 문서 모델에 반영한 뒤 HWPX 또는 HWP를
파이썬에서 직접 생성합니다. Windows나 한컴오피스는 필요하지 않습니다. 본문,
글자·문단 서식, 표, 그림 등 공통 모델이 지원하는 요소를 변환합니다. 지원하지 않는
Word 전용 요소는 변환 시 `UserWarning`으로 알립니다. `import_warnings`에서도
해당 항목을 확인할 수 있습니다.

글자 효과(굵게, 기울임, 밑줄, 취소선, 위·아래 첨자), 글꼴 이름과 너비 비율,
표 셀 안쪽 여백은 HWPX와 HWP 작성 경로에 전달합니다. 모든 DOCX 기능의
시각적 동일성을 뜻하지는 않습니다.
HWPX 경로에서는 표 셀과 머리말·꼬리말의 글자별 서식, 링크, 그림, 중첩 표를
함께 변환합니다. 기본 표 스타일의 셀 여백과 정렬도 반영하며, 탭 위치·채움선과
문단 스타일 이름, 생성·수정 시각도 왕복 변환합니다.
직접 지정한 표 테두리·셀 음영과 선형 수식·분수·위/아래 첨자·근호를 변환합니다. 자동 번호·글머리표는
표시 문자를 정적 텍스트로 옮기므로 편집 후 자동 갱신되지 않습니다.
첫 페이지 전용 머리말·꼬리말, 각주 등은 현재 손실될 수 있습니다. 지원하지 않는
수식 구조는 텍스트로 평탄화하지 않고 HWPX 변환을 거부합니다. 원본 DOCX 저장은
이 경우에도 원본 수식을 보존합니다. HWP 출력의 셀 내부 개체 보존은 HWPX 경로와 다릅니다.
재현 검사 문서는 [DOCX 변환 재현 기준](./docx-fidelity-benchmark.md)을 참고하세요.

외부 변환기를 사용하려면 `converter(source_path, output_path, output_format)` 콜백을
명시적으로 주입할 수 있습니다. `preserve_formatting=True`를 지정한 경우에도 명시적인
변환기가 필요합니다. 이 옵션은 DOCX 패키지를 외부 변환기에 전달합니다.

```python
from jakal_hwpx import Document

doc = Document.open("input.docx")
doc.save("output.hwpx")
doc.save("output.hwp")
```

`doc.to_common_document()`는 변환에 쓰이는 공통 모델을 별도로 얻을 때 사용합니다.
이 경로는 `import_warnings`에 기록된 Word 전용 기능을 모두 표현하지 못할 수 있습니다.

## HWP/HWPX에서 DOCX 내보내기

같은 `docx` extra로 역방향 변환도 사용할 수 있습니다. 한컴오피스를 자동 실행하지
않으며, 기본 경로는 공통 모델에서 편집 가능한 Word 문서를 생성합니다.

```python
from jakal_hwpx import Document, HwpxDocument, HwpDocument

Document.open("input.hwpx").save("from_hwpx.docx")
Document.open("input.hwp").save("from_hwp.docx")

word = Document.open("input.hwpx").to_docx_document()
word.append_paragraph("Word에서 계속 편집할 내용")
word.save("edited.docx")

HwpxDocument.open("input.hwpx").save_as_docx("low_level.docx")
HwpDocument.open("input.hwp").save_as_docx("binary.docx")
```

본문의 문단별 묶음과 글자 서식, 용지·여백, 머리말·꼬리말, 표 병합·셀 여백,
그림, 하이퍼링크와 책갈피를 변환합니다. HWP/HWPX 표는 셀 안의 글자 서식, 너비,
셀 색상과 테두리도 읽습니다. 셀의 `cell_texts`를 수정하면 수정된 텍스트를 우선하며,
이때 기존 셀의 글자별 서식은 그대로 적용하지 않습니다.

원본과 조판이 완전히 같다는 의미는 아닙니다. HWP의 공통 모델로 읽히지 않은 서식,
자동 번호·글머리표·페이지 번호, 활성 필드와 도형·각주 배치는 보존되지 않을 수
있습니다. 필드·도형·각주는 표시 텍스트로 바꾸고 경고합니다. 아직 변환하지 못하는
수식 구문·차트·OLE 등은 오류를 내므로 외부 변환기를 사용해야 합니다. 지원하는
선형식·분수·위/아래 첨자·근호는 편집 가능한 Word 수식(OMML)으로 작성합니다.

```python
from jakal_hwpx import Document, HancomOfficeConverter

doc = Document.open("input.hwpx")
doc.save("native.docx", converter=HancomOfficeConverter())
```

이 선택 경로는 Windows 한컴오피스와 등록된 파일 접근 보안 모듈이 필요합니다.
공통 모델을 수정하지 않았으면 원본 HWP/HWPX를 그대로 전달합니다. 수정했거나
새 모델이면 현재 상태를 HWPX로 생성해 전달하므로 공통 모델 밖의 원본 세부 기능은
보존을 보장하지 않습니다. `preserve_formatting=True`만 주고 변환기를 생략하면
오류를 반환합니다. `preserve_formatting=False`는 변환기가 있어도 파이썬 출력을 선택합니다.
`HancomOfficeConverter`는 [한컴의 DOCX 저장 필터](https://forum.developer.hancom.com/t/saveas/660)
이름인 `OOXML`로 연결합니다.

2026-10-28 역방향 검증에서는 글자 서식 문서와 병합·음영 표 문서 각각의 HWP 및
HWPX 입력을 파이썬·한컴 경로로 변환했습니다. 출력 8개 모두 실제 Microsoft Word로
열어 PDF 출력에 성공했고, 모두 Letter 크기 1페이지였습니다. 파이썬과 한컴 출력의
PDF 텍스트는 공백을 정규화한 뒤 일치했습니다. 144 DPI에서 측정한 변경 픽셀 비율은
다음과 같습니다. 이는 한컴이 내보낸 DOCX와 파이썬 DOCX를 **동일한 Word 렌더러**로
비교한 결과입니다.

| 입력 | 글자 서식 문서 | 표 문서 |
| --- | --- | --- |
| HWPX | 0.183% | 1.889% |
| HWP | 0.196% | 1.889% |

표 너비·행 높이 등 배치 차이는 남습니다. 이 검증은 두 샘플에 한정되며 전체 문서의
서식 또는 화면 동일성을 보장하지 않습니다. 역방향 회귀 검증은
`tests/test_docx_export.py`에서 실행합니다.

## 시각 검증 범위

DOCX 재저장은 원본 Word와 같은 렌더러로 비교할 수 있습니다. 파이썬 변환으로 만든
HWPX/HWP는 Word의 모든 기능과 페이지 배치를 보장하지 않습니다. 페이지 수·용지
크기·텍스트·이미지 차이는 별도로 측정해야 합니다.

다음은 2026-10-28에 **외부 한컴 변환기를 사용했던 경로**의 참고 결과입니다.
파이썬 기본 변환 경로의 시각적 검증 결과로 해석하면 안 됩니다. 글자 서식 문서와
병합·테두리·음영 표/머리말/꼬리말 문서 각 1페이지를 실제 Microsoft Word와
한컴오피스로 PDF 출력하여 144 DPI로 비교했습니다.

| 경로 | 결과 |
| --- | --- |
| 수정 없는 DOCX 재저장 | 두 문서 모두 원본 바이트 및 Word 렌더링 일치 |
| 텍스트 수정 DOCX와 Word 찾기/바꾸기 결과 | 텍스트·페이지 수·크기 일치, 변경 픽셀 비율 각각 0.014%, 0% |
| HWPX 및 HWP와 한컴이 원본 DOCX를 직접 연 결과 | 두 문서 모두 렌더링 픽셀 일치 |
| HWPX와 원본 Word 출력 | 변경 픽셀 비율 1.750%, 1.945%; PDF 높이 790pt 대 792pt |

변경 픽셀 비율은 페이지 전체에서 RGB 채널 차이가 24를 넘는 픽셀의 비율입니다.
용지 크기가 다르면 크기 불일치를 별도로 기록하고, 확대·축소 없이 흰 여백을 덧대어
비교했습니다. 이 수치는 두 검증 문서에 대한 결과이며 임의 DOCX의 전체 서식이나
Word와 한컴 간 화면 배치 동일성을 보장하지 않습니다.

자동 회귀 검증은 `tests/test_word_document.py`, `tests/test_docx_preservation.py`,
`tests/test_docx_import.py`에서 공통 메서드의 Word 연결, 원본 패키지/주석/수식 보존,
지원하지 않는 수정의 거부와 파이썬 기본 변환을 확인합니다.
