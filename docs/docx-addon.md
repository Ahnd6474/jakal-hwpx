# DOCX addon (beta)

DOCX 지원은 `jakal-hwpx`에 포함되는 선택 설치 addon입니다. 본체와 같은 버전의
wheel/sdist로 배포하며, `docx` extra로 Word 백엔드 의존성을 설치합니다.
`1.1.0b1`에서는 베타 API로 제공되며, 정식 `1.1.0` 전까지 피드백에 따라 동작과
문서 변환 충실도 정책이 조정될 수 있습니다.

```sh
pip install "jakal-hwpx[docx]"
```

저장소에서 개발할 때는 `pip install -e ".[docx]"`를 사용합니다. 기본 설치의
HWP/HWPX 읽기·편집·저장은 `python-docx`나 Microsoft Word 없이 동작합니다.
일반적인 본체 import는 DOCX addon을 로드하지 않습니다. Word 작업을 실행할 때
백엔드가 없으면 위 설치 명령을 포함한 오류를 제공합니다.

## 공통 문서 객체에서 사용

```python
from jakal_hwpx import Document
from jakal_hwpx.addons.docx import WordDocument

word = Document.open("input.docx")
assert isinstance(word, WordDocument)
word.replace_text("이전", "이후")
word.save("edited.docx")
word.save("converted.hwpx")

Document.open("input.hwpx").save("converted.docx")
```

`Document.open()`과 확장자에 따른 `save()`가 addon에 연결되므로 별도 등록 호출은
필요하지 않습니다. 공통 `Document` 모델을 그대로 사용하고, Word 전용 패키지와
세부 편집은 `WordDocument`가 담당합니다. `WordDocument.blank()`로 새 Word 문서를
만들거나 `edit_word()`에서 python-docx 백엔드를 직접 편집할 수 있습니다.

## 구현 경계와 호환성

| 모듈 | 역할 |
| --- | --- |
| `jakal_hwpx.addons.docx.document` | WordDocument, Word 편집 연결, 형식별 저장 |
| `jakal_hwpx.addons.docx.editing` | 위치·범위·표·그림·영역의 트랜잭션 편집 API |
| `jakal_hwpx.addons.docx.reader` | DOCX를 공통 모델로 투영 |
| `jakal_hwpx.addons.docx.writer` | 공통 모델을 DOCX로 작성 |
| `jakal_hwpx.addons.docx.preservation` | 원본 ZIP/OPC 패키지 보존과 부분 수정 |
| `jakal_hwpx.addons.docx.equations` | 지원하는 OMML ↔ 한글 수식 문법 |
| `jakal_hwpx.addons.docx.features` | Word 세부 속성·콘텐츠 컨트롤·custom XML 왕복 보존 |
| `jakal_hwpx.addons.docx.formatting` | 테마·조건부 표 서식·글자 간격·음영 패턴 매핑 |
| `jakal_hwpx.addons.docx.pictures` | 그림 배치·자르기·회전·외부 링크 매핑 |
| `jakal_hwpx.addons.docx.native_effects` | 한글 자동 목록·필드 매핑과 Word 목록 복원 |
| `jakal_hwpx.addons.docx.fields` | 중첩·여러 문단 필드의 명령과 결과 범위 보존 |
| `jakal_hwpx.addons.docx.textboxes` | Word 텍스트 상자 ↔ 편집 가능한 한글 글상자 |
| `jakal_hwpx.addons.docx.editing` | 원본 객체를 수정하는 트랜잭션과 편집 핸들 |
| `jakal_hwpx.addons.docx.editing_tables` | 병합 그리드의 행·열 추가·삭제와 셀 분할 |
| `jakal_hwpx.addons.docx.editing_format` | 세부 글자·문단·셀 테두리 서식 검증과 적용 |
| `jakal_hwpx.addons.docx.editing_stories` | 페이지·다단·섹션과 각주·미주 편집 |
| `jakal_hwpx.addons.docx.editing_fields` | 자동 목록 인스턴스와 필드 명령·결과 편집 |
| `jakal_hwpx.addons.docx.editing_ranges` | 서식을 유지하는 문단 간 범위·복사·이동·치환 |
| `jakal_hwpx.addons.docx.editing_navigation` | 링크·책갈피와 연결된 참조 갱신 |
| `jakal_hwpx.addons.docx.editing_styles` | 스타일 생성·관계·삭제와 테마 편집 |
| `jakal_hwpx.addons.docx.editing_table_layout` | 떠 있는 표·행 분할·표 나누기·합치기 |
| `jakal_hwpx.addons.docx.editing_math` | 구조화 수식 선택·추가·수정·삭제 |
| `jakal_hwpx.addons.docx.identity` | 영속 UUID·native 앵커·ID 기록 |
| `jakal_hwpx.addons.docx.endpoint` | 버전 비교·원자적 편집과 선택적 HTTP adapter |
| `jakal_hwpx.addons.docx.dependencies` | Word 백엔드의 선택적 로딩과 설치 안내 |

기존 `from jakal_hwpx import WordDocument`와
`from jakal_hwpx.word_document import WordDocument`는 같은 클래스에 연결됩니다.
이전 내부 변환 모듈 경로도 호환 import로 유지합니다. Word를 쓰지 않는 문서의
형식 분기와 기본 의존성은 그대로입니다.

addon에는 양방향 HWPX/DOCX 변환, 원본 DOCX 무수정 바이트 보존, 첫 페이지 영역의
대체 표현, 본문 참조번호와 각주·미주 내부 번호의 독립적인 서식 보존이 포함됩니다.
그림의 배치·자르기·회전, 조건부 표 스타일과 테마 색상, 글자 간격·세로 위치를
한글 속성에 반영하며, 추가 Word 속성은 작은 기능별 기록으로 HWPX 안에 보존합니다.
원본 DOCX 전체를 HWPX에 숨겨 넣는 방식은 사용하지 않습니다.

Word 콘텐츠 컨트롤의 본문은 한글에서 편집할 수 있는 내용으로 투영합니다.
DOCX로 돌아올 때 해당 내용을 원래 컨트롤 속성으로 감싸고, 연결된 custom XML의
요소·속성 값을 갱신합니다. 한글 안에서 Word의 드롭다운·날짜 선택 UI를 제공하지는 않습니다.

Word 자체의 세부 편집에는 [`edit_word_details()`](./word-editing.md)를 사용할 수 있습니다.
Word 원본 XML에 삽입·삭제·이동·서식 수정을 적용하고, 바인딩된 텍스트의 custom XML도
함께 갱신합니다. 편집 범위와 단위, 원본 보존 방식은 해당 가이드에 설명합니다.

## 한글에서 자동 갱신되는 목록과 필드

일반 숫자·로마 숫자·영문자 목록과 글머리표는 한글의 문단 번호·글머리표로 작성합니다.
목록 스타일의 다음 문단도 같은 번호 모양을 따르므로 한글에서 Enter로 추가한 항목이
자동으로 이어집니다. 새 항목은 별도 Word 보존 기록이 없어도 DOCX의 같은 자동 목록으로
복원합니다. 여러 수준의 숫자와 서로 다른 글머리표는 각각의 한글 자원으로 매핑합니다.

| Word 명령 | 한글 표현 | 확인한 동작 |
| --- | --- | --- |
| `PAGE`, `NUMPAGES` | 쪽 번호·전체 쪽 수 | 실제 2페이지에서 `i / 2`, `ii / 2` 출력 |
| `DATE \@ "yyyy-MM-dd"`, `yyyy/MM/dd` | 날짜 필드 | 한글에서 현재 날짜로 갱신 |
| `TITLE`, `AUTHOR`, 지원하는 `DOCPROPERTY` | 문서 요약 정보 필드 | 문서 속성으로 표시값 갱신 |
| `REF`, `PAGEREF` | 범위 책갈피·상호 참조 | 대상 글 수정 후 참조 내용과 쪽 번호 갱신 |
| 호환되는 `= SUM(LEFT)` 등 표 계산식 | 계산식 필드 | 셀 `10 → 15` 직접 입력 후 합계 `35.00`으로 갱신 |
| 단순 `MERGEFIELD` | 메일 머지 필드 | 한글 필드로 작성; 데이터 원본 연결·실행은 별도 |

원래 Word 명령과 숫자 서식은 DOCX 복원 시 유지합니다. 표 계산식 검증에서는
한글의 키보드 입력 직후 갱신과 저장 후 `15`, `20`, `35.00` 복원을 확인했습니다.
같은 문서에서 오토메이션 `InsertText`로 값을 바꾼 실험은 재계산을 유발하지 않았으므로,
이 API 호출을 키보드 입력과 같은 자동 갱신 검증으로 취급하면 안 됩니다.

모든 Word 필드와 번호 형식을 자동 갱신하는 것은 아닙니다. 예를 들어 `decimalZero`,
일부 로캘 번호 형식, 지원 범위 밖의 날짜 서식·필드 명령은 표시값과 원래 정의를
유지하고 경고합니다. `REF`의 문단 번호 옵션과 여러 문단에 걸친 책갈피 범위도 제한이
있습니다. 그림의 tight/through 윤곽 감싸기는 한글에서 사각 감싸기로 근사하며 원래
Word 윤곽 속성은 복원합니다. 굵은 밑줄·가는 해치 같은 일부 효과도 엔진 간 근사가 남습니다.

## 열기 실패와 내용 손실 회귀 수정

목록이 없는 유효한 DOCX는 `numbering.xml`을 생략할 수 있습니다. 번호 정의가 없는
파일을 읽을 때 백엔드의 미구현 생성 함수를 호출하지 않도록 수정했습니다.

중첩 필드와 여러 문단에 걸친 필드는 시작·명령·결과·끝을 구분해 보존합니다.
한글의 편집 가능한 문단·링크·그림·표와 내부 필드에 현재 결과를 두고, 빈 경계
필드에는 Word 명령과 속성만 기록합니다. DOCX로 돌아올 때 현재 내용에 복합 필드
구조를 다시 적용하므로 목차 명령이나 내부 필드의 글자가 사라지지 않습니다.
`IF`·`TOC` 같은 바깥 필드의 한글 내 재계산은 지원하지 않으며 경고합니다.
내부 필드가 지원하는 한글 필드로 매핑되면 해당 내부 필드의 자동 갱신은 유지합니다.

Word 텍스트 상자는 편집 가능한 한글 글상자로 작성하고, DOCX에서는 현재 내용을
상자 안에 복원합니다. 문단별 서식·링크·그림·표·필드를 함께 변환하며 Word의 대체
표현을 중복 수입하지 않습니다. 원래 텍스트를 비밀 사본으로 보관하지 않으며,
그림 속성 템플릿의 본문과 VML 보조 캐시도 제거합니다. Word 호환 설정을 함께
보존해 텍스트 상자의 해석이 호환 모드에 따라 달라지는 문제를 방지합니다.
여러 텍스트 상자를 하나로 묶은 그룹 도형은 별도 매핑이 필요하며 명시적으로 오류를 냅니다.

지원하지 않는 기능과 순수 파이썬 HWP 출력의 제한은
[WordDocument](./word-document.md)와 [서식 재현 검증](./docx-fidelity-benchmark.md)에 기록합니다.
Microsoft Word·한컴 자동화는 명시적으로 선택한 외부 변환/검증 경로이며 기본 변환에
필요하지 않습니다.

## 검증

```sh
python -m pytest tests/test_docx_addon.py tests/test_note_reference_formatting.py tests/test_docx_roundtrip.py tests/test_docx_extended_features.py tests/test_docx_native_effects.py tests/test_docx_loss_regressions.py
```

기존 import의 클래스 동일성, 공통 객체의 자동 분기, 원본 보존과 양방향 변환,
`python-docx`를 차단한 새 프로세스에서의 본체 동작과 설치 안내를 검사합니다.
배포 검사에는 addon의 모든 구현 파일이 wheel과 sdist에 포함되는지도 포함합니다.
한글 필드 매개변수 정규화, 번호·글머리표 혼합 수준, 새 목록 항목, 원래 Word 필드 명령의
3회 왕복도 검사합니다. 자동 테스트와 실제 한글·Word에서 측정한 시각 검증은 구분하며,
측정 조건과 남은 차이는 [서식 재현 검증](./docx-fidelity-benchmark.md)에 기록합니다.
