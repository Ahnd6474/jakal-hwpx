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
| `jakal_hwpx.addons.docx.reader` | DOCX를 공통 모델로 투영 |
| `jakal_hwpx.addons.docx.writer` | 공통 모델을 DOCX로 작성 |
| `jakal_hwpx.addons.docx.preservation` | 원본 ZIP/OPC 패키지 보존과 부분 수정 |
| `jakal_hwpx.addons.docx.equations` | 지원하는 OMML ↔ 한글 수식 문법 |
| `jakal_hwpx.addons.docx.dependencies` | Word 백엔드의 선택적 로딩과 설치 안내 |

기존 `from jakal_hwpx import WordDocument`와
`from jakal_hwpx.word_document import WordDocument`는 같은 클래스에 연결됩니다.
이전 내부 변환 모듈 경로도 호환 import로 유지합니다. Word를 쓰지 않는 문서의
형식 분기와 기본 의존성은 그대로입니다.

addon에는 양방향 HWPX/DOCX 변환, 원본 DOCX 무수정 바이트 보존, 첫 페이지 영역·목록·필드의
대체 표현, 본문 참조번호와 각주·미주 내부 번호의 독립적인 서식 보존이 함께 포함됩니다.
지원하지 않는 기능과 순수 파이썬 HWP 출력의 제한은
[WordDocument](./word-document.md)와 [서식 재현 검증](./docx-fidelity-benchmark.md)에 기록합니다.
Microsoft Word·한컴 자동화는 명시적으로 선택한 외부 변환/검증 경로이며 기본 변환에
필요하지 않습니다.

## 검증

```sh
python -m pytest tests/test_docx_addon.py tests/test_note_reference_formatting.py tests/test_docx_roundtrip.py
```

기존 import의 클래스 동일성, 공통 객체의 자동 분기, 원본 보존과 양방향 변환,
`python-docx`를 차단한 새 프로세스에서의 본체 동작과 설치 안내를 검사합니다.
배포 검사에는 addon의 모든 구현 파일이 wheel과 sdist에 포함되는지도 포함합니다.
