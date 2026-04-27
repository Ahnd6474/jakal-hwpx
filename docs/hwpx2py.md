# hwpx2py

`hwpx2py`는 HWPX 파일을 읽어서 같은 문서를 만드는 Python 스크립트를 생성합니다.

기본값은 `semantic` 모드입니다. 생성된 스크립트는 원본 HWPX의 XML이나 바이너리 part를 그대로 숨겨 넣지 않고, `HancomDocument.blank()`에서 시작해 `append_paragraph()`, `append_table()`, `append_field()` 같은 공개 API 호출로 문서를 다시 조립합니다.

Picture와 OLE block은 문서 구조를 다시 만들기 위해 payload만 base64 문자열로 넣습니다. 그래도 원본 HWPX package를 통째로 복사하는 방식은 아닙니다.

`authoring` 모드는 같은 의미 모델을 더 사람이 읽기 좋은 Python DSL로 출력합니다. 생성된 스크립트에는 `control(doc, name, ...)` 디스패처가 들어가고, 본문은 `p(doc, "...")`, 수식은 `eq(doc, r"...")`, 표는 `table(doc, [[...]])`처럼 짧은 control alias로 나옵니다. 섹션별 내용은 `section_1(doc)` 같은 함수로 나뉩니다.

`macro` 모드는 `authoring`보다 한 단계 더 LaTeX에 가까운 출력입니다. 생성된 스크립트는 `document()`, `problem()`, `text()`, `math()`, `tabular()`, `draw()` 같은 macro tree를 우선 사용합니다. 아직 문항 번호나 선택지 의미를 완전히 추론하는 단계는 아니지만, 사람이 문서 의도를 함수 단위로 편집하기 쉬운 형태를 목표로 합니다.

`raw` 모드는 원본 package part를 base64로 심는 디버그 기준선입니다. 문서를 의미적으로 재생성하는 성능 지표로 보지 않습니다.

HWP 입력에서 같은 목적의 스크립트를 만들 때는 [`hwp2py`](./hwp2py.md)를 사용합니다. `hwp2py`는 raw 모드 없이 HWP를 `HancomDocument`로 읽어 HWP 작성 스크립트를 만듭니다.

## 기본 사용

```bash
hwpx2py input.hwpx -o recreate.py --strict
python recreate.py output.hwpx
```

`recreate.py`에는 `build_document()`와 `write_hwpx()`가 들어갑니다.

```python
from recreate import build_document

doc = build_document()
doc.append_paragraph("추가 문단")
doc.write_to_hwpx("edited.hwpx")
```

## Python API

```python
from jakal_hwpx import generate_hwpx_script, write_hwpx_script

script_text = generate_hwpx_script("input.hwpx", strict=True)
script_path = write_hwpx_script("input.hwpx", "recreate.py", default_output_name="output.hwpx")
```

### `generate_hwpx_script()`

```python
generate_hwpx_script(
    source_path,
    *,
    default_output_name=None,
    validate_input=True,
    strict=False,
    include_binary_assets=True,
    mode="semantic",
) -> str
```

HWPX를 읽고 Python 소스 문자열을 반환합니다.

| 인자 | 설명 |
|---|---|
| `source_path` | 입력 HWPX 경로 |
| `default_output_name` | 생성된 스크립트가 기본으로 저장할 HWPX 파일명 |
| `validate_input` | 생성 전에 `HwpxDocument.validate()`를 실행할지 여부 |
| `strict` | 생성 전에 `strict_lint_errors()`까지 확인할지 여부 |
| `include_binary_assets` | Picture/OLE payload를 생성 스크립트에 포함할지 여부 |
| `mode` | `"semantic"`은 공개 API로 재조립, `"authoring"`은 편집용 DSL 출력, `"macro"`는 LaTeX식 macro tree 출력, `"raw"`는 package part를 그대로 embed |

### `write_hwpx_script()`

```python
write_hwpx_script(
    source_path,
    script_path=None,
    *,
    default_output_name=None,
    validate_input=True,
    strict=False,
    include_binary_assets=True,
    mode="semantic",
) -> Path
```

스크립트를 파일로 저장하고 저장 경로를 반환합니다. `script_path`를 생략하면 입력 파일 옆에 `{stem}_hwpx2py.py`가 만들어집니다.

## CLI 옵션

| 옵션 | 설명 |
|---|---|
| `-o`, `--output` | 생성할 Python 파일 경로 |
| `--default-output` | 생성된 Python 파일의 기본 HWPX 출력 경로 |
| `--no-validate-input` | 입력 HWPX 기본 검증을 건너뜀 |
| `--strict` | strict lint를 통과한 경우에만 스크립트 생성 |
| `--mode semantic\|authoring\|macro\|raw` | 생성 방식 선택. `dsl`은 `authoring`, `latex`는 `macro`의 alias |
| `--skip-binary-assets` | Picture/OLE payload를 제외하고 주석만 남김 |

## 재현 범위

`semantic` 모드는 `HancomDocument`가 이해하는 문서 구조를 코드로 복원합니다.

| 대상 | 처리 방식 |
|---|---|
| metadata | `doc.metadata` 대입으로 복원 |
| section 설정 | `SectionSettings` 필드 대입으로 복원 |
| paragraph/table/header/footer | `HancomDocument` append API로 복원 |
| field/bookmark/hyperlink/note/auto number | 의미 모델 기준으로 복원 |
| style/numbering/bullet/memo shape 정의 | 공개 append API로 복원 |
| shape/equation/form/memo/chart | 공개 append API와 추가 필드 대입으로 복원 |
| picture/OLE | payload를 base64로 넣고 `append_picture()`/`append_ole()`로 복원 |
| 알 수 없는 block | 생성 스크립트에 skip 주석으로 남김 |

`semantic`, `authoring`, `macro` 모드는 byte-for-byte roundtrip 도구가 아닙니다. 손상 원인을 좁히거나, 샘플 문서를 테스트 fixture로 바꾸거나, HWPX 편집 로직을 사람이 읽을 수 있는 코드로 분리할 때 쓰는 도구입니다.

`authoring` 출력은 정확한 저수준 속성이 필요할 때도 실행 가능한 스크립트여야 합니다. 다만 기본 스타일 ID `0`이나 기본 셀 border ID처럼 편집에 노이즈가 큰 값은 helper 기본값으로 숨깁니다.

직접 수정할 때는 alias 대신 `control()`을 써도 됩니다.

```python
control(doc, "paragraph", "본문")
control(doc, "equation", r"\frac{1}{2}+x")
control(doc, "table", [["A", "B"]])
control(doc, "label", "problem-1")
```

`macro` 출력은 보통 다음처럼 시작합니다.

```python
document(
    doc,
    problem(
        text("본문"),
        math(r"\frac{1}{2}+x"),
        tabular([["A", "B"]]),
    ),
)
```

바이너리 payload가 큰 문서에서는 생성된 Python 파일도 커집니다. 구조만 비교할 목적이면 `--skip-binary-assets`를 쓰면 됩니다.

`raw` 모드는 모든 part를 embed하므로 `--skip-binary-assets`와 같이 쓸 수 없습니다. 이 모드는 복사에 가까운 보존/비교용 경로이며, `hwpx2py`의 주된 목표는 `semantic` 모드의 재생성 범위를 넓히는 것입니다.
