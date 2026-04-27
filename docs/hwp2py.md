# hwp2py

`hwp2py`는 HWP 파일을 읽어서 같은 HWP 문서를 만드는 Python 스크립트를 생성합니다.

HWP는 HWPX처럼 zip package가 아니기 때문에 `raw` 모드를 제공하지 않습니다. 원본 바이너리를 그대로 심는 방식은 문서 편집 품질을 확인하는 데 의미가 작습니다. `hwp2py`는 항상 HWP를 `HancomDocument` 공개 모델로 해석한 뒤, 그 모델을 다시 만드는 스크립트를 출력합니다.

생성된 스크립트의 기본 출력은 HWP입니다. 즉 `input.hwp`를 분석해서 `build_document()`를 만들고, 실행하면 `write_hwp()`로 `output.hwp`를 저장합니다. 같은 문서 모델을 HWPX로 확인하고 싶을 때는 생성된 스크립트 안의 `write_hwpx()`를 호출하면 됩니다.

`--mode authoring`을 쓰면 출력이 저작용 DSL에 가까워집니다. 생성된 스크립트에는 `control(doc, name, ...)` 디스패처가 들어가고, 본문은 `p(doc, "...")`, 수식은 `eq(doc, r"...")`, 표는 `table(doc, [[...]])`처럼 짧은 control alias로 나옵니다. 문서 내용은 섹션별 함수로 분리됩니다.

`--mode macro`는 한 단계 더 LaTeX에 가까운 출력입니다. `document()`, `problem()`, `text()`, `math()`, `tabular()` 같은 macro tree를 우선 사용하고, 복잡한 블록은 실행 가능한 authoring 문장으로 폴백합니다.

## 기본 사용

```bash
hwp2py input.hwp -o recreate.py
python recreate.py output.hwp
```

`recreate.py`에는 `build_document()`, `write_hwp()`, `write_hwpx()`가 들어갑니다.

```python
from recreate import build_document

doc = build_document()
doc.append_paragraph("추가 문단")
doc.write_to_hwp("edited.hwp")
```

## Python API

```python
from jakal_hwpx import generate_hwp_script, write_hwp_script

script_text = generate_hwp_script("input.hwp")
script_path = write_hwp_script("input.hwp", "recreate.py", default_output_name="output.hwp")
```

### `generate_hwp_script()`

```python
generate_hwp_script(
    source_path,
    *,
    default_output_name=None,
    include_binary_assets=True,
    converter=None,
    mode="semantic",
) -> str
```

HWP를 읽고 Python 소스 문자열을 반환합니다.

| 인자 | 설명 |
|---|---|
| `source_path` | 입력 HWP 경로 |
| `default_output_name` | 생성된 스크립트가 기본으로 저장할 HWP 파일명 |
| `include_binary_assets` | Picture/OLE payload를 생성 스크립트에 포함할지 여부 |
| `converter` | 필요한 경우 HWP/HWPX 변환 함수. 기본값은 순수 Python HWP reader입니다. |
| `mode` | `"semantic"`은 공개 API 호출 중심, `"authoring"`은 편집용 DSL 출력, `"macro"`는 LaTeX식 macro tree 출력 |

### `write_hwp_script()`

```python
write_hwp_script(
    source_path,
    script_path=None,
    *,
    default_output_name=None,
    include_binary_assets=True,
    converter=None,
    mode="semantic",
) -> Path
```

스크립트를 파일로 저장하고 저장 경로를 반환합니다. `script_path`를 생략하면 입력 파일 옆에 `{stem}_hwp2py.py`가 만들어집니다.

## CLI 옵션

| 옵션 | 설명 |
|---|---|
| `-o`, `--output` | 생성할 Python 파일 경로 |
| `--default-output` | 생성된 Python 파일의 기본 HWP 출력 경로 |
| `--mode semantic\|authoring\|macro` | 생성 방식 선택. `dsl`은 `authoring`, `latex`는 `macro`의 alias |
| `--skip-binary-assets` | Picture/OLE payload를 제외하고 주석만 남김 |

## 재현 범위

`hwp2py`의 재현 범위는 HWP reader가 `HancomDocument`로 올릴 수 있는 구조와 같습니다.

| 대상 | 처리 방식 |
|---|---|
| metadata, section 설정 | 공개 모델 필드로 복원 |
| paragraph/table/header/footer | `HancomDocument` append API로 복원 |
| field/bookmark/hyperlink/note/auto number | 의미 모델 기준으로 복원 |
| style/numbering/bullet/memo shape 정의 | 공개 append API로 복원 |
| shape/equation/form/memo/chart | 공개 append API와 추가 필드 대입으로 복원 |
| picture/OLE | payload를 base64로 넣고 `append_picture()`/`append_ole()`로 복원 |
| 알 수 없는 HWP record/control | 현재 공개 모델로 올리지 못하면 생성 스크립트에도 나오지 않음 |

HWP 원본을 byte-for-byte로 재생성하는 도구가 아닙니다. 손상되는 HWP를 공개 모델로 얼마나 읽을 수 있는지 확인하거나, HWP 샘플을 사람이 읽을 수 있는 생성 코드로 바꿀 때 쓰는 도구입니다.
