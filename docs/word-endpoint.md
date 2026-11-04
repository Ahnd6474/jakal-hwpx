# Word 객체 ID와 편집 endpoint

`jakal-hwpx[docx]`만 설치하면 Python 서비스와 표준 라이브러리 HTTP adapter를 사용할 수
있습니다. Microsoft Word나 웹 프레임워크가 실행에 필요하지 않습니다.

## Python 서비스

```python
from jakal_hwpx.addons.docx import WordDocument, WordEndpoint, VersionConflict

doc = WordDocument.open("input.docx")
service = WordEndpoint(doc)  # 최초에 영속 객체 ID 메타데이터를 활성화
state = service.snapshot()
target = next(o["id"] for o in state["objects"] if o["kind"] == "p" and o["text"])

result = service.apply(state["version"], [
    {"target": "body", "method": "insert_paragraph", "args": [0, "추가 문단"], "as": "$new"},
    {"target": target, "method": "format_text", "kwargs": {"bold": True}},
    {"target": "$new", "method": "set_text", "args": ["삽입 후 수정한 문단"]},
])
print(result["version"], result["created"], result["changed"], result["deleted"])
doc.save("edited.docx")

# 이전 버전으로 수정하면 VersionConflict가 발생합니다.
try:
    service.apply(state["version"], [])
except VersionConflict:
    changes = service.changes(state["version"])
    # 아래의 캐시 갱신 예제처럼 동기화하고 수정 내용을 재검토합니다.
```

버전은 현재 DOCX 바이트의 SHA-256입니다. 일괄 수정은 같은 잠금 안에서 버전을 비교하고,
한 작업이라도 실패하면 전체를 되돌립니다. 오래된 요청을 자동으로 새 위치에 적용하지
않습니다. 응답의 `objects`는 최신 스냅샷이며 각 객체에 ID, 종류, 파트, 부모 ID,
서식까지 반영하는 `fingerprint`와 문단 텍스트 또는 표 크기가 있습니다.

`target`에는 UUID 또는 본문 별칭 `body`를 줍니다. `as`는 반환 객체를 같은 요청에서
다시 선택하는 `$name`입니다. `args`는 배열, `kwargs`는 객체입니다. 명시적으로 허용한
편집 메서드만 실행하며 임의 Python 코드·속성 접근·파일 경로는 받지 않습니다.

다른 스레드에서도 수정할 때는 같은 서비스 인스턴스를 사용하세요. 이 잠금은 서비스 밖의
직접 XML 수정을 직렬화하지 않습니다. 직접 수정 후에는 변경분 조회나 스냅샷으로 버전을 갱신합니다.

## 매번 ID를 다시 조회하지 않는 방법

최초 스냅샷의 ID를 클라이언트에 보관하고 수정 요청에도 그대로 사용합니다.
새로 만든 객체는 응답의 `results`에 ID가 있으며, 같은 일괄 요청 안에서는 `$별칭`으로
바로 수정할 수 있습니다. `compact=True`는 전체 `objects` 대신 `upserted`와 `deleted`를
반환합니다. 다음 예제는 최초 조회 이후 변경된 객체만 클라이언트 캐시에 반영합니다.

```python
state = service.snapshot()
cache = {o["id"]: o for o in state["objects"]}
version = state["version"]
target = next(o["id"] for o in cache.values() if o["kind"] == "p" and o["text"])

result = service.apply(version, [
    {"target": target, "method": "replace_text", "args": ["초안", "검토 완료"]},
], compact=True)
for identity in result["deleted"]:
    cache.pop(identity, None)
cache.update((o["id"], o) for o in result["upserted"])
version = result["version"]

# 다른 요청이나 직접 편집으로 변경된 경우에도 전체 ID 목록을 다시 받을 필요가 없습니다.
delta = service.changes(version)
if delta["requires_snapshot"]:
    state = service.snapshot()
    cache = {o["id"]: o for o in state["objects"]}
    version = state["version"]
else:
    for identity in delta["deleted"]:
        cache.pop(identity, None)
    cache.update((o["id"], o) for o in delta["upserted"])
    version = delta["version"]

selected = service.get_objects([target])  # 존재하지 않는 ID는 missing 배열
```

변경 이력은 서비스 인스턴스마다 최근 관측한 16개 버전을 메모리에 유지합니다.
이력이 만료되거나 서버를 다시 시작하면 `requires_snapshot=True`로 최초 동기화를
요구합니다. 변경 없는 읽기는 객체 목록과 ID 인덱스를 재사용하므로 편집 트랜잭션과
ID 재부여를 반복하지 않습니다. 외부 직접 편집 감지를 위한 문서 바이트·해시 확인은 남습니다.
반환값은 복사본이므로 클라이언트가 수정해도 서버 캐시는 바뀌지 않습니다.
버전 충돌(409)은 그대로 검사하며 오래된 편집을 자동 재실행하지 않습니다.

## HTTP 실행

저장소 루트에서 다음과 같이 실행합니다. 기본 호스트는 `127.0.0.1`이며 서버가 사용할
Bearer 토큰을 출력합니다. `--token`으로 직접 지정할 수도 있습니다.

```sh
python examples/word_endpoint.py --serve --input input.docx --port 8765
```

| 요청 | 결과 |
| --- | --- |
| `GET /document` | 버전과 객체 목록 |
| `GET /version` | 현재 버전만 반환 |
| `GET /changes?since={version}` | 변경 객체 `upserted`, 삭제 ID `deleted` 또는 `requires_snapshot=true` |
| `GET /objects/{id}` | 버전과 해당 객체 |
| `POST /objects` | `{"ids": ["UUID", "UUID"]}`로 선택 조회, `missing` 포함 |
| `POST /edits` | 버전을 비교한 원자적 일괄 수정 |
| `GET /export.docx` | 현재 DOCX 바이트 |
| `GET /export.hwpx` | 현재 내용을 변환한 HWPX 바이트 |

`POST /edits`는 `Content-Type: application/json`과 다음 본문을 받습니다.

```json
{
  "expected_version": "GET /document에서 받은 version",
  "compact": true,
  "operations": [
    {"target": "문단 UUID", "method": "replace_text", "args": ["초안", "검토 완료"]},
    {"target": "문단 UUID", "method": "format_text", "kwargs": {"bold": true}}
  ]
}
```

토큰을 지정한 서버에는 `Authorization: Bearer 토큰`을 보냅니다. 버전 충돌은 HTTP 409,
없는 객체는 404, 잘못된 작업은 422, 인증 실패는 401입니다. 한 요청은 최대 1MiB·1000개
작업을 받습니다. 서버에 파일을 여는/저장하는 경로 API는 없으며 변경된 파일은 export로
가져옵니다. Python에서 직접 사용할 때는 `make_http_server(service, port=8765, token=...)`가
아직 실행하지 않은 서버를 반환하므로 호출자가 실행·종료 수명을 관리합니다.

## 지원 작업

문단은 텍스트·범위 서식·링크·필드·각주·수식 추가·삭제·이동을 지원하고, 표는 셀·행·열·
병합·분할·테두리·배치·표 나누기·합치기를 지원합니다. 그림은 크기·자르기·회전·배치·
삭제를 지원합니다. 링크·책갈피·필드·수식은 각각의 전용 편집 메서드를 호출합니다.
`target="document"`에서는 스타일 생성·관계 변경·삭제·서식·테마, 목록 생성, 섹션 나누기를
지원합니다. 상세 인수는 [편집 API](./word-editing.md)를 참고하세요.
`update_fields`도 문서 작업으로 호출하며 결과의 상태 보고서는 `results[].value`로
반환합니다. 필드에는 `replace_result_text`·`rebuild_toc`를, 컨테이너에는
`insert_control`·`wrap_control`을 사용할 수 있습니다. 콘텐츠 컨트롤(`kind="sdt"`)의
UUID를 대상으로 `configure`·`set_text`·`unwrap`·`delete`·`move`와 내부 블록 삽입·치환을
호출합니다. 여러 문단·표를 포함한 필드 결과의 컨텍스트 편집은 Python API를 사용합니다.

| 연결 작업 | `method`와 `kwargs` |
| --- | --- |
| 표 셀 안에서 작업 | 표를 `target`으로 지정하고 작업에 `"cell": [행, 열]` 추가 |
| 목록 적용 | `set_list`, `{"listing": "목록 UUID 또는 $별칭", "level": 0}` |
| 표 합치기 | `join`, `{"following": "다음 표 UUID"}` |
| 각주 참조 이동 | `move_reference`, `{"paragraph": "대상 문단 UUID"}` |
| 문단·표 이동 | `move`, `{"index": 0, "destination": "body"}` |
| 여러 문단 범위 치환 | `replace_range`, `{"start": 2, "end_id": "끝 문단 UUID", "end": 4, "text": "새 글"}` |
| 범위 복사·이동 | `copy_range`/`move_range`, 위의 범위 인수와 `destination`, `offset` |
| 범위 책갈피·링크 | `bookmark_range`/`hyperlink_range`, 범위 인수와 `name`/`target` |

범위의 시작 문단은 작업의 `target`입니다. `end_id`를 생략하면 같은 문단을 사용합니다.
범위는 편집 가능한 텍스트 기준이며 필드 결과·수식·그림은 문자 인덱스에 포함하지 않습니다.
표·섹션 경계와 잘못 끊긴 필드·책갈피는 거부합니다.

## ID 보존과 변경 감지

ID를 활성화한 문서는 OOXML 속성, 숨김 책갈피 앵커, custom XML의 ID·앵커·해시 기록을
포함합니다. 기록에 원래 본문이나 수식 피연산자의 사본을 넣지 않습니다. 일반 저장·다시
열기·문단 삽입·이동에서 UUID를 유지합니다. 문단·표·수식 ID와 문단의 native identity도
HWPX 왕복에 보존합니다. `_JAKAL_` 뒤에 32자리 UUID가 붙는 숨김 책갈피 이름은
내부 식별용으로 예약하며 일반 책갈피 API·텍스트 범위에서 제외합니다.
공통 모델에 읽을 때도 내부 ID 책갈피를 본문 소유 필드로 만들지 않습니다. 실제 본문은
일반 문단으로 편집하고, DOCX 출력 시 내부 앵커만 별도로 복원합니다.

Word 앱이 재저장에서 알 수 없는 속성·`paraId`를 다시 만들고 책갈피 시작을 문단 밖으로
이동하는 경우도 앵커로 복구합니다. 실제 Word SaveAs로 문단 ID 유지와 편집된 내용의
재조회를 확인했습니다. 외부 앱이 메타데이터를 삭제하거나 객체 구조 자체를 바꾸면 새
객체가 될 수 있습니다. Word가 `w:hyperlink`를 리터럴 `HYPERLINK` 필드로 바꾸거나
반대로 바꾸어도 같은 문단 앵커·대상·표시 텍스트가 유일하게 대응하면 기존 링크 ID를
유지합니다. `kind="hyperlink"`와 `set_target`·`set_text`·`unlink`도 유지하며 필드의 실제
저장 형식은 `storage_kind`로 표시합니다. 대상·표시 텍스트는 기록에 해시로만 저장합니다.
이전 버전에서 생성한 링크 기록에는 대상 해시가 없으므로, 이 기능을 쓰려면 외부 앱으로
재저장하기 전에 현재 서비스로 한 번 열고 DOCX를 저장해야 합니다.
삭제·중복 링크·모호한 병합은 인덱스로 추측해 다른 객체에 적용하지 않습니다.

## 실행 예제와 검사

```sh
python examples/word_endpoint.py
python -m pytest tests/test_word_interaction.py
```

예제는 기본적으로 서버를 시작하지 않고 `build/word-interaction-example`에 DOCX·HWPX·
복원 DOCX와 스냅샷·요청·응답 JSON을 생성합니다. 범위 편집, 참조 연결, 스타일,
표 나누기·합치기, 행렬·합계·적분을 포함합니다. 이 예제의 Word 왕복 렌더링은 120dpi에서
전체 픽셀 및 흰 배경 제외 영역이 일치했습니다. 한글에서도 편집 가능한 수식 객체를
확인했으며 두 앱의 직접 배치에는 차이가 남습니다.
