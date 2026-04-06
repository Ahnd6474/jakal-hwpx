# jakal-hwpx

`jakal-hwpx`는 `HWPX` 문서를 읽고, 수정하고, 검증하고, 저장하기 위한 Python 도구 모음입니다. 현재는 `PDF <-> HWPX` 브릿지 기능도 초기 형태로 함께 포함되어 있습니다.

English version: [README.md](./README.md)

이 문서는 레포 전체 구조와 설치, 테스트, 샘플 위치를 설명합니다. 모듈과 API 상세 설명은 [HWPX_MODULE.md](./HWPX_MODULE.md)에서 따로 관리합니다.

## 레포 구성

- `src/jakal_hwpx`: 실제 Python 패키지
- `examples/samples`: 샘플 `hwpx`, `hwp`, `pdf` 문서
- `examples/output_smoke`: 테스트용 커밋된 smoke corpus
- `examples/output`: showcase 산출물
- `tools`: `.hwp -> .hwpx` 변환에 쓰는 Java 도구

## 설치

Python 3.11 이상이 필요합니다.

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

doc = HwpxDocument.open("examples/samples/hwpx/AI와_특이점_보고서.hwpx")
doc.replace_text("before", "after", count=1)
doc.save("build/edited.hwpx")
```

`PdfDocument`, `pdf_to_hwpx()`, `hwpx_to_pdf()` 같은 상세 사용법은 [HWPX_MODULE.md](./HWPX_MODULE.md)를 참고하세요.

## 샘플 파일 위치

루트에 흩어진 예시 문서 대신 샘플 입력은 아래에 정리합니다.

- `examples/samples/hwpx/`
- `examples/samples/hwp/`
- `examples/samples/pdf/`

생성 산출물과 검증 결과는 `build/validation/` 같은 빌드 디렉터리에 두는 것을 권장합니다.

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

- [HWPX_MODULE.md](./HWPX_MODULE.md): 모듈 구조와 API 설명
- [examples/SHOWCASE.md](./examples/SHOWCASE.md): showcase 생성 흐름
- [RELEASING.md](./RELEASING.md): 배포 체크리스트
- [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md): 샘플 문서, 번들 도구, HWPX 관련 명칭 범위 안내

## 라이선스

이 저장소의 프로젝트 작성 원본 소스 코드는 [MIT License](./LICENSE)로 제공합니다.

다만 샘플 문서, 커밋된 산출물, `tools/` 아래 번들 도구 자산은 별도 권리나 상위 라이선스를 따를 수 있습니다. 재배포 전에는 [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md)를 함께 확인하세요.
