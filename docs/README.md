# jakal_hwpx docs

이 폴더는 `jakal_hwpx`의 공개 API 문서를 보관합니다. 루트 [HWPX_MODULE.md](../HWPX_MODULE.md)는 전체 인덱스이고, 이 폴더의 문서는 클래스별 세부 설명을 맡습니다.

## 문서 목록

| 문서 | 설명 |
|---|---|
| [hancom-document.md](./hancom-document.md) | `HancomDocument` 기준의 공통 문서 모델을 설명합니다. 새 문서를 만들거나 HWP/HWPX를 같은 코드로 다룰 때 먼저 볼 문서입니다. |
| [hwpx-document.md](./hwpx-document.md) | `HwpxDocument`의 HWPX package, XML part, selector, append API, validation API를 설명합니다. HWPX 내부 구조를 직접 다룰 때 봅니다. |
| [hwp-document.md](./hwp-document.md) | `HwpDocument`의 HWP 전용 객체 API를 설명합니다. 기존 `.hwp`를 직접 편집하거나 HWP native wrapper를 써야 할 때 봅니다. |
| [bridge-and-binary.md](./bridge-and-binary.md) | `HwpHwpxBridge`와 `HwpBinaryDocument`를 설명합니다. 변환 경로를 고정하거나 HWP stream/record를 조사할 때 봅니다. |

## 이미지

| 파일 | 설명 |
|---|---|
| [images/basic-concept.svg](./images/basic-concept.svg) | README에서 쓰는 기본 사용 흐름 그림입니다. |
| [images/module-layers.svg](./images/module-layers.svg) | `HancomDocument`, `HwpxDocument`, `HwpDocument`의 레이어 관계를 보여 주는 그림입니다. |

## 읽는 순서

처음 사용하는 경우에는 [hancom-document.md](./hancom-document.md)를 먼저 보면 됩니다. 포맷별 세부 제어가 필요해지면 [hwpx-document.md](./hwpx-document.md)나 [hwp-document.md](./hwp-document.md)로 내려가고, 손상 원인 분석이나 역공학 작업은 [bridge-and-binary.md](./bridge-and-binary.md)에서 시작하면 됩니다.
