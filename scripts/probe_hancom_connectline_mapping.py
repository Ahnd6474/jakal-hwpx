from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any


REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = REPO_ROOT / "src"

if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from jakal_hwpx import HwpDocument, HwpxDocument  # noqa: E402
from jakal_hwpx._hancom import convert_document  # noqa: E402
from jakal_hwpx.hwp_binary import TAG_CTRL_DATA, RecordNode, hwp_tag_name  # noqa: E402


def _default_output_root() -> Path:
    return REPO_ROOT / ".codex-temp" / "connectline_probe"


def _first_diff_offset(left: bytes, right: bytes) -> int | None:
    limit = min(len(left), len(right))
    for index in range(limit):
        if left[index] != right[index]:
            return index
    if len(left) != len(right):
        return limit
    return None


def _payload_summary(node: RecordNode | None) -> dict[str, Any] | None:
    if node is None:
        return None
    return {
        "tag_id": node.tag_id,
        "tag_name": hwp_tag_name(node.tag_id),
        "payload_len": len(node.payload),
        "payload_hex": node.payload.hex(),
    }


def _json_safe(value: Any) -> Any:
    if isinstance(value, bytes):
        return {
            "type": "bytes",
            "hex": value.hex(),
        }
    if isinstance(value, dict):
        return {str(key): _json_safe(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_json_safe(item) for item in value]
    return value


def _find_shape_by_comment(document: HwpDocument, shape_comment: str):
    for shape in document.shapes():
        if getattr(shape, "shape_comment", "") == shape_comment:
            return shape
    return None


def _shape_summary(shape) -> dict[str, Any]:
    if shape is None:
        return {
            "present": False,
        }
    component = shape._shape_component_record()
    specific = shape._shape_specific_record()
    ctrl_data = next((child for child in shape.control_node.children if child.tag_id == TAG_CTRL_DATA), None)
    return {
        "present": True,
        "parsed_kind": shape.kind,
        "shape_comment": getattr(shape, "shape_comment", None),
        "size": shape.size(),
        "specific_fields": _json_safe(shape.specific_fields()),
        "descendant_tag_ids": shape.descendant_tag_ids(),
        "control_header": _payload_summary(shape.control_node),
        "component_record": _payload_summary(component),
        "specific_record": _payload_summary(specific),
        "ctrl_data": _payload_summary(ctrl_data),
    }


def _comparison(name: str, left: dict[str, Any] | None, right: dict[str, Any] | None) -> dict[str, Any]:
    left_payload = bytes.fromhex(left["payload_hex"]) if left is not None and "payload_hex" in left else b""
    right_payload = bytes.fromhex(right["payload_hex"]) if right is not None and "payload_hex" in right else b""
    return {
        "name": name,
        "same_tag": (left or {}).get("tag_name") == (right or {}).get("tag_name"),
        "same_length": len(left_payload) == len(right_payload),
        "same_payload": left_payload == right_payload,
        "first_diff_offset": _first_diff_offset(left_payload, right_payload),
    }


def _build_probe_source(kind: str, output_path: Path) -> Path:
    document = HwpxDocument.blank()
    comment = f"PROBE-{kind.upper()}"
    shape = document.append_shape(
        kind=kind,
        width=3600,
        height=900,
        line_color="#304050",
        shape_comment=comment,
    )
    if kind in {"line", "connectLine"}:
        points = shape.element.xpath("./hc:pt0 | ./hc:pt1", namespaces={"hc": "http://www.hancom.co.kr/hwpml/2011/core"})
        if len(points) >= 2:
            points[0].set("x", "100")
            points[0].set("y", "200")
            points[1].set("x", "3100")
            points[1].set("y", "700")
            shape.section.mark_modified()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    return output_path


def run_probe(output_root: Path, *, timeout_seconds: int) -> dict[str, Any]:
    output_root.mkdir(parents=True, exist_ok=True)

    line_hwpx = _build_probe_source("line", output_root / "probe_line.hwpx")
    connect_hwpx = _build_probe_source("connectLine", output_root / "probe_connectline.hwpx")

    line_hwp = convert_document(line_hwpx, output_root / "probe_line_from_hancom.hwp", "HWP", timeout_seconds=timeout_seconds)
    connect_hwp = convert_document(
        connect_hwpx,
        output_root / "probe_connectline_from_hancom.hwp",
        "HWP",
        timeout_seconds=timeout_seconds,
    )

    line_doc = HwpDocument.open(line_hwp)
    connect_doc = HwpDocument.open(connect_hwp)
    line_shape = _find_shape_by_comment(line_doc, "PROBE-LINE")
    connect_shape = _find_shape_by_comment(connect_doc, "PROBE-CONNECTLINE")

    line_summary = _shape_summary(line_shape)
    connect_summary = _shape_summary(connect_shape)

    report = {
        "paths": {
            "line_hwpx": str(line_hwpx),
            "connect_hwpx": str(connect_hwpx),
            "line_hwp": str(line_hwp),
            "connect_hwp": str(connect_hwp),
        },
        "line_shape_count": len(line_doc.shapes()),
        "connectline_shape_count": len(connect_doc.shapes()),
        "line": line_summary,
        "connectLine": connect_summary,
        "comparisons": {
            "control_header": _comparison(
                "control_header",
                line_summary.get("control_header"),
                connect_summary.get("control_header"),
            ),
            "component_record": _comparison(
                "component_record",
                line_summary.get("component_record"),
                connect_summary.get("component_record"),
            ),
            "specific_record": _comparison(
                "specific_record",
                line_summary.get("specific_record"),
                connect_summary.get("specific_record"),
            ),
            "ctrl_data": _comparison(
                "ctrl_data",
                line_summary.get("ctrl_data"),
                connect_summary.get("ctrl_data"),
            ),
        },
    }
    report_path = output_root / "report.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description="Probe Hancom's native HWP mapping for connectLine by converting HWPX sources.")
    parser.add_argument(
        "--output-root",
        default=str(_default_output_root()),
        help="Directory where generated HWPX/HWP files and the JSON report will be written.",
    )
    parser.add_argument(
        "--timeout-seconds",
        type=int,
        default=120,
        help="Hancom conversion timeout in seconds.",
    )
    args = parser.parse_args()

    report = run_probe(Path(args.output_root).expanduser().resolve(), timeout_seconds=args.timeout_seconds)
    print(json.dumps(report, ensure_ascii=True, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
