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
    return REPO_ROOT / ".codex-temp" / "graphic_layout_flags_probe"


def _first_diff_offset(left: bytes, right: bytes) -> int | None:
    limit = min(len(left), len(right))
    for index in range(limit):
        if left[index] != right[index]:
            return index
    if len(left) != len(right):
        return limit
    return None


def _byte_diffs(left: bytes, right: bytes) -> list[dict[str, int]]:
    result: list[dict[str, int]] = []
    limit = min(len(left), len(right))
    for index in range(limit):
        if left[index] != right[index]:
            result.append({"offset": index, "left": left[index], "right": right[index]})
    if len(left) != len(right):
        longer = left if len(left) > len(right) else right
        for index in range(limit, len(longer)):
            result.append(
                {
                    "offset": index,
                    "left": left[index] if index < len(left) else -1,
                    "right": right[index] if index < len(right) else -1,
                }
            )
    return result


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
        return {"type": "bytes", "hex": value.hex()}
    if isinstance(value, dict):
        return {str(key): _json_safe(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_json_safe(item) for item in value]
    return value


def _subtree_payloads(control_node: RecordNode) -> list[dict[str, Any]]:
    return [
        {
            "tag_id": node.tag_id,
            "tag_name": hwp_tag_name(node.tag_id),
            "payload_len": len(node.payload),
            "payload_hex": node.payload.hex(),
        }
        for node in control_node.iter_preorder()
    ]


def _first_ctrl_data(control_node: RecordNode) -> RecordNode | None:
    return next((child for child in control_node.children if child.tag_id == TAG_CTRL_DATA), None)


def _comparison(name: str, left: dict[str, Any] | None, right: dict[str, Any] | None) -> dict[str, Any]:
    left_payload = bytes.fromhex(left["payload_hex"]) if left is not None and "payload_hex" in left else b""
    right_payload = bytes.fromhex(right["payload_hex"]) if right is not None and "payload_hex" in right else b""
    return {
        "name": name,
        "same_tag": (left or {}).get("tag_name") == (right or {}).get("tag_name"),
        "same_length": len(left_payload) == len(right_payload),
        "same_payload": left_payload == right_payload,
        "first_diff_offset": _first_diff_offset(left_payload, right_payload),
        "byte_diffs": _byte_diffs(left_payload, right_payload),
    }


def _find_picture_by_comment(document: HwpDocument, comment: str):
    for picture in document.pictures():
        if picture.shape_comment == comment:
            return picture
    return None


def _picture_summary(picture) -> dict[str, Any]:
    if picture is None:
        return {"present": False}
    component = picture._shape_component_record()
    specific = picture._shape_picture_record()
    return {
        "present": True,
        "shape_comment": picture.shape_comment,
        "size": picture.size(),
        "layout": picture.layout(),
        "control_header": _payload_summary(picture.control_node),
        "ctrl_data": _payload_summary(_first_ctrl_data(picture.control_node)),
        "component_record": _payload_summary(component),
        "specific_record": _payload_summary(specific),
        "subtree": _subtree_payloads(picture.control_node),
    }


def _build_variant(output_path: Path, *, variant: str) -> Path:
    image_bytes = HwpDocument.blank().binary_document().read_stream("BinData/BIN0001.bmp", decompress=False)
    document = HwpxDocument.blank()
    picture = document.append_picture("probe-layout-flags.bmp", image_bytes, width=2400, height=1600)
    picture.shape_comment = f"PROBE-LAYOUT-{variant.upper()}"
    if variant == "text_flow":
        picture.set_layout(text_flow="RIGHT_ONLY")
    elif variant == "treat_as_char":
        picture.set_layout(treat_as_char=True)
    elif variant == "vert_align":
        picture.set_layout(vert_align="CENTER")
    elif variant == "horz_align":
        picture.set_layout(horz_align="RIGHT")
    elif variant != "baseline":
        raise ValueError(f"Unknown variant: {variant}")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    return output_path


def _variant_report(name: str, baseline: dict[str, Any], current: dict[str, Any]) -> dict[str, Any]:
    return {
        "baseline": _json_safe(baseline),
        "mutated": _json_safe(current),
        "comparisons": {
            "control_header": _comparison(
                f"{name}.control_header",
                baseline.get("control_header"),
                current.get("control_header"),
            ),
            "ctrl_data": _comparison(
                f"{name}.ctrl_data",
                baseline.get("ctrl_data"),
                current.get("ctrl_data"),
            ),
            "component_record": _comparison(
                f"{name}.component_record",
                baseline.get("component_record"),
                current.get("component_record"),
            ),
            "specific_record": _comparison(
                f"{name}.specific_record",
                baseline.get("specific_record"),
                current.get("specific_record"),
            ),
        },
    }


def run_probe(output_root: Path, *, timeout_seconds: int) -> dict[str, Any]:
    output_root.mkdir(parents=True, exist_ok=True)
    variants = ("baseline", "text_flow", "treat_as_char", "vert_align", "horz_align")
    hwpx_paths = {variant: _build_variant(output_root / f"{variant}.hwpx", variant=variant) for variant in variants}
    hwp_paths = {
        variant: convert_document(
            hwpx_paths[variant],
            output_root / f"{variant}_from_hancom.hwp",
            "HWP",
            timeout_seconds=timeout_seconds,
        )
        for variant in variants
    }
    docs = {variant: HwpDocument.open(path) for variant, path in hwp_paths.items()}
    summaries = {
        variant: _picture_summary(_find_picture_by_comment(docs[variant], f"PROBE-LAYOUT-{variant.upper()}"))
        for variant in variants
    }
    baseline = summaries["baseline"]
    report = {
        "paths": {f"{variant}_hwpx": str(hwpx_paths[variant]) for variant in variants}
        | {f"{variant}_hwp": str(hwp_paths[variant]) for variant in variants},
        "text_flow": _variant_report("picture.text_flow", baseline, summaries["text_flow"]),
        "treat_as_char": _variant_report("picture.treat_as_char", baseline, summaries["treat_as_char"]),
        "vert_align": _variant_report("picture.vert_align", baseline, summaries["vert_align"]),
        "horz_align": _variant_report("picture.horz_align", baseline, summaries["horz_align"]),
    }
    report_path = output_root / "report.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Probe Hancom's native HWP mapping for common graphic layout flags by converting HWPX picture sources."
    )
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
