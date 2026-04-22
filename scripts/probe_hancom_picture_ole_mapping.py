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
    return REPO_ROOT / ".codex-temp" / "picture_ole_probe"


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


def _sequence_comparison(left: list[dict[str, Any]], right: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    limit = max(len(left), len(right))
    for index in range(limit):
        left_item = left[index] if index < len(left) else None
        right_item = right[index] if index < len(right) else None
        result.append(
            {
                "index": index,
                "left_tag": None if left_item is None else left_item["tag_name"],
                "right_tag": None if right_item is None else right_item["tag_name"],
                "left_len": None if left_item is None else left_item["payload_len"],
                "right_len": None if right_item is None else right_item["payload_len"],
                "first_diff_offset": _first_diff_offset(
                    b"" if left_item is None else bytes.fromhex(left_item["payload_hex"]),
                    b"" if right_item is None else bytes.fromhex(right_item["payload_hex"]),
                ),
            }
        )
    return result


def _first_ctrl_data(control_node: RecordNode) -> RecordNode | None:
    return next((child for child in control_node.children if child.tag_id == TAG_CTRL_DATA), None)


def _find_picture_by_comment(document: HwpDocument, comment: str):
    for picture in document.pictures():
        if picture.shape_comment == comment:
            return picture
    return None


def _find_ole_by_comment(document: HwpDocument, comment: str):
    for ole in document.oles():
        if ole.shape_comment == comment:
            return ole
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
        "line_style": picture.line_style(),
        "layout": picture.layout(),
        "out_margins": picture.out_margins(),
        "rotation": picture.rotation(),
        "image_adjustment": picture.image_adjustment(),
        "crop": picture.crop(),
        "control_header": _payload_summary(picture.control_node),
        "ctrl_data": _payload_summary(_first_ctrl_data(picture.control_node)),
        "component_record": _payload_summary(component),
        "specific_record": _payload_summary(specific),
        "subtree": _subtree_payloads(picture.control_node),
    }


def _ole_summary(ole) -> dict[str, Any]:
    if ole is None:
        return {"present": False}
    component = ole._shape_component_record()
    specific = ole._ole_record()
    return {
        "present": True,
        "shape_comment": ole.shape_comment,
        "size": ole.size(),
        "object_type": ole.object_type,
        "draw_aspect": ole.draw_aspect,
        "has_moniker": ole.has_moniker,
        "eq_baseline": ole.eq_baseline,
        "line_style": ole.line_style(),
        "layout": ole.layout(),
        "out_margins": ole.out_margins(),
        "rotation": ole.rotation(),
        "extent": ole.extent(),
        "control_header": _payload_summary(ole.control_node),
        "ctrl_data": _payload_summary(_first_ctrl_data(ole.control_node)),
        "component_record": _payload_summary(component),
        "specific_record": _payload_summary(specific),
        "subtree": _subtree_payloads(ole.control_node),
    }


def _build_picture_probe(output_path: Path, *, mutated: bool) -> Path:
    image_bytes = HwpDocument.blank().binary_document().read_stream("BinData/BIN0001.bmp", decompress=False)
    document = HwpxDocument.blank()
    picture = document.append_picture("probe-picture.bmp", image_bytes, width=2400, height=1600)
    picture.shape_comment = "PROBE-PICTURE-MUTATED" if mutated else "PROBE-PICTURE-BASELINE"
    if mutated:
        picture.set_layout(
            text_wrap="SQUARE",
            vert_rel_to="PAPER",
            horz_rel_to="PAPER",
            vert_align="TOP",
            horz_align="LEFT",
            vert_offset=101,
            horz_offset=202,
        )
        picture.set_out_margins(left=13, right=24, top=35, bottom=46)
        picture.set_rotation(angle=15, center_x=111, center_y=222, rotate_image=False)
        picture.set_image_adjustment(bright=5, contrast=6, effect="GRAY_SCALE", alpha=7)
        picture.set_crop(left=1, right=2399, top=2, bottom=1598)
        picture.set_line_style(color="#112233", width=55, style="SOLID")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    return output_path


def _build_ole_probe(output_path: Path, *, mutated: bool) -> Path:
    document = HwpxDocument.blank()
    ole = document.append_ole("probe-ole.ole", b"OLE-PROBE-DATA", width=42001, height=13501)
    ole.shape_comment = "PROBE-OLE-MUTATED" if mutated else "PROBE-OLE-BASELINE"
    if mutated:
        ole.set_layout(
            text_wrap="SQUARE",
            vert_rel_to="PAPER",
            horz_rel_to="PAPER",
            vert_align="TOP",
            horz_align="LEFT",
            vert_offset=77,
            horz_offset=88,
        )
        ole.set_out_margins(left=9, right=10, top=11, bottom=12)
        ole.set_rotation(angle=45, center_x=777, center_y=888)
        ole.set_line_style(color="#445566", width=77, style="DASH")
        ole.set_object_metadata(object_type="LINK", draw_aspect="ICON", has_moniker=True, eq_baseline=12)
        ole.set_extent(x=43000, y=14000)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    return output_path


def run_probe(output_root: Path, *, timeout_seconds: int) -> dict[str, Any]:
    output_root.mkdir(parents=True, exist_ok=True)

    picture_baseline_hwpx = _build_picture_probe(output_root / "picture_baseline.hwpx", mutated=False)
    picture_mutated_hwpx = _build_picture_probe(output_root / "picture_mutated.hwpx", mutated=True)
    ole_baseline_hwpx = _build_ole_probe(output_root / "ole_baseline.hwpx", mutated=False)
    ole_mutated_hwpx = _build_ole_probe(output_root / "ole_mutated.hwpx", mutated=True)

    picture_baseline_hwp = convert_document(
        picture_baseline_hwpx,
        output_root / "picture_baseline_from_hancom.hwp",
        "HWP",
        timeout_seconds=timeout_seconds,
    )
    picture_mutated_hwp = convert_document(
        picture_mutated_hwpx,
        output_root / "picture_mutated_from_hancom.hwp",
        "HWP",
        timeout_seconds=timeout_seconds,
    )
    ole_baseline_hwp = convert_document(
        ole_baseline_hwpx,
        output_root / "ole_baseline_from_hancom.hwp",
        "HWP",
        timeout_seconds=timeout_seconds,
    )
    ole_mutated_hwp = convert_document(
        ole_mutated_hwpx,
        output_root / "ole_mutated_from_hancom.hwp",
        "HWP",
        timeout_seconds=timeout_seconds,
    )

    picture_baseline_doc = HwpDocument.open(picture_baseline_hwp)
    picture_mutated_doc = HwpDocument.open(picture_mutated_hwp)
    ole_baseline_doc = HwpDocument.open(ole_baseline_hwp)
    ole_mutated_doc = HwpDocument.open(ole_mutated_hwp)

    picture_baseline = _picture_summary(_find_picture_by_comment(picture_baseline_doc, "PROBE-PICTURE-BASELINE"))
    picture_mutated = _picture_summary(_find_picture_by_comment(picture_mutated_doc, "PROBE-PICTURE-MUTATED"))
    ole_baseline = _ole_summary(_find_ole_by_comment(ole_baseline_doc, "PROBE-OLE-BASELINE"))
    ole_mutated = _ole_summary(_find_ole_by_comment(ole_mutated_doc, "PROBE-OLE-MUTATED"))

    report = {
        "paths": {
            "picture_baseline_hwpx": str(picture_baseline_hwpx),
            "picture_mutated_hwpx": str(picture_mutated_hwpx),
            "picture_baseline_hwp": str(picture_baseline_hwp),
            "picture_mutated_hwp": str(picture_mutated_hwp),
            "ole_baseline_hwpx": str(ole_baseline_hwpx),
            "ole_mutated_hwpx": str(ole_mutated_hwpx),
            "ole_baseline_hwp": str(ole_baseline_hwp),
            "ole_mutated_hwp": str(ole_mutated_hwp),
        },
        "picture": {
            "baseline": _json_safe(picture_baseline),
            "mutated": _json_safe(picture_mutated),
            "comparisons": {
                "control_header": _comparison(
                    "picture.control_header",
                    picture_baseline.get("control_header"),
                    picture_mutated.get("control_header"),
                ),
                "ctrl_data": _comparison(
                    "picture.ctrl_data",
                    picture_baseline.get("ctrl_data"),
                    picture_mutated.get("ctrl_data"),
                ),
                "component_record": _comparison(
                    "picture.component_record",
                    picture_baseline.get("component_record"),
                    picture_mutated.get("component_record"),
                ),
                "specific_record": _comparison(
                    "picture.specific_record",
                    picture_baseline.get("specific_record"),
                    picture_mutated.get("specific_record"),
                ),
                "subtree": _sequence_comparison(
                    picture_baseline.get("subtree", []),
                    picture_mutated.get("subtree", []),
                ),
            },
        },
        "ole": {
            "baseline": _json_safe(ole_baseline),
            "mutated": _json_safe(ole_mutated),
            "comparisons": {
                "control_header": _comparison(
                    "ole.control_header",
                    ole_baseline.get("control_header"),
                    ole_mutated.get("control_header"),
                ),
                "ctrl_data": _comparison(
                    "ole.ctrl_data",
                    ole_baseline.get("ctrl_data"),
                    ole_mutated.get("ctrl_data"),
                ),
                "component_record": _comparison(
                    "ole.component_record",
                    ole_baseline.get("component_record"),
                    ole_mutated.get("component_record"),
                ),
                "specific_record": _comparison(
                    "ole.specific_record",
                    ole_baseline.get("specific_record"),
                    ole_mutated.get("specific_record"),
                ),
                "subtree": _sequence_comparison(
                    ole_baseline.get("subtree", []),
                    ole_mutated.get("subtree", []),
                ),
            },
        },
    }
    report_path = output_root / "report.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Probe Hancom's native HWP mapping for picture crop/layout and OLE extent/layout by converting HWPX sources."
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
