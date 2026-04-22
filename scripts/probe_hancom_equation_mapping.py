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
from jakal_hwpx.hwp_binary import TAG_CTRL_DATA, TAG_EQEDIT, RecordNode, hwp_tag_name  # noqa: E402


def _default_output_root() -> Path:
    return REPO_ROOT / ".codex-temp" / "equation_probe"


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
        return {"type": "bytes", "hex": value.hex()}
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


def _eqedit_record(control_node: RecordNode) -> RecordNode | None:
    return next((child for child in control_node.iter_descendants() if child.tag_id == TAG_EQEDIT), None)


def _first_equation(document: HwpDocument):
    equations = document.equations()
    return equations[0] if equations else None


def _equation_summary(equation) -> dict[str, Any]:
    if equation is None:
        return {"present": False}
    return {
        "present": True,
        "shape_comment": equation.shape_comment,
        "script": equation.script,
        "size": equation.size(),
        "font": equation.font,
        "layout": equation.layout(),
        "out_margins": equation.out_margins(),
        "rotation": equation.rotation(),
        "control_header": _payload_summary(equation.control_node),
        "ctrl_data": _payload_summary(_first_ctrl_data(equation.control_node)),
        "eqedit": _payload_summary(_eqedit_record(equation.control_node)),
        "subtree": _subtree_payloads(equation.control_node),
    }


def _build_equation_probe(output_path: Path, *, variant: str) -> Path:
    document = HwpxDocument.blank()
    equation = document.append_equation("x=y+1", width=3000, height=1800, font="Batang")
    if variant == "layout":
        equation.set_layout(
            text_wrap="SQUARE",
            vert_rel_to="PAPER",
            horz_rel_to="PAPER",
            vert_align="TOP",
            horz_align="LEFT",
            vert_offset=12,
            horz_offset=34,
        )
    elif variant == "margins":
        equation.set_out_margins(left=5, right=6, top=7, bottom=8)
    elif variant == "rotation":
        equation.set_rotation(angle=35, center_x=555, center_y=666)
    elif variant != "baseline":
        raise ValueError(f"Unknown equation probe variant: {variant}")
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
            "eqedit": _comparison(
                f"{name}.eqedit",
                baseline.get("eqedit"),
                current.get("eqedit"),
            ),
            "subtree": _sequence_comparison(
                baseline.get("subtree", []),
                current.get("subtree", []),
            ),
        },
    }


def run_probe(output_root: Path, *, timeout_seconds: int) -> dict[str, Any]:
    output_root.mkdir(parents=True, exist_ok=True)

    hwpx_paths = {
        variant: _build_equation_probe(output_root / f"equation_{variant}.hwpx", variant=variant)
        for variant in ("baseline", "layout", "margins", "rotation")
    }
    hwp_paths = {
        variant: convert_document(
            hwpx_paths[variant],
            output_root / f"equation_{variant}_from_hancom.hwp",
            "HWP",
            timeout_seconds=timeout_seconds,
        )
        for variant in ("baseline", "layout", "margins", "rotation")
    }

    docs = {variant: HwpDocument.open(path) for variant, path in hwp_paths.items()}
    summaries = {variant: _equation_summary(_first_equation(docs[variant])) for variant in ("baseline", "layout", "margins", "rotation")}

    report = {
        "paths": {
            f"{variant}_hwpx": str(hwpx_paths[variant])
            for variant in ("baseline", "layout", "margins", "rotation")
        }
        | {
            f"{variant}_hwp": str(hwp_paths[variant])
            for variant in ("baseline", "layout", "margins", "rotation")
        },
        "layout": _variant_report("equation.layout", summaries["baseline"], summaries["layout"]),
        "margins": _variant_report("equation.margins", summaries["baseline"], summaries["margins"]),
        "rotation": _variant_report("equation.rotation", summaries["baseline"], summaries["rotation"]),
    }
    report_path = output_root / "report.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Probe Hancom's native HWP mapping for equation layout, out-margins, and rotation by converting HWPX sources."
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
