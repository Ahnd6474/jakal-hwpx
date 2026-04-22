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
    return REPO_ROOT / ".codex-temp" / "equation_flags_probe"


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


def _first_ctrl_data(control_node: RecordNode) -> RecordNode | None:
    return next((child for child in control_node.children if child.tag_id == TAG_CTRL_DATA), None)


def _eqedit_record(control_node: RecordNode) -> RecordNode | None:
    return next((child for child in control_node.iter_descendants() if child.tag_id == TAG_EQEDIT), None)


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


def _first_equation(document: HwpDocument):
    equations = document.equations()
    return equations[0] if equations else None


def _equation_summary(equation) -> dict[str, Any]:
    if equation is None:
        return {"present": False}
    return {
        "present": True,
        "script": equation.script,
        "size": equation.size(),
        "font": equation.font,
        "layout": equation.layout(),
        "control_header": _payload_summary(equation.control_node),
        "ctrl_data": _payload_summary(_first_ctrl_data(equation.control_node)),
        "eqedit": _payload_summary(_eqedit_record(equation.control_node)),
    }


def _build_variant(output_path: Path, *, variant: str) -> Path:
    document = HwpxDocument.blank()
    equation = document.append_equation("x=y+1", width=3000, height=1800, font="Batang")
    if variant == "baseline_no_char":
        equation.set_layout(treat_as_char=False)
    elif variant == "top_text_flow":
        equation.set_layout(text_flow="RIGHT_ONLY")
    elif variant == "top_text_flow_left":
        equation.set_layout(text_flow="LEFT_ONLY")
    elif variant == "top_text_flow_no_char":
        equation.set_layout(treat_as_char=False, text_flow="RIGHT_ONLY")
    elif variant == "top_text_flow_left_no_char":
        equation.set_layout(treat_as_char=False, text_flow="LEFT_ONLY")
    elif variant == "square_baseline":
        equation.set_layout(text_wrap="SQUARE")
    elif variant == "square_no_char":
        equation.set_layout(text_wrap="SQUARE", treat_as_char=False)
    elif variant == "text_flow":
        equation.set_layout(text_wrap="SQUARE", text_flow="RIGHT_ONLY")
    elif variant == "text_flow_left":
        equation.set_layout(text_wrap="SQUARE", text_flow="LEFT_ONLY")
    elif variant == "text_flow_no_char":
        equation.set_layout(text_wrap="SQUARE", treat_as_char=False, text_flow="RIGHT_ONLY")
    elif variant == "text_flow_left_no_char":
        equation.set_layout(text_wrap="SQUARE", treat_as_char=False, text_flow="LEFT_ONLY")
    elif variant == "square_treat_as_char":
        equation.set_layout(text_wrap="SQUARE", treat_as_char=True)
    elif variant == "square_vert_align":
        equation.set_layout(text_wrap="SQUARE", vert_align="CENTER")
    elif variant == "square_horz_align":
        equation.set_layout(text_wrap="SQUARE", horz_align="RIGHT")
    elif variant == "treat_as_char":
        equation.set_layout(treat_as_char=True)
    elif variant == "vert_align":
        equation.set_layout(vert_align="CENTER")
    elif variant == "horz_align":
        equation.set_layout(horz_align="RIGHT")
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
            "eqedit": _comparison(
                f"{name}.eqedit",
                baseline.get("eqedit"),
                current.get("eqedit"),
            ),
        },
    }


def run_probe(output_root: Path, *, timeout_seconds: int) -> dict[str, Any]:
    output_root.mkdir(parents=True, exist_ok=True)
    variants = (
        "baseline",
        "baseline_no_char",
        "top_text_flow",
        "top_text_flow_left",
        "top_text_flow_no_char",
        "top_text_flow_left_no_char",
        "square_baseline",
        "square_no_char",
        "text_flow",
        "text_flow_left",
        "text_flow_no_char",
        "text_flow_left_no_char",
        "square_treat_as_char",
        "square_vert_align",
        "square_horz_align",
        "treat_as_char",
        "vert_align",
        "horz_align",
    )
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
    summaries = {variant: _equation_summary(_first_equation(docs[variant])) for variant in variants}
    report = {
        "paths": {f"{variant}_hwpx": str(hwpx_paths[variant]) for variant in variants}
        | {f"{variant}_hwp": str(hwp_paths[variant]) for variant in variants},
        "baseline_no_char": _variant_report("equation.baseline_no_char", summaries["baseline"], summaries["baseline_no_char"]),
        "top_text_flow": _variant_report("equation.top_text_flow", summaries["baseline"], summaries["top_text_flow"]),
        "top_text_flow_left": _variant_report("equation.top_text_flow_left", summaries["baseline"], summaries["top_text_flow_left"]),
        "top_text_flow_no_char": _variant_report(
            "equation.top_text_flow_no_char",
            summaries["baseline_no_char"],
            summaries["top_text_flow_no_char"],
        ),
        "top_text_flow_left_no_char": _variant_report(
            "equation.top_text_flow_left_no_char",
            summaries["baseline_no_char"],
            summaries["top_text_flow_left_no_char"],
        ),
        "square_baseline": _variant_report("equation.square_baseline", summaries["baseline"], summaries["square_baseline"]),
        "square_no_char": _variant_report("equation.square_no_char", summaries["baseline_no_char"], summaries["square_no_char"]),
        "text_flow": _variant_report("equation.text_flow", summaries["square_baseline"], summaries["text_flow"]),
        "text_flow_left": _variant_report("equation.text_flow_left", summaries["square_baseline"], summaries["text_flow_left"]),
        "text_flow_no_char": _variant_report(
            "equation.text_flow_no_char",
            summaries["square_no_char"],
            summaries["text_flow_no_char"],
        ),
        "text_flow_left_no_char": _variant_report(
            "equation.text_flow_left_no_char",
            summaries["square_no_char"],
            summaries["text_flow_left_no_char"],
        ),
        "square_treat_as_char": _variant_report(
            "equation.square_treat_as_char",
            summaries["square_baseline"],
            summaries["square_treat_as_char"],
        ),
        "square_vert_align": _variant_report(
            "equation.square_vert_align",
            summaries["square_baseline"],
            summaries["square_vert_align"],
        ),
        "square_horz_align": _variant_report(
            "equation.square_horz_align",
            summaries["square_baseline"],
            summaries["square_horz_align"],
        ),
        "treat_as_char": _variant_report("equation.treat_as_char", summaries["baseline"], summaries["treat_as_char"]),
        "vert_align": _variant_report("equation.vert_align", summaries["baseline"], summaries["vert_align"]),
        "horz_align": _variant_report("equation.horz_align", summaries["baseline"], summaries["horz_align"]),
    }
    report_path = output_root / "report.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Probe Hancom's native HWP mapping for equation layout flags by converting HWPX sources."
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
