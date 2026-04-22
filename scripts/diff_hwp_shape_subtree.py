from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
from typing import Any

from jakal_hwpx.hwp_binary import TAG_CTRL_HEADER, RecordNode, _decode_control_id_payload, hwp_tag_name
from jakal_hwpx.hwp_document import HwpDocument


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _node_summary(node: RecordNode) -> dict[str, Any]:
    summary: dict[str, Any] = {
        "tag_id": node.tag_id,
        "tag_name": hwp_tag_name(node.tag_id),
        "level": node.level,
        "payload_len": len(node.payload),
        "payload_sha256": _sha256(node.payload),
    }
    if node.tag_id == TAG_CTRL_HEADER:
        summary["control_id"] = _decode_control_id_payload(node.payload)
    return summary


def _flatten_shape_nodes(shape) -> list[RecordNode]:
    return [shape.control_node, *list(shape.control_node.iter_descendants())]


def _shape_semantics(shape) -> dict[str, Any]:
    return {
        "kind": shape.kind,
        "text": shape.text,
        "size": shape.size(),
        "fill_color": getattr(shape, "fill_color", None),
        "line_color": getattr(shape, "line_color", None),
        "shape_comment": getattr(shape, "shape_comment", None),
        "descendant_tag_ids": shape.descendant_tag_ids(),
    }


def compare_shape_subtrees(base_path: Path, other_path: Path, *, shape_index: int = 0) -> dict[str, Any]:
    base_doc = HwpDocument.open(base_path)
    other_doc = HwpDocument.open(other_path)
    base_shape = base_doc.shapes()[shape_index]
    other_shape = other_doc.shapes()[shape_index]
    base_nodes = _flatten_shape_nodes(base_shape)
    other_nodes = _flatten_shape_nodes(other_shape)

    differences: list[dict[str, Any]] = []
    max_len = max(len(base_nodes), len(other_nodes))
    for index in range(max_len):
        base_node = base_nodes[index] if index < len(base_nodes) else None
        other_node = other_nodes[index] if index < len(other_nodes) else None
        if base_node is None or other_node is None:
            differences.append(
                {
                    "index": index,
                    "kind": "missing_node",
                    "base": _node_summary(base_node) if base_node is not None else None,
                    "other": _node_summary(other_node) if other_node is not None else None,
                }
            )
            continue
        if (
            base_node.tag_id == other_node.tag_id
            and base_node.level == other_node.level
            and base_node.payload == other_node.payload
        ):
            continue
        first_diff = None
        limit = min(len(base_node.payload), len(other_node.payload))
        for offset in range(limit):
            if base_node.payload[offset] != other_node.payload[offset]:
                first_diff = offset
                break
        if first_diff is None and len(base_node.payload) != len(other_node.payload):
            first_diff = limit
        differences.append(
            {
                "index": index,
                "kind": "node_diff",
                "base": _node_summary(base_node),
                "other": _node_summary(other_node),
                "payload_first_diff": first_diff,
            }
        )

    return {
        "base_path": str(base_path),
        "other_path": str(other_path),
        "shape_index": shape_index,
        "base_semantics": _shape_semantics(base_shape),
        "other_semantics": _shape_semantics(other_shape),
        "base_node_count": len(base_nodes),
        "other_node_count": len(other_nodes),
        "differences": differences,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Compare native HWP shape control subtrees between two documents.")
    parser.add_argument("base", help="Baseline HWP path.")
    parser.add_argument("other", help="Comparison HWP path.")
    parser.add_argument("--shape-index", type=int, default=0, help="Shape index to compare.")
    parser.add_argument("--output", help="Optional JSON output path.")
    args = parser.parse_args()

    comparison = compare_shape_subtrees(
        Path(args.base).expanduser().resolve(),
        Path(args.other).expanduser().resolve(),
        shape_index=args.shape_index,
    )
    text = json.dumps(comparison, ensure_ascii=False, indent=2)
    if args.output:
        output_path = Path(args.output).expanduser().resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(text, encoding="utf-8")
    else:
        print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
