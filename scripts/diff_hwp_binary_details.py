from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
from typing import Any

from jakal_hwpx.hwp_binary import (
    TAG_CTRL_HEADER,
    HwpBinaryDocument,
    HwpRecord,
    _decode_control_id_payload,
    hwp_tag_name,
)


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _first_diff_offset(left: bytes, right: bytes) -> int | None:
    limit = min(len(left), len(right))
    for index in range(limit):
        if left[index] != right[index]:
            return index
    if len(left) != len(right):
        return limit
    return None


def _record_summary(record: HwpRecord) -> dict[str, Any]:
    summary: dict[str, Any] = {
        "tag_id": record.tag_id,
        "tag_name": hwp_tag_name(record.tag_id),
        "level": record.level,
        "size": record.size,
        "payload_len": len(record.payload),
        "payload_sha256": _sha256(record.payload),
    }
    if record.tag_id == TAG_CTRL_HEADER:
        summary["control_id"] = _decode_control_id_payload(record.payload)
    return summary


def _record_identity(record: HwpRecord) -> tuple[int, int, str]:
    return (record.tag_id, record.level, _sha256(record.payload))


def _records_for_stream(document: HwpBinaryDocument, path: str) -> list[HwpRecord] | None:
    if path == "DocInfo":
        return document.docinfo_records()
    if path.startswith("BodyText/Section"):
        section_index = int(path.rsplit("Section", 1)[1])
        return document.section_records(section_index)
    return None


def _compare_records(base_records: list[HwpRecord], other_records: list[HwpRecord], *, max_diffs: int) -> dict[str, Any]:
    prefix = 0
    while (
        prefix < len(base_records)
        and prefix < len(other_records)
        and _record_identity(base_records[prefix]) == _record_identity(other_records[prefix])
    ):
        prefix += 1

    suffix = 0
    max_suffix = min(len(base_records), len(other_records)) - prefix
    while (
        suffix < max_suffix
        and _record_identity(base_records[len(base_records) - 1 - suffix])
        == _record_identity(other_records[len(other_records) - 1 - suffix])
    ):
        suffix += 1

    differences: list[dict[str, Any]] = []
    max_len = max(len(base_records), len(other_records))
    for index in range(max_len):
        if len(differences) >= max_diffs:
            break
        base_record = base_records[index] if index < len(base_records) else None
        other_record = other_records[index] if index < len(other_records) else None
        if base_record is None or other_record is None:
            differences.append(
                {
                    "index": index,
                    "kind": "missing_record",
                    "base": _record_summary(base_record) if base_record is not None else None,
                    "other": _record_summary(other_record) if other_record is not None else None,
                }
            )
            continue
        if (
            base_record.tag_id == other_record.tag_id
            and base_record.level == other_record.level
            and base_record.payload == other_record.payload
        ):
            continue
        differences.append(
            {
                "index": index,
                "kind": "record_diff",
                "base": _record_summary(base_record),
                "other": _record_summary(other_record),
                "payload_first_diff": _first_diff_offset(base_record.payload, other_record.payload),
            }
        )

    insertion_window: dict[str, Any] | None = None
    if prefix + suffix == len(base_records) or prefix + suffix == len(other_records):
        base_middle = base_records[prefix : len(base_records) - suffix if suffix else len(base_records)]
        other_middle = other_records[prefix : len(other_records) - suffix if suffix else len(other_records)]
        insertion_window = {
            "common_prefix_records": prefix,
            "common_suffix_records": suffix,
            "base_middle_count": len(base_middle),
            "other_middle_count": len(other_middle),
            "base_middle": [_record_summary(record) for record in base_middle[:max_diffs]],
            "other_middle": [_record_summary(record) for record in other_middle[:max_diffs]],
        }

    return {
        "base_record_count": len(base_records),
        "other_record_count": len(other_records),
        "equal": len(base_records) == len(other_records) and not differences,
        "common_prefix_records": prefix,
        "common_suffix_records": suffix,
        "middle_window": insertion_window,
        "differences": differences,
    }


def _layout_entry_map(document: HwpBinaryDocument) -> dict[str, dict[str, Any]]:
    layout = getattr(document, "_compound_layout", None)
    if layout is None:
        return {}
    entries: dict[str, dict[str, Any]] = {}
    for entry in layout.entries:
        key = entry.path if entry.path is not None else f"sid:{entry.sid}"
        entries[key] = {
            "sid": entry.sid,
            "name": entry.name,
            "entry_type": entry.entry_type,
            "left": entry.left,
            "right": entry.right,
            "child": entry.child,
            "start_sector": entry.start_sector,
            "size": entry.size,
            "color": entry.color,
            "state_bits": entry.state_bits,
            "create_time": entry.create_time,
            "modify_time": entry.modify_time,
        }
    return entries


def _compare_layout(base: HwpBinaryDocument, other: HwpBinaryDocument, *, max_diffs: int) -> dict[str, Any]:
    base_entries = _layout_entry_map(base)
    other_entries = _layout_entry_map(other)
    all_keys = sorted(set(base_entries) | set(other_entries))
    differences: list[dict[str, Any]] = []
    for key in all_keys:
        if len(differences) >= max_diffs:
            break
        base_entry = base_entries.get(key)
        other_entry = other_entries.get(key)
        if base_entry == other_entry:
            continue
        differences.append({"path": key, "base": base_entry, "other": other_entry})
    return {
        "base_entry_count": len(base_entries),
        "other_entry_count": len(other_entries),
        "equal": base_entries == other_entries,
        "differences": differences,
    }


def compare_documents(base_path: Path, other_path: Path, *, max_record_diffs: int = 20, max_layout_diffs: int = 20) -> dict[str, Any]:
    base = HwpBinaryDocument.open(base_path)
    other = HwpBinaryDocument.open(other_path)
    stream_paths = sorted(set(base.list_stream_paths()) | set(other.list_stream_paths()))
    stream_differences: list[dict[str, Any]] = []

    for path in stream_paths:
        base_has = base.has_stream(path)
        other_has = other.has_stream(path)
        if not base_has or not other_has:
            stream_differences.append(
                {
                    "path": path,
                    "kind": "missing_stream",
                    "base_has": base_has,
                    "other_has": other_has,
                }
            )
            continue

        base_raw = base.read_stream(path, decompress=False)
        other_raw = other.read_stream(path, decompress=False)
        base_logical = base.read_stream(path)
        other_logical = other.read_stream(path)

        raw_equal = base_raw == other_raw
        logical_equal = base_logical == other_logical
        if raw_equal and logical_equal:
            continue

        entry: dict[str, Any] = {
            "path": path,
            "kind": "stream_diff",
            "base_compressed_size": len(base_raw),
            "other_compressed_size": len(other_raw),
            "base_logical_size": len(base_logical),
            "other_logical_size": len(other_logical),
            "base_compressed_sha256": _sha256(base_raw),
            "other_compressed_sha256": _sha256(other_raw),
            "base_logical_sha256": _sha256(base_logical),
            "other_logical_sha256": _sha256(other_logical),
            "raw_equal": raw_equal,
            "logical_equal": logical_equal,
            "compressed_first_diff": _first_diff_offset(base_raw, other_raw),
            "logical_first_diff": _first_diff_offset(base_logical, other_logical),
        }
        records_base = _records_for_stream(base, path)
        records_other = _records_for_stream(other, path)
        if records_base is not None and records_other is not None:
            entry["record_compare"] = _compare_records(records_base, records_other, max_diffs=max_record_diffs)
        stream_differences.append(entry)

    return {
        "base_path": str(base_path),
        "other_path": str(other_path),
        "stream_difference_count": len(stream_differences),
        "stream_differences": stream_differences,
        "layout_compare": _compare_layout(base, other, max_diffs=max_layout_diffs),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Compare low-level HWP binary stream and record differences.")
    parser.add_argument("base", help="Baseline HWP path.")
    parser.add_argument("others", nargs="+", help="One or more HWP paths to compare against the baseline.")
    parser.add_argument("--output", help="Optional JSON output path.")
    parser.add_argument("--max-record-diffs", type=int, default=20, help="Maximum per-stream record diffs to report.")
    parser.add_argument("--max-layout-diffs", type=int, default=20, help="Maximum compound-layout diffs to report.")
    args = parser.parse_args()

    base_path = Path(args.base).expanduser().resolve()
    comparisons = [
        compare_documents(
            base_path,
            Path(other).expanduser().resolve(),
            max_record_diffs=args.max_record_diffs,
            max_layout_diffs=args.max_layout_diffs,
        )
        for other in args.others
    ]
    text = json.dumps(comparisons, ensure_ascii=False, indent=2)
    if args.output:
        output_path = Path(args.output).expanduser().resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(text, encoding="utf-8")
    else:
        print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
