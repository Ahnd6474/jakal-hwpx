from __future__ import annotations

import zlib
from collections import Counter
from dataclasses import dataclass
from pathlib import Path

import olefile


@dataclass(frozen=True)
class HwpDonorSummary:
    path: Path
    control_counts: dict[str, int]
    tag_counts: dict[int, int]

    @property
    def has_table(self) -> bool:
        return self.control_counts.get("tbl ", 0) > 0

    @property
    def has_picture(self) -> bool:
        return self.control_counts.get("gso ", 0) > 0 and self.tag_counts.get(85, 0) > 0

    @property
    def has_hyperlink(self) -> bool:
        return self.control_counts.get("%hlk", 0) > 0


def scan_hwp_collection(root: str | Path) -> list[HwpDonorSummary]:
    collection_root = Path(root).expanduser().resolve()
    summaries: list[HwpDonorSummary] = []
    for path in sorted(collection_root.rglob("*.hwp")):
        if not olefile.isOleFile(str(path)):
            continue
        try:
            summaries.append(_summarize_hwp(path))
        except Exception:
            continue
    return summaries


def find_best_table_donor(root: str | Path) -> HwpDonorSummary | None:
    return _pick_best(scan_hwp_collection(root), predicate=lambda item: item.has_table)


def find_best_picture_donor(root: str | Path) -> HwpDonorSummary | None:
    return _pick_best(scan_hwp_collection(root), predicate=lambda item: item.has_picture)


def find_best_hyperlink_donor(root: str | Path) -> HwpDonorSummary | None:
    return _pick_best(scan_hwp_collection(root), predicate=lambda item: item.has_hyperlink)


def find_best_combo_donor(root: str | Path) -> HwpDonorSummary | None:
    items = scan_hwp_collection(root)
    candidates = [item for item in items if item.has_table and item.has_picture and item.has_hyperlink]
    if not candidates:
        return None
    candidates.sort(
        key=lambda item: (
            -(item.control_counts.get("tbl ", 0) + item.tag_counts.get(85, 0) + item.control_counts.get("%hlk", 0)),
            item.path.name,
        )
    )
    return candidates[0]


def _pick_best(items: list[HwpDonorSummary], *, predicate) -> HwpDonorSummary | None:
    candidates = [item for item in items if predicate(item)]
    if not candidates:
        return None
    candidates.sort(
        key=lambda item: (
            -(item.control_counts.get("tbl ", 0) if item.has_table else 0),
            -(item.tag_counts.get(85, 0) if item.has_picture else 0),
            -(item.control_counts.get("%hlk", 0) if item.has_hyperlink else 0),
            item.path.name,
        )
    )
    return candidates[0]


def _summarize_hwp(path: Path) -> HwpDonorSummary:
    with olefile.OleFileIO(str(path)) as ole:
        compressed = bool(int.from_bytes(ole.openstream("FileHeader").read()[36:40], "little") & 1)
        control_counts: Counter[str] = Counter()
        tag_counts: Counter[int] = Counter()

        for entry in ole.listdir(streams=True, storages=False):
            if not entry or entry[0] != "BodyText":
                continue
            raw = ole.openstream(entry).read()
            if compressed:
                raw = zlib.decompress(raw, -15)
            offset = 0
            while offset + 4 <= len(raw):
                header = int.from_bytes(raw[offset : offset + 4], "little")
                tag_id = header & 0x3FF
                size = (header >> 20) & 0xFFF
                header_size = 4
                if size == 0xFFF:
                    size = int.from_bytes(raw[offset + 4 : offset + 8], "little")
                    header_size = 8
                payload_offset = offset + header_size
                payload = raw[payload_offset : payload_offset + size]
                tag_counts[tag_id] += 1
                if tag_id == 71 and len(payload) >= 4:
                    control_id = payload[:4][::-1].decode("latin1", errors="replace")
                    control_counts[control_id] += 1
                offset = payload_offset + size

    return HwpDonorSummary(
        path=path,
        control_counts=dict(control_counts),
        tag_counts=dict(tag_counts),
    )
