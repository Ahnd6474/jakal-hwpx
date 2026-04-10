from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path

from .hwp_binary import HwpBinaryDocument, HwpRecord
from .hwp_collection import find_best_combo_donor, scan_hwp_collection
from .hwp_template_lab import CONTROL_BY_FEATURE, _paragraph_ranges, build_minimal_control_candidate, find_control_occurrences


BUNDLED_PROFILE_DIR = Path(__file__).resolve().with_name("bundled_hwp_profile")


@dataclass(frozen=True)
class HwpPureProfile:
    root: Path
    base_path: Path
    template_paths: dict[str, Path]
    donor_path: Path
    target_section_index: int
    template_section_indices: dict[str, int]

    @classmethod
    def load(cls, root: str | Path) -> "HwpPureProfile":
        resolved_root = Path(root).expanduser().resolve()
        metadata = json.loads((resolved_root / "profile.json").read_text(encoding="utf-8"))
        return cls(
            root=resolved_root,
            base_path=resolved_root / metadata["base_path"],
            template_paths={key: resolved_root / value for key, value in metadata["template_paths"].items()},
            donor_path=Path(metadata.get("donor_path", metadata.get("donor_name", ""))),
            target_section_index=int(metadata["target_section_index"]),
            template_section_indices={key: int(value) for key, value in metadata["template_section_indices"].items()},
        )

    @classmethod
    def load_bundled(cls) -> "HwpPureProfile":
        return cls.load(BUNDLED_PROFILE_DIR)


def build_hwp_pure_profile(collection_root: str | Path, output_dir: str | Path) -> HwpPureProfile:
    donor = find_best_combo_donor(collection_root)
    if donor is None:
        raise RuntimeError("No donor in the collection contains table, picture, and hyperlink controls together.")

    target_root = Path(output_dir).expanduser().resolve()
    target_root.mkdir(parents=True, exist_ok=True)

    donor_document = HwpBinaryDocument.open(donor.path)
    target_section_index = _pick_target_section_index(donor_document)
    base_path = target_root / "base.hwp"
    _build_base_document(donor_document, base_path, target_section_index=target_section_index)

    template_paths: dict[str, Path] = {}
    template_section_indices: dict[str, int] = {}
    collection_summaries = scan_hwp_collection(collection_root)
    for feature, control_id in CONTROL_BY_FEATURE.items():
        best_source = _pick_best_feature_source(collection_summaries, control_id)
        if best_source is None:
            continue
        feature_document = HwpBinaryDocument.open(best_source[0].path)
        occurrence = best_source[1]
        template_path = target_root / f"{feature}.hwp"
        build_minimal_control_candidate(feature_document, occurrence, template_path, keep_first_paragraph=False)
        template_paths[feature] = template_path
        template_section_indices[feature] = occurrence.section_index

    metadata = {
        "donor_name": donor.path.name,
        "base_path": base_path.name,
        "template_paths": {key: value.name for key, value in template_paths.items()},
        "target_section_index": target_section_index,
        "template_section_indices": template_section_indices,
    }
    (target_root / "profile.json").write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
    return HwpPureProfile(
        root=target_root,
        base_path=base_path,
        template_paths=template_paths,
        donor_path=donor.path,
        target_section_index=target_section_index,
        template_section_indices=template_section_indices,
    )


def bundled_hwp_profile_root() -> Path:
    return BUNDLED_PROFILE_DIR


def append_feature_from_profile(document: HwpBinaryDocument, profile: HwpPureProfile, feature: str) -> None:
    if feature not in profile.template_paths:
        raise KeyError(feature)
    template_document = HwpBinaryDocument.open(profile.template_paths[feature])
    template_section_index = profile.template_section_indices[feature]
    template_records = template_document.section_records(template_section_index)
    template_ranges = _paragraph_ranges(template_records)
    if not template_ranges:
        raise RuntimeError(f"Template for {feature} does not contain any paragraph range.")

    append_ranges = template_ranges[1:] if len(template_ranges) > 1 else template_ranges
    body_records = document.section_records(profile.target_section_index)
    for start, end in append_ranges:
        body_records.extend(_clone_records(template_records[start : end + 1]))
    document.replace_section_records(profile.target_section_index, body_records)


def _build_base_document(document: HwpBinaryDocument, output_path: Path, *, target_section_index: int) -> Path:
    base_document = HwpBinaryDocument.open(document.source_path)
    records = base_document.section_records(target_section_index)
    ranges = _paragraph_ranges(records)
    if not ranges:
        raise RuntimeError("Donor document does not contain any top-level paragraph range.")
    first_start, first_end = ranges[0]
    base_document.replace_section_records(target_section_index, _clone_records(records[first_start : first_end + 1]))
    base_document.save(output_path)
    return output_path


def _pick_target_section_index(document: HwpBinaryDocument) -> int:
    section_sets: list[set[int]] = []
    for control_id in CONTROL_BY_FEATURE.values():
        section_sets.append({occurrence.section_index for occurrence in find_control_occurrences(document, control_id)})
    common = set.intersection(*section_sets) if section_sets else {0}
    if not common:
        raise RuntimeError("No section in the donor contains all required pure-profile features.")
    return max(common, key=lambda index: document.stream_size(document.section_stream_paths()[index]))


def _clone_records(records: list[HwpRecord]) -> list[HwpRecord]:
    return [
        HwpRecord(
            tag_id=record.tag_id,
            level=record.level,
            size=record.size,
            header_size=record.header_size,
            offset=record.offset,
            payload=bytes(record.payload),
        )
        for record in records
    ]


def _pick_best_occurrence(occurrences):
    return sorted(
        occurrences,
        key=lambda occurrence: (
            occurrence.paragraph_end_index - occurrence.paragraph_start_index,
            occurrence.level,
            occurrence.section_index,
            occurrence.record_index,
        ),
    )[0]


def _pick_best_feature_source(collection_summaries, control_id: str):
    candidates: list[tuple[object, object]] = []
    for summary in collection_summaries:
        if summary.control_counts.get(control_id, 0) <= 0:
            continue
        document = HwpBinaryDocument.open(summary.path)
        occurrences = find_control_occurrences(document, control_id)
        if not occurrences:
            continue
        candidates.append((summary, _pick_best_occurrence(occurrences)))
    if not candidates:
        return None
    return sorted(
        candidates,
        key=lambda item: (
            item[1].paragraph_end_index - item[1].paragraph_start_index,
            item[1].level,
            item[1].section_index,
            item[1].record_index,
            item[0].path.name,
        ),
    )[0]
