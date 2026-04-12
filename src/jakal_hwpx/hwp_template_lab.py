from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from ._hancom import convert_document
from .hwp_binary import HwpBinaryDocument, HwpRecord
from .hwp_collection import (
    HwpDonorSummary,
    find_best_hyperlink_donor,
    find_best_picture_donor,
    find_best_table_donor,
)


Validator = Callable[[str | Path, str | Path, str], Path]


CONTROL_BY_FEATURE = {
    "table": "tbl ",
    "picture": "gso ",
    "hyperlink": "%hlk",
}


@dataclass(frozen=True)
class HwpControlOccurrence:
    control_id: str
    section_index: int
    record_index: int
    level: int
    paragraph_start_index: int
    paragraph_end_index: int


@dataclass(frozen=True)
class HwpTemplateCandidate:
    source_path: Path
    feature: str
    control_id: str
    section_index: int
    occurrence_index: int
    output_path: Path
    paragraph_window: int
    kept_paragraph_ranges: tuple[tuple[int, int], ...]
    valid: bool | None
    validation_detail: str | None = None


def find_control_occurrences(document: HwpBinaryDocument, control_id: str) -> list[HwpControlOccurrence]:
    occurrences: list[HwpControlOccurrence] = []
    for section_index, _path in enumerate(document.section_stream_paths()):
        records = document.section_records(section_index)
        paragraph_ranges = _paragraph_ranges(records)
        for record_index, record in enumerate(records):
            if record.tag_id != 71 or len(record.payload) < 4:
                continue
            current_control_id = record.payload[:4][::-1].decode("latin1", errors="replace")
            if current_control_id != control_id:
                continue
            paragraph_index = _paragraph_index_for_record(record_index, paragraph_ranges)
            if paragraph_index is None:
                continue
            paragraph_start, paragraph_end = paragraph_ranges[paragraph_index]
            occurrences.append(
                HwpControlOccurrence(
                    control_id=control_id,
                    section_index=section_index,
                    record_index=record_index,
                    level=record.level,
                    paragraph_start_index=paragraph_start,
                    paragraph_end_index=paragraph_end,
                )
            )
    return occurrences


def build_minimal_control_candidate(
    document: HwpBinaryDocument,
    occurrence: HwpControlOccurrence,
    output_path: str | Path,
    *,
    paragraph_window: int = 0,
    keep_first_paragraph: bool = True,
) -> Path:
    target_path = Path(output_path).expanduser().resolve()
    target_path.parent.mkdir(parents=True, exist_ok=True)

    records = document.section_records(occurrence.section_index)
    paragraph_ranges = _paragraph_ranges(records)
    target_paragraph_index = _paragraph_index_for_record(occurrence.record_index, paragraph_ranges)
    if target_paragraph_index is None:
        raise IndexError("The requested control occurrence is not contained in a paragraph range.")

    keep_indices = {target_paragraph_index}
    if keep_first_paragraph:
        keep_indices.add(0)
    for delta in range(1, paragraph_window + 1):
        keep_indices.add(max(0, target_paragraph_index - delta))
        keep_indices.add(min(len(paragraph_ranges) - 1, target_paragraph_index + delta))

    kept_ranges = [paragraph_ranges[index] for index in sorted(keep_indices)]
    reduced_records: list[HwpRecord] = []
    for start, end in kept_ranges:
        reduced_records.extend(records[start : end + 1])

    reduced = HwpBinaryDocument.open(document.source_path)
    reduced.replace_section_records(occurrence.section_index, reduced_records)
    reduced.save(target_path)
    return target_path


def pick_best_donor_for_feature(root: str | Path, feature: str) -> HwpDonorSummary | None:
    feature_key = feature.lower()
    if feature_key == "table":
        return find_best_table_donor(root)
    if feature_key == "picture":
        return find_best_picture_donor(root)
    if feature_key == "hyperlink":
        return find_best_hyperlink_donor(root)
    raise ValueError(f"Unsupported feature: {feature}")


def run_template_lab(
    collection_root: str | Path,
    output_dir: str | Path,
    *,
    features: list[str] | None = None,
    paragraph_window: int = 0,
    validate_with_hancom: bool = False,
    validator: Validator | None = None,
) -> list[HwpTemplateCandidate]:
    target_root = Path(output_dir).expanduser().resolve()
    target_root.mkdir(parents=True, exist_ok=True)
    selected_features = features or ["table", "picture", "hyperlink"]
    validate = validator or convert_document
    candidates: list[HwpTemplateCandidate] = []

    for feature in selected_features:
        donor = pick_best_donor_for_feature(collection_root, feature)
        if donor is None:
            continue
        document = HwpBinaryDocument.open(donor.path)
        control_id = CONTROL_BY_FEATURE[feature]
        occurrences = find_control_occurrences(document, control_id)
        if not occurrences:
            continue

        occurrence = occurrences[0]
        candidate_path = target_root / f"{feature}_template_minimal.hwp"
        build_minimal_control_candidate(document, occurrence, candidate_path, paragraph_window=paragraph_window)

        valid: bool | None = None
        detail: str | None = None
        if validate_with_hancom:
            valid, detail = _validate_candidate(candidate_path, validate)

        paragraph_ranges = _paragraph_ranges(document.section_records(occurrence.section_index))
        paragraph_index = _paragraph_index_for_record(occurrence.record_index, paragraph_ranges)
        kept_indices = {0, paragraph_index if paragraph_index is not None else 0}
        for delta in range(1, paragraph_window + 1):
            if paragraph_index is not None:
                kept_indices.add(max(0, paragraph_index - delta))
                kept_indices.add(min(len(paragraph_ranges) - 1, paragraph_index + delta))
        kept_ranges = tuple(paragraph_ranges[index] for index in sorted(kept_indices))

        candidates.append(
            HwpTemplateCandidate(
                source_path=donor.path,
                feature=feature,
                control_id=control_id,
                section_index=occurrence.section_index,
                occurrence_index=occurrence.record_index,
                output_path=candidate_path,
                paragraph_window=paragraph_window,
                kept_paragraph_ranges=kept_ranges,
                valid=valid,
                validation_detail=detail,
            )
        )

    return candidates


def _paragraph_ranges(records: list[HwpRecord]) -> list[tuple[int, int]]:
    ranges: list[tuple[int, int]] = []
    start: int | None = None
    for index, record in enumerate(records):
        if record.tag_id == 66 and record.level == 0:
            if start is not None:
                ranges.append((start, index - 1))
            start = index
    if start is not None:
        ranges.append((start, len(records) - 1))
    return ranges


def _paragraph_index_for_record(record_index: int, paragraph_ranges: list[tuple[int, int]]) -> int | None:
    for index, (start, end) in enumerate(paragraph_ranges):
        if start <= record_index <= end:
            return index
    return None


def _validate_candidate(path: Path, validator: Validator) -> tuple[bool, str | None]:
    validation_dir = path.parent / "_validation"
    validation_dir.mkdir(parents=True, exist_ok=True)
    output_path = validation_dir / f"{path.stem}.validated.hwpx"
    try:
        validator(path, output_path, "HWPX")
    except Exception as exc:
        return False, str(exc)
    return True, None
