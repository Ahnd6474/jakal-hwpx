from __future__ import annotations

from pathlib import Path

from jakal_hwpx import (
    HwpBinaryDocument,
    build_minimal_control_candidate,
    find_control_occurrences,
    run_template_lab,
)


REPO_ROOT = Path(__file__).resolve().parents[1]
COLLECTION_ROOT = REPO_ROOT / "hwp_collection"


def test_find_control_occurrences_finds_table_controls_in_collection_donor() -> None:
    donor_path = next(path for path in sorted(COLLECTION_ROOT.glob("*.hwp")) if "RFP-100개" in path.name)
    document = HwpBinaryDocument.open(donor_path)

    occurrences = find_control_occurrences(document, "tbl ")

    assert occurrences
    assert occurrences[0].control_id == "tbl "
    assert occurrences[0].paragraph_start_index <= occurrences[0].record_index <= occurrences[0].paragraph_end_index


def test_build_minimal_control_candidate_rewrites_to_smaller_section(tmp_path: Path) -> None:
    donor_path = next(path for path in sorted(COLLECTION_ROOT.glob("*.hwp")) if "RFP-100개" in path.name)
    document = HwpBinaryDocument.open(donor_path)
    occurrence = find_control_occurrences(document, "tbl ")[0]

    output_path = tmp_path / "table_minimal.hwp"
    build_minimal_control_candidate(document, occurrence, output_path)

    reopened = HwpBinaryDocument.open(output_path)
    assert reopened.section_records(occurrence.section_index)
    assert len(reopened.section_records(occurrence.section_index)) < len(document.section_records(occurrence.section_index))


def test_run_template_lab_generates_minimized_candidates_with_fake_validator(tmp_path: Path) -> None:
    def fake_validator(input_path: str | Path, output_path: str | Path, output_format: str) -> Path:
        target = Path(output_path)
        target.write_bytes(b"validated")
        return target

    candidates = run_template_lab(
        COLLECTION_ROOT,
        tmp_path,
        features=["table", "picture", "hyperlink"],
        validate_with_hancom=True,
        validator=fake_validator,
    )

    assert candidates
    assert {candidate.feature for candidate in candidates} == {"table", "picture", "hyperlink"}
    assert all(candidate.output_path.exists() for candidate in candidates)
    assert all(candidate.valid is True for candidate in candidates)
