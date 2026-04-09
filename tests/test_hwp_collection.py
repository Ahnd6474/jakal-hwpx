from __future__ import annotations

from pathlib import Path

from jakal_hwpx import (
    HwpDonorSummary,
    find_best_hyperlink_donor,
    find_best_picture_donor,
    find_best_table_donor,
    scan_hwp_collection,
)


REPO_ROOT = Path(__file__).resolve().parents[1]
COLLECTION_ROOT = REPO_ROOT / "hwp_collection"


def test_scan_hwp_collection_finds_legacy_hwp_donors() -> None:
    donors = scan_hwp_collection(COLLECTION_ROOT)

    assert donors
    assert all(isinstance(item, HwpDonorSummary) for item in donors)
    assert any(item.has_table for item in donors)
    assert any(item.has_picture for item in donors)
    assert any(item.has_hyperlink for item in donors)


def test_find_best_donors_pick_expected_feature_rich_documents() -> None:
    table_donor = find_best_table_donor(COLLECTION_ROOT)
    picture_donor = find_best_picture_donor(COLLECTION_ROOT)
    hyperlink_donor = find_best_hyperlink_donor(COLLECTION_ROOT)

    assert table_donor is not None
    assert table_donor.has_table is True

    assert picture_donor is not None
    assert picture_donor.has_picture is True

    assert hyperlink_donor is not None
    assert hyperlink_donor.has_hyperlink is True
