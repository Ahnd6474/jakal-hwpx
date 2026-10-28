"""White padding and blank pages must not inflate document fidelity scores."""
import importlib.util
from pathlib import Path

import pytest


spec = importlib.util.spec_from_file_location("render_comparison", Path(__file__).resolve().parents[1] / "scripts/compare_document_renders.py")
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
compare = module.compare_rgb_pixels
WHITE, BLACK = b"\xff\xff\xff", b"\x00\x00\x00"


def test_white_padding_does_not_dilute_foreground_difference():
    small = compare(BLACK + WHITE * 9, WHITE * 10, 10, 1)
    large = compare(BLACK + WHITE * 999, WHITE * 1000, 1000, 1)
    assert small["whole_page_changed_fraction"] == .1
    assert large["whole_page_changed_fraction"] == .001
    assert small["foreground_changed_fraction"] == large["foreground_changed_fraction"] == 1
    assert small["ink_iou"] == large["ink_iou"] == 0


def test_extra_content_and_displaced_content_count_in_the_union():
    added = compare(WHITE * 10, WHITE * 9 + BLACK, 10, 1)
    shifted = compare(BLACK + WHITE, WHITE + BLACK, 2, 1)
    assert added["foreground_changed_fraction"] == 1
    assert added["candidate_ink_pixels"] == 1
    assert shifted["changed_pixels"] == shifted["ink_union_pixels"] == 2
    assert shifted["foreground_changed_fraction"] == 1


def test_blank_pages_have_no_foreground_score():
    result = compare(WHITE * 10, WHITE * 10, 10, 1)
    assert result["pixel_exact"]
    assert result["foreground_changed_fraction"] is None
    assert result["ink_iou"] is None
    assert not result["foreground_comparison_available"]


def test_identical_visible_content_has_exact_pixels_and_full_ink_overlap():
    result = compare(WHITE + BLACK, WHITE + BLACK, 2, 1)
    assert result["pixel_exact"]
    assert result["foreground_changed_fraction"] == 0
    assert result["ink_iou"] == 1


def test_channel_tolerance_does_not_mislabel_pixel_exactness():
    result = compare(BLACK, bytes([1, 0, 0]), 1, 1)
    assert result["foreground_changed_fraction"] == 0
    assert not result["pixel_exact"]
    assert result["maximum_channel_difference"] == 1


def test_mismatched_buffers_are_rejected():
    with pytest.raises(ValueError, match="same positive page dimensions"):
        compare(BLACK, BLACK * 2, 1, 1)
