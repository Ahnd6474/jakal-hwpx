from __future__ import annotations

from pathlib import Path

from jakal_hwpx import HwpBinaryDocument, HwpDocument, HwpPureProfile, build_hwp_pure_profile, bundled_hwp_profile_root


REPO_ROOT = Path(__file__).resolve().parents[1]
COLLECTION_ROOT = REPO_ROOT / "hwp_collection"


def test_build_hwp_pure_profile_creates_base_and_feature_templates(tmp_path: Path) -> None:
    profile = build_hwp_pure_profile(COLLECTION_ROOT, tmp_path)

    assert profile.base_path.exists()
    assert set(profile.template_paths) == {"table", "picture", "hyperlink"}
    assert all(path.exists() for path in profile.template_paths.values())
    assert (tmp_path / "profile.json").exists()
    assert profile.target_section_index >= 0


def test_hwp_document_blank_from_profile_can_append_pure_templates(tmp_path: Path) -> None:
    profile = build_hwp_pure_profile(COLLECTION_ROOT, tmp_path / "profile")
    loaded_profile = HwpPureProfile.load(profile.root)

    document = HwpDocument.blank_from_profile(loaded_profile.root)
    base_record_count = len(document.binary_document().section_records(loaded_profile.target_section_index))

    document.append_table_pure()
    after_table_count = len(document.binary_document().section_records(loaded_profile.target_section_index))
    assert after_table_count > base_record_count

    document.append_picture_pure()
    after_picture_count = len(document.binary_document().section_records(loaded_profile.target_section_index))
    assert after_picture_count > after_table_count

    document.append_hyperlink_pure()
    after_hyperlink_count = len(document.binary_document().section_records(loaded_profile.target_section_index))
    assert after_hyperlink_count > after_picture_count

    output_path = tmp_path / "pure_profile_output.hwp"
    document.save(output_path)
    reopened = HwpBinaryDocument.open(output_path)
    assert len(reopened.section_records(loaded_profile.target_section_index)) == after_hyperlink_count


def test_bundled_profile_supports_blank_document_and_auto_loaded_pure_append(tmp_path: Path) -> None:
    bundled_root = bundled_hwp_profile_root()
    assert (bundled_root / "profile.json").exists()

    blank_document = HwpDocument.blank()
    blank_document.append_table_pure()
    blank_document.append_picture_pure()
    blank_document.append_hyperlink_pure()

    output_path = tmp_path / "bundled_blank_output.hwp"
    blank_document.save(output_path)

    reopened = HwpBinaryDocument.open(output_path)
    assert reopened.section_records(HwpPureProfile.load_bundled().target_section_index)
