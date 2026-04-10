from __future__ import annotations

from pathlib import Path

import pytest

from jakal_hwpx import (
    DocInfoModel,
    DocumentPropertiesRecord,
    HwpBinaryDocument,
    HwpBinaryFileHeader,
    HwpParagraph,
    HwpRecord,
    HwpStreamCapacity,
    HwpPureProfile,
    ParagraphHeaderRecord,
    ParagraphTextRecord,
    SectionModel,
)


REPO_ROOT = Path(__file__).resolve().parents[1]
HWP_SAMPLE_DIR = REPO_ROOT / "examples" / "samples" / "hwp"


@pytest.fixture(scope="session")
def sample_hwp_path() -> Path:
    paths = sorted(HWP_SAMPLE_DIR.glob("*.hwp"))
    assert paths, f"No .hwp samples were found under {HWP_SAMPLE_DIR}"
    return paths[0]


def test_open_hwp_binary_document_exposes_file_header_and_streams(sample_hwp_path: Path) -> None:
    document = HwpBinaryDocument.open(sample_hwp_path)

    file_header = document.file_header()
    assert isinstance(file_header, HwpBinaryFileHeader)
    assert file_header.signature == "HWP Document File"
    assert file_header.version == "5.1.1.0"
    assert file_header.compressed is True

    stream_paths = document.list_stream_paths()
    assert "FileHeader" in stream_paths
    assert "DocInfo" in stream_paths
    assert "BodyText/Section0" in stream_paths
    assert "PrvText" in stream_paths

    preview = document.preview_text()
    assert "2027" in preview


def test_docinfo_and_section_records_parse_from_sample_hwp(sample_hwp_path: Path) -> None:
    document = HwpBinaryDocument.open(sample_hwp_path)

    docinfo_records = document.docinfo_records()
    assert docinfo_records
    assert isinstance(docinfo_records[0], HwpRecord)
    assert docinfo_records[0].tag_id == 0x010

    paragraphs = document.paragraphs(0)
    assert paragraphs
    assert isinstance(paragraphs[0], HwpParagraph)
    assert any("2027" in paragraph.text for paragraph in paragraphs)
    assert "2027" in document.get_document_text()


def test_hwp_binary_document_can_update_preview_and_same_length_body_text(
    sample_hwp_path: Path,
    tmp_path: Path,
) -> None:
    document = HwpBinaryDocument.open(sample_hwp_path)
    document.set_preview_text("JAKAL HWP PREVIEW")

    replacements = document.replace_text_same_length("2027", "2028", count=1)
    assert replacements == 1

    output_path = tmp_path / "edited_sample.hwp"
    document.save(output_path)

    reopened = HwpBinaryDocument.open(output_path)
    assert reopened.preview_text().startswith("JAKAL HWP PREVIEW")
    assert "2028" in reopened.get_document_text()


def test_docinfo_model_supports_typed_access_and_roundtrip(
    sample_hwp_path: Path,
    tmp_path: Path,
) -> None:
    document = HwpBinaryDocument.open(sample_hwp_path)

    model = document.docinfo_model()
    assert isinstance(model, DocInfoModel)
    properties_record = model.document_properties_record()
    assert isinstance(properties_record, DocumentPropertiesRecord)
    original_properties = properties_record.to_properties()
    properties_record.properties = original_properties.__class__(
        section_count=original_properties.section_count,
        page_start_number=original_properties.page_start_number,
        footnote_start_number=original_properties.footnote_start_number,
        endnote_start_number=original_properties.endnote_start_number,
        picture_start_number=original_properties.picture_start_number + 1,
        table_start_number=original_properties.table_start_number,
        equation_start_number=original_properties.equation_start_number,
        list_id=original_properties.list_id,
        paragraph_id=original_properties.paragraph_id,
        character_unit_position=original_properties.character_unit_position,
    )
    assert model.tag_counts()["document_properties"] == 1

    document.replace_docinfo_model(model)
    output_path = tmp_path / "docinfo_model_roundtrip.hwp"
    document.save(output_path)

    reopened = HwpBinaryDocument.open(output_path)
    assert reopened.document_properties().picture_start_number == original_properties.picture_start_number + 1


def test_section_model_builds_paragraph_tree_and_roundtrips_same_length_text_edit(
    sample_hwp_path: Path,
    tmp_path: Path,
) -> None:
    document = HwpBinaryDocument.open(sample_hwp_path)

    model = document.section_model(0)
    assert isinstance(model, SectionModel)
    paragraphs = model.paragraphs()
    assert paragraphs
    assert isinstance(paragraphs[0].header, ParagraphHeaderRecord)
    assert isinstance(paragraphs[0].text_record(), (ParagraphTextRecord, type(None)))
    assert [record.tag_id for record in model.to_records()] == [record.tag_id for record in document.section_records(0)]
    assert any(paragraph.control_nodes() for paragraph in paragraphs)

    target = next(paragraph for paragraph in paragraphs if "2027" in paragraph.text)
    replaced = target.replace_text_same_length("2027", "2031", count=1)
    assert replaced == 1

    document.replace_section_model(0, model)
    output_path = tmp_path / "section_model_roundtrip.hwp"
    document.save(output_path)

    reopened = HwpBinaryDocument.open(output_path)
    assert "2031" in reopened.get_document_text()


def test_hwp_binary_document_append_paragraph_and_capacity_report(sample_hwp_path: Path, tmp_path: Path) -> None:
    document = HwpBinaryDocument.open(sample_hwp_path)

    before_capacity = document.section_capacity(0)
    assert isinstance(before_capacity, HwpStreamCapacity)
    assert before_capacity.fits is True

    paragraph = document.append_paragraph("Pure paragraph")
    assert paragraph.text == "Pure paragraph"

    after_capacity = document.section_capacity(0)
    assert after_capacity.current_size >= before_capacity.current_size
    assert after_capacity.fits is True

    output_path = tmp_path / "append_paragraph_roundtrip.hwp"
    document.save(output_path)
    reopened = HwpBinaryDocument.open(output_path)
    assert "Pure paragraph" in reopened.get_document_text()


def test_hwp_binary_document_can_append_bundled_profile_features(tmp_path: Path) -> None:
    document = HwpBinaryDocument.open(REPO_ROOT / "src" / "jakal_hwpx" / "bundled_hwp_profile" / "base.hwp")
    profile = HwpPureProfile.load_bundled()

    before_count = len(document.section_records(profile.target_section_index))
    document.append_table()
    document.append_picture()
    document.append_hyperlink()
    after_count = len(document.section_records(profile.target_section_index))
    assert after_count > before_count

    output_path = tmp_path / "binary_append_controls.hwp"
    document.save(output_path)
    reopened = HwpBinaryDocument.open(output_path)
    assert len(reopened.section_records(profile.target_section_index)) == after_count
