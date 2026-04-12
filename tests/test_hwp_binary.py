from __future__ import annotations

from pathlib import Path

import pytest

from jakal_hwpx import (
    BinDataRecord,
    DocInfoModel,
    DocumentPropertiesRecord,
    HwpBinaryDocument,
    HwpBinaryFileHeader,
    HwpParagraph,
    HwpRecord,
    HwpStreamCapacity,
    HwpPureProfile,
    IdMappingsRecord,
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
    return next((path for path in paths if not path.name.startswith("generated_")), paths[0])


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
    control_ids = [record.payload[:4][::-1].decode("latin1", errors="replace") for record in reopened.section_records(profile.target_section_index) if record.tag_id == 71]
    assert "tbl " in control_ids
    assert "gso " in control_ids
    assert "%hlk" in control_ids


def test_hwp_binary_document_can_grow_streams_with_custom_cfb_writer(sample_hwp_path: Path, tmp_path: Path) -> None:
    document = HwpBinaryDocument.open(sample_hwp_path)
    value = "LONG PREVIEW " * 50
    document.set_preview_text(value)

    output_path = tmp_path / "grown_streams.hwp"
    document.save(output_path)

    reopened = HwpBinaryDocument.open(output_path)
    assert reopened.preview_text().startswith("LONG PREVIEW")
    assert len(reopened.preview_text()) >= len(value.rstrip())


def test_docinfo_model_can_add_and_remove_bindata_records() -> None:
    document = HwpBinaryDocument.open(REPO_ROOT / "src" / "jakal_hwpx" / "bundled_hwp_profile" / "base.hwp")
    model = document.docinfo_model()

    id_record = model.id_mappings_record()
    assert isinstance(id_record, IdMappingsRecord)
    before = id_record.bin_data_count
    bindata = model.add_embedded_bindata("png")
    assert isinstance(bindata, BinDataRecord)
    assert bindata.storage_id == before + 1
    assert model.id_mappings_record().bin_data_count == before + 1
    assert model.remove_bindata(bindata.storage_id or 0) is True
    assert model.id_mappings_record().bin_data_count == before


def test_hwp_binary_document_can_add_bindata_and_append_picture_from_bytes(tmp_path: Path) -> None:
    document = HwpBinaryDocument.open(REPO_ROOT / "src" / "jakal_hwpx" / "bundled_hwp_profile" / "base.hwp")
    profile = HwpPureProfile.load_bundled()
    image_bytes = document.read_stream("BinData/BIN0001.bmp", decompress=False)
    before_bindata_count = document.docinfo_model().id_mappings_record().bin_data_count
    before_section_count = len(document.section_records(profile.target_section_index))

    document.append_picture(image_bytes, extension="bmp")

    after_model = document.docinfo_model()
    after_bindata_count = after_model.id_mappings_record().bin_data_count
    assert after_bindata_count == before_bindata_count + 1
    assert f"BinData/BIN{after_bindata_count:04d}.bmp" in document.bindata_stream_paths()
    assert len(document.section_records(profile.target_section_index)) > before_section_count

    output_path = tmp_path / "append_picture_bytes.hwp"
    document.save(output_path)
    reopened = HwpBinaryDocument.open(output_path)
    assert reopened.docinfo_model().id_mappings_record().bin_data_count == after_bindata_count
    assert f"BinData/BIN{after_bindata_count:04d}.bmp" in reopened.bindata_stream_paths()


def test_hwp_binary_document_can_append_table_with_cell_text_and_hyperlink_with_text(tmp_path: Path) -> None:
    document = HwpBinaryDocument.open(REPO_ROOT / "src" / "jakal_hwpx" / "bundled_hwp_profile" / "base.hwp")
    profile = HwpPureProfile.load_bundled()

    document.append_table("CELL-1")
    document.append_hyperlink("https://example.com/demo", text="EXAMPLE-LINK")

    output_path = tmp_path / "custom_table_hyperlink.hwp"
    document.save(output_path)

    reopened = HwpBinaryDocument.open(output_path)
    text = reopened.get_document_text()
    assert "CELL-1" in text
    assert "EXAMPLE-LINK" in text
    section_records = reopened.section_records(profile.target_section_index)
    hyperlink_ctrl = next(
        record for record in section_records if record.tag_id == 71 and record.payload[:4][::-1].decode("latin1", errors="replace") == "%hlk"
    )
    assert "https\\://example.com/demo;1;0;0;".encode("utf-16-le") in hyperlink_ctrl.payload


def test_hwp_binary_document_can_append_multi_cell_table_and_custom_hyperlink_metadata(tmp_path: Path) -> None:
    document = HwpBinaryDocument.open(REPO_ROOT / "src" / "jakal_hwpx" / "bundled_hwp_profile" / "base.hwp")
    profile = HwpPureProfile.load_bundled()

    document.append_table(
        rows=2,
        cols=3,
        cell_texts=[
            ["R1C1", "R1C2", "R1C3"],
            ["R2C1", "R2C2", "R2C3"],
        ],
    )
    document.append_hyperlink(
        "https://example.com/meta",
        text="META-LINK",
        metadata_fields=[9, 7, "bookmark"],
    )

    output_path = tmp_path / "multi_table_hyperlink_metadata.hwp"
    document.save(output_path)
    reopened = HwpBinaryDocument.open(output_path)
    text = reopened.get_document_text()
    for value in ("R1C1", "R1C2", "R1C3", "R2C1", "R2C2", "R2C3", "META-LINK"):
        assert value in text
    section_records = reopened.section_records(profile.target_section_index)
    table_record = [record for record in section_records if record.tag_id == 77][-1]
    assert int.from_bytes(table_record.payload[4:6], "little") == 2
    assert int.from_bytes(table_record.payload[6:8], "little") == 3
    cell_headers = [record for record in section_records if record.tag_id == 72 and len(record.payload) == 47]
    assert len(cell_headers) >= 6
    hyperlink_ctrl = next(
        record for record in section_records if record.tag_id == 71 and record.payload[:4][::-1].decode("latin1", errors="replace") == "%hlk"
    )
    assert "https\\://example.com/meta;9;7;bookmark;".encode("utf-16-le") in hyperlink_ctrl.payload


def test_hwp_binary_document_can_encode_table_spans_row_heights_and_cell_border_fills(tmp_path: Path) -> None:
    document = HwpBinaryDocument.open(REPO_ROOT / "src" / "jakal_hwpx" / "bundled_hwp_profile" / "base.hwp")
    profile = HwpPureProfile.load_bundled()

    document.append_table(
        rows=3,
        cols=3,
        cell_texts=[
            ["A", "", "B"],
            ["C", "D", ""],
            ["E", "", ""],
        ],
        row_heights=[111, 222, 333],
        col_widths=[1000, 2000, 3000],
        cell_spans={
            (0, 0): (1, 2),
            (1, 1): (2, 2),
        },
        cell_border_fill_ids={
            (0, 0): 7,
            (1, 1): 9,
        },
        table_border_fill_id=5,
    )

    output_path = tmp_path / "table_geometry.hwp"
    document.save(output_path)
    reopened = HwpBinaryDocument.open(output_path)
    section_records = reopened.section_records(profile.target_section_index)
    table_record = [record for record in section_records if record.tag_id == 77][-1]
    assert int.from_bytes(table_record.payload[4:6], "little") == 3
    assert int.from_bytes(table_record.payload[6:8], "little") == 3
    assert [int.from_bytes(table_record.payload[offset : offset + 2], "little") for offset in (18, 20, 22)] == [111, 222, 333]
    assert int.from_bytes(table_record.payload[24:26], "little") == 5

    last_table_index = max(index for index, record in enumerate(section_records) if record.tag_id == 77)
    cell_headers: list[bytes] = []
    for record in section_records[last_table_index + 1 :]:
        if record.level < 2:
            break
        if record.tag_id == 72 and len(record.payload) == 47:
            cell_headers.append(record.payload)
    assert len(cell_headers) == 5

    def parse_cell(payload: bytes) -> tuple[int, int, int, int, int, int, int]:
        return (
            int.from_bytes(payload[8:10], "little"),
            int.from_bytes(payload[10:12], "little"),
            int.from_bytes(payload[12:14], "little"),
            int.from_bytes(payload[14:16], "little"),
            int.from_bytes(payload[16:20], "little"),
            int.from_bytes(payload[20:24], "little"),
            int.from_bytes(payload[32:34], "little"),
        )

    parsed = [parse_cell(payload) for payload in cell_headers]
    assert (0, 0, 2, 1, 3000, 111, 7) in parsed
    assert (2, 0, 1, 1, 3000, 111, 5) in parsed
    assert (0, 1, 1, 1, 1000, 222, 5) in parsed
    assert (1, 1, 2, 2, 5000, 555, 9) in parsed
    assert (0, 2, 1, 1, 1000, 333, 5) in parsed


def test_hwp_binary_template_free_table_and_hyperlink_builders_do_not_require_profile_files(tmp_path: Path) -> None:
    document = HwpBinaryDocument.open(REPO_ROOT / "src" / "jakal_hwpx" / "bundled_hwp_profile" / "base.hwp")
    document.append_table("NO-PROFILE", profile_root=tmp_path / "missing-profile")
    document.append_hyperlink(
        "https://example.com/no-profile",
        text="NO-PROFILE-LINK",
        profile_root=tmp_path / "missing-profile",
    )
    output_path = tmp_path / "template_free_controls.hwp"
    document.save(output_path)
    reopened = HwpBinaryDocument.open(output_path)
    text = reopened.get_document_text()
    assert "NO-PROFILE" in text
    assert "NO-PROFILE-LINK" in text


def test_hwp_binary_template_free_picture_builder_does_not_require_profile_files(tmp_path: Path) -> None:
    document = HwpBinaryDocument.open(REPO_ROOT / "src" / "jakal_hwpx" / "bundled_hwp_profile" / "base.hwp")
    before_records = len(document.section_records(0))

    document.append_picture(profile_root=tmp_path / "missing-profile")

    output_path = tmp_path / "template_free_picture.hwp"
    document.save(output_path)
    reopened = HwpBinaryDocument.open(output_path)
    after_records = reopened.section_records(0)
    assert len(after_records) > before_records
    control_ids = [record.payload[:4][::-1].decode("latin1", errors="replace") for record in after_records if record.tag_id == 71]
    assert "gso " in control_ids
