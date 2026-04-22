from __future__ import annotations

from pathlib import Path

import pytest

import jakal_hwpx.hwp_binary as hwp_binary_module
from jakal_hwpx import (
    AutoNumberControlRecord,
    BinDataRecord,
    BookmarkControlRecord,
    BulletRecord,
    ChartDataRecord,
    CharShapeRecord,
    ControlHeaderRecord,
    DocInfoModel,
    DocumentPropertiesRecord,
    EquationControlRecord,
    FieldControlRecord,
    FaceNameRecord,
    FormObjectRecord,
    FootnoteShapeRecord,
    HeaderFooterControlRecord,
    HyperlinkControlRecord,
    HwpBinaryDocument,
    HwpBinaryFileHeader,
    HwpParagraph,
    HwpDocument,
    HwpRecord,
    HwpStreamCapacity,
    HwpPureProfile,
    IdMappingsRecord,
    NumberingRecord,
    NoteControlRecord,
    OleControlRecord,
    PageBorderFillRecord,
    PageDefRecord,
    PageNumberControlRecord,
    ParagraphHeaderRecord,
    ParagraphTextRecord,
    ParaShapeRecord,
    PictureControlRecord,
    PictureShapeComponentRecord,
    SectionDefinitionControlRecord,
    SectionModel,
    ShapeControlRecord,
    StyleRecord,
    TabDefRecord,
    MemoListRecord,
)


REPO_ROOT = Path(__file__).resolve().parents[1]
HWP_SAMPLE_DIR = REPO_ROOT / "examples" / "samples" / "hwp"


@pytest.fixture(scope="session")
def sample_hwp_path() -> Path:
    paths = sorted(HWP_SAMPLE_DIR.glob("*.hwp"))
    assert paths, f"No .hwp samples were found under {HWP_SAMPLE_DIR}"
    return next((path for path in paths if not path.name.startswith("generated_")), paths[0])


def _layout_chain(document: HwpBinaryDocument, path: str) -> tuple[int, ...]:
    return next(entry.sector_chain for entry in document._compound_layout.entries if entry.path == path)


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


def test_docinfo_model_exposes_typed_docinfo_records(sample_hwp_path: Path) -> None:
    document = HwpBinaryDocument.open(sample_hwp_path)
    model = document.docinfo_model()

    assert isinstance(model.face_name_records()[0], FaceNameRecord)
    assert isinstance(model.char_shape_records()[0], CharShapeRecord)
    assert isinstance(model.para_shape_records()[0], ParaShapeRecord)
    assert isinstance(model.style_records()[0], StyleRecord)
    assert isinstance(model.tab_def_records()[0], TabDefRecord)
    assert isinstance(model.numbering_records()[0], NumberingRecord)
    if model.bullet_records():
        assert isinstance(model.bullet_records()[0], BulletRecord)


def test_docinfo_model_exposes_named_docinfo_payload_fields(sample_hwp_path: Path) -> None:
    document = HwpBinaryDocument.open(sample_hwp_path)
    model = document.docinfo_model()

    face = model.face_name_records()[0].fields()
    char_shape = model.char_shape_records()[0].fields()
    para_shape = model.para_shape_records()[0].fields()
    style = model.style_records()[0].fields(style_id=0)
    numbering = model.numbering_records()[0].fields()

    assert isinstance(face["name"], str)
    assert "flag_bits" in face
    assert "font_refs" in char_shape and len(char_shape["font_refs"]) == 7
    assert "attribute_words" in char_shape
    assert "attribute_sections" in char_shape
    assert "probable_effect_flags" in char_shape
    assert "colors" in char_shape
    assert "alignment_horizontal" in para_shape
    assert "attribute_bits" in para_shape
    assert "line_spacing_primary" in para_shape
    assert "tab_def_id" in para_shape
    assert "border_fill_id" in para_shape
    assert "heading" in para_shape
    assert style["style_id"] == 0
    assert "style_word" in style
    assert "style_word_bytes" in style
    assert isinstance(numbering["levels"], list)
    assert all("level_index" in level for level in numbering["levels"])
    assert all("para_head_flags" in level for level in numbering["levels"])
    if model.border_fill_records():
        border_fill = model.border_fill_records()[0].fields()
        assert "border_flag_bits" in border_fill
        assert "left" in border_fill or "fill" in border_fill
        if "fill" in border_fill:
            assert "fill_type_flags" in border_fill["fill"]
    if model.tab_def_records():
        tab_def = model.tab_def_records()[0].fields()
        assert "tabs" in tab_def
    if model.bullet_records():
        bullet = model.bullet_records()[0].fields()
        assert "bullet_char" in bullet
        assert "bullet_char_codepoint" in bullet
        assert "flag_bits" in bullet


def test_native_char_and_para_shape_payload_parsers_expose_tail_and_bitfield_structure() -> None:
    char_payload = bytes.fromhex(
        "010001000100010000000200000064646464646464000000000000006464646464646400000000000000"
        "4c040000020000000a0a0000000000000000ffffffffc0c0c000040000000000"
    )
    para_payload = bytes.fromhex(
        "800100000000000000000000000000000000000000000000a000000000000000050000000000000000000000000000000000a0000000"
    )

    char_fields = CharShapeRecord.from_record(
        HwpRecord(tag_id=0x015, level=0, size=len(char_payload), header_size=4, offset=0, payload=char_payload)
    ).fields()
    para_fields = ParaShapeRecord.from_record(
        HwpRecord(tag_id=0x019, level=0, size=len(para_payload), header_size=4, offset=0, payload=para_payload)
    ).fields()

    assert char_fields["text_color"] == "#000000"
    assert char_fields["tail_gap_58_60"] == b"\x00\x00"
    assert char_fields["shade_color"] == "none"
    assert char_fields["shadow_color"] == "#C0C0C0"
    assert char_fields["border_fill_id"] == 4
    assert char_fields["tail_reserved"] == b"\x00\x00"
    assert char_fields["hwpx_compatible_flags"] == {"useFontSpace": False, "useKerning": False}

    assert para_fields["alignment_horizontal"] == "JUSTIFY"
    assert para_fields["attribute_bits"]["snap_to_grid"] is True
    assert para_fields["attribute_bits"]["word_break_keep_word"] is True
    assert para_fields["line_spacing_primary"] == 160
    assert para_fields["border_fill_id"] == 5
    assert para_fields["line_spacing"] == 160


def test_hwp_binary_document_noop_save_copy_is_byte_identical(sample_hwp_path: Path, tmp_path: Path) -> None:
    document = HwpBinaryDocument.open(sample_hwp_path)

    output_path = tmp_path / "noop_copy.hwp"
    document.save_copy(output_path)

    assert output_path.read_bytes() == sample_hwp_path.read_bytes()


def test_hwp_binary_document_noop_reencode_preserves_raw_streams(sample_hwp_path: Path, tmp_path: Path) -> None:
    document = HwpBinaryDocument.open(sample_hwp_path)

    output_path = tmp_path / "noop_reencode.hwp"
    document.save_copy(output_path, preserve_original_bytes=False)

    reopened = HwpBinaryDocument.open(output_path)
    assert reopened.list_stream_paths() == document.list_stream_paths()
    for stream_path in document.list_stream_paths():
        assert reopened.read_stream(stream_path, decompress=False) == document.read_stream(stream_path, decompress=False)
        assert reopened.read_stream(stream_path) == document.read_stream(stream_path)


def test_hwp_binary_document_noop_reencode_is_byte_identical(sample_hwp_path: Path, tmp_path: Path) -> None:
    document = HwpBinaryDocument.open(sample_hwp_path)

    output_path = tmp_path / "noop_reencode_exact.hwp"
    document.save_copy(output_path, preserve_original_bytes=False)

    assert output_path.read_bytes() == sample_hwp_path.read_bytes()


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

    exact_copy_path = tmp_path / "edited_sample_copy.hwp"
    reopened.save_copy(exact_copy_path)
    assert exact_copy_path.read_bytes() == output_path.read_bytes()


def test_section_model_exposes_typed_control_ast(tmp_path: Path) -> None:
    document = HwpBinaryDocument.open(REPO_ROOT / "src" / "jakal_hwpx" / "bundled_hwp_profile" / "base.hwp")
    target_section_index = HwpPureProfile.load_bundled().target_section_index
    document.append_bookmark("AST-BOOKMARK")
    document.append_hyperlink("https://example.com/ast", text="AST-LINK", metadata_fields=[9, 2, "anchor"])
    document.append_field(field_type="MAILMERGE", display_text="AST-FIELD", name="AST_FIELD", parameters={"FieldName": "AST_FIELD"})
    document.append_auto_number(kind="newNum", number=7)
    document.append_header("AST-HEADER", apply_page_type="EVEN")
    document.append_footer("AST-FOOTER", apply_page_type="ODD")
    document.append_footnote("AST-NOTE")
    document.append_equation("x+y=z", width=3456, height=2100, font="Batang", shape_comment="AST-EQ")
    model = document.section_model(target_section_index)
    equation_node = next(control for control in model.controls() if isinstance(control, EquationControlRecord))
    equation_node.payload = hwp_binary_module._set_graphic_control_layout_payload(
        equation_node.payload,
        text_wrap="SQUARE",
        text_flow="RIGHT_ONLY",
        treat_as_char=False,
        vert_rel_to="PAPER",
        horz_rel_to="PAPER",
        vert_align="CENTER",
        horz_align="RIGHT",
        vert_offset=12,
        horz_offset=34,
    )
    document.replace_section_model(target_section_index, model)
    document.set_section_page_numbers(target_section_index, [{"pos": "TOP_RIGHT", "formatType": "ROMAN_SMALL", "sideChar": "*"}])

    model = document.section_model(target_section_index)
    controls = model.controls()
    assert controls
    assert all(isinstance(control, ControlHeaderRecord) for control in controls)

    bookmark = next(control for control in controls if isinstance(control, BookmarkControlRecord))
    hyperlink = next(control for control in controls if isinstance(control, HyperlinkControlRecord))
    field = next(control for control in controls if isinstance(control, FieldControlRecord) and control.control_id != "%hlk")
    auto_number = next(control for control in controls if isinstance(control, AutoNumberControlRecord))
    header = next(control for control in controls if isinstance(control, HeaderFooterControlRecord) and control.kind() == "header")
    footer = next(control for control in controls if isinstance(control, HeaderFooterControlRecord) and control.kind() == "footer")
    note = next(control for control in controls if isinstance(control, NoteControlRecord))
    page_number = next(control for control in controls if isinstance(control, PageNumberControlRecord))
    equation = next(control for control in controls if isinstance(control, EquationControlRecord))
    hyperlink_paragraph = next(paragraph for paragraph in model.paragraphs() if hyperlink in paragraph.control_nodes())
    section_definition = model.section_definition_control()

    assert bookmark.name() == "AST-BOOKMARK"
    assert hyperlink.url() == "https://example.com/ast"
    assert hyperlink.metadata_fields() == ["9", "2", "anchor"]
    assert field.field_type() == "MAILMERGE"
    assert field.native_field_type() == "%mai"
    assert field.name() == "AST_FIELD"
    assert field.parameters()["FieldName"] == "AST_FIELD"
    assert field.is_mail_merge() is True
    assert field.is_doc_property() is False
    assert field.is_date() is False
    assert field.merge_field_name() == "AST_FIELD"
    assert "AST-LINK" in hyperlink_paragraph.text
    assert auto_number.kind() == "newNum"
    assert auto_number.number() == "7"
    assert auto_number.fields()["number_type"] == "PAGE"
    assert header.apply_page_type() == "EVEN"
    assert header.paragraph_texts() == ["AST-HEADER"]
    assert footer.apply_page_type() == "ODD"
    assert footer.paragraph_texts() == ["AST-FOOTER"]
    assert note.kind() == "footNote"
    assert note.paragraph_texts() == ["AST-NOTE"]
    assert page_number.pos() == "TOP_RIGHT"
    assert page_number.format_type() == "ROMAN_SMALL"
    assert page_number.side_char() == "*"
    assert bookmark.name() == "AST-BOOKMARK"
    assert hyperlink.fields()["url"] == "https://example.com/ast"
    assert hyperlink.fields()["metadata_fields"] == ["9", "2", "anchor"]
    assert field.fields()["document_property_name"] is None
    assert header.fields()["text"] == "AST-HEADER"
    assert footer.fields()["paragraph_texts"] == ["AST-FOOTER"]
    assert note.fields()["text"] == "AST-NOTE"
    assert page_number.fields()["pos"] == "TOP_RIGHT"
    assert page_number.fields()["side_char"] == "*"
    assert equation.script() == "x+y=z"
    assert equation.font() == "Batang"
    assert equation.graphic_size() == {"width": 3456, "height": 2100}
    assert equation.layout()["textWrap"] == "SQUARE"
    assert equation.layout()["textFlow"] == "RIGHT_ONLY"
    assert equation.layout()["treatAsChar"] == "0"
    assert equation.layout()["vertRelTo"] == "PAPER"
    assert equation.layout()["horzRelTo"] == "PAPER"
    assert equation.layout()["vertAlign"] == "CENTER"
    assert equation.layout()["horzAlign"] == "RIGHT"
    assert equation.layout()["vertOffset"] == "12"
    assert equation.layout()["horzOffset"] == "34"
    assert section_definition is not None
    assert isinstance(section_definition, SectionDefinitionControlRecord)
    page_def = section_definition.page_def_record()
    assert isinstance(page_def, PageDefRecord)
    assert page_def.page_settings()["page_width"] > 0
    assert all(isinstance(record, PageBorderFillRecord) for record in section_definition.page_border_fill_records())
    assert all(isinstance(record, FootnoteShapeRecord) for record in section_definition.note_shape_records())


def test_section_model_specializes_graphic_controls_by_native_kind(tmp_path: Path) -> None:
    document = HwpBinaryDocument.open(REPO_ROOT / "src" / "jakal_hwpx" / "bundled_hwp_profile" / "base.hwp")
    image_bytes = document.read_stream("BinData/BIN0001.bmp", decompress=False)
    target_section_index = HwpPureProfile.load_bundled().target_section_index

    document.append_picture(image_bytes, extension="bmp", width=2400, height=1800)
    document.append_shape(kind="line", width=1500, height=900, line_color="#010203")
    document.append_shape(kind="rect", text="AST-SHAPE", width=3600, height=1800, fill_color="#ABCDEF", line_color="#123456")
    document.append_shape(kind="polygon", width=2200, height=1400, fill_color="#FEDCBA", line_color="#654321")
    document.append_shape(kind="textart", text="AST-TEXTART", width=2600, height=1300, fill_color="#0A0B0C", line_color="#0D0E0F")
    document.append_ole(
        "ast.ole",
        b"AST-OLE",
        width=3900,
        height=2100,
        shape_comment="AST-OLE-COMMENT",
        object_type="LINK",
        draw_aspect="ICON",
        has_moniker=True,
        eq_baseline=12,
        line_color="#445566",
        line_width=77,
    )

    model = document.section_model(target_section_index)
    pictures = [control for control in model.controls() if isinstance(control, PictureControlRecord)]
    shapes = [control for control in model.controls() if isinstance(control, ShapeControlRecord)]
    oles = [control for control in model.controls() if isinstance(control, OleControlRecord)]

    picture = max(pictures, key=lambda control: control.storage_id())
    line = next(control for control in shapes if control.graphic_kind() == "line")
    shape = next(control for control in shapes if control.text() == "AST-SHAPE")
    polygon = next(control for control in shapes if control.graphic_kind() == "polygon")
    textart = next(control for control in shapes if control.graphic_kind() == "textart")
    ole = next(control for control in oles if control.line_width() == 77)

    assert picture.storage_id() > 0
    assert picture.picture_size() == {"width": 2400, "height": 1800}
    assert isinstance(picture.picture_record(), PictureShapeComponentRecord)
    assert line.specific_fields()["start"] == {"x": 0, "y": 0}
    assert line.specific_fields()["end"] == {"x": 1500, "y": 900}
    assert shape.graphic_kind() == "rect"
    assert shape.text() == "AST-SHAPE"
    assert shape.shape_size() == {"width": 3600, "height": 1800}
    assert shape.fill_color() == "#ABCDEF"
    assert shape.line_color() == "#123456"
    assert "u16_values" in polygon.specific_fields()
    assert textart.text() == "AST-TEXTART"
    assert ole.storage_id() > 0
    assert ole.object_type() == "LINK"
    assert ole.draw_aspect() == "ICON"
    assert ole.has_moniker() is True
    assert ole.eq_baseline() == 12
    assert ole.line_color() == "#445566"
    assert ole.line_width() == 77


def test_section_model_detects_hancom_connectline_signature_without_ctrl_data() -> None:
    document = HwpBinaryDocument.open(REPO_ROOT / "src" / "jakal_hwpx" / "bundled_hwp_profile" / "base.hwp")
    target_section_index = HwpPureProfile.load_bundled().target_section_index
    document.append_shape(kind="line", width=1500, height=900, line_color="#010203")

    model = document.section_model(target_section_index)
    line_control = next(control for control in model.controls() if isinstance(control, ShapeControlRecord) and control.graphic_kind() == "line")
    component = line_control.shape_component_record()
    specific = line_control.specific_shape_record()
    assert component is not None
    assert specific is not None

    component.payload = b"loc$loc$" + component.payload[8:]
    specific.payload = specific.payload + (b"\x00" * 13)
    document.replace_section_model(target_section_index, model)

    reopened = document.section_model(target_section_index)
    connect_line = next(control for control in reopened.controls() if isinstance(control, ShapeControlRecord) and control.graphic_kind() == "connectLine")
    assert connect_line.graphic_kind() == "connectLine"
    assert connect_line.specific_fields()["variant"] == "hancom_connectline"
    assert connect_line.specific_fields()["raw_payload"]


def test_section_model_parses_native_picture_and_ole_probe_backed_fields(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    image_bytes = document.binary_document().read_stream("BinData/BIN0001.bmp", decompress=False)
    section = document.section(0)
    section.append_picture(image_bytes, extension="bmp", width=2400, height=1600)
    section.append_ole("probe.ole", b"OLE-PROBE", width=42001, height=13501)

    picture = section.pictures()[0]
    ole = section.oles()[0]
    picture.set_layout(
        text_wrap="SQUARE",
        text_flow="RIGHT_ONLY",
        treat_as_char=False,
        vert_rel_to="PAPER",
        horz_rel_to="PAPER",
        vert_align="CENTER",
        horz_align="RIGHT",
        vert_offset=101,
        horz_offset=202,
    )
    picture.set_out_margins(left=13, right=24, top=35, bottom=46)
    picture.set_rotation(angle=15, center_x=111, center_y=222, rotate_image=False)
    picture.set_image_adjustment(bright=5, contrast=6, effect="GRAY_SCALE", alpha=7)
    picture.set_crop(left=1, right=2399, top=2, bottom=1598)
    picture.set_line_style(color="#112233", width=55)
    ole.set_layout(
        text_wrap="SQUARE",
        text_flow="LEFT_ONLY",
        treat_as_char=False,
        vert_rel_to="PAPER",
        horz_rel_to="PAPER",
        vert_align="CENTER",
        horz_align="RIGHT",
        vert_offset=77,
        horz_offset=88,
    )
    ole.set_out_margins(left=9, right=10, top=11, bottom=12)
    ole.set_rotation(angle=45, center_x=777, center_y=888)
    ole.set_extent(x=43000, y=14000)
    ole.set_line_style(color="#445566", width=77)

    output_path = tmp_path / "native_picture_ole_probe_fields.hwp"
    document.save(output_path)

    reopened = HwpBinaryDocument.open(output_path)
    model = reopened.section_model(0)
    picture_control = next(control for control in model.controls() if isinstance(control, PictureControlRecord))
    ole_control = next(control for control in model.controls() if isinstance(control, OleControlRecord))
    picture_record = picture_control.picture_record()
    ole_record = ole_control.ole_record()

    assert picture_control.layout()["textWrap"] == "SQUARE"
    assert picture_control.layout()["textFlow"] == "RIGHT_ONLY"
    assert picture_control.layout()["treatAsChar"] == "0"
    assert picture_control.layout()["vertRelTo"] == "PAPER"
    assert picture_control.layout()["horzRelTo"] == "PAPER"
    assert picture_control.layout()["vertAlign"] == "CENTER"
    assert picture_control.layout()["horzAlign"] == "RIGHT"
    assert picture_control.layout()["vertOffset"] == "101"
    assert picture_control.layout()["horzOffset"] == "202"
    assert picture_control.out_margins() == {"left": 13, "right": 24, "top": 35, "bottom": 46}
    assert picture_control.rotation() == {
        "angle": "15",
        "centerX": "111",
        "centerY": "222",
        "rotateimage": "0",
    }
    assert picture_control.image_adjustment() == {
        "bright": "5",
        "contrast": "6",
        "effect": "GRAY_SCALE",
        "alpha": "7",
    }
    assert picture_control.crop() == {"left": 1, "right": 2399, "top": 2, "bottom": 1598}
    assert picture_record is not None
    assert picture_record.line_color() == "#112233"
    assert picture_record.line_width() == 55
    assert picture_record.payload[0:4] == bytes.fromhex("11223300")
    assert int.from_bytes(picture_control.payload[8:12], "little", signed=True) == 101
    assert int.from_bytes(picture_control.payload[12:16], "little", signed=True) == 202
    assert picture_control.payload[4] & 0x20
    assert picture_control.payload[4] & 0x01 == 0
    assert picture_control.payload[5] & 0x08
    assert picture_control.payload[7] & 0x03 == 0x02
    assert int.from_bytes(picture_control.shape_component_record().payload[40:42], "little", signed=False) == 15
    assert int.from_bytes(picture_control.shape_component_record().payload[42:46], "little", signed=True) == 111
    assert int.from_bytes(picture_control.shape_component_record().payload[46:50], "little", signed=True) == 222
    assert int.from_bytes(picture_control.payload[28:30], "little", signed=False) == 13
    assert int.from_bytes(picture_control.payload[30:32], "little", signed=False) == 24
    assert int.from_bytes(picture_control.payload[32:34], "little", signed=False) == 35
    assert int.from_bytes(picture_control.payload[34:36], "little", signed=False) == 46

    assert ole_control.layout()["textWrap"] == "SQUARE"
    assert ole_control.layout()["textFlow"] == "LEFT_ONLY"
    assert ole_control.layout()["treatAsChar"] == "0"
    assert ole_control.layout()["vertRelTo"] == "PAPER"
    assert ole_control.layout()["horzRelTo"] == "PAPER"
    assert ole_control.layout()["vertAlign"] == "CENTER"
    assert ole_control.layout()["horzAlign"] == "RIGHT"
    assert ole_control.layout()["vertOffset"] == "77"
    assert ole_control.layout()["horzOffset"] == "88"
    assert ole_control.shape_size() == {"width": 42001, "height": 13501}
    assert ole_control.rotation()["angle"] == "45"
    assert ole_control.extent() == {"x": 43000, "y": 14000}
    assert ole_control.line_color() == "#445566"
    assert ole_control.line_width() == 77
    assert ole_record is not None
    assert int.from_bytes(ole_control.payload[8:12], "little", signed=True) == 77
    assert int.from_bytes(ole_control.payload[12:16], "little", signed=True) == 88
    assert ole_control.payload[4] & 0x20
    assert ole_control.payload[4] & 0x01 == 0
    assert ole_control.payload[5] & 0x08
    assert ole_control.payload[7] & 0x03 == 0x01
    assert ole_record.payload[14:18] == bytes.fromhex("44556600")
    assert int.from_bytes(ole_control.shape_component_record().payload[40:42], "little", signed=False) == 45
    assert int.from_bytes(ole_record.payload[4:8], "little", signed=True) == 43000
    assert int.from_bytes(ole_record.payload[8:12], "little", signed=True) == 14000
    assert int.from_bytes(ole_control.payload[28:30], "little", signed=False) == 9
    assert int.from_bytes(ole_control.payload[30:32], "little", signed=False) == 10
    assert int.from_bytes(ole_control.payload[32:34], "little", signed=False) == 11
    assert int.from_bytes(ole_control.payload[34:36], "little", signed=False) == 12


def test_section_model_parses_native_generic_shape_rotation_payload(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    section = document.section(0)
    section.append_shape(kind="rect", text="ROT-RECT", width=5100, height=2200, fill_color="#ABCDEF", line_color="#112233")

    shape = next(item for item in section.shapes() if item.kind == "rect")
    shape.set_rotation(angle=25, center_x=333, center_y=444)

    output_path = tmp_path / "native_shape_rotation_payload.hwp"
    document.save(output_path)

    reopened = HwpBinaryDocument.open(output_path)
    model = reopened.section_model(0)
    shape_control = next(
        control for control in model.controls() if isinstance(control, ShapeControlRecord) and control.graphic_kind() == "rect"
    )
    component = shape_control.shape_component_record()

    assert shape_control.rotation() == {"angle": "25", "centerX": "333", "centerY": "444"}
    assert component is not None
    assert int.from_bytes(component.payload[40:42], "little", signed=False) == 25
    assert int.from_bytes(component.payload[42:46], "little", signed=True) == 333
    assert int.from_bytes(component.payload[46:50], "little", signed=True) == 444


def test_record_node_from_record_maps_form_memo_list_and_chart_records() -> None:
    form_record = HwpRecord(
        tag_id=hwp_binary_module.TAG_FORM_OBJECT,
        level=2,
        size=len("FORM".encode("utf-16-le")),
        header_size=4,
        offset=-1,
        payload="FORM".encode("utf-16-le"),
    )
    memo_record = HwpRecord(
        tag_id=hwp_binary_module.TAG_MEMO_LIST,
        level=2,
        size=len("MEMO".encode("utf-16-le")),
        header_size=4,
        offset=-1,
        payload="MEMO".encode("utf-16-le"),
    )
    chart_record = HwpRecord(
        tag_id=hwp_binary_module.TAG_CHART_DATA,
        level=3,
        size=len("CHART".encode("utf-16-le")),
        header_size=4,
        offset=-1,
        payload="CHART".encode("utf-16-le"),
    )

    form_node = hwp_binary_module.record_node_from_record(form_record)
    memo_node = hwp_binary_module.record_node_from_record(memo_record)
    chart_node = hwp_binary_module.record_node_from_record(chart_record)

    assert isinstance(form_node, FormObjectRecord)
    assert form_node.fields()["utf16_text"] == "FORM"
    assert form_node.fields()["raw_payload"] == "FORM".encode("utf-16-le")

    assert isinstance(memo_node, MemoListRecord)
    assert memo_node.fields()["utf16_text"] == "MEMO"
    assert memo_node.fields()["raw_payload"] == "MEMO".encode("utf-16-le")

    assert isinstance(chart_node, ChartDataRecord)
    assert chart_node.fields()["utf16_text"] == "CHART"
    assert chart_node.fields()["raw_payload"] == "CHART".encode("utf-16-le")


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


def test_hwp_binary_reencode_is_byte_stable_after_first_rewrite(sample_hwp_path: Path, tmp_path: Path) -> None:
    document = HwpBinaryDocument.open(sample_hwp_path)
    first_path = tmp_path / "reencode_once.hwp"
    second_path = tmp_path / "reencode_twice.hwp"

    document.save_copy(first_path, preserve_original_bytes=False)
    first_bytes = first_path.read_bytes()
    HwpBinaryDocument.open(first_path).save_copy(second_path, preserve_original_bytes=False)

    assert first_bytes == second_path.read_bytes()


def test_hwp_binary_reencode_with_preview_growth_preserves_unmodified_section_chain(
    sample_hwp_path: Path,
    tmp_path: Path,
) -> None:
    document = HwpBinaryDocument.open(sample_hwp_path)
    original_section_chain = _layout_chain(document, "BodyText/Section0")
    first_path = tmp_path / "grown_preview_reencode_once.hwp"
    second_path = tmp_path / "grown_preview_reencode_twice.hwp"

    document.set_preview_text("PREVIEW-" * 900)
    document.save_copy(first_path, preserve_original_bytes=False)
    first_bytes = first_path.read_bytes()

    reopened = HwpBinaryDocument.open(first_path)
    reopened.save_copy(second_path, preserve_original_bytes=False)

    assert first_bytes == second_path.read_bytes()
    assert reopened.preview_text().startswith("PREVIEW-PREVIEW")
    assert len(reopened.preview_text()) >= 7000
    assert _layout_chain(reopened, "BodyText/Section0") == original_section_chain


def test_hwp_binary_reencode_with_preview_shrink_is_byte_stable_after_first_rewrite(
    sample_hwp_path: Path,
    tmp_path: Path,
) -> None:
    grown_path = tmp_path / "preview_grown_once.hwp"
    shrunk_path = tmp_path / "preview_shrunk_once.hwp"
    shrunk_copy_path = tmp_path / "preview_shrunk_twice.hwp"

    grown = HwpBinaryDocument.open(sample_hwp_path)
    grown.set_preview_text("PREVIEW-" * 900)
    grown.save_copy(grown_path, preserve_original_bytes=False)

    shrunk = HwpBinaryDocument.open(grown_path)
    shrunk.set_preview_text("SHORT")
    shrunk.save_copy(shrunk_path, preserve_original_bytes=False)
    shrunk_bytes = shrunk_path.read_bytes()

    reopened = HwpBinaryDocument.open(shrunk_path)
    reopened.save_copy(shrunk_copy_path, preserve_original_bytes=False)

    assert shrunk_bytes == shrunk_copy_path.read_bytes()
    assert reopened.preview_text() == "SHORT"


def test_hwp_binary_reencode_with_added_stream_is_byte_stable_and_preserves_section_chain(
    sample_hwp_path: Path,
    tmp_path: Path,
) -> None:
    document = HwpBinaryDocument.open(sample_hwp_path)
    original_section_chain = _layout_chain(document, "BodyText/Section0")
    first_path = tmp_path / "added_stream_once.hwp"
    second_path = tmp_path / "added_stream_twice.hwp"

    document.add_stream("Jakal/Added.bin", b"ADDED-DATA" * 200)
    document.save_copy(first_path, preserve_original_bytes=False)
    first_bytes = first_path.read_bytes()

    reopened = HwpBinaryDocument.open(first_path)
    reopened.save_copy(second_path, preserve_original_bytes=False)

    assert reopened.has_stream("Jakal/Added.bin") is True
    assert reopened.read_stream("Jakal/Added.bin", decompress=False) == b"ADDED-DATA" * 200
    assert first_bytes == second_path.read_bytes()
    assert _layout_chain(reopened, "BodyText/Section0") == original_section_chain


def test_hwp_binary_reencode_with_removed_stream_is_byte_stable_after_first_rewrite(
    sample_hwp_path: Path,
    tmp_path: Path,
) -> None:
    document = HwpBinaryDocument.open(sample_hwp_path)
    original_section_chain = _layout_chain(document, "BodyText/Section0")
    first_path = tmp_path / "removed_stream_once.hwp"
    second_path = tmp_path / "removed_stream_twice.hwp"

    document.remove_stream("PrvText")
    document.save_copy(first_path, preserve_original_bytes=False)
    first_bytes = first_path.read_bytes()

    reopened = HwpBinaryDocument.open(first_path)
    reopened.save_copy(second_path, preserve_original_bytes=False)

    assert reopened.has_stream("PrvText") is False
    assert first_bytes == second_path.read_bytes()
    assert _layout_chain(reopened, "BodyText/Section0") == original_section_chain


def test_hwp_binary_reencode_with_large_added_stream_uses_difat_and_is_byte_stable(
    sample_hwp_path: Path,
    tmp_path: Path,
) -> None:
    document = HwpBinaryDocument.open(sample_hwp_path)
    first_path = tmp_path / "large_added_stream_once.hwp"
    second_path = tmp_path / "large_added_stream_twice.hwp"
    large_payload = b"L" * (9 * 1024 * 1024)

    document.add_stream("Jakal/Large.bin", large_payload)
    document.save_copy(first_path, preserve_original_bytes=False)
    first_bytes = first_path.read_bytes()

    reopened = HwpBinaryDocument.open(first_path)
    reopened.save_copy(second_path, preserve_original_bytes=False)

    assert reopened.has_stream("Jakal/Large.bin") is True
    assert reopened.read_stream("Jakal/Large.bin", decompress=False) == large_payload
    assert reopened._compound_layout.num_difat_sectors > 0
    assert len(reopened._compound_layout.fat_sector_ids) > 109
    assert first_bytes == second_path.read_bytes()


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
