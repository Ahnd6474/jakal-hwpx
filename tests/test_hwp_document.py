from __future__ import annotations

import struct
from pathlib import Path

import pytest

import jakal_hwpx.hwp_document as hwp_document_module
from jakal_hwpx.hwp_binary import (
    TAG_CHART_DATA,
    TAG_CTRL_DATA,
    TAG_CTRL_HEADER,
    TAG_FORM_OBJECT,
    TAG_LIST_HEADER,
    TAG_MEMO_LIST,
    TAG_PARA_HEADER,
    TAG_SHAPE_COMPONENT,
    TAG_SHAPE_COMPONENT_LINE,
    TAG_SHAPE_COMPONENT_PICTURE,
    TAG_TABLE,
)

from jakal_hwpx import (
    DocInfoModel,
    HwpAutoNumberObject,
    HwpArcShapeObject,
    HwpBinaryDocument,
    HwpBookmarkObject,
    HwpChartObject,
    HwpConnectLineShapeObject,
    HwpContainerShapeObject,
    HwpCurveShapeObject,
    HwpDocument,
    HwpDocumentProperties,
    HwpEllipseShapeObject,
    HwpEquationObject,
    HwpFieldObject,
    HwpFormObject,
    HwpHeaderFooterObject,
    HwpHyperlinkObject,
    HwpLineShapeObject,
    HwpMemoObject,
    HwpNoteObject,
    HwpOleObject,
    HwpParagraphObject,
    HwpPageNumObject,
    HwpPictureObject,
    HwpPolygonShapeObject,
    HwpPureProfile,
    HwpRectangleShapeObject,
    HwpSection,
    HwpShapeObject,
    HwpTableObject,
    HwpTextArtShapeObject,
    HwpxValidationError,
    RecordNode,
    HwpxDocument,
    SectionModel,
    SectionParagraphModel,
)


REPO_ROOT = Path(__file__).resolve().parents[1]
HWP_SAMPLE_DIR = REPO_ROOT / "examples" / "samples" / "hwp"


@pytest.fixture(scope="session")
def sample_hwp_path() -> Path:
    paths = sorted(HWP_SAMPLE_DIR.glob("*.hwp"))
    assert paths, f"No .hwp samples were found under {HWP_SAMPLE_DIR}"
    return next((path for path in paths if not path.name.startswith("generated_")), paths[0])


def test_hwp_document_exposes_object_model_for_sections_and_paragraphs(sample_hwp_path: Path) -> None:
    document = HwpDocument.open(sample_hwp_path)

    properties = document.document_properties()
    assert isinstance(properties, HwpDocumentProperties)
    assert properties.section_count >= 1

    sections = document.sections()
    assert sections
    assert isinstance(sections[0], HwpSection)

    paragraphs = sections[0].paragraphs()
    assert paragraphs
    assert isinstance(paragraphs[0], HwpParagraphObject)
    assert any("2027" in paragraph.text for paragraph in paragraphs)


def test_hwp_document_exposes_docinfo_and_section_models(sample_hwp_path: Path) -> None:
    document = HwpDocument.open(sample_hwp_path)

    docinfo_model = document.docinfo_model()
    assert isinstance(docinfo_model, DocInfoModel)
    assert docinfo_model.document_properties().section_count >= 1

    section_model = document.section_model(0)
    assert isinstance(section_model, SectionModel)
    assert section_model.paragraphs()
    assert document.section(0).model().paragraphs()


def test_hwp_document_pure_append_paragraph_roundtrips(sample_hwp_path: Path, tmp_path: Path) -> None:
    document = HwpDocument.open(sample_hwp_path)
    paragraph = document.append_paragraph("Pure document paragraph")
    assert isinstance(paragraph, HwpParagraphObject)
    assert paragraph.text == "Pure document paragraph"

    output_path = tmp_path / "pure_append_paragraph.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)
    assert "Pure document paragraph" in reopened.get_document_text()


def test_hwp_document_blank_starts_without_profile_body_text(tmp_path: Path) -> None:
    document = HwpDocument.blank()

    assert document.get_document_text().strip() == ""
    assert document.preview_text() == ""

    output_path = tmp_path / "blank_document.hwp"
    document.save(output_path)

    reopened = HwpDocument.open(output_path)
    assert reopened.get_document_text().strip() == ""
    assert reopened.preview_text() == ""


def test_hwp_document_append_paragraph_can_auto_grow_sections(tmp_path: Path) -> None:
    document = HwpDocument.blank()

    document.append_paragraph("SECTION-0", section_index=0)
    document.append_paragraph("SECTION-2", section_index=2)
    document.append_paragraph("SECTION-3", section_index=3)

    output_path = tmp_path / "multi_section_append.hwp"
    document.save(output_path)

    reopened = HwpDocument.open(output_path)
    assert len(reopened.sections()) == 4
    assert any(paragraph.text.strip() == "SECTION-0" for paragraph in reopened.section(0).paragraphs())
    assert any(paragraph.text.strip() == "SECTION-2" for paragraph in reopened.section(2).paragraphs())
    assert any(paragraph.text.strip() == "SECTION-3" for paragraph in reopened.section(3).paragraphs())
    assert "목 차" not in reopened.get_document_text()


def test_hwp_document_bridge_section_settings_roundtrip_to_native_hwp(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    settings = document.section_settings(0)
    settings.set_page_size(width=123456, height=98765, landscape="LANDSCAPE")
    settings.set_margins(left=1111, right=2222, top=3333, bottom=4444, header=5555, footer=6666, gutter=7777)

    output_path = tmp_path / "bridge_section_settings.hwp"
    document.save(output_path)

    reopened = HwpDocument.open(output_path).to_hwpx_document()
    reopened_settings = reopened.section_settings(0)
    assert reopened_settings.page_width == 123456
    assert reopened_settings.page_height == 98765
    assert reopened_settings.margins()["left"] == 1111
    assert reopened_settings.margins()["right"] == 2222
    assert reopened_settings.margins()["top"] == 3333
    assert reopened_settings.margins()["bottom"] == 4444
    assert reopened_settings.margins()["header"] == 5555
    assert reopened_settings.margins()["footer"] == 6666
    assert reopened_settings.margins()["gutter"] == 7777


def test_hwp_document_encodes_extended_style_fields_into_native_hwp(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    para_style = document.append_paragraph_style(style_id="41", alignment_horizontal="CENTER", line_spacing=180)
    para_style.element.set("tabPrIDRef", "3")
    para_style.element.set("snapToGrid", "1")
    para_style.header_part.mark_modified()
    para_style.set_break_setting(keep_with_next=True, keep_lines=True, page_break_before=True, line_wrap="SQUEEZE")
    para_style.set_auto_spacing(e_asian_eng=True, e_asian_num=False)

    char_style = document.append_character_style(style_id="42", text_color="#224466", height=1250)
    char_style.set_effects(shade_color="#EEEEEE", use_font_space=True, use_kerning=True, sym_mark="DOT_ABOVE")

    output_path = tmp_path / "native_style_encoding.hwp"
    document.save(output_path)

    reopened = HwpBinaryDocument.open(output_path)
    para_fields = reopened.docinfo_model().para_shape_records()[41].fields()
    char_fields = reopened.docinfo_model().char_shape_records()[42].fields()

    assert para_fields["alignment_horizontal"] == "CENTER"
    assert para_fields["tab_def_id"] == 3
    assert para_fields["attribute_bits"]["snap_to_grid"] is True
    assert para_fields["attribute_bits"]["keep_with_next"] is True
    assert para_fields["attribute_bits"]["keep_lines"] is True
    assert para_fields["attribute_bits"]["page_break_before"] is True
    assert para_fields["attribute_bits"]["line_wrap_squeeze"] is True
    assert para_fields["auto_spacing"] == {"eAsianEng": "1", "eAsianNum": "0"}
    assert para_fields["line_spacing"] == 180

    assert char_fields["text_color"] == "#224466"
    assert char_fields["shade_color"] == "#EEEEEE"
    assert char_fields["hwpx_compatible_flags"] == {"useFontSpace": True, "useKerning": True}


def test_hwp_document_can_roundtrip_section_page_border_fills_to_native_hwp(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    document.apply_section_page_border_fills(
        [
            {
                "type": "BOTH",
                "borderFillIDRef": "1",
                "textBorder": "BORDER",
                "headerInside": "1",
                "footerInside": "1",
                "fillArea": "BORDER",
                "left": 2222,
                "right": 3333,
                "top": 4444,
                "bottom": 5555,
            },
            {
                "type": "EVEN",
                "borderFillIDRef": "1",
                "textBorder": "PAPER",
                "headerInside": "0",
                "footerInside": "0",
                "fillArea": "BORDER",
                "left": 123,
                "right": 456,
                "top": 789,
                "bottom": 987,
            },
            {
                "type": "ODD",
                "borderFillIDRef": "1",
                "textBorder": "BORDER",
                "headerInside": "0",
                "footerInside": "0",
                "fillArea": "PAPER",
                "left": 11,
                "right": 22,
                "top": 33,
                "bottom": 44,
            },
        ]
    )

    output_path = tmp_path / "bridge_section_page_border_fill.hwp"
    document.save(output_path)

    reopened = HwpDocument.open(output_path).to_hwpx_document()
    both = reopened.section_xml(0).find(".//hp:pageBorderFill[@type='BOTH']")
    even = reopened.section_xml(0).find(".//hp:pageBorderFill[@type='EVEN']")
    odd = reopened.section_xml(0).find(".//hp:pageBorderFill[@type='ODD']")
    assert both is not None
    assert even is not None
    assert odd is not None
    assert both.get_attr("textBorder") == "BORDER"
    assert both.get_attr("headerInside") == "1"
    assert both.get_attr("footerInside") == "1"
    assert both.get_attr("fillArea") == "BORDER"
    assert both.find("./hp:offset").get_attr("left") == "2222"
    assert both.find("./hp:offset").get_attr("right") == "3333"
    assert even.get_attr("fillArea") == "BORDER"
    assert even.find("./hp:offset").get_attr("top") == "789"
    assert odd.get_attr("textBorder") == "BORDER"
    assert odd.find("./hp:offset").get_attr("bottom") == "44"


def test_hwp_document_can_roundtrip_section_definition_settings_to_native_hwp(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    settings = document.section_settings(0)
    settings.set_grid(line_grid=42, char_grid=21, wonggoji_format=True)
    settings.set_start_numbers(page_starts_on="ODD", page=7, pic=8, tbl=9, equation=10)
    settings.set_visibility(
        hide_first_header=True,
        hide_first_footer=True,
        hide_first_master_page=True,
        hide_first_page_num=True,
        hide_first_empty_line=True,
    )

    output_path = tmp_path / "bridge_section_definition_settings.hwp"
    document.save(output_path)

    reopened = HwpDocument.open(output_path).to_hwpx_document()
    reopened_settings = reopened.section_settings(0)
    assert reopened_settings.grid() == {"lineGrid": 42, "charGrid": 21, "wonggojiFormat": 1}
    assert reopened_settings.start_numbers() == {
        "pageStartsOn": "ODD",
        "page": "7",
        "pic": "8",
        "tbl": "9",
        "equation": "10",
    }
    assert reopened_settings.visibility()["hideFirstHeader"] == "1"
    assert reopened_settings.visibility()["hideFirstFooter"] == "1"
    assert reopened_settings.visibility()["hideFirstMasterPage"] == "1"
    assert reopened_settings.visibility()["hideFirstPageNum"] == "1"
    assert reopened_settings.visibility()["hideFirstEmptyLine"] == "1"


def test_hwp_document_can_roundtrip_section_numbering_shape_and_picture_size(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    image_bytes = document.binary_document().read_stream("BinData/BIN0001.bmp", decompress=False)
    document.apply_section_settings(section_index=0, numbering_shape_id="2")
    document.append_picture(image_bytes, extension="bmp", width=3600, height=2400)

    output_path = tmp_path / "section_numbering_shape_picture_size.hwp"
    document.save(output_path)

    reopened = HwpDocument.open(output_path)
    native_settings = reopened.binary_document().section_definition_settings(0)
    assert native_settings["numbering_shape_id"] == "2"
    assert reopened.pictures()[0].size() == {"width": 3600, "height": 2400}

    reopened_hwpx = reopened.to_hwpx_document()
    picture = reopened_hwpx.pictures()[0]
    assert picture.size() == {"width": 3600, "height": 2400}


def test_hwp_document_can_roundtrip_native_show_line_number_with_best_effort_line_number_shape(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    document.set_paragraph_text(0, 0, "alpha")
    document.append_paragraph("beta")
    document.append_paragraph("gamma")
    settings = document.section_settings(0)
    settings.set_visibility(show_line_number=True)
    line_number_shape = document.section_xml(0).find(".//hp:lineNumberShape")
    assert line_number_shape is not None
    line_number_shape.set_attr("restartType", 1).set_attr("countBy", 3).set_attr("distance", 150).set_attr("startNumber", 7)

    output_path = tmp_path / "native_show_line_number.hwp"
    document.save(output_path)

    reopened_hwp = HwpDocument.open(output_path)
    native_settings = reopened_hwp.binary_document().section_definition_settings(0)
    assert native_settings["visibility"]["showLineNumber"] == "1"

    reopened = reopened_hwp.to_hwpx_document()
    reopened_settings = reopened.section_settings(0)
    assert reopened_settings.visibility()["showLineNumber"] == "1"
    reopened_line_number = reopened.section_xml(0).find(".//hp:lineNumberShape")
    assert reopened_line_number is not None
    assert reopened_line_number.get_attr("restartType") == "0"
    assert reopened_line_number.get_attr("countBy") == "0"
    assert reopened_line_number.get_attr("distance") == "0"
    assert reopened_line_number.get_attr("startNumber") == "0"


def test_hwp_paragraph_object_can_replace_same_length_text_and_roundtrip(
    sample_hwp_path: Path,
    tmp_path: Path,
) -> None:
    document = HwpDocument.open(sample_hwp_path)
    target = next(paragraph for paragraph in document.section(0).paragraphs() if "2027" in paragraph.text)

    replacements = target.replace_text_same_length("2027", "2028", count=1)
    assert replacements == 1

    output_path = tmp_path / "object_model_roundtrip.hwp"
    document.save(output_path)

    reopened = HwpDocument.open(output_path)
    assert "2028" in reopened.get_document_text()


def test_hwp_section_object_can_scope_same_length_replacements(sample_hwp_path: Path, tmp_path: Path) -> None:
    document = HwpDocument.open(sample_hwp_path)
    section = document.section(0)

    replacements = section.replace_text_same_length("2027", "2029", count=1)
    assert replacements == 1

    output_path = tmp_path / "section_scope_roundtrip.hwp"
    document.save(output_path)

    reopened = HwpDocument.open(output_path)
    assert "2029" in reopened.get_document_text()


def test_hwp_document_can_bridge_to_hwpx_high_level_api(
    sample_hwp_path: Path,
    sample_hwpx_path: Path,
    tmp_path: Path,
) -> None:
    conversions: list[tuple[str, str, str]] = []

    def fake_converter(input_path: str | Path, output_path: str | Path, output_format: str) -> Path:
        source = Path(input_path)
        target = Path(output_path)
        conversions.append((str(source), str(target), output_format))
        if output_format == "HWPX":
            target.write_bytes(sample_hwpx_path.read_bytes())
        elif output_format == "HWP":
            target.write_bytes(sample_hwp_path.read_bytes())
        else:
            raise AssertionError(output_format)
        return target

    document = HwpDocument.open(sample_hwp_path, converter=fake_converter)
    bridge = document.to_hwpx_document()
    assert isinstance(bridge, HwpxDocument)
    assert document.metadata().title == bridge.metadata().title

    output_hwpx = tmp_path / "bridged.hwpx"
    document.save(output_hwpx)
    assert output_hwpx.exists()

    output_hwp = tmp_path / "bridged.hwp"
    document.save(output_hwp)
    assert output_hwp.exists()
    assert conversions == []


def test_hwp_document_exposes_direct_pure_python_control_append_methods(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    target_section_index = document._default_append_section_index
    before_count = len(document.binary_document().section_records(target_section_index))
    document.append_table()
    document.append_picture()
    document.append_hyperlink()
    after_count = len(document.binary_document().section_records(target_section_index))
    assert after_count > before_count

    output_path = tmp_path / "direct_control_append.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)
    assert len(reopened.binary_document().section_records(target_section_index)) == after_count


def test_hwp_document_can_append_picture_from_bytes(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    image_bytes = document.binary_document().read_stream("BinData/BIN0001.bmp", decompress=False)
    before_bindata_count = document.binary_document().docinfo_model().id_mappings_record().bin_data_count

    document.append_picture(image_bytes, extension="bmp")

    after_bindata_count = document.binary_document().docinfo_model().id_mappings_record().bin_data_count
    assert after_bindata_count == before_bindata_count + 1

    output_path = tmp_path / "append_picture_bytes_document.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)
    assert reopened.binary_document().docinfo_model().id_mappings_record().bin_data_count == after_bindata_count


def test_hwp_document_can_append_custom_table_and_hyperlink(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    document.append_table("DOC-CELL")
    document.append_hyperlink("https://example.com/doc", text="DOC-LINK")

    output_path = tmp_path / "custom_table_hyperlink_document.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)
    text = reopened.get_document_text()
    assert "DOC-CELL" in text
    assert "DOC-LINK" in text


def test_hwp_document_supports_multi_cell_tables_and_hyperlink_metadata(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    document.append_table(rows=2, cols=2, cell_texts=[["A1", "A2"], ["B1", "B2"]])
    document.append_hyperlink(
        "https://example.com/document-meta",
        text="DOC-META",
        metadata_fields=[3, 2, "anchor"],
    )

    output_path = tmp_path / "document_multi_cell_metadata.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)
    text = reopened.get_document_text()
    for value in ("A1", "A2", "B1", "B2", "DOC-META"):
        assert value in text


def test_hwp_document_supports_table_geometry_options(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    document.append_table(
        rows=2,
        cols=3,
        cell_texts=[["M1", "", "M2"], ["M3", "M4", ""]],
        row_heights=[150, 250],
        col_widths=[1200, 1300, 1400],
        cell_spans={(0, 0): (1, 2), (1, 1): (1, 2)},
        cell_border_fill_ids={(0, 0): 7, (1, 1): 8},
        table_border_fill_id=6,
    )

    output_path = tmp_path / "document_table_geometry.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)
    text = reopened.get_document_text()
    for value in ("M1", "M2", "M3", "M4"):
        assert value in text


def test_hwp_document_exposes_table_picture_and_hyperlink_objects(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    image_bytes = document.binary_document().read_stream("BinData/BIN0001.bmp", decompress=False)

    document.append_table(
        rows=2,
        cols=2,
        cell_texts=[["T1", "T2"], ["T3", "T4"]],
        row_heights=[101, 202],
        col_widths=[1200, 1300],
        cell_border_fill_ids={(0, 0): 7},
        table_border_fill_id=6,
    )
    document.append_picture(image_bytes, extension="bmp")
    document.append_hyperlink(
        "https://example.com/object-layer",
        text="OBJ-LINK",
        metadata_fields=[9, 4, "anchor"],
    )

    table = next(table for table in document.tables() if any(cell.text == "T4" for cell in table.cells()))
    picture = max(document.pictures(), key=lambda item: item.storage_id)
    hyperlink = next(link for link in document.hyperlinks() if link.display_text == "OBJ-LINK")

    assert isinstance(table, HwpTableObject)
    assert table.row_count == 2
    assert table.column_count == 2
    assert table.row_heights == [101, 202]
    assert table.table_border_fill_id == 6
    assert table.cell(0, 0).text == "T1"
    assert table.cell(0, 0).border_fill_id == 7
    assert table.cell_text_matrix()[1][1] == "T4"

    assert isinstance(picture, HwpPictureObject)
    assert picture.bindata_path().endswith(".bmp")
    assert picture.extension == "bmp"
    assert picture.binary_data() == image_bytes

    assert isinstance(hyperlink, HwpHyperlinkObject)
    assert hyperlink.url == "https://example.com/object-layer"
    assert hyperlink.display_text == "OBJ-LINK"
    assert hyperlink.metadata_fields == ["9", "4", "anchor"]


def test_hwp_control_objects_roundtrip_after_save(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    image_bytes = document.binary_document().read_stream("BinData/BIN0001.bmp", decompress=False)
    document.append_table("ROUNDTRIP-CELL")
    document.append_picture(image_bytes, extension="bmp")
    document.append_hyperlink("https://example.com/roundtrip", text="ROUNDTRIP-LINK")

    output_path = tmp_path / "hwp_control_objects.hwp"
    document.save(output_path)

    reopened = HwpDocument.open(output_path)
    assert any(cell.text == "ROUNDTRIP-CELL" for table in reopened.tables() for cell in table.cells())
    assert any(picture.extension == "bmp" for picture in reopened.pictures())
    assert any(link.url == "https://example.com/roundtrip" and link.display_text == "ROUNDTRIP-LINK" for link in reopened.hyperlinks())


def test_hwp_document_exposes_field_shape_and_equation_objects(sample_hwp_path: Path) -> None:
    document = HwpDocument.open(sample_hwp_path)

    shapes = document.shapes()
    equations = document.equations()

    assert any(isinstance(shape, HwpShapeObject) for shape in shapes)
    assert any(isinstance(equation, HwpEquationObject) for equation in equations)
    assert any(equation.script for equation in equations)
    assert any(shape.kind in {"rect", "line", "ellipse", "arc", "polygon", "curve", "container", "textart", "shape"} for shape in shapes)


def test_hwp_hyperlinks_are_included_in_generic_fields() -> None:
    document = HwpDocument.blank()
    document.append_hyperlink("https://example.com/field-view", text="FIELD-VIEW")

    fields = document.fields()
    assert any(isinstance(field, HwpHyperlinkObject) for field in fields)
    assert any(isinstance(field, HwpFieldObject) and getattr(field, "field_type", "") == "%hlk" for field in fields)


def test_hwp_ole_wrapper_can_classify_ole_controls_without_sample() -> None:
    document = HwpDocument.blank()
    paragraph = SectionParagraphModel(section_index=0, index=0, header=document.section_model(0).paragraphs()[0].header)
    control_node = RecordNode(tag_id=71, level=1, payload=b" osg" + (b"\x00" * 42))
    shape_component = RecordNode(tag_id=76, level=2, payload=b"")
    shape_component.add_child(RecordNode(tag_id=84, level=3, payload=b""))
    control_node.add_child(shape_component)

    ole = HwpOleObject(document, paragraph, control_node, 0)
    assert ole.kind == "ole"


def test_hwp_form_memo_and_chart_wrappers_can_classify_and_expose_payloads() -> None:
    document = HwpDocument.blank()
    section_model = document.section_model(0)
    paragraph = section_model.paragraphs()[0]

    form_control = RecordNode(tag_id=TAG_CTRL_HEADER, level=1, payload=b"mrof" + (b"\x00" * 42))
    form_control.add_child(RecordNode(tag_id=TAG_FORM_OBJECT, level=2, payload="FORM-FIELD".encode("utf-16-le")))
    paragraph.header.add_child(form_control)

    memo_control = RecordNode(tag_id=TAG_CTRL_HEADER, level=1, payload=b"omem" + (b"\x00" * 42))
    memo_control.add_child(RecordNode(tag_id=TAG_MEMO_LIST, level=2, payload="MEMO-BODY".encode("utf-16-le")))
    paragraph.header.add_child(memo_control)

    chart_control = RecordNode(tag_id=TAG_CTRL_HEADER, level=1, payload=b" osg" + (b"\x00" * 42))
    shape_component = RecordNode(tag_id=TAG_SHAPE_COMPONENT, level=2, payload=b"")
    shape_component.add_child(RecordNode(tag_id=TAG_CHART_DATA, level=3, payload="CHART-TITLE".encode("utf-16-le")))
    chart_control.add_child(shape_component)
    paragraph.header.add_child(chart_control)

    document.binary_document().replace_section_model(0, section_model)

    form = document.forms()[0]
    memo = document.memos()[0]
    chart = document.charts()[0]

    assert isinstance(form, HwpFormObject)
    assert form.kind == "form"
    assert form.fields()["utf16_text"] == "FORM-FIELD"
    assert form.raw_payload == "FORM-FIELD".encode("utf-16-le")

    assert isinstance(memo, HwpMemoObject)
    assert memo.kind == "memo"
    assert memo.text == "MEMO-BODY"
    assert memo.raw_payload == "MEMO-BODY".encode("utf-16-le")

    assert isinstance(chart, HwpChartObject)
    assert chart.kind == "chart"
    assert chart.utf16_text == "CHART-TITLE"
    assert chart.raw_payload == "CHART-TITLE".encode("utf-16-le")


def test_hwp_form_memo_and_chart_objects_support_semantic_edit_roundtrip(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    document.append_form(
        "Accept terms",
        form_type="CHECKBOX",
        name="accept",
        value="Y",
        checked=True,
        items=["Y", "N"],
        editable=True,
        locked=False,
        placeholder="choose",
    )
    document.append_memo(
        "Need review",
        author="Kim",
        memo_id="memo-1",
        anchor_id="p-1",
        order=3,
        visible=False,
    )
    document.append_chart(
        "Sales",
        chart_type="BAR",
        categories=["Q1", "Q2"],
        series=[{"name": "Revenue", "values": [10, 20]}],
        data_ref="Sheet1!A1:B3",
        legend_visible=True,
        width=9000,
        height=4200,
        shape_comment="chart comment",
    )

    form = document.forms()[-1]
    memo = document.memos()[-1]
    chart = document.charts()[-1]

    assert form.label == "Accept terms"
    assert form.form_type == "CHECKBOX"
    assert form.name == "accept"
    assert form.value == "Y"
    assert form.checked is True
    assert form.items == ["Y", "N"]
    assert form.editable is True
    assert form.locked is False
    assert form.placeholder == "choose"

    assert memo.text == "Need review"
    assert memo.author == "Kim"
    assert memo.memo_id == "memo-1"
    assert memo.anchor_id == "p-1"
    assert memo.order == 3
    assert memo.visible is False

    assert chart.title == "Sales"
    assert chart.chart_type == "BAR"
    assert chart.categories == ["Q1", "Q2"]
    assert chart.series == [{"name": "Revenue", "values": [10, 20]}]
    assert chart.data_ref == "Sheet1!A1:B3"
    assert chart.legend_visible is True
    assert chart.size() == {"width": 9000, "height": 4200}
    assert chart.shape_comment == "chart comment"

    form.set_label("Pick a value")
    form.set_form_type("COMBO")
    form.set_name("accept_state")
    form.set_value("N")
    form.set_checked(False)
    form.set_items(["A", "B", "C"])
    form.set_editable(False)
    form.set_locked(True)
    form.set_placeholder("pick one")

    memo.set_text("Reviewed")
    memo.set_author("Lee")
    memo.set_memo_id("memo-2")
    memo.set_anchor_id("p-2")
    memo.set_order(9)
    memo.set_visible(True)

    chart.set_title("Revenue")
    chart.set_chart_type("LINE")
    chart.set_categories(["Jan", "Feb", "Mar"])
    chart.set_series(
        [
            {"name": "Sales", "values": [1, 2, 3]},
            {"name": "Cost", "values": [0, 1, 1]},
        ]
    )
    chart.set_data_ref("Sheet1!A1:C4")
    chart.set_legend_visible(False)
    chart.set_size(width=9600, height=4800)

    output_path = tmp_path / "semantic_controls.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)

    reopened_form = reopened.forms()[-1]
    reopened_memo = reopened.memos()[-1]
    reopened_chart = reopened.charts()[-1]

    assert reopened_form.fields()["label"] == "Pick a value"
    assert reopened_form.form_type == "COMBO"
    assert reopened_form.name == "accept_state"
    assert reopened_form.value == "N"
    assert reopened_form.checked is False
    assert reopened_form.items == ["A", "B", "C"]
    assert reopened_form.editable is False
    assert reopened_form.locked is True
    assert reopened_form.placeholder == "pick one"
    assert reopened_form.raw_payload == "Pick a value".encode("utf-16-le")

    assert reopened_memo.fields()["text"] == "Reviewed"
    assert reopened_memo.author == "Lee"
    assert reopened_memo.memo_id == "memo-2"
    assert reopened_memo.anchor_id == "p-2"
    assert reopened_memo.order == 9
    assert reopened_memo.visible is True
    assert reopened_memo.raw_payload == "Reviewed".encode("utf-16-le")

    assert reopened_chart.fields()["title"] == "Revenue"
    assert reopened_chart.chart_type == "LINE"
    assert reopened_chart.categories == ["Jan", "Feb", "Mar"]
    assert reopened_chart.series == [
        {"name": "Sales", "values": [1, 2, 3]},
        {"name": "Cost", "values": [0, 1, 1]},
    ]
    assert reopened_chart.data_ref == "Sheet1!A1:C4"
    assert reopened_chart.legend_visible is False
    assert reopened_chart.size() == {"width": 9600, "height": 4800}
    assert reopened_chart.raw_payload == "Revenue".encode("utf-16-le")


def test_hwp_object_methods_can_modify_content_and_roundtrip(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    image_bytes = document.binary_document().read_stream("BinData/BIN0001.bmp", decompress=False)
    document.append_table(rows=1, cols=1, cell_texts=[["CELL-OLD"]])
    document.append_hyperlink("https://example.com/old", text="LINK-OLD", metadata_fields=[1, 0, 0])
    document.append_picture(image_bytes, extension="bmp")

    table = document.tables()[-1]
    hyperlink = document.hyperlinks()[-1]
    picture = document.pictures()[-1]

    table.set_cell_text(0, 0, "CELL-NEW")
    table.set_row_heights([321])
    table.set_table_border_fill_id(8)
    hyperlink.set_url("https://example.com/new")
    hyperlink.set_metadata_fields([9, 2, "dest"])
    hyperlink.set_display_text("LINK-NEW")
    picture.replace_binary(image_bytes, extension="bmp")

    output_path = tmp_path / "hwp_object_methods.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)

    assert reopened.tables()[-1].cell(0, 0).text == "CELL-NEW"
    assert reopened.tables()[-1].row_heights == [321]
    assert reopened.tables()[-1].table_border_fill_id == 8
    assert reopened.hyperlinks()[-1].url == "https://example.com/new"
    assert reopened.hyperlinks()[-1].display_text == "LINK-NEW"
    assert reopened.hyperlinks()[-1].metadata_fields == ["9", "2", "dest"]
    assert reopened.pictures()[-1].binary_data() == image_bytes


def test_hwp_field_and_control_setters_can_reconfigure_native_controls(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    section = document.section(0)
    section.apply_settings(
        page_width=60123,
        page_height=85234,
        landscape="WIDELY",
        margins={"left": 7000, "right": 7100},
        visibility={"hideFirstHeader": "1", "hideFirstPageNum": "1"},
        grid={"lineGrid": 42, "charGrid": 21, "wonggojiFormat": 1},
        start_numbers={"pageStartsOn": "ODD", "page": "7", "pic": "8", "tbl": "9", "equation": "10"},
    )
    section.apply_page_numbers([{"pos": "BOTTOM_CENTER", "formatType": "DIGIT", "sideChar": "-"}])
    section.apply_note_settings(footnote_pr={"numbering": {"type": "CONTINUOUS", "newNum": "3"}})
    section.append_doc_property_field("Subject", display_text="FIELD-OLD")
    section.append_header("HDR-OLD", apply_page_type="BOTH")
    section.append_footnote("NOTE-OLD")
    section.append_auto_number(kind="autoNum", number_type="PAGE")

    field = section.fields()[0]
    header = next(control for control in section.controls() if isinstance(control, HwpHeaderFooterObject) and control.kind == "header")
    note = next(control for control in section.controls() if isinstance(control, HwpNoteObject) and control.kind == "footNote")
    page_num = section.page_numbers()[0]
    auto_number = next(control for control in section.controls() if isinstance(control, HwpAutoNumberObject) and control.kind == "autoNum")

    field.configure_cross_reference("bookmark_target", display_text="FIELD-NEW")
    header.set_text("HDR-NEW")
    header.set_apply_page_type("ODD")
    note.set_text("NOTE-NEW")
    note.set_number(9)
    page_num.set_pos("TOP_RIGHT")
    page_num.set_format_type("ROMAN_SMALL")
    page_num.set_side_char("*")
    auto_number.set_number_type("PAGE")

    output_path = tmp_path / "hwp_native_control_setters.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)

    reopened_field = reopened.fields()[0]
    reopened_header = next(control for control in reopened.controls() if isinstance(control, HwpHeaderFooterObject) and control.kind == "header")
    reopened_note = next(control for control in reopened.controls() if isinstance(control, HwpNoteObject) and control.kind == "footNote")
    reopened_page_num = reopened.page_numbers()[0]
    reopened_auto_number = next(
        control for control in reopened.controls() if isinstance(control, HwpAutoNumberObject) and control.kind == "autoNum"
    )

    assert reopened_field.field_type == "CROSSREF"
    assert reopened_field.bookmark_name == "bookmark_target"
    assert reopened_field.display_text == "FIELD-NEW"
    assert reopened_header.text == "HDR-NEW"
    assert reopened_header.apply_page_type == "ODD"
    assert reopened_note.text == "NOTE-NEW"
    assert reopened_note.number == "9"
    assert reopened_page_num.page_number == {"pos": "TOP_RIGHT", "formatType": "ROMAN_SMALL", "sideChar": "*"}
    assert reopened_auto_number.number_type == "PAGE"
    assert reopened.binary_document().section_page_settings(0)["page_width"] == 60123
    assert reopened.binary_document().section_definition_settings(0)["start_numbers"]["pageStartsOn"] == "ODD"


def test_hwp_paragraph_object_set_text_can_change_length(sample_hwp_path: Path, tmp_path: Path) -> None:
    document = HwpDocument.open(sample_hwp_path)
    paragraph = next(item for item in document.section(0).paragraphs() if "2027" in item.text)
    paragraph.set_text(paragraph.text.replace("2027", "2035-UPDATED"))

    output_path = tmp_path / "hwp_paragraph_set_text.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)
    assert "2035-UPDATED" in reopened.get_document_text()


def test_hwp_equation_and_shape_object_methods_can_modify_records(sample_hwp_path: Path, tmp_path: Path) -> None:
    document = HwpDocument.open(sample_hwp_path)
    equation = document.equations()[0]
    shape = document.shapes()[0]

    original_size = shape.size()
    equation.set_script("x+y")
    shape.set_size(width=original_size["width"] + 10, height=original_size["height"] + 20)

    output_path = tmp_path / "hwp_equation_shape_edit.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)

    assert reopened.equations()[0].script == "x+y"
    assert reopened.shapes()[0].size()["width"] == original_size["width"] + 10
    assert reopened.shapes()[0].size()["height"] == original_size["height"] + 20


def test_hwp_shape_wrapper_exposes_specific_fields_for_line_and_polygon(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    document.append_shape(kind="line", width=1500, height=900, line_color="#102030")
    document.append_shape(kind="polygon", text="NATIVE-POLYGON", width=4200, height=2400, fill_color="#DDEEFF", line_color="#304050")

    output_path = tmp_path / "hwp_shape_specific_fields.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)

    line = next(shape for shape in reopened.shapes() if shape.kind == "line")
    polygon = next(shape for shape in reopened.shapes() if shape.kind == "polygon")

    assert "start" in line.specific_fields()
    assert "end" in line.specific_fields()
    assert line.specific_fields()["start"] == {"x": 0, "y": 0}
    assert "points" in polygon.specific_fields()
    assert "u16_values" in polygon.specific_fields()


def test_hwp_document_dispatches_gso_subtypes_to_wrapper_classes(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    image_bytes = document.binary_document().read_stream("BinData/BIN0001.bmp", decompress=False)
    section = document.section(0)
    section.append_shape(kind="line", width=1500, height=900, line_color="#102030")
    section.append_shape(kind="rect", width=2200, height=1200, fill_color="#E0E8F0")
    section.append_shape(kind="ellipse", width=2400, height=1400, fill_color="#AACCEE")
    section.append_shape(kind="arc", width=2600, height=1600, line_color="#204060")
    section.append_shape(kind="polygon", text="POLYGON", width=2800, height=1800, fill_color="#DDEEFF")
    section.append_shape(kind="textart", text="TEXTART", width=3000, height=1900, fill_color="#FFE0AA")
    document.append_picture(image_bytes, extension="bmp")
    document.append_ole("wrapped.ole", b"OLE-DATA")

    output_path = tmp_path / "hwp_gso_subtype_wrappers.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)
    reopened_section = reopened.section(0)

    assert any(isinstance(shape, HwpLineShapeObject) for shape in reopened.shapes())
    assert any(isinstance(shape, HwpRectangleShapeObject) for shape in reopened.shapes())
    assert any(isinstance(shape, HwpEllipseShapeObject) for shape in reopened.shapes())
    assert any(isinstance(shape, HwpArcShapeObject) for shape in reopened.shapes())
    assert any(isinstance(shape, HwpPolygonShapeObject) for shape in reopened.shapes())
    assert any(isinstance(shape, HwpTextArtShapeObject) for shape in reopened.shapes())

    line = reopened.lines()[0]
    rectangle = reopened.rectangles()[0]
    ellipse = reopened.ellipses()[0]
    arc = reopened.arcs()[0]
    polygon = reopened.polygons()[0]
    textart = reopened.textarts()[0]

    assert line.start == {"x": 0, "y": 0}
    assert line.end["x"] == 1500
    assert rectangle.corner_radius >= 0
    assert ellipse.center["x"] >= 0
    assert arc.axis1["x"] >= 0
    assert polygon.point_count >= 0
    assert isinstance(polygon.points, list)
    assert textart.native_text == "TEXTART"
    assert isinstance(reopened.pictures()[0], HwpPictureObject)
    assert isinstance(reopened.oles()[0], HwpOleObject)
    assert isinstance(reopened_section.lines()[0], HwpLineShapeObject)
    assert isinstance(reopened_section.textarts()[0], HwpTextArtShapeObject)


def test_hwp_gso_subtype_setters_roundtrip_native_payloads(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    section = document.section(0)
    section.append_shape(kind="line", width=1500, height=900, line_color="#102030")
    section.append_shape(kind="rect", width=2200, height=1200, fill_color="#E0E8F0")
    section.append_shape(kind="ellipse", width=2400, height=1400, fill_color="#AACCEE")
    section.append_shape(kind="arc", width=2600, height=1600, line_color="#204060")
    section.append_shape(kind="polygon", text="POLYGON", width=2800, height=1800, fill_color="#DDEEFF")
    section.append_shape(kind="textart", text="TEXTART", width=3000, height=1900, fill_color="#FFE0AA")

    line = section.lines()[0]
    rectangle = section.rectangles()[0]
    ellipse = section.ellipses()[0]
    arc = section.arcs()[0]
    polygon = section.polygons()[0]
    textart = section.textarts()[0]

    line.set_endpoints(start={"x": 10, "y": 20}, end={"x": 310, "y": 410}, attributes=55)
    rectangle.set_corner_radius(77)
    ellipse.set_geometry(
        arc_flags=5,
        center={"x": 120, "y": 80},
        axis1={"x": 60, "y": 0},
        axis2={"x": 0, "y": 40},
        start={"x": 240, "y": 80},
        end={"x": 120, "y": 0},
        start_2={"x": 240, "y": 80},
        end_2={"x": 120, "y": 160},
    )
    arc.set_geometry(
        arc_type=7,
        center={"x": 130, "y": 90},
        axis1={"x": 70, "y": 0},
        axis2={"x": 0, "y": 50},
    )
    polygon.set_points([{"x": 0, "y": 0}, {"x": 100, "y": 0}, {"x": 50, "y": 80}])
    textart.set_text("TEXTART-UPDATED")
    textart.set_style(font_name="Batang", font_style="Italic")

    output_path = tmp_path / "hwp_gso_subtype_setters.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)

    reopened_line = reopened.lines()[0]
    reopened_rectangle = reopened.rectangles()[0]
    reopened_ellipse = reopened.ellipses()[0]
    reopened_arc = reopened.arcs()[0]
    reopened_polygon = reopened.polygons()[0]
    reopened_textart = reopened.textarts()[0]

    assert reopened_line.start == {"x": 10, "y": 20}
    assert reopened_line.end == {"x": 310, "y": 410}
    assert reopened_line.attributes == 55
    assert reopened_rectangle.corner_radius == 77
    assert reopened_ellipse.arc_flags == 5
    assert reopened_ellipse.center == {"x": 120, "y": 80}
    assert reopened_ellipse.axis1 == {"x": 60, "y": 0}
    assert reopened_ellipse.axis2 == {"x": 0, "y": 40}
    assert reopened_arc.arc_type == 7
    assert reopened_arc.center == {"x": 130, "y": 90}
    assert reopened_arc.axis1 == {"x": 70, "y": 0}
    assert reopened_arc.axis2 == {"x": 0, "y": 50}
    assert reopened_polygon.point_count == 3
    assert reopened_polygon.points == [{"x": 0, "y": 0}, {"x": 100, "y": 0}, {"x": 50, "y": 80}]
    assert reopened_textart.native_text == "TEXTART-UPDATED"
    assert reopened_textart.font_name == "Batang"
    assert reopened_textart.font_style == "Italic"


def test_hwp_connect_line_wrapper_roundtrips_and_stays_out_of_plain_lines(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    section = document.section(0)
    section.append_shape(kind="line", width=1500, height=900, line_color="#102030")
    section.append_shape(kind="connectLine", width=1700, height=950, line_color="#304050")

    connect_line = section.connect_lines()[0]
    connect_line.set_endpoints(start={"x": 15, "y": 25}, end={"x": 215, "y": 325}, attributes=9)

    output_path = tmp_path / "hwp_connect_line_wrapper.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)

    assert len(reopened.lines()) == 1
    assert len(reopened.connect_lines()) == 1
    assert len(reopened.section(0).lines()) == 1
    assert len(reopened.section(0).connect_lines()) == 1

    reopened_connect_line = reopened.connect_lines()[0]
    assert isinstance(reopened_connect_line, HwpConnectLineShapeObject)
    assert reopened_connect_line.kind == "connectLine"
    assert reopened_connect_line.start == {"x": 15, "y": 25}
    assert reopened_connect_line.end == {"x": 215, "y": 325}
    assert reopened_connect_line.attributes == 9


def test_hwp_document_detects_hancom_connectline_signature_without_ctrl_data(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    document.append_shape(kind="line", width=1700, height=950, line_color="#304050", shape_comment="HANCOM-CONNECTLINE")

    section_model = document.section_model(0)
    line_control = next(
        control_node
        for paragraph in section_model.paragraphs()
        for control_node in paragraph.control_nodes()
        if control_node.payload[:4][::-1].decode("latin1", errors="replace") == "gso "
    )
    component = next(node for node in line_control.iter_descendants() if node.tag_id == TAG_SHAPE_COMPONENT)
    specific = next(node for node in line_control.iter_descendants() if node.tag_id == TAG_SHAPE_COMPONENT_LINE)
    component.payload = b"loc$loc$" + component.payload[8:]
    specific.payload = specific.payload + (b"\x00" * 13)
    document.binary_document().replace_section_model(0, section_model)

    output_path = tmp_path / "hwp_hancom_connectline_signature.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)

    assert reopened.lines() == []
    assert len(reopened.connect_lines()) == 1
    connect_line = reopened.connect_lines()[0]
    assert isinstance(connect_line, HwpConnectLineShapeObject)
    assert connect_line.kind == "connectLine"
    assert connect_line.shape_comment == "HANCOM-CONNECTLINE"
    assert connect_line.specific_fields()["variant"] == "hancom_connectline"
    assert connect_line.specific_fields()["raw_payload"]


def test_hwp_strict_lint_reports_binary_and_control_subtree_issues(tmp_path: Path) -> None:
    source = HwpDocument.blank()
    image_bytes = source.binary_document().read_stream("BinData/BIN0001.bmp", decompress=False)
    source.append_picture(image_bytes, extension="bmp", width=2400, height=1800)
    source.append_table(rows=1, cols=1, cell_texts=[["A1"]])
    source.append_field(field_type="DOCPROPERTY", display_text="TITLE", name="Title")

    source_path = tmp_path / "hwp_strict_lint_source.hwp"
    source.save(source_path)

    document = HwpDocument.open(source_path)
    picture_control = None
    picture_section_index = None
    picture_section_model = None
    table_control = None
    table_section_index = None
    table_section_model = None
    field_control = None
    field_section_index = None
    field_section_model = None

    for section_index in range(len(document.sections())):
        section_model = document.section_model(section_index)
        for paragraph in section_model.paragraphs():
            for control_node in paragraph.control_nodes():
                control_id = control_node.payload[:4][::-1].decode("latin1", errors="replace")
                if picture_control is None and control_id == "gso " and any(
                    node.tag_id == TAG_SHAPE_COMPONENT_PICTURE for node in control_node.iter_descendants()
                ):
                    picture_control = control_node
                    picture_section_index = section_index
                    picture_section_model = section_model
                elif table_control is None and control_id == "tbl ":
                    table_control = control_node
                    table_section_index = section_index
                    table_section_model = section_model
                elif field_control is None and control_id.startswith("%") and control_id != "%hlk":
                    field_control = control_node
                    field_section_index = section_index
                    field_section_model = section_model

    assert picture_control is not None
    assert picture_section_index is not None
    assert picture_section_model is not None
    assert table_control is not None
    assert table_section_index is not None
    assert table_section_model is not None
    assert field_control is not None
    assert field_section_index is not None
    assert field_section_model is not None

    picture_record = next(node for node in picture_control.iter_descendants() if node.tag_id == TAG_SHAPE_COMPONENT_PICTURE)
    storage_id = int.from_bytes(picture_record.payload[71:73], "little", signed=False)
    document.binary_document().remove_stream(f"BinData/BIN{storage_id:04d}.bmp")
    table_control.children = [child for child in table_control.children if child.tag_id != TAG_TABLE]
    field_control.children = [child for child in field_control.children if child.tag_id != TAG_CTRL_DATA]
    document.binary_document().replace_section_model(table_section_index, table_section_model)
    if picture_section_index != table_section_index:
        document.binary_document().replace_section_model(picture_section_index, picture_section_model)
    if field_section_index not in {table_section_index, picture_section_index}:
        document.binary_document().replace_section_model(field_section_index, field_section_model)

    issues = document.strict_lint_errors()
    issue_codes = {issue.code for issue in issues}
    issue_messages = "\n".join(issue.message for issue in issues)
    report = document.strict_lint_report()
    formatted = document.format_strict_lint_errors()

    assert "binary_closure" in issue_codes
    assert "control_subtree" in issue_codes
    assert "DocInfo BinData id" in issue_messages
    assert "Table control does not contain a TAG_TABLE record." in issue_messages
    assert "Field control is missing TAG_CTRL_DATA parameters." in issue_messages
    assert any(item["code"] == "binary_closure" and "hint" in item for item in report)
    assert "Hint: Keep DocInfo BinData records" in formatted
    assert "control[tbl ]" in formatted

    with pytest.raises(HwpxValidationError):
        document.strict_validate()


def test_hwp_strict_lint_reports_docinfo_count_and_paragraph_references() -> None:
    document = HwpDocument.blank()
    docinfo_model = document.docinfo_model()
    id_mappings = docinfo_model.id_mappings_record()
    id_mappings.set_count(9, 0)
    document.binary_document().replace_docinfo_model(docinfo_model)

    section_model = document.section_model(0)
    first_paragraph = section_model.paragraphs()[0]
    first_paragraph.header.para_shape_id = len(document.docinfo_model().para_shape_records()) + 1
    first_paragraph.header.style_id = len(document.docinfo_model().style_records()) + 1
    document.binary_document().replace_section_model(0, section_model)

    issues = document.strict_lint_errors()
    issue_codes = {issue.code for issue in issues}
    issue_messages = "\n".join(issue.message for issue in issues)

    assert "docinfo_mapping" in issue_codes
    assert "docinfo_reference" in issue_codes
    assert "id_mappings.char_shapes=0 is smaller than actual char_shapes record count" in issue_messages
    assert "Paragraph references missing para shape id" in issue_messages
    assert "Paragraph references missing style id" in issue_messages


def test_hwp_picture_ole_curve_container_and_table_richer_setters_roundtrip(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    image_bytes = document.binary_document().read_stream("BinData/BIN0001.bmp", decompress=False)
    section = document.section(0)
    section.append_table(
        rows=2,
        cols=2,
        cell_texts=[["A1", "A2"], ["B1", "B2"]],
        row_heights=[200, 300],
        col_widths=[1000, 2000],
        table_border_fill_id=6,
    )
    section.append_picture(image_bytes, extension="bmp", width=2400, height=1800)
    section.append_shape(kind="curve", width=1800, height=900, shape_comment="CURVE-COMMENT")
    section.append_shape(kind="container", width=2200, height=1100, shape_comment="CONTAINER-COMMENT")
    section.append_ole("editable.ole", b"OLE-DATA", width=3900, height=2100)

    table = section.tables()[0]
    picture = section.pictures()[0]
    curve = section.curves()[0]
    container = section.containers()[0]
    ole = section.oles()[0]

    table.set_column_widths([1111, 2222])
    table.set_table_margins(left=600, right=610, top=150, bottom=160)
    table.set_cell_spacing(33)
    table.set_cell_border_fill_id(0, 1, 9)
    table.set_cell_margins(1, 0, left=11, right=22, top=33, bottom=44)

    picture.set_shape_comment("PICTURE-COMMENT")
    picture.set_size(width=2800, height=1900)
    picture.set_layout(
        text_wrap="SQUARE",
        vert_rel_to="PAPER",
        horz_rel_to="PAPER",
        vert_align="TOP",
        horz_align="LEFT",
        vert_offset=101,
        horz_offset=202,
    )
    picture.set_out_margins(left=13, right=24, top=35, bottom=46)
    picture.set_rotation(angle=15, center_x=111, center_y=222, rotate_image=False)
    picture.set_image_adjustment(bright=5, contrast=6, effect="GRAY_SCALE", alpha=7)
    picture.set_crop(left=1, right=2399, top=2, bottom=1798)
    picture.set_line_style(color="#223344", width=66)

    curve.set_raw_payload(b"\x01\x02\x03\x04")
    container.set_raw_payload(b"\x05\x06")

    ole.set_shape_comment("OLE-COMMENT")
    ole.set_size(width=4100, height=2300)
    ole.set_layout(
        text_wrap="SQUARE",
        vert_rel_to="PAPER",
        horz_rel_to="PAPER",
        vert_align="TOP",
        horz_align="LEFT",
        vert_offset=77,
        horz_offset=88,
    )
    ole.set_out_margins(left=9, right=10, top=11, bottom=12)
    ole.set_rotation(angle=45, center_x=777, center_y=888)
    ole.set_object_metadata(object_type="LINK", draw_aspect="ICON", has_moniker=True, eq_baseline=12)
    ole.set_extent(x=43000, y=14000)
    ole.set_line_style(color="#445566", width=77)

    output_path = tmp_path / "hwp_richer_picture_ole_curve_container_table.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)

    reopened_table = reopened.tables()[0]
    reopened_picture = reopened.pictures()[0]
    reopened_curve = reopened.curves()[0]
    reopened_container = reopened.containers()[0]
    reopened_ole = reopened.oles()[0]

    assert reopened_table.column_widths == [1111, 2222]
    assert reopened_table.table_margins == (600, 610, 150, 160)
    assert reopened_table.cell_spacing == 33
    assert reopened_table.cell(0, 1).border_fill_id == 9
    assert reopened_table.cell(1, 0).margins == (11, 22, 33, 44)

    assert reopened_picture.shape_comment == "PICTURE-COMMENT"
    assert reopened_picture.size() == {"width": 2800, "height": 1900}
    assert reopened_picture.layout()["textWrap"] == "SQUARE"
    assert reopened_picture.out_margins() == {"left": 13, "right": 24, "top": 35, "bottom": 46}
    assert reopened_picture.rotation()["angle"] == "15"
    assert reopened_picture.rotation()["centerX"] == "111"
    assert reopened_picture.rotation()["centerY"] == "222"
    assert reopened_picture.image_adjustment()["effect"] == "GRAY_SCALE"
    assert reopened_picture.crop() == {"left": 1, "right": 2399, "top": 2, "bottom": 1798}
    assert reopened_picture.line_color == "#223344"
    assert reopened_picture.line_width == 66

    assert isinstance(reopened_curve, HwpCurveShapeObject)
    assert reopened_curve.raw_payload == b"\x01\x02\x03\x04"
    assert isinstance(reopened_container, HwpContainerShapeObject)
    assert reopened_container.raw_payload == b"\x05\x06"

    assert reopened_ole.shape_comment == "OLE-COMMENT"
    assert reopened_ole.size() == {"width": 4100, "height": 2300}
    assert reopened_ole.object_type == "LINK"
    assert reopened_ole.draw_aspect == "ICON"
    assert reopened_ole.has_moniker is True
    assert reopened_ole.eq_baseline == 12
    assert reopened_ole.layout()["textWrap"] == "SQUARE"
    assert reopened_ole.out_margins() == {"left": 9, "right": 10, "top": 11, "bottom": 12}
    assert reopened_ole.rotation()["angle"] == "45"
    assert reopened_ole.extent() == {"x": 43000, "y": 14000}
    assert reopened_ole.line_color == "#445566"
    assert reopened_ole.line_width == 77


def test_hwp_strict_lint_reports_form_memo_and_chart_native_semantic_breakage() -> None:
    document = HwpDocument.blank()
    document.append_form("Accept", form_type="INPUT", name="accept", value="Y")
    document.append_memo("Review needed")
    document.append_chart(
        "Sales",
        chart_type="BAR",
        categories=["Q1"],
        series=[{"name": "Revenue", "values": [10]}],
    )
    document.apply_section_settings(section_index=0, memo_shape_id="1")

    section_model = document.section_model(0)
    controls_by_id: dict[str, RecordNode] = {}
    for paragraph in section_model.paragraphs():
        for control_node in paragraph.control_nodes():
            control_id = hwp_document_module._control_id(control_node)
            controls_by_id.setdefault(control_id, control_node)

    form_control = controls_by_id["mrof"]
    memo_control = controls_by_id["omem"]
    chart_control = controls_by_id["gso "]

    form_control.children = [child for child in form_control.children if child.tag_id != TAG_FORM_OBJECT]

    memo_node = next(child for child in memo_control.children if child.tag_id == TAG_MEMO_LIST)
    memo_node.payload = b"\x61"

    shape_component = next(child for child in chart_control.children if child.tag_id == TAG_SHAPE_COMPONENT)
    shape_component.children = [child for child in shape_component.children if child.tag_id != TAG_CHART_DATA]

    document.binary_document().replace_section_model(0, section_model)

    issues = document.strict_lint_errors()
    issue_codes = {issue.code for issue in issues}
    issue_messages = "\n".join(issue.message for issue in issues)
    report = document.strict_lint_report()

    assert "control_subtree" in issue_codes
    assert "payload_sanity" in issue_codes
    assert "docinfo_mapping" in issue_codes
    assert "Form control does not contain a TAG_FORM_OBJECT record." in issue_messages
    assert "Memo payload is not UTF-16 aligned." in issue_messages
    assert "Chart control metadata is present but TAG_CHART_DATA is missing." in issue_messages
    assert "Section memo_shape_id 1 points to missing TAG_MEMO_SHAPE record." in issue_messages
    assert all("hint" in item for item in report)

    with pytest.raises(HwpxValidationError):
        document.strict_validate()


def test_hwp_strict_lint_reports_invalid_form_memo_and_chart_metadata_schema() -> None:
    document = HwpDocument.blank()
    document.append_form("Accept", form_type="INPUT", name="accept", value="Y", checked=True, items=["Y"])
    document.append_memo("Review needed", order=1, visible=True)
    document.append_chart(
        "Sales",
        chart_type="BAR",
        categories=["Q1"],
        series=[{"name": "Revenue", "values": [10]}],
        data_ref="dataset-1",
        legend_visible=True,
    )

    section_model = document.section_model(0)
    controls_by_id: dict[str, RecordNode] = {}
    for paragraph in section_model.paragraphs():
        for control_node in paragraph.control_nodes():
            control_id = hwp_document_module._control_id(control_node)
            controls_by_id.setdefault(control_id, control_node)

    form_control = controls_by_id["mrof"]
    memo_control = controls_by_id["omem"]
    chart_control = controls_by_id["gso "]

    form_metadata_node = next(child for child in form_control.children if child.tag_id == TAG_CTRL_DATA)
    form_metadata_node.payload = "JAKAL_CTRL_META;form.checked=maybe;form.items=[oops".encode("utf-16-le")

    memo_metadata_node = next(child for child in memo_control.children if child.tag_id == TAG_CTRL_DATA)
    memo_metadata_node.payload = "JAKAL_CTRL_META;memo.order=NaN;memo.visible=maybe".encode("utf-16-le")

    chart_metadata_node = next(child for child in chart_control.children if child.tag_id == TAG_CTRL_DATA)
    chart_metadata_node.payload = (
        "JAKAL_SHAPE_META;chart.legendVisible=maybe;chart.categories={oops};chart.series=[1,2]"
    ).encode("utf-16-le")

    document.binary_document().replace_section_model(0, section_model)

    issues = document.strict_lint_errors()
    issue_codes = {issue.code for issue in issues}
    issue_messages = "\n".join(issue.message for issue in issues)
    formatted = document.format_strict_lint_errors()

    assert "metadata_schema" in issue_codes
    assert "form.checked metadata must be a bool-like literal." in issue_messages
    assert "form.items metadata must be valid JSON" in issue_messages
    assert "memo.order metadata must be an integer." in issue_messages
    assert "memo.visible metadata must be a bool-like literal." in issue_messages
    assert "chart.legendVisible metadata must be a bool-like literal." in issue_messages
    assert "chart.categories metadata must be valid JSON" in issue_messages
    assert "chart.series metadata must decode to a JSON list of objects." in issue_messages
    assert "Hint: Use the semantic setter API" in formatted

    with pytest.raises(HwpxValidationError):
        document.strict_validate()


def test_hwp_document_append_methods_create_controls_without_bridge(tmp_path: Path) -> None:
    document = HwpDocument.blank()

    document.append_field(field_type="DOCPROPERTY", display_text="FIELD-TEXT", name="Subject")
    document.append_equation("x+y=z", width=3456, height=2100, font="Batang")
    document.append_shape(kind="rect", text="SHAPE-TEXT", width=3600, height=1800)
    document.append_ole("embedded.ole", b"OLE-DATA")

    assert any(field.field_type == "DOCPROPERTY" for field in document.fields())
    assert "FIELD-TEXT" in document.get_document_text()
    assert any(equation.script == "x+y=z" for equation in document.equations())
    assert any(shape.kind == "rect" and shape.size()["width"] == 3600 for shape in document.shapes())
    assert len(document.oles()) == 1
    assert any(path.endswith(".ole") for path in document.bindata_stream_paths())

    output_path = tmp_path / "hwp_direct_append_controls.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)
    field = next(field for field in reopened.fields() if field.field_type == "DOCPROPERTY")
    assert field.native_field_type == "%doc"
    assert field.name == "Subject"
    assert field.display_text == "FIELD-TEXT"
    assert field.parameters["FieldName"] == "Subject"
    assert "FIELD-TEXT" in reopened.get_document_text()
    assert any(equation.script == "x+y=z" for equation in reopened.equations())
    assert any(shape.kind == "rect" and shape.size()["width"] == 3600 for shape in reopened.shapes())
    assert len(reopened.oles()) == 1


def test_hwp_document_native_field_wrapper_exposes_name_and_flags(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    document.append_field(
        field_type="MAILMERGE",
        display_text="FIELD-NATIVE",
        name="FIELD_NATIVE_NAME",
        parameters={"FieldName": "FIELD_NATIVE_NAME", "Format": "upper"},
        editable=True,
        dirty=True,
    )

    output_path = tmp_path / "hwp_native_field_wrapper.hwp"
    document.save(output_path)

    reopened = HwpDocument.open(output_path)
    field = reopened.fields()[0]

    assert field.field_type == "MAILMERGE"
    assert field.native_field_type == "%mai"
    assert field.name == "FIELD_NATIVE_NAME"
    assert field.display_text == "FIELD-NATIVE"
    assert field.parameters["FieldName"] == "FIELD_NATIVE_NAME"
    assert field.parameters["Format"] == "upper"
    assert field.editable is True
    assert field.dirty is True


def test_hwp_document_native_field_helpers_group_semantic_subtypes(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    document.append_mail_merge_field("customer_name", display_text="CUSTOMER")
    document.append_calculation_field("40+2", display_text="42")
    document.append_cross_reference("bookmark_a", display_text="Bookmark A")
    document.append_field(field_type="DOCPROPERTY", display_text="TITLE", name="Title")
    document.append_field(field_type="DATE", display_text="2026-04-13")

    output_path = tmp_path / "hwp_native_field_helpers.hwp"
    document.save(output_path)

    reopened = HwpDocument.open(output_path)

    assert len(reopened.mail_merge_fields()) == 1
    assert reopened.mail_merge_fields()[0].merge_field_name == "customer_name"
    assert reopened.mail_merge_fields()[0].display_text == "CUSTOMER"
    assert len(reopened.calculation_fields()) == 1
    assert reopened.calculation_fields()[0].expression == "40+2"
    assert len(reopened.cross_references()) == 1
    assert reopened.cross_references()[0].bookmark_name == "bookmark_a"
    assert len(reopened.doc_property_fields()) == 1
    assert reopened.doc_property_fields()[0].document_property_name == "Title"
    assert len(reopened.date_fields()) == 1
    assert reopened.date_fields()[0].native_field_type == "%dat"


def test_hwp_document_native_control_append_writes_richer_equation_shape_and_ole_payloads(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    document.append_equation("x+y=z", width=3456, height=2100, font="Batang", shape_comment="NATIVE-EQ")
    document.append_shape(
        kind="rect",
        text="NATIVE-SHAPE",
        width=3600,
        height=1800,
        fill_color="#ABCDEF",
        line_color="#123456",
        shape_comment="NATIVE-SHAPE-COMMENT",
    )
    document.append_ole(
        "native.ole",
        b"NATIVE-OLE-DATA",
        width=3900,
        height=2200,
        shape_comment="NATIVE-OLE-COMMENT",
        object_type="LINK",
        draw_aspect="ICON",
        has_moniker=True,
        eq_baseline=12,
        line_color="#445566",
        line_width=77,
    )
    equation = document.equations()[0]
    equation.set_layout(
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
    equation.set_out_margins(left=5, right=6, top=7, bottom=8)
    equation.set_rotation(angle=35, center_x=555, center_y=666)
    shape = document.shapes()[0]
    shape.set_rotation(angle=25, center_x=333, center_y=444)

    output_path = tmp_path / "hwp_native_richer_controls.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)

    equation = reopened.equations()[0]
    shape = reopened.shapes()[0]
    ole = reopened.oles()[0]

    assert equation.script == "x+y=z"
    assert equation.size() == {"width": 3456, "height": 2100}
    assert equation.font == "Batang"
    assert equation.shape_comment == "NATIVE-EQ"
    assert equation.layout()["textWrap"] == "SQUARE"
    assert equation.layout()["textFlow"] == "RIGHT_ONLY"
    assert equation.layout()["treatAsChar"] == "0"
    assert equation.layout()["vertAlign"] == "CENTER"
    assert equation.layout()["horzAlign"] == "RIGHT"
    assert equation.layout()["vertOffset"] == "12"
    assert equation.out_margins() == {"left": 5, "right": 6, "top": 7, "bottom": 8}
    assert equation.rotation()["angle"] == "35"
    assert equation.rotation()["centerX"] == "555"
    assert equation.rotation()["centerY"] == "666"
    assert shape.text == "NATIVE-SHAPE"
    assert shape.size() == {"width": 3600, "height": 1800}
    assert shape.shape_comment == "NATIVE-SHAPE-COMMENT"
    assert shape.fill_color == "#ABCDEF"
    assert shape.line_color == "#123456"
    assert shape.rotation()["angle"] == "25"
    assert shape.rotation()["centerX"] == "333"
    assert shape.rotation()["centerY"] == "444"
    assert TAG_LIST_HEADER in shape.descendant_tag_ids()
    assert TAG_PARA_HEADER in shape.descendant_tag_ids()
    assert ole.text == "native.ole"
    assert ole.size() == {"width": 3900, "height": 2200}
    assert ole.shape_comment == "NATIVE-OLE-COMMENT"
    assert ole.object_type == "LINK"
    assert ole.draw_aspect == "ICON"
    assert ole.has_moniker is True
    assert ole.eq_baseline == 12
    assert ole.line_color == "#445566"
    assert ole.line_width == 77
    assert ole.storage_id > 0
    assert TAG_LIST_HEADER in ole.descendant_tag_ids()


def test_hwp_document_native_shape_append_writes_richer_ellipse_arc_polygon_and_textart_payloads(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    document.append_shape(
        kind="ellipse",
        text="NATIVE-ELLIPSE",
        width=3800,
        height=2000,
        fill_color="#A1B2C3",
        line_color="#102030",
        shape_comment="ELLIPSE-COMMENT",
    )
    document.append_shape(
        kind="arc",
        text="NATIVE-ARC",
        width=4000,
        height=2200,
        fill_color="#C3B2A1",
        line_color="#203040",
        shape_comment="ARC-COMMENT",
    )
    document.append_shape(
        kind="polygon",
        text="NATIVE-POLYGON",
        width=4200,
        height=2400,
        fill_color="#DDEEFF",
        line_color="#304050",
        shape_comment="POLYGON-COMMENT",
    )
    document.append_shape(
        kind="textart",
        text="NATIVE-TEXTART",
        width=4400,
        height=2600,
        fill_color="#FFEEDD",
        line_color="#405060",
        shape_comment="TEXTART-COMMENT",
    )

    output_path = tmp_path / "hwp_native_richer_shape_kinds.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)

    shapes = reopened.shapes()
    ellipse = next(shape for shape in shapes if shape.kind == "ellipse")
    arc = next(shape for shape in shapes if shape.kind == "arc")
    polygon = next(shape for shape in shapes if shape.kind == "polygon")
    textart = next(shape for shape in shapes if shape.kind == "textart")

    assert ellipse.text == "NATIVE-ELLIPSE"
    assert ellipse.size() == {"width": 3800, "height": 2000}
    assert ellipse.fill_color == "#A1B2C3"
    assert ellipse.line_color == "#102030"
    assert ellipse.shape_comment == "ELLIPSE-COMMENT"
    assert sum(1 for child in ellipse._paragraph.header.children if child.tag_id == TAG_CTRL_HEADER) >= 3

    assert arc.text == "NATIVE-ARC"
    assert arc.size() == {"width": 4000, "height": 2200}
    assert arc.fill_color == "#C3B2A1"
    assert arc.line_color == "#203040"
    assert arc.shape_comment == "ARC-COMMENT"

    assert polygon.text == "NATIVE-POLYGON"
    assert polygon.size() == {"width": 4200, "height": 2400}
    assert polygon.fill_color == "#DDEEFF"
    assert polygon.line_color == "#304050"
    assert polygon.shape_comment == "POLYGON-COMMENT"

    assert textart.text == "NATIVE-TEXTART"
    assert textart.size() == {"width": 4400, "height": 2600}
    assert textart.fill_color == "#FFEEDD"
    assert textart.line_color == "#405060"
    assert textart.shape_comment == "TEXTART-COMMENT"
    assert sum(1 for child in textart._paragraph.header.children if child.tag_id == TAG_CTRL_HEADER) >= 3
    assert TAG_LIST_HEADER not in textart.descendant_tag_ids()
    assert TAG_PARA_HEADER not in textart.descendant_tag_ids()


def test_hwp_document_shape_component_payloads_follow_native_template_storage_rules(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    document.append_shape(kind="rect", width=3601, height=1802, text="R")
    document.append_shape(kind="ellipse", width=3811, height=2013, text="E")
    document.append_shape(kind="arc", width=4022, height=2224, text="A")
    document.append_shape(kind="polygon", width=4233, height=2435, text="P")

    output_path = tmp_path / "hwp_template_shape_component_sizes.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)

    expected_component_sizes = {
        "rect": (3601, 1802, 3601, 1802),
        "ellipse": (0, 0, 0, 0),
        "arc": (0, 0, 0, 0),
        "polygon": (0, 0, 0, 0),
    }
    expected_shape_sizes = {
        "rect": {"width": 3601, "height": 1802},
        "ellipse": {"width": 3811, "height": 2013},
        "arc": {"width": 4022, "height": 2224},
        "polygon": {"width": 4233, "height": 2435},
    }

    for kind, expected in expected_component_sizes.items():
        shape = next(item for item in reopened.shapes() if item.kind == kind)
        component = shape._shape_component_record()
        assert component is not None
        assert int.from_bytes(component.payload[20:24], "little", signed=False) == expected[0]
        assert int.from_bytes(component.payload[24:28], "little", signed=False) == expected[1]
        assert int.from_bytes(component.payload[28:32], "little", signed=False) == expected[2]
        assert int.from_bytes(component.payload[32:36], "little", signed=False) == expected[3]
        assert shape.size() == expected_shape_sizes[kind]


def test_hwp_document_template_shapes_match_clean_hancom_donor_payloads(tmp_path: Path) -> None:
    def _build_colorref(value: str) -> bytes:
        normalized = value.lstrip("#")
        return bytes((int(normalized[4:6], 16), int(normalized[2:4], 16), int(normalized[0:2], 16), 0))

    def _expected_component(
        payload_hex: str,
        *,
        width: int,
        height: int,
        line_color: str,
        line_width: int = 33,
        fill_color: str | None = None,
        store_size_in_component: bool = True,
    ) -> str:
        payload = bytearray.fromhex(payload_hex)
        if store_size_in_component:
            payload[20:24] = int(width).to_bytes(4, "little", signed=False)
            payload[24:28] = int(height).to_bytes(4, "little", signed=False)
            payload[28:32] = int(width).to_bytes(4, "little", signed=False)
            payload[32:36] = int(height).to_bytes(4, "little", signed=False)
        payload[196:200] = _build_colorref(line_color)
        payload[200:202] = int(line_width).to_bytes(2, "little", signed=False)
        if fill_color is not None and len(payload) >= 217:
            payload[213:217] = _build_colorref(fill_color)
        return payload.hex()

    document = HwpDocument.blank()
    document.append_shape(
        kind="rect",
        width=3800,
        height=2000,
        fill_color="#A1B2C3",
        line_color="#302010",
        shape_comment="사각형입니다.",
    )
    document.append_shape(
        kind="ellipse",
        width=3800,
        height=2000,
        fill_color="#A1B2C3",
        line_color="#302010",
        shape_comment="타원입니다.",
    )
    document.append_shape(
        kind="arc",
        width=3800,
        height=2000,
        fill_color="#A1B2C3",
        line_color="#302010",
        shape_comment="호입니다.",
    )
    document.append_shape(
        kind="polygon",
        width=3800,
        height=2000,
        fill_color="#A1B2C3",
        line_color="#302010",
        shape_comment="다각형입니다.",
    )
    output_path = tmp_path / "hwp_template_shape_donor_payloads.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)

    expected_payloads = {
        "rect": {
            "ctrl": "206f736711222a000000000000000000d80e0000d007000000000000000000000000000076a7417f000000000700acc001ac15d685c7c8b2e4b22e00",
            "component": _expected_component(
                "63657224636572240000000000000000000001000000000000000000000000000000000000000800000000000000000000000100000000000000f03f000000000000000000000000000000000000000000000000000000000000f03f0000000000000000000000000000f03f000000000000000000000000000000000000000000000000000000000000f03f0000000000000000000000000000f03f000000000000000000000000000000000000000000000000000000000000f03f000000000000000030201000780000000000000000000000000000000000000000b2b2b200000000000000000077a7413f0000",
                width=3800,
                height=2000,
                line_color="#302010",
            ),
            "specific": "102030002100400000d10001000000c3b2a10010203000000000000000000000000000000000000000000000000000000000000000000000000000000000000000",
        },
        "ellipse": {
            "ctrl": "206f736711222a000000000000000000d80e0000d00700000000000000000000000000008da7417f000000000600c0d0d0c685c7c8b2e4b22e00",
            "component": _expected_component(
                "6c6c65246c6c65240000000000000000000001000000000000000000000000000000000000000800000000000000000000000100000000000000f03f000000000000000000000000000000000000000000000000000000000000f03f0000000000000000000000000000f03f000000000000000000000000000000000000000000000000000000000000f03f0000000000000000000000000000f03f000000000000000000000000000000000000000000000000000000000000f03f000000000000000030201000780000000000000000000000000000000000000000b2b2b20000000000000000004f35443f0000",
                width=3800,
                height=2000,
                line_color="#302010",
                line_width=120,
                store_size_in_component=False,
            ),
            "specific": "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000",
        },
        "arc": {
            "ctrl": "206f736711222a000000000000000000d80e0000d0070000000000000000000000000000a1a7417f00000000050038d685c7c8b2e4b22e00",
            "component": _expected_component(
                "63726124637261240000000000000000000001000000000000000000000000000000000000000800000000000000000000000100000000000000f03f000000000000000000000000000000000000000000000000000000000000f03f0000000000000000000000000000f03f000000000000000000000000000000000000000000000000000000000000f03f0000000000000000000000000000f03f000000000000000000000000000000000000000000000000000000000000f03f000000000000000030201000780000000000000000000000000000000000000000b2b2b20000000000000000006735443f0000",
                width=3800,
                height=2000,
                line_color="#302010",
                line_width=120,
                store_size_in_component=False,
            ),
            "specific": "00000000000000000000000000000000000000000000000000",
        },
        "polygon": {
            "ctrl": "206f736711222a000000000000000000d80e0000d0070000000000000000000000000000b6a7417f000000000700e4b201ac15d685c7c8b2e4b22e00",
            "component": _expected_component(
                "6c6f70246c6f70240000000000000000000001000000000000000000000000000000000000000800000000000000000000000100000000000000f03f000000000000000000000000000000000000000000000000000000000000f03f0000000000000000000000000000f03f000000000000000000000000000000000000000000000000000000000000f03f0000000000000000000000000000f03f000000000000000000000000000000000000000000000000000000000000f03f000000000000000030201000780000000000000000000000000000000000000000b2b2b20000000000000000007c35443f0000",
                width=3800,
                height=2000,
                line_color="#302010",
                line_width=120,
                store_size_in_component=False,
            ),
            "specific": "0000000000000000",
        },
    }

    actual_shapes = {shape.kind: shape for shape in reopened.shapes()}
    for kind, expected in expected_payloads.items():
        shape = actual_shapes[kind]
        component = shape._shape_component_record()
        specific = shape._shape_specific_record()
        assert component is not None
        assert specific is not None
        assert shape.control_node.payload.hex() == expected["ctrl"]
        assert component.payload.hex() == expected["component"]
        assert specific.payload.hex() == expected["specific"]


def test_hwp_document_textart_payload_matches_normalized_hancom_reference(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    document.append_shape(
        kind="textart",
        text="COM-TEXTART",
        width=18000,
        height=6000,
        fill_color="#FFEEDD",
        line_color="#665544",
        shape_comment="글맵시입니다.",
    )

    output_path = tmp_path / "hwp_textart_normalized_reference.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)
    textart = next(shape for shape in reopened.shapes() if shape.kind == "textart")
    component = textart._shape_component_record()
    specific = textart._shape_specific_record()
    assert component is not None
    assert specific is not None

    assert textart.control_node.payload[4:8].hex() == "11220a04"
    assert int.from_bytes(textart.control_node.payload[16:20], "little", signed=False) == 18000
    assert int.from_bytes(textart.control_node.payload[20:24], "little", signed=False) == 6000

    reference_extent = int.from_bytes(component.payload[20:24], "little", signed=False)
    assert reference_extent == 14173
    assert int.from_bytes(component.payload[24:28], "little", signed=False) == 14173
    assert int.from_bytes(component.payload[28:32], "little", signed=False) == 18000
    assert int.from_bytes(component.payload[32:36], "little", signed=False) == 6000
    assert int.from_bytes(component.payload[36:40], "little", signed=False) == 0x00080000
    assert int.from_bytes(component.payload[40:44], "little", signed=False) == 18000 << 15
    assert int.from_bytes(component.payload[44:48], "little", signed=False) == 6000 << 15
    assert struct.unpack("<d", component.payload[100:108])[0] == pytest.approx(18000 / reference_extent)
    assert struct.unpack("<d", component.payload[132:140])[0] == pytest.approx(6000 / reference_extent)
    assert component.payload[196:200].hex() == "44556600"
    assert int.from_bytes(component.payload[200:202], "little", signed=False) == 33
    assert component.payload[213:217].hex() == "ddeeff00"

    text_length = int.from_bytes(specific.payload[32:34], "little", signed=False)
    text_end = 34 + text_length * 2
    assert text_length == len("COM-TEXTART")
    assert specific.payload[34:text_end].decode("utf-16-le", errors="ignore") == "COM-TEXTART"
    assert int.from_bytes(specific.payload[text_end : text_end + 2], "little", signed=False) == len("Malgun Gothic")


def test_hwp_document_default_ellipse_arc_polygon_do_not_emit_ctrl_data_metadata(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    document.append_shape(kind="ellipse", width=3800, height=2000, line_color="#304050", shape_comment="ELLIPSE")
    document.append_shape(kind="arc", width=4000, height=2200, line_color="#304050", shape_comment="ARC")
    document.append_shape(kind="polygon", width=4200, height=2400, line_color="#304050", shape_comment="POLYGON")

    output_path = tmp_path / "hwp_default_complex_shape_metadata.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)

    for kind in ("ellipse", "arc", "polygon"):
        shape = next(item for item in reopened.shapes() if item.kind == kind)
        assert TAG_CTRL_DATA not in shape.descendant_tag_ids()


def test_hwp_document_native_table_picture_hyperlink_append_do_not_require_profile_reload(
    tmp_path: Path,
    monkeypatch,
) -> None:
    document = HwpDocument.blank()
    image_bytes = document.binary_document().read_stream("BinData/BIN0001.bmp", decompress=False)
    document._pure_profile = None

    def _fail_load_bundled():
        raise AssertionError("append path should not reload bundled pure profile")

    monkeypatch.setattr(hwp_document_module.HwpPureProfile, "load_bundled", staticmethod(_fail_load_bundled))

    document.append_table(rows=1, cols=2, cell_texts=[["NATIVE-1", "NATIVE-2"]])
    document.append_picture(image_bytes, extension="bmp")
    document.append_hyperlink("https://example.com/native", text="NATIVE-LINK")

    output_path = tmp_path / "hwp_native_append_no_profile_reload.hwp"
    document.save(output_path)
    monkeypatch.undo()
    reopened = HwpDocument.open(output_path)

    assert any(table.row_count == 1 and table.column_count == 2 for table in reopened.tables())
    assert any(picture.extension == "bmp" for picture in reopened.pictures())
    assert any(link.display_text == "NATIVE-LINK" for link in reopened.hyperlinks())


def test_hwp_section_append_helpers_create_controls_in_target_section() -> None:
    document = HwpDocument.blank()
    section = document.section(0)

    section.append_field(field_type="DATE")
    section.append_bookmark("SEC-BOOKMARK")
    section.append_footnote("SEC-FOOTNOTE")
    section.append_endnote("SEC-ENDNOTE")
    section.append_auto_number(number=7, kind="newNum")
    section.append_header("SEC-HEADER", apply_page_type="EVEN")
    section.append_footer("SEC-FOOTER", apply_page_type="ODD")
    section.append_equation("a+b")
    section.append_shape(kind="ellipse", text="SEC-SHAPE")
    section.append_ole("sec.ole", b"section-ole")

    assert any(field.field_type == "DATE" and field.native_field_type == "%dat" for field in section.fields())
    assert any(isinstance(control, HwpBookmarkObject) and control.name == "SEC-BOOKMARK" for control in section.controls())
    assert any(isinstance(control, HwpNoteObject) and control.kind == "footNote" and control.text == "SEC-FOOTNOTE" for control in section.controls())
    assert any(isinstance(control, HwpNoteObject) and control.kind == "endNote" and control.text == "SEC-ENDNOTE" for control in section.controls())
    assert any(isinstance(control, HwpAutoNumberObject) and control.kind == "newNum" for control in section.controls())
    assert any(
        isinstance(control, HwpHeaderFooterObject)
        and control.kind == "header"
        and control.text == "SEC-HEADER"
        and control.apply_page_type == "EVEN"
        for control in section.controls()
    )
    assert any(
        isinstance(control, HwpHeaderFooterObject)
        and control.kind == "footer"
        and control.text == "SEC-FOOTER"
        and control.apply_page_type == "ODD"
        for control in section.controls()
    )
    assert any(equation.script == "a+b" for equation in section.equations())
    assert any(shape.kind == "ellipse" for shape in section.shapes())
    assert len(section.oles()) == 1


def test_hwp_document_can_append_header_footer_and_new_number_and_bridge_back(tmp_path: Path) -> None:
    document = HwpDocument.blank()

    document.append_header("PURE-HEADER", apply_page_type="EVEN")
    document.append_footer("PURE-FOOTER", apply_page_type="ODD")
    document.append_auto_number(number=7, kind="newNum")

    output_path = tmp_path / "hwp_header_footer_number.hwp"
    document.save(output_path)

    reopened = HwpDocument.open(output_path)
    controls = reopened.section(0).controls()
    assert any(
        isinstance(control, HwpHeaderFooterObject)
        and control.kind == "header"
        and control.text == "PURE-HEADER"
        and control.apply_page_type == "EVEN"
        for control in controls
    )
    assert any(
        isinstance(control, HwpHeaderFooterObject)
        and control.kind == "footer"
        and control.text == "PURE-FOOTER"
        and control.apply_page_type == "ODD"
        for control in controls
    )
    assert any(isinstance(control, HwpAutoNumberObject) and control.kind == "newNum" and control.number == "7" for control in controls)

    bridged = reopened.to_hwpx_document()
    assert bridged.headers()[0].text == "PURE-HEADER"
    assert bridged.headers()[0].apply_page_type == "EVEN"
    assert bridged.footers()[0].text == "PURE-FOOTER"
    assert bridged.footers()[0].apply_page_type == "ODD"
    assert any(item.kind == "newNum" and item.number == "7" for item in bridged.auto_numbers())


def test_hwp_document_can_append_bookmark_and_notes_and_bridge_back(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    document.apply_section_note_settings(
        section_index=0,
        footnote_pr={"numbering": {"type": "CONTINUOUS", "newNum": "3"}},
        endnote_pr={"numbering": {"type": "CONTINUOUS", "newNum": "8"}},
    )

    document.append_paragraph("BASE")
    document.append_bookmark("py_anchor")
    document.append_footnote("Pure footnote", number=1)
    document.append_endnote("Pure endnote", number=2)

    output_path = tmp_path / "hwp_bookmark_notes.hwp"
    document.save(output_path)

    reopened = HwpDocument.open(output_path)
    assert any(isinstance(control, HwpBookmarkObject) and control.name == "py_anchor" for control in reopened.controls())
    assert any(
        isinstance(control, HwpNoteObject)
        and control.kind == "footNote"
        and control.text == "Pure footnote"
        and control.number == "3"
        for control in reopened.controls()
    )
    assert any(
        isinstance(control, HwpNoteObject)
        and control.kind == "endNote"
        and control.text == "Pure endnote"
        and control.number == "8"
        for control in reopened.controls()
    )

    bridged = reopened.to_hwpx_document()
    assert any(item.name == "py_anchor" for item in bridged.bookmarks())
    assert any(item.kind == "footNote" and item.text == "Pure footnote" and item.number == "3" for item in bridged.notes())
    assert any(item.kind == "endNote" and item.text == "Pure endnote" and item.number == "8" for item in bridged.notes())


def test_hwp_document_pure_bridge_extracts_native_header_and_auto_number_controls(sample_hwp_path: Path) -> None:
    document = HwpDocument.open(sample_hwp_path)
    bridged = document.to_hwpx_document()

    assert bridged.headers()
    assert any((header.text or "").strip() for header in bridged.headers())
    assert any(item.kind == "autoNum" for item in bridged.auto_numbers())


def test_hwp_document_append_methods_support_paragraph_index() -> None:
    document = HwpDocument.blank()
    section = document.section(0)
    section.append_paragraph("FIRST")
    section.append_paragraph("SECOND")
    section.append_paragraph("THIRD")
    initial_texts = [paragraph.text for paragraph in section.paragraphs()]
    first_index = initial_texts.index("FIRST")

    document.append_field(field_type="DOCPROPERTY", display_text="FIELD-MID", paragraph_index=first_index + 1)
    document.append_equation("EQ-MID", paragraph_index=first_index + 2)
    document.append_shape(kind="rect", text="SHAPE-MID", paragraph_index=first_index + 3)
    document.append_ole("OLE-MID", b"ole-mid", paragraph_index=first_index + 4)

    texts = [paragraph.text for paragraph in section.paragraphs()]
    updated_first_index = texts.index("FIRST")
    equation_index = next(equation.paragraph_index for equation in document.equations() if equation.script == "EQ-MID")
    rect_index = next(shape.paragraph_index for shape in document.shapes() if shape.kind == "rect" and shape.text == "SHAPE-MID")
    ole_index = next(ole.paragraph_index for ole in document.oles() if ole.text == "OLE-MID")

    assert "FIELD-MID" in texts[updated_first_index + 1]
    assert texts[equation_index] == "\r"
    assert "OLE-MID" in texts[ole_index]
    assert texts[rect_index] == "\r"
    assert texts[-2:] == ["SECOND", "THIRD"]
    assert updated_first_index + 1 < equation_index < ole_index < rect_index


def test_hwp_table_object_can_append_row_and_roundtrip(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    document.append_table(
        rows=1,
        cols=2,
        cell_texts=[["R0C0", "R0C1"]],
        row_heights=[222],
        col_widths=[1300, 1700],
        table_border_fill_id=6,
    )

    table = document.tables()[-1]
    new_row = table.append_row()

    assert table.row_count == 2
    assert table.row_heights == [222, 222]
    assert len(new_row) == 2
    assert [cell.column for cell in new_row] == [0, 1]
    assert all(cell.row == 1 for cell in new_row)

    output_path = tmp_path / "hwp_table_append_row.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)
    assert reopened.tables()[-1].row_count == 2
    assert reopened.tables()[-1].row_heights == [222, 222]


def test_hwp_table_object_can_merge_cells_and_roundtrip(tmp_path: Path) -> None:
    document = HwpDocument.blank()
    document.append_table(
        rows=2,
        cols=2,
        cell_texts=[["MERGE", ""], ["", ""]],
        row_heights=[180, 220],
        col_widths=[1100, 1900],
        table_border_fill_id=4,
    )

    table = document.tables()[-1]
    merged = table.merge_cells(0, 0, 1, 1)

    assert merged.row_span == 2
    assert merged.col_span == 2
    assert merged.width == 3000
    assert merged.height == 400
    assert table.cell_text_matrix()[0][0] == "MERGE"
    assert len(table.cells()) == 1

    output_path = tmp_path / "hwp_table_merge_cells.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)
    reopened_cell = reopened.tables()[-1].cell(0, 0)
    assert reopened_cell.row_span == 2
    assert reopened_cell.col_span == 2
    assert reopened_cell.text == "MERGE"
