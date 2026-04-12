from __future__ import annotations

from pathlib import Path

import pytest

import jakal_hwpx.hwp_document as hwp_document_module
from jakal_hwpx.hwp_binary import TAG_CTRL_HEADER, TAG_LIST_HEADER, TAG_PARA_HEADER

from jakal_hwpx import (
    DocInfoModel,
    HwpAutoNumberObject,
    HwpBookmarkObject,
    HwpDocument,
    HwpDocumentProperties,
    HwpEquationObject,
    HwpFieldObject,
    HwpHeaderFooterObject,
    HwpHyperlinkObject,
    HwpNoteObject,
    HwpOleObject,
    HwpParagraphObject,
    HwpPictureObject,
    HwpPureProfile,
    HwpSection,
    HwpShapeObject,
    HwpTableObject,
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


def test_hwp_document_append_methods_create_controls_without_bridge(tmp_path: Path) -> None:
    document = HwpDocument.blank()

    document.append_field(field_type="DOCPROPERTY", display_text="FIELD-TEXT", name="Subject")
    document.append_equation("x+y=z", width=3456, height=2100, font="Batang")
    document.append_shape(kind="rect", text="SHAPE-TEXT", width=3600, height=1800)
    document.append_ole("embedded.ole", b"OLE-DATA")

    assert any(field.field_type == "%doc" for field in document.fields())
    assert "FIELD-TEXT" in document.get_document_text()
    assert any(equation.script == "x+y=z" for equation in document.equations())
    assert any(shape.kind == "rect" and shape.size()["width"] == 3600 for shape in document.shapes())
    assert len(document.oles()) == 1
    assert any(path.endswith(".ole") for path in document.bindata_stream_paths())

    output_path = tmp_path / "hwp_direct_append_controls.hwp"
    document.save(output_path)
    reopened = HwpDocument.open(output_path)
    assert any(field.field_type == "%doc" for field in reopened.fields())
    assert "FIELD-TEXT" in reopened.get_document_text()
    assert any(equation.script == "x+y=z" for equation in reopened.equations())
    assert any(shape.kind == "rect" and shape.size()["width"] == 3600 for shape in reopened.shapes())
    assert len(reopened.oles()) == 1


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
    assert shape.text == "NATIVE-SHAPE"
    assert shape.size() == {"width": 3600, "height": 1800}
    assert shape.shape_comment == "NATIVE-SHAPE-COMMENT"
    assert shape.fill_color == "#ABCDEF"
    assert shape.line_color == "#123456"
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
    section.append_header("SEC-HEADER")
    section.append_footer("SEC-FOOTER")
    section.append_equation("a+b")
    section.append_shape(kind="ellipse", text="SEC-SHAPE")
    section.append_ole("sec.ole", b"section-ole")

    assert any(field.field_type == "%dat" for field in section.fields())
    assert any(isinstance(control, HwpBookmarkObject) and control.name == "SEC-BOOKMARK" for control in section.controls())
    assert any(isinstance(control, HwpNoteObject) and control.kind == "footNote" and control.text == "SEC-FOOTNOTE" for control in section.controls())
    assert any(isinstance(control, HwpNoteObject) and control.kind == "endNote" and control.text == "SEC-ENDNOTE" for control in section.controls())
    assert any(isinstance(control, HwpAutoNumberObject) and control.kind == "newNum" for control in section.controls())
    assert any(isinstance(control, HwpHeaderFooterObject) and control.kind == "header" and control.text == "SEC-HEADER" for control in section.controls())
    assert any(isinstance(control, HwpHeaderFooterObject) and control.kind == "footer" and control.text == "SEC-FOOTER" for control in section.controls())
    assert any(equation.script == "a+b" for equation in section.equations())
    assert any(shape.kind == "ellipse" for shape in section.shapes())
    assert len(section.oles()) == 1


def test_hwp_document_can_append_header_footer_and_new_number_and_bridge_back(tmp_path: Path) -> None:
    document = HwpDocument.blank()

    document.append_header("PURE-HEADER")
    document.append_footer("PURE-FOOTER")
    document.append_auto_number(number=7, kind="newNum")

    output_path = tmp_path / "hwp_header_footer_number.hwp"
    document.save(output_path)

    reopened = HwpDocument.open(output_path)
    controls = reopened.section(0).controls()
    assert any(isinstance(control, HwpHeaderFooterObject) and control.kind == "header" and control.text == "PURE-HEADER" for control in controls)
    assert any(isinstance(control, HwpHeaderFooterObject) and control.kind == "footer" and control.text == "PURE-FOOTER" for control in controls)
    assert any(isinstance(control, HwpAutoNumberObject) and control.kind == "newNum" and control.number == "7" for control in controls)

    bridged = reopened.to_hwpx_document()
    assert bridged.headers()[0].text == "PURE-HEADER"
    assert bridged.footers()[0].text == "PURE-FOOTER"
    assert any(item.kind == "newNum" and item.number == "7" for item in bridged.auto_numbers())


def test_hwp_document_can_append_bookmark_and_notes_and_bridge_back(tmp_path: Path) -> None:
    document = HwpDocument.blank()

    document.append_paragraph("BASE")
    document.append_bookmark("py_anchor")
    document.append_footnote("Pure footnote", number=1)
    document.append_endnote("Pure endnote", number=2)

    output_path = tmp_path / "hwp_bookmark_notes.hwp"
    document.save(output_path)

    reopened = HwpDocument.open(output_path)
    assert any(isinstance(control, HwpBookmarkObject) and control.name == "py_anchor" for control in reopened.controls())
    assert any(isinstance(control, HwpNoteObject) and control.kind == "footNote" and control.text == "Pure footnote" for control in reopened.controls())
    assert any(isinstance(control, HwpNoteObject) and control.kind == "endNote" and control.text == "Pure endnote" for control in reopened.controls())

    bridged = reopened.to_hwpx_document()
    assert any(item.name == "py_anchor" for item in bridged.bookmarks())
    assert any(item.kind == "footNote" and item.text == "Pure footnote" for item in bridged.notes())
    assert any(item.kind == "endNote" and item.text == "Pure endnote" for item in bridged.notes())


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
