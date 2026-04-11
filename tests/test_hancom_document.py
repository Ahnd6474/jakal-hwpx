from __future__ import annotations

from pathlib import Path

import pytest

from jakal_hwpx import (
    AutoNumber,
    Bookmark,
    Equation,
    Field,
    HancomDocument,
    HwpDocument,
    HwpxDocument,
    Note,
    Ole,
    Shape,
)


REPO_ROOT = Path(__file__).resolve().parents[1]
HWP_SAMPLE_DIR = REPO_ROOT / "examples" / "samples" / "hwp"
NATIVE_HWP_PATH = (
    REPO_ROOT
    / "examples"
    / "output_bridge_stability_lab"
    / "hancom"
    / "from_hwp_save_native_hwp"
    / "native.hwp"
)


@pytest.fixture(scope="session")
def sample_hwp_path() -> Path:
    paths = sorted(HWP_SAMPLE_DIR.glob("*.hwp"))
    assert paths, f"No .hwp samples were found under {HWP_SAMPLE_DIR}"
    return paths[0]


def test_hancom_document_authoring_can_write_hwpx(tmp_path: Path) -> None:
    image_bytes = HwpDocument.blank().binary_document().read_stream("BinData/BIN0001.bmp", decompress=False)
    ole_bytes = b"FAKE-OLE-BINARY"

    document = HancomDocument.blank()
    document.metadata.title = "HANCOM-IR"
    document.append_paragraph_style(style_id="11", alignment_horizontal="CENTER", line_spacing=180)
    document.append_character_style(style_id="12", text_color="#224466", height=1100)
    document.append_style(
        "IRStyle",
        style_id="13",
        english_name="IRStyle",
        para_pr_id="11",
        char_pr_id="12",
    )
    document.sections[0].settings.page_width = 150000
    document.sections[0].settings.page_height = 90000
    document.sections[0].settings.visibility["hideFirstHeader"] = "1"
    document.append_header("IR-HEADER", section_index=0, apply_page_type="BOTH")
    document.append_footer("IR-FOOTER", section_index=0, apply_page_type="BOTH")
    document.append_paragraph("IR-PARAGRAPH")
    document.append_bookmark("IR-BOOKMARK")
    document.append_field(
        field_type="MAILMERGE",
        display_text="IR-FIELD",
        name="IR_FIELD_NAME",
        parameters={"FieldName": "IR_FIELD_NAME"},
    )
    document.append_auto_number(number=7, number_type="PAGE", kind="newNum")
    document.append_table(rows=2, cols=2, cell_texts=[["A1", "A2"], ["B1", "B2"]])
    document.append_hyperlink("https://example.com/ir", display_text="IR-LINK")
    document.append_picture("ir.bmp", image_bytes, extension="bmp", width=3600, height=2400)
    document.append_note("IR-NOTE", kind="footNote", number=3)
    document.append_equation("x+y", width=3200, height=1800)
    document.append_shape(kind="rect", text="IR-SHAPE", width=5000, height=2200, fill_color="#FFEEAA", line_color="#112233")
    document.append_ole("ir.ole", ole_bytes, width=4100, height=2100)

    output_path = tmp_path / "hancom_ir.hwpx"
    document.write_to_hwpx(output_path)

    reopened = HwpxDocument.open(output_path)
    assert reopened.metadata().title == "HANCOM-IR"
    text = reopened.get_document_text()
    for value in ("IR-PARAGRAPH", "A1", "A2", "B1", "B2", "IR-FIELD", "IR-LINK", "IR-NOTE", "IR-SHAPE"):
        assert value in text
    assert reopened.headers()[0].text == "IR-HEADER"
    assert reopened.footers()[0].text == "IR-FOOTER"
    assert reopened.bookmarks()[0].name == "IR-BOOKMARK"
    assert any(field.name == "IR_FIELD_NAME" for field in reopened.fields())
    assert any(number.number == "7" for number in reopened.auto_numbers())
    assert any(style.style_id == "13" and style.name == "IRStyle" for style in reopened.styles())
    assert any(style.style_id == "11" for style in reopened.paragraph_styles())
    assert any(style.style_id == "12" for style in reopened.character_styles())
    assert reopened.tables()
    assert reopened.hyperlinks()
    assert reopened.pictures()
    assert reopened.notes()
    assert reopened.equations()
    assert reopened.shapes()
    assert reopened.oles()
    settings = reopened.section_settings(0)
    assert settings.page_width == 150000
    assert settings.page_height == 90000
    assert settings.visibility()["hideFirstHeader"] == "1"


def test_hancom_document_can_read_hwpx_and_roundtrip_hwpx(sample_hwpx_path: Path, tmp_path: Path) -> None:
    document = HancomDocument.read_hwpx(sample_hwpx_path)
    assert document.source_format == "hwpx"
    assert document.sections

    output_path = tmp_path / "hancom_from_hwpx.hwpx"
    document.write_to_hwpx(output_path)

    reopened = HwpxDocument.open(output_path)
    assert reopened.get_document_text()


def test_hancom_document_can_read_new_control_blocks_from_hwpx(tmp_path: Path) -> None:
    source = HwpxDocument.blank()
    source.set_metadata(title="IR-SOURCE")
    source.append_paragraph_style(style_id="21", alignment_horizontal="RIGHT", line_spacing=160)
    source.append_character_style(style_id="22", text_color="#445566", height=1200)
    source.append_style("SRC-STYLE", style_id="23", para_pr_id="21", char_pr_id="22")
    source.append_header("SRC-HEADER")
    source.append_footer("SRC-FOOTER")
    source.append_paragraph("SRC-P")
    source.append_bookmark("SRC-BOOKMARK")
    source.append_field(field_type="MAILMERGE", display_text="SRC-FIELD", name="SRC_FIELD", parameters={"FieldName": "SRC_FIELD"})
    source.append_auto_number(number=9, number_type="PAGE", kind="newNum")
    source.append_footnote("SRC-NOTE", number=7)
    source.append_equation("a+b", width=2800, height=1700)
    source.append_shape(kind="rect", text="SRC-SHAPE", width=4400, height=2100)
    source.append_ole("src.ole", b"SRC-OLE", width=3900, height=2200)
    source.section_settings(0).set_page_size(width=120000, height=88000)

    source_path = tmp_path / "source_controls.hwpx"
    source.save(source_path)

    document = HancomDocument.read_hwpx(source_path)
    assert document.metadata.title == "IR-SOURCE"
    assert any(block.kind == "header" and block.text == "SRC-HEADER" for block in document.sections[0].header_footer_blocks)
    assert any(block.kind == "footer" and block.text == "SRC-FOOTER" for block in document.sections[0].header_footer_blocks)
    blocks = document.sections[0].blocks
    assert any(getattr(block, "text", None) == "SRC-P" for block in blocks)
    assert any(isinstance(block, Bookmark) and getattr(block, "name", None) == "SRC-BOOKMARK" for block in blocks)
    assert any(isinstance(block, Field) and getattr(block, "name", None) == "SRC_FIELD" for block in blocks)
    assert any(isinstance(block, AutoNumber) and str(getattr(block, "number", "")) == "9" for block in blocks)
    assert any(isinstance(block, Note) and getattr(block, "text", None) == "SRC-NOTE" for block in blocks)
    assert any(isinstance(block, Equation) and getattr(block, "script", None) == "a+b" for block in blocks)
    assert any(isinstance(block, Shape) and getattr(block, "text", None) == "SRC-SHAPE" for block in blocks)
    assert any(isinstance(block, Ole) and getattr(block, "name", None) == "src.ole" for block in blocks)
    assert any(style.style_id == "23" and style.name == "SRC-STYLE" for style in document.style_definitions)
    assert any(style.style_id == "21" for style in document.paragraph_styles)
    assert any(style.style_id == "22" for style in document.character_styles)
    assert document.sections[0].settings.page_width == 120000
    assert document.sections[0].settings.page_height == 88000


def test_hancom_document_can_read_hwp_without_hancom_and_export_hwpx(sample_hwp_path: Path, tmp_path: Path) -> None:
    document = HancomDocument.read_hwp(sample_hwp_path)
    assert document.source_format == "hwp"
    assert document.sections

    output_path = tmp_path / "hancom_from_hwp.hwpx"
    document.write_to_hwpx(output_path)

    reopened = HwpxDocument.open(output_path)
    assert reopened.get_document_text()


def test_hancom_document_pure_hwp_roundtrip_preserves_header_footer_and_new_number(tmp_path: Path) -> None:
    document = HancomDocument.blank()
    document.append_header("PURE-HEADER")
    document.append_footer("PURE-FOOTER")
    document.append_auto_number(number=7, number_type="PAGE", kind="newNum")

    output_path = tmp_path / "hancom_ir_controls.hwp"
    document.write_to_hwp(output_path)

    reopened = HwpDocument.open(output_path).to_hwpx_document()
    assert reopened.headers()[0].text == "PURE-HEADER"
    assert reopened.footers()[0].text == "PURE-FOOTER"
    assert any(item.kind == "newNum" and item.number == "7" for item in reopened.auto_numbers())


def test_hancom_document_pure_hwp_roundtrip_preserves_bookmark_and_notes(tmp_path: Path) -> None:
    document = HancomDocument.blank()
    document.append_paragraph("PURE-BODY")
    document.append_bookmark("pure_anchor")
    document.append_footnote("PURE-FOOTNOTE", number=1)
    document.append_endnote("PURE-ENDNOTE", number=2)

    output_path = tmp_path / "hancom_ir_bookmark_note.hwp"
    document.write_to_hwp(output_path)

    reopened = HwpDocument.open(output_path).to_hwpx_document()
    assert any(item.name == "pure_anchor" for item in reopened.bookmarks())
    assert any(item.kind == "footNote" and item.text == "PURE-FOOTNOTE" for item in reopened.notes())
    assert any(item.kind == "endNote" and item.text == "PURE-ENDNOTE" for item in reopened.notes())


def test_hancom_document_pure_hwp_roundtrip_preserves_multi_section_layout_without_profile_leak(tmp_path: Path) -> None:
    document = HancomDocument.blank()
    document.metadata.title = "PURE-MULTI"
    document.sections[0].settings.page_width = 111111
    document.sections[0].settings.page_height = 222222
    document.sections[0].settings.margins = {
        "left": 1001,
        "right": 1002,
        "top": 1003,
        "bottom": 1004,
        "header": 1005,
        "footer": 1006,
        "gutter": 1007,
    }
    document.append_paragraph("SECTION-0", section_index=0)
    document.append_section()
    document.sections[1].settings.page_width = 333333
    document.sections[1].settings.page_height = 444444
    document.append_paragraph("SECTION-1", section_index=1)
    document.append_section()
    document.append_paragraph("SECTION-2", section_index=2)

    output_path = tmp_path / "hancom_ir_multi_section.hwp"
    document.write_to_hwp(output_path)

    reopened_hwp = HwpDocument.open(output_path)
    assert len(reopened_hwp.sections()) == 3
    assert reopened_hwp.preview_text() == "PURE-MULTI"
    assert "목 차" not in reopened_hwp.get_document_text()

    reopened_ir = HancomDocument.read_hwp(output_path)
    assert len(reopened_ir.sections) == 3
    assert reopened_ir.sections[0].settings.page_width == 111111
    assert reopened_ir.sections[0].settings.page_height == 222222
    assert reopened_ir.sections[0].settings.margins["left"] == 1001
    assert reopened_ir.sections[0].settings.margins["right"] == 1002
    assert reopened_ir.sections[0].settings.margins["top"] == 1003
    assert reopened_ir.sections[0].settings.margins["bottom"] == 1004
    assert reopened_ir.sections[0].settings.margins["header"] == 1005
    assert reopened_ir.sections[0].settings.margins["footer"] == 1006
    assert reopened_ir.sections[0].settings.margins["gutter"] == 1007
    assert reopened_ir.sections[1].settings.page_width == 333333
    assert reopened_ir.sections[1].settings.page_height == 444444
    assert any(getattr(block, "text", None) == "SECTION-0" for block in reopened_ir.sections[0].blocks)
    assert any(getattr(block, "text", None) == "SECTION-1" for block in reopened_ir.sections[1].blocks)
    assert any(getattr(block, "text", None) == "SECTION-2" for block in reopened_ir.sections[2].blocks)

    reopened = reopened_ir.to_hwpx_document()
    assert len(reopened.sections) == 3
    assert "SECTION-0" in reopened.get_document_text()
    assert "SECTION-1" in reopened.get_document_text()
    assert "SECTION-2" in reopened.get_document_text()


def test_hancom_document_pure_hwp_roundtrip_preserves_page_border_fills(tmp_path: Path) -> None:
    source = HwpxDocument.blank()
    sec_pr = source.section_xml(0).find(".//hp:secPr")
    assert sec_pr is not None
    sec_pr.append_xml(
        '<hp:pageBorderFill xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" '
        'type="BOTH" borderFillIDRef="1" textBorder="BORDER" headerInside="1" footerInside="1" fillArea="BORDER">'
        '<hp:offset left="2222" right="3333" top="4444" bottom="5555"/>'
        "</hp:pageBorderFill>"
    )
    sec_pr.append_xml(
        '<hp:pageBorderFill xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" '
        'type="EVEN" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="BORDER">'
        '<hp:offset left="123" right="456" top="789" bottom="987"/>'
        "</hp:pageBorderFill>"
    )
    sec_pr.append_xml(
        '<hp:pageBorderFill xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" '
        'type="ODD" borderFillIDRef="1" textBorder="BORDER" headerInside="0" footerInside="0" fillArea="PAPER">'
        '<hp:offset left="11" right="22" top="33" bottom="44"/>'
        "</hp:pageBorderFill>"
    )

    source_path = tmp_path / "page_border_fill_source.hwpx"
    source.save(source_path)

    document = HancomDocument.read_hwpx(source_path)
    assert document.sections[0].settings.page_border_fills

    output_path = tmp_path / "page_border_fill_roundtrip.hwp"
    document.write_to_hwp(output_path)

    reopened_ir = HancomDocument.read_hwp(output_path)
    reopened = reopened_ir.to_hwpx_document()
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
    assert both.find("./hp:offset").get_attr("bottom") == "5555"
    assert even.get_attr("fillArea") == "BORDER"
    assert even.find("./hp:offset").get_attr("right") == "456"
    assert odd.get_attr("textBorder") == "BORDER"
    assert odd.find("./hp:offset").get_attr("top") == "33"


def test_hancom_document_pure_hwp_roundtrip_preserves_section_definition_settings(tmp_path: Path) -> None:
    source = HwpxDocument.blank()
    settings = source.section_settings(0)
    settings.set_grid(line_grid=42, char_grid=21, wonggoji_format=True)
    settings.set_start_numbers(page_starts_on="ODD", page=7, pic=8, tbl=9, equation=10)
    settings.set_visibility(
        hide_first_header=True,
        hide_first_footer=True,
        hide_first_master_page=True,
        hide_first_page_num=True,
        hide_first_empty_line=True,
    )

    source_path = tmp_path / "section_definition_source.hwpx"
    source.save(source_path)

    document = HancomDocument.read_hwpx(source_path)
    output_path = tmp_path / "section_definition_roundtrip.hwp"
    document.write_to_hwp(output_path)

    reopened_ir = HancomDocument.read_hwp(output_path)
    assert reopened_ir.sections[0].settings.grid == {"lineGrid": 42, "charGrid": 21, "wonggojiFormat": 1}
    assert reopened_ir.sections[0].settings.start_numbers == {
        "pageStartsOn": "ODD",
        "page": "7",
        "pic": "8",
        "tbl": "9",
        "equation": "10",
    }
    assert reopened_ir.sections[0].settings.visibility["hideFirstHeader"] == "1"
    assert reopened_ir.sections[0].settings.visibility["hideFirstFooter"] == "1"
    assert reopened_ir.sections[0].settings.visibility["hideFirstMasterPage"] == "1"
    assert reopened_ir.sections[0].settings.visibility["hideFirstPageNum"] == "1"
    assert reopened_ir.sections[0].settings.visibility["hideFirstEmptyLine"] == "1"

    reopened = reopened_ir.to_hwpx_document()
    reopened_settings = reopened.section_settings(0)
    assert reopened_settings.grid() == {"lineGrid": 42, "charGrid": 21, "wonggojiFormat": 1}
    assert reopened_settings.start_numbers()["pageStartsOn"] == "ODD"
    assert reopened_settings.start_numbers()["page"] == "7"
    assert reopened_settings.visibility()["hideFirstFooter"] == "1"
    assert reopened_settings.visibility()["hideFirstPageNum"] == "1"


def test_hancom_document_pure_hwp_roundtrip_preserves_page_numbers_note_settings_and_visibility_enums(tmp_path: Path) -> None:
    source = HwpxDocument.blank()
    settings = source.section_settings(0)
    settings.set_visibility(
        hide_first_header=True,
        border="HIDE_FIRST",
        fill="SHOW_FIRST",
        show_line_number=True,
    )
    line_number_shape = source.section_xml(0).find(".//hp:lineNumberShape")
    assert line_number_shape is not None
    line_number_shape.set_attr("restartType", 1).set_attr("countBy", 3).set_attr("distance", 150).set_attr("startNumber", 7)
    sec_pr = source.section_xml(0).find(".//hp:secPr")
    assert sec_pr is not None
    sec_pr.append_xml(
        '<hp:footNotePr xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">'
        '<hp:autoNumFormat type="ROMAN_SMALL" userChar="*" prefixChar="[" suffixChar="]" supscript="1"/>'
        '<hp:noteLine length="1234" type="DASH" width="0.25 mm" color="#112233"/>'
        '<hp:noteSpacing betweenNotes="444" belowLine="333" aboveLine="222"/>'
        '<hp:numbering type="ON_PAGE" newNum="9"/>'
        '<hp:placement place="RIGHT_MOST_COLUMN" beneathText="1"/>'
        '</hp:footNotePr>'
    )
    sec_pr.append_xml(
        '<hp:endNotePr xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">'
        '<hp:autoNumFormat type="LATIN_SMALL" userChar="" prefixChar="(" suffixChar=")" supscript="0"/>'
        '<hp:noteLine length="4321" type="DOT" width="0.4 mm" color="#445566"/>'
        '<hp:noteSpacing betweenNotes="777" belowLine="666" aboveLine="555"/>'
        '<hp:numbering type="ON_SECTION" newNum="5"/>'
        '<hp:placement place="END_OF_SECTION" beneathText="0"/>'
        '</hp:endNotePr>'
    )
    source.append_control_xml(
        '<hp:pageNum xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" pos="BOTTOM_CENTER" formatType="DIGIT" sideChar="-"/>'
    )

    source_path = tmp_path / "page_num_note_visibility_source.hwpx"
    source.save(source_path)

    document = HancomDocument.read_hwpx(source_path)
    assert document.sections[0].settings.visibility["border"] == "HIDE_FIRST"
    assert document.sections[0].settings.visibility["fill"] == "SHOW_FIRST"
    assert document.sections[0].settings.visibility["showLineNumber"] == "1"
    assert document.sections[0].settings.line_number_shape == {
        "restartType": "1",
        "countBy": "3",
        "distance": "150",
        "startNumber": "7",
    }
    assert document.sections[0].settings.page_numbers == [
        {"pos": "BOTTOM_CENTER", "formatType": "DIGIT", "sideChar": "-"}
    ]
    assert document.sections[0].settings.footnote_pr["numbering"]["type"] == "ON_PAGE"
    assert document.sections[0].settings.footnote_pr["noteLine"]["width"] == "0.25 mm"
    assert document.sections[0].settings.endnote_pr["placement"]["place"] == "END_OF_SECTION"

    direct_roundtrip = document.to_hwpx_document()
    direct_visibility = direct_roundtrip.section_settings(0).visibility()
    assert direct_visibility["border"] == "HIDE_FIRST"
    assert direct_visibility["fill"] == "SHOW_FIRST"
    assert direct_visibility["showLineNumber"] == "1"
    direct_line_number = direct_roundtrip.section_xml(0).find(".//hp:lineNumberShape")
    assert direct_line_number is not None
    assert direct_line_number.get_attr("countBy") == "3"
    assert direct_roundtrip.section_xml(0).find(".//hp:pageNum").get_attr("pos") == "BOTTOM_CENTER"
    assert direct_roundtrip.section_xml(0).find(".//hp:footNotePr/hp:numbering").get_attr("type") == "ON_PAGE"
    assert direct_roundtrip.section_xml(0).find(".//hp:endNotePr/hp:placement").get_attr("place") == "END_OF_SECTION"

    output_path = tmp_path / "page_num_note_visibility_roundtrip.hwp"
    document.write_to_hwp(output_path)

    reopened_hwp = HwpDocument.open(output_path)
    native_visibility = reopened_hwp.binary_document().section_definition_settings(0)["visibility"]
    native_page_numbers = reopened_hwp.binary_document().section_page_numbers(0)
    native_note_settings = reopened_hwp.binary_document().section_note_settings(0)
    assert native_visibility["border"] == "HIDE_FIRST"
    assert native_visibility["fill"] == "SHOW_FIRST"
    assert native_page_numbers == [{"pos": "BOTTOM_CENTER", "formatType": "DIGIT", "sideChar": "-"}]
    assert native_note_settings["footNotePr"]["numbering"]["type"] == "ON_PAGE"
    assert native_note_settings["footNotePr"]["noteLine"]["length"] == "1234"
    assert native_note_settings["footNotePr"]["noteLine"]["type"] == "DASH"
    assert native_note_settings["footNotePr"]["noteLine"]["width"] == "0.25 mm"
    assert native_note_settings["footNotePr"]["noteLine"]["color"] == "#112233"
    assert native_note_settings["footNotePr"]["placement"]["place"] == "RIGHT_MOST_COLUMN"
    assert native_note_settings["endNotePr"]["placement"]["place"] == "END_OF_SECTION"
    assert native_note_settings["endNotePr"]["noteLine"]["width"] == "0.4 mm"
    assert native_note_settings["endNotePr"]["noteLine"]["color"] == "#445566"

    reopened_ir = HancomDocument.read_hwp(output_path)
    reopened = reopened_ir.to_hwpx_document()
    assert reopened.section_settings(0).visibility()["border"] == "HIDE_FIRST"
    assert reopened.section_settings(0).visibility()["fill"] == "SHOW_FIRST"
    assert reopened.section_xml(0).find(".//hp:pageNum").get_attr("formatType") == "DIGIT"
    assert reopened.section_xml(0).find(".//hp:footNotePr/hp:noteLine").get_attr("width") == "0.25 mm"
    assert reopened.section_xml(0).find(".//hp:endNotePr/hp:noteLine").get_attr("color") == "#445566"


def test_hancom_document_can_extract_native_hwp_header_and_auto_number_controls(tmp_path: Path) -> None:
    if not NATIVE_HWP_PATH.exists():
        pytest.skip("native.hwp sample is not available in this workspace.")

    document = HancomDocument.read_hwp(NATIVE_HWP_PATH)
    assert document.sections
    assert document.sections[0].header_footer_blocks
    assert any(block.kind == "header" and block.text for block in document.sections[0].header_footer_blocks)
    assert any(isinstance(block, AutoNumber) and block.kind == "autoNum" for block in document.sections[0].blocks)

    output_path = tmp_path / "native_hwp_extract.hwpx"
    document.write_to_hwpx(output_path)
    reopened = HwpxDocument.open(output_path)
    assert reopened.headers()
    assert any(item.kind == "autoNum" for item in reopened.auto_numbers())


def test_hancom_document_can_write_hwp_via_converter(
    sample_hwp_path: Path,
    sample_hwpx_path: Path,
    tmp_path: Path,
) -> None:
    conversions: list[str] = []

    def fake_converter(input_path: str | Path, output_path: str | Path, output_format: str) -> Path:
        target = Path(output_path)
        conversions.append(output_format.upper())
        if output_format.upper() == "HWP":
            target.write_bytes(sample_hwp_path.read_bytes())
        elif output_format.upper() == "HWPX":
            target.write_bytes(sample_hwpx_path.read_bytes())
        else:
            raise AssertionError(output_format)
        return target

    document = HancomDocument.blank(converter=fake_converter)
    document.metadata.title = "WRITE-HWP"
    document.append_paragraph("WRITE-HWP-PARAGRAPH")

    output_path = tmp_path / "hancom_ir.hwp"
    document.write_to_hwp(output_path)

    assert output_path.exists()
    assert conversions == []
