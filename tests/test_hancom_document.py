from __future__ import annotations

from pathlib import Path

import pytest

from jakal_hwpx import (
    AutoNumber,
    Bookmark,
    Equation,
    Field,
    HancomDocument,
    Hyperlink,
    HwpDocument,
    HwpxDocument,
    Note,
    Ole,
    Paragraph,
    Picture,
    Shape,
)
from jakal_hwpx.hwp_binary import TAG_BULLET, TAG_NUMBERING
from jakal_hwpx.namespaces import NS


REPO_ROOT = Path(__file__).resolve().parents[1]
HWP_SAMPLE_DIR = REPO_ROOT / "examples" / "samples" / "hwp"


@pytest.fixture(scope="session")
def sample_hwp_path() -> Path:
    paths = sorted(HWP_SAMPLE_DIR.glob("*.hwp"))
    assert paths, f"No .hwp samples were found under {HWP_SAMPLE_DIR}"
    return next((path for path in paths if not path.name.startswith("generated_")), paths[0])


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


def test_hancom_document_preserves_paragraph_style_refs_for_hwpx_roundtrip(tmp_path: Path) -> None:
    source = HwpxDocument.blank()
    source.append_paragraph_style(style_id="31", alignment_horizontal="CENTER", line_spacing=180)
    source.append_character_style(style_id="32", text_color="#334455", height=1200)
    source.append_style("PARA-STYLE", style_id="33", para_pr_id="31", char_pr_id="32")
    source.append_paragraph("Styled paragraph", style_id="33", para_pr_id="31", char_pr_id="32")

    document = HancomDocument.from_hwpx_document(source)
    paragraph_blocks = [block for block in document.sections[0].blocks if isinstance(block, Paragraph)]
    styled = next(block for block in paragraph_blocks if block.text == "Styled paragraph")

    assert styled.style_id == "33"
    assert styled.para_pr_id == "31"
    assert styled.char_pr_id == "32"

    output_path = tmp_path / "styled_roundtrip.hwpx"
    document.write_to_hwpx(output_path)

    reopened = HwpxDocument.open(output_path)
    paragraphs = reopened.sections[0].root_element.xpath("./hp:p", namespaces=NS)
    target = next(
        paragraph
        for paragraph in paragraphs
        if "".join(paragraph.xpath("./hp:run/hp:t/text()", namespaces=NS)).strip() == "Styled paragraph"
    )
    assert target.get("styleIDRef") == "33"
    assert target.get("paraPrIDRef") == "31"
    run_ids = target.xpath("./hp:run/@charPrIDRef", namespaces=NS)
    assert run_ids and all(value == "32" for value in run_ids)


def test_hancom_document_preserves_extended_style_fields_for_hwpx_roundtrip() -> None:
    source = HwpxDocument.blank()
    para_style = source.append_paragraph_style(style_id="41", alignment_horizontal="JUSTIFY")
    para_style.set_alignment(vertical="CENTER")
    para_style.set_line_spacing(175, spacing_type="PERCENT")
    heading = para_style.element.xpath("./hh:heading", namespaces=NS)[0]
    heading.set("type", "NUMBER")
    heading.set("idRef", "5")
    heading.set("level", "2")
    para_style.set_margin(intent=100, left=200, right=300, prev=400, next=500, unit="HWPUNIT")
    para_style.set_break_setting(
        break_latin_word="KEEP_WORD",
        break_non_latin_word="BREAK_WORD",
        widow_orphan=True,
        keep_with_next=True,
        keep_lines=False,
        page_break_before=True,
        line_wrap="SQUEEZE",
    )
    para_style.set_auto_spacing(e_asian_eng=True, e_asian_num=False)

    char_style = source.append_character_style(style_id="42", text_color="#123456", height=1300)
    char_style.set_font_refs(hangul="1", latin="2", hanja="3", japanese="4", other="5", symbol="6", user="7")

    document = HancomDocument.from_hwpx_document(source)
    extracted_para = next(style for style in document.paragraph_styles if style.style_id == "41")
    extracted_char = next(style for style in document.character_styles if style.style_id == "42")

    assert extracted_para.line_spacing == "175"
    assert extracted_para.line_spacing_type == "PERCENT"
    assert extracted_para.heading == {"type": "NUMBER", "idRef": "5", "level": "2"}
    assert extracted_para.margins == {
        "intent": "100",
        "left": "200",
        "right": "300",
        "prev": "400",
        "next": "500",
        "unit": "HWPUNIT",
    }
    assert extracted_para.break_setting == {
        "breakLatinWord": "KEEP_WORD",
        "breakNonLatinWord": "BREAK_WORD",
        "widowOrphan": "1",
        "keepWithNext": "1",
        "keepLines": "0",
        "pageBreakBefore": "1",
        "lineWrap": "SQUEEZE",
    }
    assert extracted_para.auto_spacing == {"eAsianEng": "1", "eAsianNum": "0"}
    assert extracted_char.font_refs == {
        "hangul": "1",
        "latin": "2",
        "hanja": "3",
        "japanese": "4",
        "other": "5",
        "symbol": "6",
        "user": "7",
    }

    reopened = document.to_hwpx_document()
    reopened_para = next(style for style in reopened.paragraph_styles() if style.style_id == "41")
    reopened_char = next(style for style in reopened.character_styles() if style.style_id == "42")

    assert reopened_para.element.xpath(".//hh:lineSpacing/@type", namespaces=NS) == ["PERCENT"]
    assert reopened_para.element.xpath("./hh:heading/@type", namespaces=NS) == ["NUMBER"]
    assert reopened_para.element.xpath("./hh:heading/@idRef", namespaces=NS) == ["5"]
    assert reopened_para.element.xpath("./hh:heading/@level", namespaces=NS) == ["2"]
    assert reopened_para.element.xpath("./hh:margin/hc:left/@value", namespaces=NS) == ["200"]
    assert reopened_para.element.xpath("./hh:breakSetting/@keepWithNext", namespaces=NS) == ["1"]
    assert reopened_para.element.xpath("./hh:autoSpacing/@eAsianNum", namespaces=NS) == ["0"]
    assert reopened_char.element.xpath("./hh:fontRef/@latin", namespaces=NS) == ["2"]
    assert reopened_char.element.xpath("./hh:fontRef/@user", namespaces=NS) == ["7"]


def test_hancom_document_preserves_extended_control_fields_for_hwpx_roundtrip(tmp_path: Path) -> None:
    source = HwpxDocument.blank()
    source.append_equation(
        "a^2+b^2",
        width=2800,
        height=1700,
        shape_comment="EQ-COMMENT",
        text_color="#13579B",
        base_unit=1250,
        font="Cambria Math",
    )
    source.append_shape(
        kind="rect",
        text="CTRL-SHAPE",
        width=4400,
        height=2100,
        fill_color="#ABCDEF",
        line_color="#123456",
        shape_comment="SHAPE-COMMENT",
    )
    source.append_ole("ctrl.ole", b"CTRL-OLE", width=3900, height=2200, shape_comment="OLE-COMMENT")

    source_path = tmp_path / "extended_controls.hwpx"
    source.save(source_path)

    document = HancomDocument.read_hwpx(source_path)
    equation = next(block for block in document.sections[0].blocks if isinstance(block, Equation))
    shape = next(block for block in document.sections[0].blocks if isinstance(block, Shape))
    ole = next(block for block in document.sections[0].blocks if isinstance(block, Ole))

    assert equation.shape_comment == "EQ-COMMENT"
    assert equation.text_color == "#13579B"
    assert equation.base_unit == 1250
    assert equation.font == "Cambria Math"
    assert shape.shape_comment == "SHAPE-COMMENT"
    assert ole.shape_comment == "OLE-COMMENT"

    reopened = document.to_hwpx_document()
    reopened_equation = reopened.equations()[0]
    reopened_shape = reopened.shapes()[0]
    reopened_ole = reopened.oles()[0]

    assert reopened_equation.shape_comment == "EQ-COMMENT"
    assert reopened_equation.element.get("textColor") == "#13579B"
    assert reopened_equation.element.get("baseUnit") == "1250"
    assert reopened_equation.element.get("font") == "Cambria Math"
    assert reopened_shape.shape_comment == "SHAPE-COMMENT"
    assert reopened_ole.shape_comment == "OLE-COMMENT"


def test_hancom_document_pure_hwp_roundtrip_preserves_style_definitions_and_docinfo_refs(tmp_path: Path) -> None:
    document = HancomDocument.blank()
    document.append_paragraph_style(style_id="430", alignment_horizontal="CENTER", line_spacing=180)
    document.append_character_style(style_id="425", text_color="#112233", height=1100)
    document.append_style(
        "HWP-IR-STYLE",
        style_id="33",
        english_name="HWP-IR-STYLE",
        para_pr_id="430",
        char_pr_id="425",
        next_style_id="33",
    )
    paragraph = document.append_paragraph("HWP-STYLED")
    paragraph.style_id = "33"
    paragraph.para_pr_id = "430"
    paragraph.char_pr_id = "425"

    output_path = tmp_path / "hancom_ir_style_docinfo.hwp"
    document.write_to_hwp(output_path)

    reopened_hwp = HwpDocument.open(output_path)
    styled_paragraph = next(paragraph for paragraph in reopened_hwp.paragraphs() if paragraph.text == "HWP-STYLED")
    assert styled_paragraph.style_id == 33
    assert styled_paragraph.para_shape_id == 430
    docinfo = reopened_hwp.docinfo_model()
    assert len(docinfo.style_records()) >= 34
    assert len(docinfo.char_shape_records()) >= 426
    assert len(docinfo.para_shape_records()) >= 431

    reopened_ir = HancomDocument.read_hwp(output_path)
    assert any(style.style_id == "33" and style.name == "HWP-IR-STYLE" for style in reopened_ir.style_definitions)
    assert any(style.style_id == "430" for style in reopened_ir.paragraph_styles)
    assert any(style.style_id == "425" for style in reopened_ir.character_styles)
    ir_paragraph = next(block for block in reopened_ir.sections[0].blocks if isinstance(block, Paragraph) and block.text == "HWP-STYLED")
    assert ir_paragraph.style_id == "33"
    assert ir_paragraph.para_pr_id == "430"
    assert ir_paragraph.char_pr_id == "425"

    reopened_hwpx = reopened_hwp.to_hwpx_document()
    assert any(style.style_id == "33" and style.name == "HWP-IR-STYLE" for style in reopened_hwpx.styles())
    assert any(style.style_id == "430" for style in reopened_hwpx.paragraph_styles())
    assert any(style.style_id == "425" for style in reopened_hwpx.character_styles())
    paragraphs = reopened_hwpx.sections[0].root_element.xpath("./hp:p", namespaces=NS)
    target = next(
        paragraph_node
        for paragraph_node in paragraphs
        if "".join(paragraph_node.xpath("./hp:run/hp:t/text()", namespaces=NS)).strip() == "HWP-STYLED"
    )
    assert target.get("styleIDRef") == "33"
    assert target.get("paraPrIDRef") == "430"
    run_ids = target.xpath("./hp:run/@charPrIDRef", namespaces=NS)
    assert run_ids and all(value == "425" for value in run_ids)


def test_hancom_document_pure_hwp_roundtrip_preserves_numbering_bullet_picture_size_and_hyperlink_metadata(
    tmp_path: Path,
) -> None:
    document = HancomDocument.blank()
    image_bytes = HwpDocument.blank().binary_document().read_stream("BinData/BIN0001.bmp", decompress=False)
    document.append_numbering_definition(style_id="2")
    document.append_bullet_definition(style_id="0")
    document.sections[0].settings.numbering_shape_id = "2"
    document.append_picture("native.bmp", image_bytes, extension="bmp", width=3600, height=2400)
    document.append_hyperlink(
        "https://example.com/native",
        display_text="NATIVE-LINK",
        metadata_fields=[9, 2, "anchor"],
    )

    output_path = tmp_path / "hancom_ir_numbering_picture_link.hwp"
    document.write_to_hwp(output_path)

    reopened_hwp = HwpDocument.open(output_path)
    native_settings = reopened_hwp.binary_document().section_definition_settings(0)
    docinfo = reopened_hwp.docinfo_model()
    picture = reopened_hwp.pictures()[0]
    hyperlink = next(link for link in reopened_hwp.hyperlinks() if link.display_text == "NATIVE-LINK")

    assert native_settings["numbering_shape_id"] == "2"
    assert len(docinfo.records_by_tag_id(TAG_NUMBERING)) >= 3
    assert len(docinfo.records_by_tag_id(TAG_BULLET)) >= 1
    assert picture.size() == {"width": 3600, "height": 2400}
    assert hyperlink.metadata_fields == ["9", "2", "anchor"]

    reopened_ir = HancomDocument.read_hwp(output_path)
    assert reopened_ir.sections[0].settings.numbering_shape_id == "2"
    assert any(definition.style_id == "2" for definition in reopened_ir.numbering_definitions)
    assert any(definition.style_id == "0" for definition in reopened_ir.bullet_definitions)
    picture_block = next(block for block in reopened_ir.sections[0].blocks if isinstance(block, Picture))
    hyperlink_block = next(block for block in reopened_ir.sections[0].blocks if isinstance(block, Hyperlink))
    assert picture_block.width == 3600
    assert picture_block.height == 2400
    assert hyperlink_block.metadata_fields == ["9", "2", "anchor"]


def test_hancom_document_pure_hwp_roundtrip_generates_numbering_and_bullet_payloads_from_ir(tmp_path: Path) -> None:
    document = HancomDocument.blank()
    document.append_character_style(style_id="12", text_color="#224466", height=1100)
    document.append_numbering_definition(
        style_id="5",
        formats=["^1.", "(^2)", "Part ^3"],
        para_heads=[12, 268, 44],
        width_adjusts=[0, 8, 16],
        text_offsets=[60, 70, 80],
        char_pr_ids=[12, None, 12],
        start_numbers=[1, 2, 3, 4, 5, 6, 7],
    )
    document.append_bullet_definition(
        style_id="1",
        bullet_char="*",
        width_adjust=14,
        text_offset=64,
        char_pr_id=12,
    )
    document.sections[0].settings.numbering_shape_id = "5"
    document.append_paragraph("NUMBERING-PAYLOAD")

    output_path = tmp_path / "hancom_ir_numbering_payloads.hwp"
    document.write_to_hwp(output_path)

    reopened_hwp = HwpDocument.open(output_path)
    docinfo = reopened_hwp.docinfo_model()
    assert len(docinfo.records_by_tag_id(TAG_NUMBERING)) >= 6
    assert len(docinfo.records_by_tag_id(TAG_BULLET)) >= 2

    reopened_ir = HancomDocument.read_hwp(output_path)
    numbering = next(definition for definition in reopened_ir.numbering_definitions if definition.style_id == "5")
    bullet = next(definition for definition in reopened_ir.bullet_definitions if definition.style_id == "1")

    assert reopened_ir.sections[0].settings.numbering_shape_id == "5"
    assert numbering.formats[:3] == ["^1.", "(^2)", "Part ^3"]
    assert numbering.para_heads[:3] == [12, 268, 44]
    assert numbering.width_adjusts[:3] == [0, 8, 16]
    assert numbering.text_offsets[:3] == [60, 70, 80]
    assert numbering.char_pr_ids[:3] == [12, None, 12]
    assert numbering.start_numbers[:7] == [1, 2, 3, 4, 5, 6, 7]
    assert bullet.bullet_char == "*"
    assert bullet.width_adjust == 14
    assert bullet.text_offset == 64
    assert bullet.char_pr_id == 12


def test_hancom_document_pure_hwp_roundtrip_generates_para_and_char_payloads_from_ir(tmp_path: Path) -> None:
    document = HancomDocument.blank()
    document.append_paragraph_style(
        style_id="430",
        alignment_horizontal="RIGHT",
        line_spacing=175,
        margins={
            "intent": "120",
            "left": "240",
            "right": "360",
            "prev": "480",
            "next": "600",
            "unit": "HWPUNIT",
        },
    )
    document.append_character_style(
        style_id="425",
        text_color="#123456",
        height=1300,
        font_refs={
            "hangul": "2",
            "latin": "1",
            "hanja": "1",
            "japanese": "1",
            "other": "0",
            "symbol": "2",
            "user": "0",
        },
    )
    document.append_style("DIRECT-PAYLOAD-STYLE", style_id="33", para_pr_id="430", char_pr_id="425")
    paragraph = document.append_paragraph("DIRECT-PAYLOAD")
    paragraph.style_id = "33"
    paragraph.para_pr_id = "430"
    paragraph.char_pr_id = "425"

    output_path = tmp_path / "hancom_ir_direct_style_payloads.hwp"
    document.write_to_hwp(output_path)

    reopened_ir = HancomDocument.read_hwp(output_path)
    para_style = next(style for style in reopened_ir.paragraph_styles if style.style_id == "430")
    char_style = next(style for style in reopened_ir.character_styles if style.style_id == "425")

    assert para_style.alignment_horizontal == "RIGHT"
    assert para_style.line_spacing == "175"
    assert para_style.margins == {
        "intent": "120",
        "left": "240",
        "right": "360",
        "prev": "480",
        "next": "600",
        "unit": "HWPUNIT",
    }
    assert char_style.text_color == "#123456"
    assert char_style.height == "1300"
    assert char_style.font_refs == {
        "hangul": "2",
        "latin": "1",
        "hanja": "1",
        "japanese": "1",
        "other": "0",
        "symbol": "2",
        "user": "0",
    }


def test_hancom_document_reads_richer_native_hwp_equation_shape_and_ole_controls(tmp_path: Path) -> None:
    source = HwpDocument.blank()
    source.append_equation("x+y=z", width=3456, height=2100, font="Batang", shape_comment="NATIVE-EQ")
    source.append_shape(
        kind="rect",
        text="NATIVE-SHAPE",
        width=3600,
        height=1800,
        fill_color="#ABCDEF",
        line_color="#123456",
        shape_comment="NATIVE-SHAPE-COMMENT",
    )
    source.append_ole(
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

    output_path = tmp_path / "hancom_native_control_rich.hwp"
    source.save(output_path)

    document = HancomDocument.read_hwp(output_path)
    equation = next(block for block in document.sections[0].blocks if isinstance(block, Equation))
    shape = next(block for block in document.sections[0].blocks if isinstance(block, Shape))
    ole = next(block for block in document.sections[0].blocks if isinstance(block, Ole))

    assert equation.script == "x+y=z"
    assert equation.width == 3456
    assert equation.height == 2100
    assert equation.font == "Batang"
    assert equation.shape_comment == "NATIVE-EQ"
    assert shape.text == "NATIVE-SHAPE"
    assert shape.width == 3600
    assert shape.height == 1800
    assert shape.shape_comment == "NATIVE-SHAPE-COMMENT"
    assert shape.fill_color == "#ABCDEF"
    assert shape.line_color == "#123456"
    assert ole.name == "native.ole"
    assert ole.width == 3900
    assert ole.height == 2200
    assert ole.shape_comment == "NATIVE-OLE-COMMENT"
    assert ole.object_type == "LINK"
    assert ole.draw_aspect == "ICON"
    assert ole.has_moniker is True
    assert ole.eq_baseline == 12
    assert ole.line_color == "#445566"
    assert ole.line_width == 77

    reopened_hwpx = document.to_hwpx_document()
    assert reopened_hwpx.equations()[0].script == "x+y=z"
    assert reopened_hwpx.equations()[0].size() == {"width": 3456, "height": 2100}
    assert reopened_hwpx.equations()[0].element.get("font") == "Batang"
    assert reopened_hwpx.equations()[0].shape_comment == "NATIVE-EQ"
    assert reopened_hwpx.shapes()[0].text == "NATIVE-SHAPE"
    assert reopened_hwpx.shapes()[0].shape_comment == "NATIVE-SHAPE-COMMENT"
    assert reopened_hwpx.shapes()[0].fill_style()["faceColor"] == "#ABCDEF"
    assert reopened_hwpx.shapes()[0].line_style()["color"] == "#123456"
    assert reopened_hwpx.oles()[0].binary_data() == b"NATIVE-OLE-DATA"
    assert reopened_hwpx.oles()[0].shape_comment == "NATIVE-OLE-COMMENT"
    assert reopened_hwpx.oles()[0].object_type == "LINK"
    assert reopened_hwpx.oles()[0].draw_aspect == "ICON"
    assert reopened_hwpx.oles()[0].has_moniker is True
    assert reopened_hwpx.oles()[0].line_style()["color"] == "#445566"
    assert reopened_hwpx.oles()[0].line_style()["width"] == "77"


def test_hancom_document_pure_hwp_roundtrip_preserves_native_shape_and_ole_metadata(tmp_path: Path) -> None:
    document = HancomDocument.blank()
    document.append_equation("a+b", width=3200, height=1800, font="Batang", shape_comment="IR-EQ")
    document.append_shape(
        kind="rect",
        text="IR-SHAPE",
        width=4100,
        height=1900,
        fill_color="#FEDCBA",
        line_color="#234567",
        shape_comment="IR-SHAPE-COMMENT",
    )
    document.append_ole(
        "ir.ole",
        b"IR-OLE-DATA",
        width=4200,
        height=2300,
        shape_comment="IR-OLE-COMMENT",
        object_type="LINK",
        draw_aspect="ICON",
        has_moniker=True,
        eq_baseline=9,
        line_color="#556677",
        line_width=55,
    )

    output_path = tmp_path / "hancom_native_full_parity.hwp"
    document.write_to_hwp(output_path)

    reopened = HancomDocument.read_hwp(output_path)
    equation = next(block for block in reopened.sections[0].blocks if isinstance(block, Equation))
    shape = next(block for block in reopened.sections[0].blocks if isinstance(block, Shape))
    ole = next(block for block in reopened.sections[0].blocks if isinstance(block, Ole))

    assert equation.shape_comment == "IR-EQ"
    assert shape.shape_comment == "IR-SHAPE-COMMENT"
    assert shape.fill_color == "#FEDCBA"
    assert shape.line_color == "#234567"
    assert ole.shape_comment == "IR-OLE-COMMENT"
    assert ole.object_type == "LINK"
    assert ole.draw_aspect == "ICON"
    assert ole.has_moniker is True
    assert ole.eq_baseline == 9
    assert ole.line_color == "#556677"
    assert ole.line_width == 55


def test_hancom_document_native_hwp_roundtrip_preserves_richer_shape_kinds(tmp_path: Path) -> None:
    source = HwpDocument.blank()
    source.append_shape(
        kind="ellipse",
        text="NATIVE-ELLIPSE",
        width=3800,
        height=2000,
        fill_color="#A1B2C3",
        line_color="#102030",
        shape_comment="ELLIPSE-COMMENT",
    )
    source.append_shape(
        kind="arc",
        text="NATIVE-ARC",
        width=4000,
        height=2200,
        fill_color="#C3B2A1",
        line_color="#203040",
        shape_comment="ARC-COMMENT",
    )
    source.append_shape(
        kind="polygon",
        text="NATIVE-POLYGON",
        width=4200,
        height=2400,
        fill_color="#DDEEFF",
        line_color="#304050",
        shape_comment="POLYGON-COMMENT",
    )
    source.append_shape(
        kind="textart",
        text="NATIVE-TEXTART",
        width=4400,
        height=2600,
        fill_color="#FFEEDD",
        line_color="#405060",
        shape_comment="TEXTART-COMMENT",
    )

    output_path = tmp_path / "hancom_native_shape_kinds.hwp"
    source.save(output_path)

    document = HancomDocument.read_hwp(output_path)
    shapes = [block for block in document.sections[0].blocks if isinstance(block, Shape)]

    ellipse = next(shape for shape in shapes if shape.kind == "ellipse")
    arc = next(shape for shape in shapes if shape.kind == "arc")
    polygon = next(shape for shape in shapes if shape.kind == "polygon")
    textart = next(shape for shape in shapes if shape.kind == "textart")

    assert ellipse.fill_color == "#A1B2C3"
    assert ellipse.line_color == "#102030"
    assert ellipse.shape_comment == "ELLIPSE-COMMENT"
    assert arc.fill_color == "#C3B2A1"
    assert arc.line_color == "#203040"
    assert arc.shape_comment == "ARC-COMMENT"
    assert polygon.fill_color == "#DDEEFF"
    assert polygon.line_color == "#304050"
    assert polygon.shape_comment == "POLYGON-COMMENT"
    assert textart.text == "NATIVE-TEXTART"
    assert textart.fill_color == "#FFEEDD"
    assert textart.line_color == "#405060"
    assert textart.shape_comment == "TEXTART-COMMENT"

    reopened_hwpx = document.to_hwpx_document()
    reopened_shapes = reopened_hwpx.shapes()
    reopened_ellipse = next(shape for shape in reopened_shapes if shape.kind == "ellipse")
    reopened_arc = next(shape for shape in reopened_shapes if shape.kind == "arc")
    reopened_polygon = next(shape for shape in reopened_shapes if shape.kind == "polygon")
    reopened_textart = next(shape for shape in reopened_shapes if shape.kind == "textart")

    assert reopened_ellipse.fill_style()["faceColor"] == "#A1B2C3"
    assert reopened_ellipse.line_style()["color"] == "#102030"
    assert reopened_ellipse.shape_comment == "ELLIPSE-COMMENT"
    assert reopened_arc.fill_style()["faceColor"] == "#C3B2A1"
    assert reopened_arc.line_style()["color"] == "#203040"
    assert reopened_arc.shape_comment == "ARC-COMMENT"
    assert reopened_polygon.fill_style()["faceColor"] == "#DDEEFF"
    assert reopened_polygon.line_style()["color"] == "#304050"
    assert reopened_polygon.shape_comment == "POLYGON-COMMENT"
    assert reopened_textart.text == "NATIVE-TEXTART"
    assert reopened_textart.fill_style()["faceColor"] == "#FFEEDD"
    assert reopened_textart.line_style()["color"] == "#405060"
    assert reopened_textart.shape_comment == "TEXTART-COMMENT"


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


def test_hancom_document_can_extract_native_hwp_header_and_auto_number_controls(
    sample_hwp_path: Path,
    tmp_path: Path,
) -> None:
    document = HancomDocument.read_hwp(sample_hwp_path)
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
