from __future__ import annotations

from pathlib import Path

import pytest

from jakal_hwpx import (
    AutoNumber,
    Bookmark,
    Chart,
    Equation,
    Field,
    Form,
    HancomDocument,
    Hyperlink,
    HwpConnectLineShapeObject,
    HwpDocument,
    HwpxDocument,
    Memo,
    Note,
    Ole,
    Paragraph,
    Picture,
    Shape,
    Table,
)
from jakal_hwpx.hwp_binary import TAG_BULLET, TAG_NUMBERING
from jakal_hwpx.namespaces import NS


REPO_ROOT = Path(__file__).resolve().parents[1]
HWP_SAMPLE_DIR = REPO_ROOT / "examples" / "samples" / "hwp"
MATH_DEBUG_SAMPLE_DIR = REPO_ROOT / "math debug sample"


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


def test_hancom_document_does_not_materialize_default_hwpx_memo_shape_in_hwp_bridge(tmp_path: Path) -> None:
    source = HwpxDocument.blank()

    document = HancomDocument.from_hwpx_document(source)
    assert document.sections[0].settings.memo_shape_id is None

    hwp_document = document.to_hwp_document()
    assert hwp_document.docinfo_model().memo_shape_records() == []
    assert hwp_document.binary_document().section_definition_settings(0).get("memo_shape_id") is None

    roundtrip = HancomDocument.from_hwp_document(hwp_document).to_hwpx_document()
    output_path = tmp_path / "blank_memo_shape_roundtrip.hwpx"
    roundtrip.save(output_path)

    reopened = HwpxDocument.open(output_path)
    assert reopened.memo_shapes() == []
    assert reopened.section_settings().memo_shape_id == "0"


@pytest.mark.parametrize("sample_name", ["mock_exam_before.hwpx", "mock_exam_1_10 after.hwpx"])
def test_math_debug_sample_preserves_equation_sizes_through_hwpx_bridge(tmp_path: Path, sample_name: str) -> None:
    sample_path = MATH_DEBUG_SAMPLE_DIR / sample_name
    source = HancomDocument.from_hwpx_document(HwpxDocument.open(sample_path))
    before = [
        (block.script, block.width, block.height)
        for section in source.sections
        for block in section.blocks
        if isinstance(block, Equation)
    ]

    output_path = tmp_path / f"{sample_path.stem}_math_debug_roundtrip.hwpx"
    source.to_hwpx_document().save(output_path)
    restored = HancomDocument.from_hwpx_document(HwpxDocument.open(output_path))
    after = [
        (block.script, block.width, block.height)
        for section in restored.sections
        for block in section.blocks
        if isinstance(block, Equation)
    ]

    assert after == before


@pytest.mark.parametrize("sample_name", ["mock_exam_before.hwpx", "mock_exam_1_10 after.hwpx"])
def test_math_debug_sample_preserves_equation_sizes_through_hwp_bridge(tmp_path: Path, sample_name: str) -> None:
    sample_path = MATH_DEBUG_SAMPLE_DIR / sample_name
    source = HancomDocument.from_hwpx_document(HwpxDocument.open(sample_path))
    before = [
        (block.script, block.width, block.height)
        for section in source.sections
        for block in section.blocks
        if isinstance(block, Equation)
    ]

    hwp_path = tmp_path / f"{sample_path.stem}_math_debug_roundtrip.hwp"
    source.to_hwp_document().save(hwp_path)
    restored = HancomDocument.from_hwp_document(HwpDocument.open(hwp_path))
    after = [
        (block.script, block.width, block.height)
        for section in restored.sections
        for block in section.blocks
        if isinstance(block, Equation)
    ]

    assert after == before


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
    equation = source.append_equation("a+b", width=2800, height=1700)
    equation.set_layout(
        text_wrap="SQUARE",
        text_flow="RIGHT_ONLY",
        treat_as_char=False,
        vert_rel_to="PAPER",
        horz_rel_to="PAPER",
        vert_align="TOP",
        horz_align="LEFT",
        vert_offset=12,
        horz_offset=34,
    )
    equation.set_out_margins(left=5, right=6, top=7, bottom=8)
    equation.set_rotation(angle=35, center_x=555, center_y=666)
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
    extracted_equation = next(block for block in blocks if isinstance(block, Equation) and getattr(block, "script", None) == "a+b")
    assert extracted_equation.layout["textWrap"] == "SQUARE"
    assert extracted_equation.layout["textFlow"] == "RIGHT_ONLY"
    assert extracted_equation.layout["treatAsChar"] == "0"
    assert extracted_equation.layout["vertOffset"] == "12"
    assert extracted_equation.out_margins == {"left": 5, "right": 6, "top": 7, "bottom": 8}
    assert extracted_equation.rotation["angle"] == "35"
    assert extracted_equation.rotation["centerX"] == "555"
    assert extracted_equation.rotation["centerY"] == "666"
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


def test_hancom_document_bridges_form_memo_and_chart_through_hwp() -> None:
    document = HancomDocument.blank()
    document.append_form(
        "Consent",
        form_type="CHECKBOX",
        name="consent",
        value="Y",
        checked=True,
        items=["Y", "N"],
        editable=False,
        locked=True,
        placeholder="Choose",
    )
    document.append_memo(
        "Review later",
        author="Alice",
        memo_id="memo-1",
        anchor_id="anchor-1",
        order=2,
        visible=False,
    )
    chart = document.append_chart(
        "Revenue",
        chart_type="LINE",
        categories=["Q1", "Q2"],
        series=[{"name": "Sales", "values": [10, 20]}],
        data_ref="dataset-1",
        legend_visible=False,
        width=3300,
        height=2100,
        shape_comment="chart-comment",
    )
    chart.layout = {"textWrap": "SQUARE", "textFlow": "LEFT_ONLY", "treatAsChar": "0"}
    chart.out_margins = {"left": 1, "right": 2, "top": 3, "bottom": 4}
    chart.rotation = {"angle": "15", "centerX": "123", "centerY": "456"}

    hwp_document = document.to_hwp_document()

    hwp_form = hwp_document.forms()[0]
    assert hwp_form.label == "Consent"
    assert hwp_form.form_type == "CHECKBOX"
    assert hwp_form.name == "consent"
    assert hwp_form.value == "Y"
    assert hwp_form.checked is True
    assert hwp_form.items == ["Y", "N"]
    assert hwp_form.editable is False
    assert hwp_form.locked is True
    assert hwp_form.placeholder == "Choose"

    hwp_memo = hwp_document.memos()[0]
    assert hwp_memo.text == "Review later"
    assert hwp_memo.author == "Alice"
    assert hwp_memo.memo_id == "memo-1"
    assert hwp_memo.anchor_id == "anchor-1"
    assert hwp_memo.order == 2
    assert hwp_memo.visible is False

    hwp_chart = hwp_document.charts()[0]
    assert hwp_chart.title == "Revenue"
    assert hwp_chart.chart_type == "LINE"
    assert hwp_chart.categories == ["Q1", "Q2"]
    assert hwp_chart.series == [{"name": "Sales", "values": [10, 20]}]
    assert hwp_chart.data_ref == "dataset-1"
    assert hwp_chart.legend_visible is False
    assert hwp_chart.shape_comment == "chart-comment"
    assert hwp_chart.layout()["textWrap"] == "SQUARE"
    assert hwp_chart.layout()["textFlow"] == "LEFT_ONLY"
    assert hwp_chart.rotation()["angle"] == "15"

    restored = HancomDocument.from_hwp_document(hwp_document)
    restored_form = next(block for block in restored.sections[0].blocks if isinstance(block, Form))
    restored_memo = next(block for block in restored.sections[0].blocks if isinstance(block, Memo))
    restored_chart = next(block for block in restored.sections[0].blocks if isinstance(block, Chart))

    assert restored_form.label == "Consent"
    assert restored_form.checked is True
    assert restored_form.items == ["Y", "N"]
    assert restored_memo.author == "Alice"
    assert restored_memo.visible is False
    assert restored_chart.chart_type == "LINE"
    assert restored_chart.categories == ["Q1", "Q2"]
    assert restored_chart.series == [{"name": "Sales", "values": [10, 20]}]
    assert restored_chart.shape_comment == "chart-comment"


def test_hancom_document_bridges_form_memo_and_chart_through_hwpx() -> None:
    document = HancomDocument.blank()
    memo_shape = document.append_memo_shape_definition(
        width=16384,
        line_width=24,
        line_type="DOT",
        line_color="#112233",
        fill_color="#FFF4CC",
        active_color="#FFAA00",
        memo_type="USER_INSERT",
    )
    document.sections[0].settings.set_memo_shape_id(memo_shape.memo_shape_id)
    document.append_form(
        "Approval",
        form_type="INPUT",
        name="approval",
        value="pending",
        checked=False,
        items=["pending", "approved"],
        editable=True,
        locked=False,
        placeholder="Type status",
    )
    document.append_memo(
        "Check attachment",
        author="Bob",
        memo_id="memo-2",
        anchor_id="anchor-2",
        order=5,
        visible=True,
    )
    chart = document.append_chart(
        "Trend",
        chart_type="BAR",
        categories=["Jan", "Feb"],
        series=[{"name": "Count", "values": [3, 4]}],
        data_ref="dataset-2",
        legend_visible=True,
        width=3600,
        height=2000,
        shape_comment="chart-note",
    )
    chart.layout = {"textWrap": "SQUARE", "textFlow": "RIGHT_ONLY", "treatAsChar": "0"}
    chart.out_margins = {"left": 6, "right": 7, "top": 8, "bottom": 9}
    chart.rotation = {"angle": "25", "centerX": "222", "centerY": "333"}

    hwpx_document = document.to_hwpx_document()
    field_types = [field.field_type for field in hwpx_document.fields()]
    assert "JAKAL_FORM" not in field_types
    assert "JAKAL_MEMO" in field_types
    assert [memo.text for memo in hwpx_document.memos()] == ["Check attachment"]
    assert len(hwpx_document.memo_shapes()) == 1
    assert hwpx_document.memo_shapes()[0].memo_shape_id == "0"
    assert hwpx_document.memo_shapes()[0].memo_type == "USER_INSERT"
    assert hwpx_document.section_settings().memo_shape_id == "0"
    assert len(hwpx_document.charts()) == 1
    assert hwpx_document.charts()[0].chart_type == "BAR"
    assert hwpx_document.charts()[0].shape_comment == "chart-note"
    assert hwpx_document.charts()[0].data_ref == "dataset-2"
    assert hwpx_document.charts()[0].rotation()["angle"] == "25"
    assert hwpx_document.charts()[0].rotation()["centerX"] == "222"
    assert hwpx_document.charts()[0].rotation()["centerY"] == "333"
    assert hwpx_document.oles() == []
    assert hwpx_document.sections[0].root_element.xpath(".//hp:chart/hp:shapeComment[text()='chart-note']", namespaces=NS)
    assert hwpx_document.sections[0].root_element.xpath(
        ".//hp:switch[hp:case/hp:chart]/hp:default/hp:ole/hp:rotationInfo[@angle='25'][@centerX='222'][@centerY='333']",
        namespaces=NS,
    )
    assert "Chart/chart1.xml" in hwpx_document.list_part_paths()
    assert "BinData/ole1.ole" in hwpx_document.list_part_paths()
    forms = hwpx_document.forms()
    assert len(forms) == 1
    assert forms[0].form_type == "INPUT"
    assert forms[0].label == "Approval"
    assert forms[0].value == "pending"
    assert forms[0].placeholder == "Type status"
    edit_nodes = hwpx_document.sections[0].root_element.xpath(".//hp:edit", namespaces=NS)
    assert edit_nodes
    assert edit_nodes[0].xpath("./hp:caption/hp:subList/hp:p/hp:run/hp:t[text()='Approval']", namespaces=NS)
    assert "label" not in edit_nodes[0].get("command", "")

    restored = HancomDocument.from_hwpx_document(hwpx_document)
    restored_form = next(block for block in restored.sections[0].blocks if isinstance(block, Form))
    restored_memo = next(block for block in restored.sections[0].blocks if isinstance(block, Memo))
    restored_chart = next(block for block in restored.sections[0].blocks if isinstance(block, Chart))

    assert restored_form.label == "Approval"
    assert restored_form.form_type == "INPUT"
    assert restored_form.name == "approval"
    assert restored_form.value == "pending"
    assert restored_form.items == ["pending", "approved"]
    assert restored_form.placeholder == "Type status"
    assert restored_memo.text == "Check attachment"
    assert restored_memo.author == "Bob"
    assert restored_memo.memo_id == "memo-2"
    assert restored_memo.anchor_id == "anchor-2"
    assert restored_memo.order == 5
    assert restored_memo.visible is True
    assert len(restored.memo_shape_definitions) == 1
    assert restored.memo_shape_definitions[0].memo_type == "USER_INSERT"
    assert restored.sections[0].settings.memo_shape_id == "0"
    assert restored_chart.title == "Trend"
    assert restored_chart.chart_type == "BAR"
    assert restored_chart.categories == ["Jan", "Feb"]
    assert restored_chart.series == [{"name": "Count", "values": [3, 4]}]
    assert restored_chart.data_ref == "dataset-2"
    assert restored_chart.legend_visible is True
    assert restored_chart.shape_comment == "chart-note"
    assert restored_chart.layout["textFlow"] == "RIGHT_ONLY"
    assert restored_chart.rotation["angle"] == "25"
    assert restored_chart.rotation["centerX"] == "222"
    assert restored_chart.rotation["centerY"] == "333"

    hwp_document = restored.to_hwp_document()
    assert hwp_document.forms()[0].placeholder == "Type status"
    assert hwp_document.memos()[0].text == "Check attachment"
    assert hwp_document.charts()[0].chart_type == "BAR"
    assert hwp_document.charts()[0].shape_comment == "chart-note"
    assert hwp_document.charts()[0].rotation()["angle"] == "25"
    assert hwp_document.binary_document().section_definition_settings(0)["memo_shape_id"] == "0"
    memo_shape_records = hwp_document.docinfo_model().memo_shape_records()
    assert len(memo_shape_records) == 1
    assert memo_shape_records[0].fields()["memo_type"] == "USER_INSERT"

    restored_from_hwp = HancomDocument.from_hwp_document(hwp_document)
    assert len(restored_from_hwp.memo_shape_definitions) == 1
    assert restored_from_hwp.memo_shape_definitions[0].memo_shape_id == "0"
    assert restored_from_hwp.memo_shape_definitions[0].memo_type == "USER_INSERT"
    assert restored_from_hwp.sections[0].settings.memo_shape_id == "0"


def test_hancom_document_outputs_pass_strict_lint_for_both_formats(tmp_path: Path) -> None:
    document = HancomDocument.blank()
    document.metadata.title = "Strict parity"
    document.append_header("internal")
    document.append_paragraph("body")
    document.append_table(rows=2, cols=2, cell_texts=[["A", "B"], ["1", "2"]])
    document.append_bookmark("anchor")
    document.append_cross_reference("anchor", display_text="see anchor")
    document.append_form("Approval", form_type="INPUT", name="approval", value="pending", placeholder="status")
    document.append_memo("review", author="QA", memo_id="memo-1", anchor_id="anchor-1", order=1, visible=True)
    document.append_chart("Revenue", categories=["Q1"], series=[{"name": "Sales", "values": [10]}])

    hwpx_path = tmp_path / "strict_parity.hwpx"
    hwp_path = tmp_path / "strict_parity.hwp"
    document.write_to_hwpx(hwpx_path)
    document.write_to_hwp(hwp_path)

    hwpx_document = HwpxDocument.open(hwpx_path)
    hwp_document = HwpDocument.open(hwp_path)

    assert hwpx_document.strict_lint_errors() == []
    assert hwp_document.strict_lint_errors() == []

    from_hwpx = HancomDocument.read_hwpx(hwpx_path)
    from_hwp = HancomDocument.read_hwp(hwp_path)

    assert [type(block).__name__ for block in from_hwpx.sections[0].blocks] == [
        type(block).__name__ for block in from_hwp.sections[0].blocks
    ]
    assert any(isinstance(block, Chart) for block in from_hwp.sections[0].blocks)
    assert any(isinstance(block, Form) for block in from_hwpx.sections[0].blocks)


def test_hancom_document_pairs_hwpx_memo_carriers_by_text_before_index() -> None:
    source = HwpxDocument.blank()
    source.append_memo("Memo A", paragraph_index=0)
    source.append_memo("Memo B", paragraph_index=0)
    source.append_field(
        field_type="JAKAL_MEMO",
        name="memo-b",
        parameters={"Text": "Memo B", "Author": "Bob", "Order": "2", "Visible": "0"},
        paragraph_index=0,
    )
    source.append_field(
        field_type="JAKAL_MEMO",
        name="memo-a",
        parameters={"Text": "Memo A", "Author": "Alice", "Order": "1", "Visible": "1"},
        paragraph_index=0,
    )

    document = HancomDocument.from_hwpx_document(source)
    memos = [block for block in document.sections[0].blocks if isinstance(block, Memo)]

    assert [(memo.text, memo.author, memo.memo_id, memo.order, memo.visible) for memo in memos] == [
        ("Memo A", "Alice", "memo-a", 1, True),
        ("Memo B", "Bob", "memo-b", 2, False),
    ]


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


def test_hancom_document_read_hwp_preserves_field_name_parameters_and_flags(tmp_path: Path) -> None:
    source = HwpDocument.blank()
    source.append_field(
        field_type="MAILMERGE",
        display_text="NATIVE-FIELD",
        name="MAIL_NAME",
        parameters={"FieldName": "MAIL_NAME", "Format": "upper"},
        editable=True,
        dirty=True,
    )

    output_path = tmp_path / "hancom_native_field.hwp"
    source.save(output_path)

    document = HancomDocument.read_hwp(output_path)
    field = next(block for block in document.sections[0].blocks if isinstance(block, Field))

    assert field.field_type == "MAILMERGE"
    assert field.display_text == "NATIVE-FIELD"
    assert field.name == "MAIL_NAME"
    assert field.native_field_type == "%mai"
    assert field.semantic_field_type == "MAILMERGE"
    assert field.parameters["FieldName"] == "MAIL_NAME"
    assert field.parameters["Format"] == "upper"
    assert field.is_mail_merge is True
    assert field.merge_field_name == "MAIL_NAME"
    assert field.editable is True
    assert field.dirty is True


def test_hancom_document_roundtrips_native_field_subtypes_back_to_hwp(tmp_path: Path) -> None:
    source = HwpDocument.blank()
    source.append_field(
        field_type="%pag",
        display_text="PAGE-REF",
        name="PAGE_ANCHOR",
        parameters={"BookmarkName": "PAGE_ANCHOR", "Path": "PAGE_ANCHOR"},
    )

    source_path = tmp_path / "native_field_subtype_source.hwp"
    source.save(source_path)

    ir_document = HancomDocument.read_hwp(source_path)
    field = next(block for block in ir_document.sections[0].blocks if isinstance(block, Field))
    roundtrip_path = tmp_path / "native_field_subtype_roundtrip.hwp"
    ir_document.to_hwp_document().save(roundtrip_path)

    reopened = HwpDocument.open(roundtrip_path)
    reopened_field = reopened.fields()[0]

    assert field.field_type == "CROSSREF"
    assert field.native_field_type == "%pag"
    assert field.semantic_field_type == "CROSSREF"
    assert field.is_cross_reference is True
    assert field.bookmark_name == "PAGE_ANCHOR"
    assert reopened_field.field_type == "CROSSREF"
    assert reopened_field.native_field_type == "%pag"
    assert reopened_field.bookmark_name == "PAGE_ANCHOR"


def test_hancom_document_append_field_preserves_explicit_native_field_type(tmp_path: Path) -> None:
    document = HancomDocument.blank()
    field = document.append_field(
        field_type="%pag",
        display_text="PAGE-REF",
        name="PAGE_ANCHOR",
        parameters={"BookmarkName": "PAGE_ANCHOR", "Path": "PAGE_ANCHOR"},
    )
    output_path = tmp_path / "hancom_explicit_native_field.hwp"
    document.to_hwp_document().save(output_path)

    reopened = HwpDocument.open(output_path)
    reopened_field = reopened.fields()[0]

    assert field.field_type == "CROSSREF"
    assert field.native_field_type == "%pag"
    assert field.semantic_field_type == "CROSSREF"
    assert reopened_field.native_field_type == "%pag"
    assert reopened_field.bookmark_name == "PAGE_ANCHOR"


def test_hancom_document_helper_methods_configure_fields_and_settings_for_hwp_roundtrip(tmp_path: Path) -> None:
    document = HancomDocument.blank()
    settings = document.sections[0].settings
    settings.set_page_size(width=60000, height=85000, landscape="WIDELY")
    settings.set_margins(left=7000, right=7100, top=5000, bottom=5100)
    settings.set_visibility(hide_first_header=True, border="HIDE_FIRST", show_line_number=True)
    settings.set_grid(line_grid=42, char_grid=21, wonggoji_format=True)
    settings.set_start_numbers(page_starts_on="ODD", page=7, pic=8, tbl=9, equation=10)
    settings.set_page_numbers([{"pos": "TOP_RIGHT", "formatType": "ROMAN_SMALL", "sideChar": "*"}])
    settings.set_note_settings(footnote_pr={"numbering": {"type": "CONTINUOUS", "newNum": "4"}})

    field = document.append_mail_merge_field("customer_name", display_text="CUSTOMER")
    field.configure_doc_property("Subject", display_text="SUBJECT")
    header = document.append_header("HEADER", apply_page_type="BOTH")
    header.set_text("HEADER-UPDATED")
    header.set_apply_page_type("EVEN")
    auto_number = document.append_auto_number(kind="autoNum")
    auto_number.set_number_type("PAGE")

    output_path = tmp_path / "hancom_helper_roundtrip.hwp"
    document.write_to_hwp(output_path)

    reopened = HancomDocument.read_hwp(output_path)
    reopened_field = next(block for block in reopened.sections[0].blocks if isinstance(block, Field))
    reopened_header = reopened.sections[0].header_footer_blocks[0]
    reopened_auto_number = next(block for block in reopened.sections[0].blocks if isinstance(block, AutoNumber))

    assert reopened_field.field_type == "DOCPROPERTY"
    assert reopened_field.document_property_name == "Subject"
    assert reopened_field.display_text == "SUBJECT"
    assert reopened_header.text == "HEADER-UPDATED"
    assert reopened_header.apply_page_type == "EVEN"
    assert reopened.sections[0].settings.page_numbers == [{"pos": "TOP_RIGHT", "formatType": "ROMAN_SMALL", "sideChar": "*"}]
    assert reopened.sections[0].settings.start_numbers["pageStartsOn"] == "ODD"
    assert reopened.sections[0].settings.grid == {"lineGrid": 42, "charGrid": 21, "wonggojiFormat": 1}
    assert reopened.sections[0].settings.visibility["border"] == "HIDE_FIRST"
    assert reopened.sections[0].settings.footnote_pr["numbering"]["newNum"] == "4"
    assert reopened_auto_number.kind == "autoNum"
    assert reopened_auto_number.number_type == "PAGE"


def test_hancom_document_read_hwp_preserves_native_note_numbers(tmp_path: Path) -> None:
    source = HwpDocument.blank()
    source.apply_section_note_settings(
        section_index=0,
        footnote_pr={"numbering": {"type": "CONTINUOUS", "newNum": "5"}},
        endnote_pr={"numbering": {"type": "CONTINUOUS", "newNum": "9"}},
    )
    source.append_footnote("FIRST-FOOTNOTE")
    source.append_endnote("FIRST-ENDNOTE")

    output_path = tmp_path / "hancom_native_notes.hwp"
    source.save(output_path)

    document = HancomDocument.read_hwp(output_path)
    footnote = next(block for block in document.sections[0].blocks if isinstance(block, Note) and block.kind == "footNote")
    endnote = next(block for block in document.sections[0].blocks if isinstance(block, Note) and block.kind == "endNote")

    assert footnote.text == "FIRST-FOOTNOTE"
    assert footnote.number == 5
    assert endnote.text == "FIRST-ENDNOTE"
    assert endnote.number == 9


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
        unknown_short_bits={"bit_1": True, "bit_4": True},
    )
    document.append_bullet_definition(
        style_id="1",
        flags=0x11,
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
    assert numbering.unknown_short_bits["bit_1"] is True
    assert numbering.unknown_short_bits["bit_4"] is True
    assert bullet.bullet_char == "*"
    assert bullet.flags == 0x11
    assert bullet.flag_bits["bit_0"] is True
    assert bullet.flag_bits["bit_4"] is True
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


def test_hancom_document_preserves_native_style_definition_trailing_payload(tmp_path: Path) -> None:
    template_document = HwpDocument.blank()
    template_payload = bytes(template_document.docinfo_model().style_records()[0].payload) + b"\x34\x12\x78\x56"

    document = HancomDocument.blank()
    document.append_style(
        "TRAILING-STYLE",
        style_id="61",
        para_pr_id="0",
        char_pr_id="0",
        native_hwp_payload=template_payload,
    )
    document.append_paragraph("STYLE-TRAILING")

    output_path = tmp_path / "style_trailing_payload.hwp"
    document.write_to_hwp(output_path)

    reopened = HwpDocument.open(output_path)
    payload = bytes(reopened.docinfo_model().style_records()[61].payload)
    assert payload.endswith(b"\x34\x12\x78\x56")


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
    equation = source.equations()[0]
    equation.set_layout(
        text_wrap="SQUARE",
        text_flow="RIGHT_ONLY",
        treat_as_char=False,
        vert_rel_to="PAPER",
        horz_rel_to="PAPER",
        vert_align="TOP",
        horz_align="LEFT",
        vert_offset=12,
        horz_offset=34,
    )
    equation.set_out_margins(left=5, right=6, top=7, bottom=8)
    equation.set_rotation(angle=35, center_x=555, center_y=666)
    shape = source.shapes()[0]
    shape.set_layout(
        text_wrap="SQUARE",
        text_flow="LEFT_ONLY",
        treat_as_char=False,
        vert_rel_to="PAPER",
        horz_rel_to="PAPER",
        vert_align="TOP",
        horz_align="LEFT",
        vert_offset=21,
        horz_offset=43,
    )
    ole = source.oles()[0]
    ole.set_layout(
        text_wrap="SQUARE",
        text_flow="LEFT_ONLY",
        treat_as_char=False,
        vert_rel_to="PAPER",
        horz_rel_to="PAPER",
        vert_align="TOP",
        horz_align="LEFT",
        vert_offset=31,
        horz_offset=53,
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
    assert equation.layout["textWrap"] == "SQUARE"
    assert equation.layout["textFlow"] == "RIGHT_ONLY"
    assert equation.layout["treatAsChar"] == "0"
    assert equation.layout["vertOffset"] == "12"
    assert equation.out_margins == {"left": 5, "right": 6, "top": 7, "bottom": 8}
    assert equation.rotation["angle"] == "35"
    assert equation.rotation["centerX"] == "555"
    assert equation.rotation["centerY"] == "666"
    assert shape.text == "NATIVE-SHAPE"
    assert shape.width == 3600
    assert shape.height == 1800
    assert shape.shape_comment == "NATIVE-SHAPE-COMMENT"
    assert shape.fill_color == "#ABCDEF"
    assert shape.line_color == "#123456"
    assert shape.layout["textWrap"] == "SQUARE"
    assert shape.layout["textFlow"] == "LEFT_ONLY"
    assert shape.layout["treatAsChar"] == "0"
    assert shape.layout["vertOffset"] == "21"
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
    assert ole.layout["textWrap"] == "SQUARE"
    assert ole.layout["textFlow"] == "LEFT_ONLY"
    assert ole.layout["treatAsChar"] == "0"
    assert ole.layout["vertOffset"] == "31"

    reopened_hwpx = document.to_hwpx_document()
    assert reopened_hwpx.equations()[0].script == "x+y=z"
    assert reopened_hwpx.equations()[0].size() == {"width": 3456, "height": 2100}
    assert reopened_hwpx.equations()[0].element.get("font") == "Batang"
    assert reopened_hwpx.equations()[0].shape_comment == "NATIVE-EQ"
    assert reopened_hwpx.equations()[0].layout()["textWrap"] == "SQUARE"
    assert reopened_hwpx.equations()[0].layout()["textFlow"] == "RIGHT_ONLY"
    assert reopened_hwpx.equations()[0].layout()["treatAsChar"] == "0"
    assert reopened_hwpx.equations()[0].layout()["vertOffset"] == "12"
    assert reopened_hwpx.equations()[0].out_margins() == {"left": 5, "right": 6, "top": 7, "bottom": 8}
    assert reopened_hwpx.equations()[0].rotation()["angle"] == "35"
    assert reopened_hwpx.equations()[0].rotation()["centerX"] == "555"
    assert reopened_hwpx.equations()[0].rotation()["centerY"] == "666"
    assert reopened_hwpx.shapes()[0].text == "NATIVE-SHAPE"
    assert reopened_hwpx.shapes()[0].shape_comment == "NATIVE-SHAPE-COMMENT"
    assert reopened_hwpx.shapes()[0].fill_style()["faceColor"] == "#ABCDEF"
    assert reopened_hwpx.shapes()[0].line_style()["color"] == "#123456"
    assert reopened_hwpx.shapes()[0].layout()["textWrap"] == "SQUARE"
    assert reopened_hwpx.shapes()[0].layout()["textFlow"] == "LEFT_ONLY"
    assert reopened_hwpx.shapes()[0].layout()["treatAsChar"] == "0"
    assert reopened_hwpx.oles()[0].binary_data() == b"NATIVE-OLE-DATA"
    assert reopened_hwpx.oles()[0].shape_comment == "NATIVE-OLE-COMMENT"
    assert reopened_hwpx.oles()[0].object_type == "LINK"
    assert reopened_hwpx.oles()[0].draw_aspect == "ICON"
    assert reopened_hwpx.oles()[0].has_moniker is True
    assert reopened_hwpx.oles()[0].line_style()["color"] == "#445566"
    assert reopened_hwpx.oles()[0].line_style()["width"] == "77"
    assert reopened_hwpx.oles()[0].layout()["textWrap"] == "SQUARE"
    assert reopened_hwpx.oles()[0].layout()["textFlow"] == "LEFT_ONLY"
    assert reopened_hwpx.oles()[0].layout()["treatAsChar"] == "0"


def test_hancom_document_pure_hwp_roundtrip_preserves_native_shape_and_ole_metadata(tmp_path: Path) -> None:
    image_bytes = HwpDocument.blank().binary_document().read_stream("BinData/BIN0001.bmp", decompress=False)
    document = HancomDocument.blank()
    equation = document.append_equation("a+b", width=3200, height=1800, font="Batang", shape_comment="IR-EQ")
    equation.layout = {
        "textWrap": "SQUARE",
        "textFlow": "RIGHT_ONLY",
        "treatAsChar": "0",
        "vertRelTo": "PAPER",
        "horzRelTo": "PAPER",
        "vertAlign": "TOP",
        "horzAlign": "LEFT",
        "vertOffset": "12",
        "horzOffset": "34",
    }
    equation.out_margins = {"left": 5, "right": 6, "top": 7, "bottom": 8}
    equation.rotation = {"angle": "35", "centerX": "555", "centerY": "666"}
    picture = document.append_picture("ir.bmp", image_bytes, extension="bmp", width=3600, height=2400)
    picture.shape_comment = "IR-PICTURE-COMMENT"
    picture.layout = {
        "textWrap": "SQUARE",
        "textFlow": "RIGHT_ONLY",
        "treatAsChar": "0",
        "vertRelTo": "PAPER",
        "horzRelTo": "PAPER",
        "vertAlign": "TOP",
        "horzAlign": "LEFT",
        "vertOffset": "101",
        "horzOffset": "202",
    }
    picture.out_margins = {"left": 13, "right": 24, "top": 35, "bottom": 46}
    picture.rotation = {"angle": "15", "centerX": "111", "centerY": "222", "rotateimage": "0"}
    picture.image_adjustment = {"bright": "5", "contrast": "6", "effect": "GRAY_SCALE", "alpha": "7"}
    picture.crop = {"left": 1, "right": 3599, "top": 2, "bottom": 2398}
    picture.line_color = "#223344"
    picture.line_width = 66
    shape = document.append_shape(
        kind="rect",
        text="IR-SHAPE",
        width=4100,
        height=1900,
        fill_color="#FEDCBA",
        line_color="#234567",
        shape_comment="IR-SHAPE-COMMENT",
    )
    shape.layout = {
        "textWrap": "SQUARE",
        "textFlow": "LEFT_ONLY",
        "treatAsChar": "0",
        "vertRelTo": "PAPER",
        "horzRelTo": "PAPER",
        "vertAlign": "TOP",
        "horzAlign": "LEFT",
        "vertOffset": "55",
        "horzOffset": "66",
    }
    shape.out_margins = {"left": 6, "right": 7, "top": 8, "bottom": 9}
    shape.rotation = {"angle": "25", "centerX": "333", "centerY": "444"}
    ole = document.append_ole(
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
    ole.layout = {
        "textWrap": "SQUARE",
        "textFlow": "LEFT_ONLY",
        "treatAsChar": "0",
        "vertRelTo": "PAPER",
        "horzRelTo": "PAPER",
        "vertAlign": "TOP",
        "horzAlign": "LEFT",
        "vertOffset": "77",
        "horzOffset": "88",
    }
    ole.out_margins = {"left": 9, "right": 10, "top": 11, "bottom": 12}
    ole.rotation = {"angle": "45", "centerX": "777", "centerY": "888"}
    ole.extent = {"x": 43000, "y": 14000}

    output_path = tmp_path / "hancom_native_full_parity.hwp"
    document.write_to_hwp(output_path)

    reopened = HancomDocument.read_hwp(output_path)
    equation = next(block for block in reopened.sections[0].blocks if isinstance(block, Equation))
    picture = next(block for block in reopened.sections[0].blocks if isinstance(block, Picture))
    shape = next(block for block in reopened.sections[0].blocks if isinstance(block, Shape))
    ole = next(block for block in reopened.sections[0].blocks if isinstance(block, Ole))

    assert equation.shape_comment == "IR-EQ"
    assert equation.layout["textWrap"] == "SQUARE"
    assert equation.layout["textFlow"] == "RIGHT_ONLY"
    assert equation.layout["treatAsChar"] == "0"
    assert equation.layout["vertOffset"] == "12"
    assert equation.out_margins == {"left": 5, "right": 6, "top": 7, "bottom": 8}
    assert equation.rotation["angle"] == "35"
    assert equation.rotation["centerX"] == "555"
    assert equation.rotation["centerY"] == "666"
    assert picture.shape_comment == "IR-PICTURE-COMMENT"
    assert picture.layout["textWrap"] == "SQUARE"
    assert picture.layout["textFlow"] == "RIGHT_ONLY"
    assert picture.layout["treatAsChar"] == "0"
    assert picture.out_margins == {"left": 13, "right": 24, "top": 35, "bottom": 46}
    assert picture.rotation["angle"] == "15"
    assert picture.rotation["centerX"] == "111"
    assert picture.rotation["centerY"] == "222"
    assert picture.image_adjustment["effect"] == "GRAY_SCALE"
    assert picture.crop == {"left": 1, "right": 3599, "top": 2, "bottom": 2398}
    assert picture.line_color == "#223344"
    assert picture.line_width == 66
    assert shape.shape_comment == "IR-SHAPE-COMMENT"
    assert shape.fill_color == "#FEDCBA"
    assert shape.line_color == "#234567"
    assert shape.layout["textWrap"] == "SQUARE"
    assert shape.layout["textFlow"] == "LEFT_ONLY"
    assert shape.layout["treatAsChar"] == "0"
    assert shape.out_margins == {"left": 6, "right": 7, "top": 8, "bottom": 9}
    assert shape.rotation["angle"] == "25"
    assert shape.rotation["centerX"] == "333"
    assert shape.rotation["centerY"] == "444"
    assert ole.shape_comment == "IR-OLE-COMMENT"
    assert ole.object_type == "LINK"
    assert ole.draw_aspect == "ICON"
    assert ole.has_moniker is True
    assert ole.eq_baseline == 9
    assert ole.line_color == "#556677"
    assert ole.line_width == 55
    assert ole.layout["textWrap"] == "SQUARE"
    assert ole.layout["textFlow"] == "LEFT_ONLY"
    assert ole.layout["treatAsChar"] == "0"
    assert ole.out_margins == {"left": 9, "right": 10, "top": 11, "bottom": 12}
    assert ole.rotation["angle"] == "45"
    assert ole.extent == {"x": 43000, "y": 14000}

    roundtrip_hwpx = tmp_path / "hancom_native_full_parity.hwpx"
    reopened.write_to_hwpx(roundtrip_hwpx)
    reopened_hwpx = HwpxDocument.open(roundtrip_hwpx)

    reopened_picture = reopened_hwpx.pictures()[0]
    reopened_shape = reopened_hwpx.shapes()[0]
    reopened_ole = reopened_hwpx.oles()[0]
    reopened_equation = reopened_hwpx.equations()[0]

    assert reopened_equation.shape_comment == "IR-EQ"
    assert reopened_equation.layout()["textWrap"] == "SQUARE"
    assert reopened_equation.layout()["textFlow"] == "RIGHT_ONLY"
    assert reopened_equation.layout()["treatAsChar"] == "0"
    assert reopened_equation.layout()["vertOffset"] == "12"
    assert reopened_equation.out_margins() == {"left": 5, "right": 6, "top": 7, "bottom": 8}
    assert reopened_equation.rotation()["angle"] == "35"
    assert reopened_equation.rotation()["centerX"] == "555"
    assert reopened_equation.rotation()["centerY"] == "666"
    assert reopened_picture.shape_comment == "IR-PICTURE-COMMENT"
    assert reopened_picture.layout()["textWrap"] == "SQUARE"
    assert reopened_picture.layout()["textFlow"] == "RIGHT_ONLY"
    assert reopened_picture.layout()["treatAsChar"] == "0"
    assert reopened_picture.out_margins() == {"left": 13, "right": 24, "top": 35, "bottom": 46}
    assert reopened_picture.rotation()["angle"] == "15"
    assert reopened_picture.rotation()["centerX"] == "111"
    assert reopened_picture.rotation()["centerY"] == "222"
    assert reopened_picture.image_adjustment()["effect"] == "GRAY_SCALE"
    assert reopened_picture.crop() == {"left": 1, "right": 3599, "top": 2, "bottom": 2398}
    assert reopened_picture.line_style()["color"] == "#223344"
    assert reopened_picture.line_style()["width"] == "66"
    assert reopened_shape.shape_comment == "IR-SHAPE-COMMENT"
    assert reopened_shape.layout()["textWrap"] == "SQUARE"
    assert reopened_shape.layout()["textFlow"] == "LEFT_ONLY"
    assert reopened_shape.layout()["treatAsChar"] == "0"
    assert reopened_shape.out_margins() == {"left": 6, "right": 7, "top": 8, "bottom": 9}
    assert reopened_shape.rotation()["angle"] == "25"
    assert reopened_shape.rotation()["centerX"] == "333"
    assert reopened_shape.rotation()["centerY"] == "444"
    assert reopened_ole.shape_comment == "IR-OLE-COMMENT"
    assert reopened_ole.layout()["textWrap"] == "SQUARE"
    assert reopened_ole.layout()["textFlow"] == "LEFT_ONLY"
    assert reopened_ole.layout()["treatAsChar"] == "0"
    assert reopened_ole.out_margins() == {"left": 9, "right": 10, "top": 11, "bottom": 12}
    assert reopened_ole.rotation()["angle"] == "45"
    assert reopened_ole.extent() == {"x": 43000, "y": 14000}


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


def test_hancom_document_preserves_richer_hwpx_table_picture_shape_and_ole_semantics(tmp_path: Path) -> None:
    image_bytes = HwpDocument.blank().binary_document().read_stream("BinData/BIN0001.bmp", decompress=False)
    source = HwpxDocument.blank()

    table = source.append_table(2, 2, cell_texts=[["A1", "A2"], ["B1", "B2"]])
    table.set_layout(
        text_wrap="SQUARE",
        treat_as_char=False,
        vert_rel_to="PAPER",
        horz_rel_to="PAPER",
        vert_align="TOP",
        horz_align="LEFT",
        vert_offset=12,
        horz_offset=34,
    )
    table.set_out_margins(left=10, right=20, top=30, bottom=40)
    table.set_in_margins(left=1, right=2, top=3, bottom=4)
    table.set_cell_spacing(17)
    table.set_table_border_fill_id(5)
    table.set_page_break("TABLE")
    table.set_repeat_header(False)
    table.cell(0, 1).set_border_fill_id(7)
    table.cell(1, 0).set_margins(left=11, right=22, top=33, bottom=44)
    table.cell(1, 0).set_vertical_align("BOTTOM")

    picture = source.append_picture("bridge.bmp", image_bytes, width=2400, height=1600)
    picture.shape_comment = "PICTURE-COMMENT"
    picture.set_layout(
        text_wrap="SQUARE",
        text_flow="RIGHT_ONLY",
        treat_as_char=False,
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
    picture.set_crop(left=1, right=2399, top=2, bottom=1598)
    picture.set_line_style(color="#112233", width=55, style="SOLID")

    shape = source.append_shape(kind="rect", text="IR-SHAPE", width=5000, height=2000, fill_color="#ABCDEF", line_color="#123456")
    shape.shape_comment = "SHAPE-COMMENT"
    shape.set_layout(
        text_wrap="SQUARE",
        text_flow="LEFT_ONLY",
        treat_as_char=False,
        vert_rel_to="PAPER",
        horz_rel_to="PAPER",
        vert_align="TOP",
        horz_align="LEFT",
        vert_offset=55,
        horz_offset=66,
    )
    shape.set_out_margins(left=6, right=7, top=8, bottom=9)
    shape.set_rotation(angle=25, center_x=333, center_y=444)
    shape.set_text_margins(left=4, right=5, top=6, bottom=7)

    ole = source.append_ole("bridge.ole", b"OLE-BRIDGE", width=42001, height=13501)
    ole.shape_comment = "OLE-COMMENT"
    ole.set_layout(
        text_wrap="SQUARE",
        text_flow="LEFT_ONLY",
        treat_as_char=False,
        vert_rel_to="PAPER",
        horz_rel_to="PAPER",
        vert_align="TOP",
        horz_align="LEFT",
        vert_offset=77,
        horz_offset=88,
    )
    ole.set_out_margins(left=9, right=10, top=11, bottom=12)
    ole.set_rotation(angle=45, center_x=777, center_y=888)
    ole.set_line_style(color="#445566", width=77, style="DASH")
    ole.set_object_metadata(object_type="LINK", draw_aspect="ICON", has_moniker=True, eq_baseline=12)
    ole.set_extent(x=43000, y=14000)

    source_path = tmp_path / "hancom_richer_hwpx_source.hwpx"
    source.save(source_path)

    document = HancomDocument.read_hwpx(source_path)
    table_block = next(block for block in document.sections[0].blocks if isinstance(block, Table))
    picture_block = next(block for block in document.sections[0].blocks if isinstance(block, Picture))
    shape_block = next(block for block in document.sections[0].blocks if isinstance(block, Shape))
    ole_block = next(block for block in document.sections[0].blocks if isinstance(block, Ole))

    assert table_block.cell_spacing == 17
    assert table_block.table_border_fill_id == 5
    assert table_block.table_margins == {"left": 1, "right": 2, "top": 3, "bottom": 4}
    assert table_block.out_margins == {"left": 10, "right": 20, "top": 30, "bottom": 40}
    assert table_block.page_break == "TABLE"
    assert table_block.repeat_header is False
    assert table_block.cell_border_fill_ids[(0, 1)] == 7
    assert table_block.cell_margins[(1, 0)] == {"left": 11, "right": 22, "top": 33, "bottom": 44}
    assert table_block.cell_vertical_aligns[(1, 0)] == "BOTTOM"
    assert table_block.layout["textWrap"] == "SQUARE"

    assert picture_block.shape_comment == "PICTURE-COMMENT"
    assert picture_block.layout["textWrap"] == "SQUARE"
    assert picture_block.layout["textFlow"] == "RIGHT_ONLY"
    assert picture_block.layout["treatAsChar"] == "0"
    assert picture_block.rotation["angle"] == "15"
    assert picture_block.rotation["centerX"] == "111"
    assert picture_block.rotation["centerY"] == "222"
    assert picture_block.image_adjustment["effect"] == "GRAY_SCALE"
    assert picture_block.crop == {"left": 1, "right": 2399, "top": 2, "bottom": 1598}
    assert picture_block.line_color == "#112233"
    assert picture_block.line_width == 55

    assert shape_block.kind == "rect"
    assert shape_block.shape_comment == "SHAPE-COMMENT"
    assert shape_block.layout["textWrap"] == "SQUARE"
    assert shape_block.layout["textFlow"] == "LEFT_ONLY"
    assert shape_block.layout["treatAsChar"] == "0"
    assert shape_block.out_margins == {"left": 6, "right": 7, "top": 8, "bottom": 9}
    assert shape_block.rotation["angle"] == "25"
    assert shape_block.rotation["centerX"] == "333"
    assert shape_block.rotation["centerY"] == "444"
    assert shape_block.text_margins == {"left": 4, "right": 5, "top": 6, "bottom": 7}

    assert ole_block.shape_comment == "OLE-COMMENT"
    assert ole_block.object_type == "LINK"
    assert ole_block.draw_aspect == "ICON"
    assert ole_block.has_moniker is True
    assert ole_block.line_color == "#445566"
    assert ole_block.line_width == 77
    assert ole_block.layout["textWrap"] == "SQUARE"
    assert ole_block.layout["textFlow"] == "LEFT_ONLY"
    assert ole_block.layout["treatAsChar"] == "0"
    assert ole_block.extent == {"x": 43000, "y": 14000}

    roundtrip_path = tmp_path / "hancom_richer_hwpx_roundtrip.hwpx"
    document.write_to_hwpx(roundtrip_path)
    reopened = HwpxDocument.open(roundtrip_path)

    reopened_table = reopened.tables()[0]
    reopened_picture = reopened.pictures()[0]
    reopened_shape = reopened.shapes()[0]
    reopened_ole = reopened.oles()[0]

    assert reopened_table.cell_spacing == 17
    assert reopened_table.table_border_fill_id == 5
    assert reopened_table.in_margins() == {"left": 1, "right": 2, "top": 3, "bottom": 4}
    assert reopened_table.out_margins() == {"left": 10, "right": 20, "top": 30, "bottom": 40}
    assert reopened_table.page_break == "TABLE"
    assert reopened_table.repeat_header is False
    assert reopened_table.cell(0, 1).border_fill_id == 7
    assert reopened_table.cell(1, 0).margins == {"left": 11, "right": 22, "top": 33, "bottom": 44}
    assert reopened_table.cell(1, 0).vertical_align == "BOTTOM"

    assert reopened_picture.shape_comment == "PICTURE-COMMENT"
    assert reopened_picture.layout()["textWrap"] == "SQUARE"
    assert reopened_picture.layout()["textFlow"] == "RIGHT_ONLY"
    assert reopened_picture.layout()["treatAsChar"] == "0"
    assert reopened_picture.rotation()["angle"] == "15"
    assert reopened_picture.image_adjustment()["effect"] == "GRAY_SCALE"
    assert reopened_picture.rotation()["centerX"] == "111"
    assert reopened_picture.rotation()["centerY"] == "222"
    assert reopened_picture.crop() == {"left": 1, "right": 2399, "top": 2, "bottom": 1598}
    assert reopened_picture.line_style()["color"] == "#112233"
    assert reopened_picture.line_style()["width"] == "55"

    assert reopened_shape.kind == "rect"
    assert reopened_shape.shape_comment == "SHAPE-COMMENT"
    assert reopened_shape.layout()["textWrap"] == "SQUARE"
    assert reopened_shape.layout()["textFlow"] == "LEFT_ONLY"
    assert reopened_shape.layout()["treatAsChar"] == "0"
    assert reopened_shape.out_margins() == {"left": 6, "right": 7, "top": 8, "bottom": 9}
    assert reopened_shape.rotation()["angle"] == "25"
    assert reopened_shape.rotation()["centerX"] == "333"
    assert reopened_shape.rotation()["centerY"] == "444"
    assert reopened_shape.text_margins() == {"left": 4, "right": 5, "top": 6, "bottom": 7}

    assert reopened_ole.shape_comment == "OLE-COMMENT"
    assert reopened_ole.object_type == "LINK"
    assert reopened_ole.draw_aspect == "ICON"
    assert reopened_ole.has_moniker is True
    assert reopened_ole.layout()["textWrap"] == "SQUARE"
    assert reopened_ole.layout()["textFlow"] == "LEFT_ONLY"
    assert reopened_ole.layout()["treatAsChar"] == "0"
    assert reopened_ole.line_style()["color"] == "#445566"
    assert reopened_ole.line_style()["width"] == "77"
    assert reopened_ole.extent() == {"x": 43000, "y": 14000}


def test_hancom_document_preserves_connect_line_kind_through_hwp_bridge(tmp_path: Path) -> None:
    source = HwpxDocument.blank()
    connect_line = source.append_shape(kind="connectLine", width=3200, height=1400, line_color="#102030", shape_comment="CONNECT-LINE")
    connect_line.set_layout(
        text_wrap="SQUARE",
        vert_rel_to="PAPER",
        horz_rel_to="PAPER",
        vert_align="TOP",
        horz_align="LEFT",
        vert_offset=12,
        horz_offset=34,
    )

    source_path = tmp_path / "hancom_connect_line_source.hwpx"
    source.save(source_path)

    document = HancomDocument.read_hwpx(source_path)
    shape_block = next(block for block in document.sections[0].blocks if isinstance(block, Shape))
    assert shape_block.kind == "connectLine"
    assert "start" in shape_block.specific_fields
    assert "end" in shape_block.specific_fields

    output_path = tmp_path / "hancom_connect_line_roundtrip.hwp"
    document.write_to_hwp(output_path)

    reopened_hwp = HwpDocument.open(output_path)
    assert reopened_hwp.lines() == []
    assert len(reopened_hwp.connect_lines()) == 1
    reopened_connect_line = reopened_hwp.connect_lines()[0]
    assert isinstance(reopened_connect_line, HwpConnectLineShapeObject)
    assert reopened_connect_line.kind == "connectLine"
    assert reopened_connect_line.shape_comment == "CONNECT-LINE"

    reopened_ir = HancomDocument.read_hwp(output_path)
    reopened_shape_block = next(block for block in reopened_ir.sections[0].blocks if isinstance(block, Shape))
    assert reopened_shape_block.kind == "connectLine"
    reopened_hwpx = reopened_ir.to_hwpx_document()
    assert any(shape.kind == "connectLine" for shape in reopened_hwpx.shapes())


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
    document.append_header("PURE-HEADER", apply_page_type="EVEN")
    document.append_footer("PURE-FOOTER", apply_page_type="ODD")
    document.append_auto_number(number=7, number_type="PAGE", kind="newNum")

    output_path = tmp_path / "hancom_ir_controls.hwp"
    document.write_to_hwp(output_path)

    reopened = HwpDocument.open(output_path).to_hwpx_document()
    assert reopened.headers()[0].text == "PURE-HEADER"
    assert reopened.headers()[0].apply_page_type == "EVEN"
    assert reopened.footers()[0].text == "PURE-FOOTER"
    assert reopened.footers()[0].apply_page_type == "ODD"
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
