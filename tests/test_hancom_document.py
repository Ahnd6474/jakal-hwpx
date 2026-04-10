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
    assert "HWP" in conversions
