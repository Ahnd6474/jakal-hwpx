from __future__ import annotations

import zipfile
from pathlib import Path

from lxml import etree
import pytest
from jakal_hwpx import HwpxDocument


def find_sample_with_section_xpath(files: list[Path], expression: str, *, allow_missing: bool = False) -> Path | None:
    namespaces = {"hp": "http://www.hancom.co.kr/hwpml/2011/paragraph"}
    for path in files:
        with zipfile.ZipFile(path) as zf:
            for name in [n for n in zf.namelist() if n.startswith("Contents/section") and n.endswith(".xml")]:
                data = zf.read(name)
                if not data.lstrip().startswith(b"<"):
                    continue
                root = etree.fromstring(data)
                if root.xpath(expression, namespaces=namespaces):
                    return path
    if allow_missing:
        return None
    raise LookupError(f"No sample matched xpath: {expression}")


def test_header_and_footer_edit_roundtrip(valid_hwpx_files: list[Path], tmp_path: Path) -> None:
    header_source = find_sample_with_section_xpath(valid_hwpx_files, ".//hp:header")
    assert header_source is not None
    header_doc = HwpxDocument.open(header_source)
    header_block = header_doc.headers()[0]
    header_block.set_text("HEADER_EDITED_BY_TEST")
    header_out = tmp_path / "header_edit.hwpx"
    header_doc.save(header_out)
    reopened_header = HwpxDocument.open(header_out)
    assert "HEADER_EDITED_BY_TEST" in reopened_header.headers()[0].text

    footer_source = find_sample_with_section_xpath(valid_hwpx_files, ".//hp:footer", allow_missing=True)
    if footer_source is None:
        pytest.skip("No footer sample is available in the current corpus.")
    footer_doc = HwpxDocument.open(footer_source)
    footer_block = footer_doc.footers()[0]
    footer_block.set_text("FOOTER_EDITED_BY_TEST")
    footer_out = tmp_path / "footer_edit.hwpx"
    footer_doc.save(footer_out)
    reopened_footer = HwpxDocument.open(footer_out)
    assert "FOOTER_EDITED_BY_TEST" in reopened_footer.footers()[0].text


def test_table_picture_paragraph_and_style_edit_roundtrip(valid_hwpx_files: list[Path], tmp_path: Path) -> None:
    source = find_sample_with_section_xpath(
        valid_hwpx_files,
        ".//hp:tbl[.//hp:t[normalize-space()]] and .//hp:pic and .//hp:header",
    )
    document = HwpxDocument.open(source)

    table = next(table for table in document.tables() if any(cell.text.strip() for cell in table.cells()))
    cell = next(cell for cell in table.cells() if cell.text.strip())
    cell.set_text("TABLE_CELL_EDITED")

    picture = document.pictures()[0]
    original_comment = picture.shape_comment
    picture.shape_comment = (original_comment + " [edited]").strip()
    picture.replace_binary(picture.binary_data())

    style = document.styles()[0]
    para_style = document.paragraph_styles()[0]
    char_style = document.character_styles()[0]

    style.set_name("테스트스타일")
    para_style.set_alignment(horizontal="CENTER")
    para_style.set_line_spacing(180)
    char_style.set_text_color("#112233")
    char_style.set_height(1100)

    document.append_paragraph("PARAGRAPH_EDITED_BY_TEST", section_index=0)
    document.apply_style_to_paragraph(
        0,
        0,
        style_id=style.style_id,
        para_pr_id=para_style.style_id,
        char_pr_id=char_style.style_id,
    )

    output_path = tmp_path / "feature_edit.hwpx"
    document.save(output_path)

    reopened = HwpxDocument.open(output_path)
    reopened.validate()

    reopened_cell = next(
        cell
        for table in reopened.tables()
        for cell in table.cells()
        if cell.text.strip() == "TABLE_CELL_EDITED"
    )
    assert reopened_cell.text == "TABLE_CELL_EDITED"

    reopened_picture = reopened.pictures()[0]
    assert reopened_picture.shape_comment.endswith("[edited]")

    reopened_style = reopened.styles()[0]
    reopened_para_style = reopened.paragraph_styles()[0]
    reopened_char_style = reopened.character_styles()[0]
    assert reopened_style.name == "테스트스타일"
    assert reopened_para_style.alignment_horizontal == "CENTER"
    assert reopened_para_style.line_spacing == "180"
    assert reopened_char_style.text_color == "#112233"
    assert reopened_char_style.height == "1100"
    assert "PARAGRAPH_EDITED_BY_TEST" in reopened.get_document_text()

    first_paragraph = reopened.sections[0].paragraphs()[0]
    assert first_paragraph.get_attr("styleIDRef") == reopened_style.style_id
    assert first_paragraph.get_attr("paraPrIDRef") == reopened_para_style.style_id
