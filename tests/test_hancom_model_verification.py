from __future__ import annotations

import re
import tempfile
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from jakal_hwpx import HwpxDocument


def _find_sample_with_section_xpath(files: list[Path], expression: str) -> Path:
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
    raise LookupError(f"No sample matched xpath: {expression}")


def _hancom_model_root(document_path: Path) -> etree._Element:
    win32com = pytest.importorskip("win32com.client")
    try:
        hwp = win32com.gencache.EnsureDispatch("HWPFrame.HwpObject")
    except Exception as exc:  # noqa: BLE001
        pytest.skip(f"HWPFrame.HwpObject is not available: {exc}")

    try:
        try:
            hwp.XHwpWindows.Item(0).Visible = False
        except Exception:
            pass
        module_name = HwpxDocument.discover_hancom_security_module_name()
        if module_name:
            hwp.RegisterModule("FilePathCheckDLL", module_name)
        hwp.Open(str(document_path), "", "")
        hwpml = hwp.GetTextFile("HWPML2X", "")
        hwpml = re.sub(r"^<\?xml[^>]*\?>", "", hwpml)
        return etree.fromstring(hwpml.encode("utf-8"))
    finally:
        try:
            hwp.XHwpDocuments.Active_XHwpDocument.Close(False)
        except Exception:
            pass
        hwp.Quit()


def _paragraph_text(paragraph: etree._Element) -> str:
    return "".join((node.text or "") for node in paragraph.xpath(".//CHAR"))


def _section_paragraphs(root: etree._Element, section_index: int) -> list[etree._Element]:
    return list(root.xpath(f"./BODY/SECTION[{section_index + 1}]/P"))


def _first_paragraph_with_text(root: etree._Element, text: str) -> etree._Element:
    for paragraph in root.xpath(".//P"):
        if _paragraph_text(paragraph) == text:
            return paragraph
    raise LookupError(text)


def test_hancom_model_header_footer_picture_positions(valid_hwpx_files: list[Path], tmp_path: Path) -> None:
    source = _find_sample_with_section_xpath(valid_hwpx_files, ".//hp:header and .//hp:footer and .//hp:pic")
    document = HwpxDocument.open(source)
    document.headers()[0].set_text("HEADER_EDITED")
    document.footers()[0].set_text("FOOTER_EDITED")
    document.pictures()[0].shape_comment = "PICTURE_EDITED"

    output_path = tmp_path / "hancom_header_footer_picture.hwpx"
    document.save(output_path)
    root = _hancom_model_root(output_path)

    header_texts = root.xpath(".//HEADER//CHAR/text()")
    footer_texts = root.xpath(".//FOOTER//CHAR/text()")
    picture_comments = root.xpath(".//PICTURE/SHAPEOBJECT/SHAPECOMMENT/text()")

    assert "HEADER_EDITED" in header_texts
    assert "FOOTER_EDITED" in footer_texts
    assert "PICTURE_EDITED" in picture_comments


def test_hancom_model_style_table_and_positions(valid_hwpx_files: list[Path], tmp_path: Path) -> None:
    source = _find_sample_with_section_xpath(valid_hwpx_files, ".//hp:tbl[hp:tr/hp:tc[2]]")
    document = HwpxDocument.open(source)

    style_id = next((item.style_id for item in document.styles() if item.style_id not in {None, "0"}), document.styles()[0].style_id)
    para_pr_id = next(
        (item.style_id for item in document.paragraph_styles() if item.style_id not in {None, "0"}),
        document.paragraph_styles()[0].style_id,
    )
    char_pr_id = next(
        (item.style_id for item in document.character_styles() if item.style_id not in {None, "0"}),
        document.character_styles()[0].style_id,
    )

    document.append_paragraph("BATCH_TARGET_TEXT", section_index=0)
    updated = document.apply_style_batch(
        section_index=0,
        text_contains="BATCH_TARGET_TEXT",
        style_id=style_id,
        para_pr_id=para_pr_id,
        char_pr_id=char_pr_id,
    )
    assert updated == 1

    table = next(table for table in document.tables() if table.column_count >= 2)
    new_row = table.append_row()
    new_row[0].set_text("LEFT")
    new_row[1].set_text("RIGHT")
    merged = table.merge_cells(new_row[0].row, new_row[0].column, new_row[0].row, new_row[1].column)
    merged.set_text("MERGED")

    output_path = tmp_path / "hancom_style_table.hwpx"
    document.save(output_path)
    root = _hancom_model_root(output_path)

    section0_paragraphs = _section_paragraphs(root, 0)
    batch_paragraph = _first_paragraph_with_text(root, "BATCH_TARGET_TEXT")
    assert batch_paragraph is section0_paragraphs[-1]
    assert batch_paragraph.get("Style") == style_id
    assert batch_paragraph.get("ParaShape") == para_pr_id
    assert batch_paragraph.xpath("./TEXT/@CharShape")[0] == char_pr_id

    merged_cells = root.xpath(".//TABLE/ROW/CELL[@ColSpan='2'][.//CHAR='MERGED']")
    assert merged_cells
    assert merged_cells[0].get("ColAddr") == "0"


def test_hancom_model_fields_land_in_expected_paragraph_order(sample_hwpx_path: Path, tmp_path: Path) -> None:
    document = HwpxDocument.open(sample_hwpx_path)
    document.append_bookmark("python_anchor")
    document.append_hyperlink("https://example.com/hwpx", display_text="Example Link")
    document.append_mail_merge_field("customer_name", display_text="CUSTOMER_NAME")
    document.append_calculation_field("40+2", display_text="42")
    document.append_cross_reference("python_anchor", display_text="Anchor Ref")

    output_path = tmp_path / "hancom_fields.hwpx"
    document.save(output_path)
    root = _hancom_model_root(output_path)

    section0_paragraphs = _section_paragraphs(root, 0)
    last_five = section0_paragraphs[-5:]

    assert last_five[0].xpath("./TEXT/BOOKMARK/@Name") == ["python_anchor"]
    assert _paragraph_text(last_five[0]) == ""

    hyperlink = last_five[1].xpath("./TEXT/FIELDBEGIN")[0]
    assert hyperlink.get("Type") == "Hyperlink"
    assert _paragraph_text(last_five[1]) == "Example Link"

    mailmerge = last_five[2].xpath("./TEXT/FIELDBEGIN")[0]
    assert mailmerge.get("Type") == "Mailmerge"
    assert mailmerge.get("Name") == "customer_name"
    assert _paragraph_text(last_five[2]) == "CUSTOMER_NAME"

    formula = last_five[3].xpath("./TEXT/FIELDBEGIN")[0]
    assert formula.get("Type") == "Formula"
    assert formula.get("Command") == "40+2"
    assert _paragraph_text(last_five[3]) == "42"

    crossref = last_five[4].xpath("./TEXT/FIELDBEGIN")[0]
    assert crossref.get("Type") == "Crossref"
    assert crossref.get("Name") == "python_anchor"
    assert _paragraph_text(last_five[4]) == "Anchor Ref"


def test_hancom_model_notes_autonum_equation_and_shape(valid_hwpx_files: list[Path], tmp_path: Path) -> None:
    note_source = _find_sample_with_section_xpath(valid_hwpx_files, ".//hp:footNote and .//hp:endNote")
    note_doc = HwpxDocument.open(note_source)
    notes = note_doc.notes()
    notes[0].set_text("FOOTNOTE_EDITED")
    notes[1].set_text("ENDNOTE_EDITED")
    note_output = tmp_path / "hancom_notes.hwpx"
    note_doc.save(note_output)
    note_root = _hancom_model_root(note_output)
    assert "FOOTNOTE_EDITED" in note_root.xpath(".//FOOTNOTE//CHAR/text()")
    assert "ENDNOTE_EDITED" in note_root.xpath(".//ENDNOTE//CHAR/text()")

    newnum_source = _find_sample_with_section_xpath(valid_hwpx_files, ".//hp:newNum")
    newnum_doc = HwpxDocument.open(newnum_source)
    newnum_doc.auto_numbers()[0].set_number(9)
    newnum_output = tmp_path / "hancom_newnum.hwpx"
    newnum_doc.save(newnum_output)
    newnum_root = _hancom_model_root(newnum_output)
    assert "9" in newnum_root.xpath(".//NEWNUM/@Number")

    equation_source = _find_sample_with_section_xpath(valid_hwpx_files, ".//hp:equation")
    equation_doc = HwpxDocument.open(equation_source)
    equation_doc.equations()[0].script = "x=1+2"
    equation_output = tmp_path / "hancom_equation.hwpx"
    equation_doc.save(equation_output)
    equation_root = _hancom_model_root(equation_output)
    assert equation_root.xpath(".//EQUATION/SCRIPT/text()") == ["x=1+2"]

    shape_source = _find_sample_with_section_xpath(valid_hwpx_files, ".//hp:textart or .//hp:rect or .//hp:line")
    shape_doc = HwpxDocument.open(shape_source)
    shape = shape_doc.shapes()[0]
    shape.shape_comment = "SHAPE_EDITED"
    if shape.kind == "textart":
        shape.set_text("TEXTART_EDITED")
    shape_output = tmp_path / "hancom_shape.hwpx"
    shape_doc.save(shape_output)
    shape_root = _hancom_model_root(shape_output)
    assert "SHAPE_EDITED" in shape_root.xpath(".//SHAPECOMMENT/text()")
    if shape.kind == "textart":
        assert "TEXTART_EDITED" in shape_root.xpath(".//TEXTART//CHAR/text()")


def test_hancom_probe_path_reports_open_state(sample_hwpx_path: Path) -> None:
    pytest.importorskip("win32com.client")
    probe = HwpxDocument.probe_path_in_hancom(sample_hwpx_path)

    assert probe.opened is True
    assert probe.status == "opened"
    assert probe.hwpml_retrieved is True
    assert probe.hwpml_length is not None
