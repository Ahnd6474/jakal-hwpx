from __future__ import annotations

import zipfile
from pathlib import Path

from lxml import etree

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


def test_section_settings_style_batch_and_advanced_table(valid_hwpx_files: list[Path], tmp_path: Path) -> None:
    source = find_sample_with_section_xpath(valid_hwpx_files, ".//hp:tbl[hp:tr/hp:tc[2]]")
    document = HwpxDocument.open(source)

    settings = document.section_settings(0)
    original_width = settings.page_width or 0
    settings.set_page_size(width=original_width + 10)
    settings.set_margins(left=7777, right=6666)

    document.append_paragraph("BATCH_TARGET_TEXT", section_index=0)
    updated = document.apply_style_batch(
        section_index=0,
        text_contains="BATCH_TARGET_TEXT",
        style_id=document.styles()[0].style_id,
        para_pr_id=document.paragraph_styles()[0].style_id,
        char_pr_id=document.character_styles()[0].style_id,
    )
    assert updated == 1

    table = next(table for table in document.tables() if table.column_count >= 2)
    new_row = table.append_row()
    new_row[0].set_text("ROW_APPEND_LEFT")
    new_row[1].set_text("ROW_APPEND_RIGHT")
    merged = table.merge_cells(new_row[0].row, new_row[0].column, new_row[0].row, new_row[1].column)
    merged.set_text("ROW_MERGED")

    output_path = tmp_path / "section_style_table.hwpx"
    document.save(output_path)
    reopened = HwpxDocument.open(output_path)

    reopened_settings = reopened.section_settings(0)
    assert reopened_settings.page_width == original_width + 10
    assert reopened_settings.margins()["left"] == 7777
    assert reopened_settings.margins()["right"] == 6666
    assert "BATCH_TARGET_TEXT" in reopened.get_document_text()
    merged_cell = next(cell for table in reopened.tables() for cell in table.cells() if cell.text == "ROW_MERGED")
    assert merged_cell.col_span == 2


def test_notes_bookmarks_fields_numbers_equations_and_shapes(valid_hwpx_files: list[Path], tmp_path: Path) -> None:
    note_source = find_sample_with_section_xpath(valid_hwpx_files, ".//hp:footNote and .//hp:endNote", allow_missing=True)
    if note_source is not None:
        note_doc = HwpxDocument.open(note_source)
        notes = note_doc.notes()
        notes[0].set_text("FOOTNOTE_EDITED")
        notes[1].set_text("ENDNOTE_EDITED")
        note_out = tmp_path / "notes.hwpx"
        note_doc.save(note_out)
        reopened_notes = HwpxDocument.open(note_out)
        assert any(note.text == "FOOTNOTE_EDITED" for note in reopened_notes.notes())
        assert any(note.text == "ENDNOTE_EDITED" for note in reopened_notes.notes())

    bookmark_source = find_sample_with_section_xpath(valid_hwpx_files, ".//hp:bookmark", allow_missing=True)
    if bookmark_source is not None:
        bookmark_doc = HwpxDocument.open(bookmark_source)
        bookmark_doc.bookmarks()[0].rename("renamed_bookmark")
    else:
        bookmark_doc = HwpxDocument.blank()
        bookmark_doc.append_bookmark("original_bookmark")
        bookmark_doc.bookmarks()[0].rename("renamed_bookmark")
    bookmark_out = tmp_path / "bookmark.hwpx"
    bookmark_doc.save(bookmark_out)
    reopened_bookmark = HwpxDocument.open(bookmark_out)
    assert reopened_bookmark.bookmarks()[0].name == "renamed_bookmark"

    numbering_source = find_sample_with_section_xpath(valid_hwpx_files, ".//hp:autoNum or .//hp:newNum")
    assert numbering_source is not None
    numbering_doc = HwpxDocument.open(numbering_source)
    numbering_doc.auto_numbers()[0].set_number(9)
    numbering_out = tmp_path / "numbering.hwpx"
    numbering_doc.save(numbering_out)
    reopened_numbering = HwpxDocument.open(numbering_out)
    assert any(item.number == "9" for item in reopened_numbering.auto_numbers())

    field_source = find_sample_with_section_xpath(valid_hwpx_files, ".//hp:fieldBegin[@type='HYPERLINK']", allow_missing=True)
    if field_source is not None:
        field_doc = HwpxDocument.open(field_source)
        hyperlink = field_doc.hyperlinks()[0]
        hyperlink.set_hyperlink_target("https://example.com/test")
    else:
        field_doc = HwpxDocument.blank()
        field_doc.append_hyperlink("https://example.com/test", display_text="Example")
    field_out = tmp_path / "field.hwpx"
    field_doc.save(field_out)
    reopened_field = HwpxDocument.open(field_out)
    assert reopened_field.hyperlinks()[0].hyperlink_target == "https://example.com/test"

    equation_source = find_sample_with_section_xpath(valid_hwpx_files, ".//hp:equation")
    assert equation_source is not None
    equation_doc = HwpxDocument.open(equation_source)
    equation = equation_doc.equations()[0]
    equation.script = "x=1+2"
    equation_out = tmp_path / "equation.hwpx"
    equation_doc.save(equation_out)
    reopened_equation = HwpxDocument.open(equation_out)
    assert reopened_equation.equations()[0].script == "x=1+2"

    shape_source = find_sample_with_section_xpath(valid_hwpx_files, ".//hp:textart or .//hp:rect or .//hp:line")
    assert shape_source is not None
    shape_doc = HwpxDocument.open(shape_source)
    shape = shape_doc.shapes()[0]
    original_comment = shape.shape_comment
    shape.shape_comment = (original_comment + " edited-shape").strip()
    if shape.kind == "textart":
        shape.set_text("TEXTART_EDITED")
    shape_out = tmp_path / "shape.hwpx"
    shape_doc.save(shape_out)
    reopened_shape = HwpxDocument.open(shape_out)
    reopened_first_shape = reopened_shape.shapes()[0]
    assert reopened_first_shape.shape_comment.endswith("edited-shape")
    if reopened_first_shape.kind == "textart":
        assert reopened_first_shape.text == "TEXTART_EDITED"


def test_field_creation_helpers_and_reference_integrity(sample_hwpx_path: Path, tmp_path: Path) -> None:
    document = HwpxDocument.open(sample_hwpx_path)

    bookmark = document.append_bookmark("python_anchor")
    hyperlink = document.append_hyperlink("https://example.com/hwpx", display_text="Example Link")
    mail_merge = document.append_mail_merge_field("customer_name", display_text="CUSTOMER_NAME")
    calculation = document.append_calculation_field("1+2", display_text="3")
    cross_ref = document.append_cross_reference(bookmark.name or "python_anchor", display_text="Anchor Ref")

    hyperlink.set_display_text("Example Link Updated")
    mail_merge.configure_mail_merge("customer_name", display_text="CUSTOMER_NAME_UPDATED")
    calculation.configure_calculation("40+2", display_text="42")
    cross_ref.configure_cross_reference("python_anchor", display_text="Anchor Ref Updated")

    output_path = tmp_path / "field_creation.hwpx"
    document.save(output_path)
    reopened = HwpxDocument.open(output_path)

    assert any(item.name == "python_anchor" for item in reopened.bookmarks())
    assert any(item.hyperlink_target == "https://example.com/hwpx" for item in reopened.hyperlinks())
    assert any(item.display_text == "Example Link Updated" for item in reopened.hyperlinks())
    assert any(item.name == "customer_name" and item.display_text == "CUSTOMER_NAME_UPDATED" for item in reopened.mail_merge_fields())
    assert any(item.get_parameter("Expression") == "40+2" and item.display_text == "42" for item in reopened.calculation_fields())
    assert any(item.get_parameter("BookmarkName") == "python_anchor" for item in reopened.cross_references())
    assert reopened.reference_validation_errors() == []


def test_validation_layers(valid_hwpx_files: list[Path], tmp_path: Path) -> None:
    source = valid_hwpx_files[0]
    document = HwpxDocument.open(source)
    version_schema = tmp_path / "version.xsd"
    version_schema.write_text(
        """<?xml version="1.0" encoding="UTF-8"?>
<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema"
           targetNamespace="http://www.hancom.co.kr/hwpml/2011/version"
           xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version"
           elementFormDefault="qualified">
  <xs:element name="HCFVersion" type="xs:anyType"/>
</xs:schema>
""",
        encoding="utf-8",
    )

    assert document.xml_validation_errors() == []
    assert document.schema_validation_errors({"version.xml": version_schema}) == []
    assert document.reference_validation_errors() == []
    assert document.save_reopen_validation_errors() == []
