from __future__ import annotations

from pathlib import Path

from lxml import etree
import pytest

from jakal_hwpx import HwpxDocument, HwpxValidationError, ValidationIssue
from jakal_hwpx.elements import _replace_text, _set_text

from conftest import find_sample_with_section_xpath


NS = {"hp": "http://www.hancom.co.kr/hwpml/2011/paragraph"}


def test_table_cell_text_edit_invalidates_linesegarray(valid_hwpx_files: list[Path], tmp_path: Path) -> None:
    source = find_sample_with_section_xpath(valid_hwpx_files, ".//hp:tc//hp:p[hp:linesegarray and .//hp:t]")
    document = HwpxDocument.open(source)

    target_cell = next(
        cell
        for table in document.tables()
        for cell in table.cells()
        if cell.element.xpath(".//hp:p[hp:linesegarray and .//hp:t]", namespaces=NS)
    )

    target_cell.set_text("TABLE_CELL_LAYOUT_REFRESH")

    assert not target_cell.element.xpath(".//hp:p/hp:linesegarray", namespaces=NS)

    output_path = tmp_path / "cell_layout_refresh.hwpx"
    document.save(output_path)

    reopened = HwpxDocument.open(output_path)
    reopened_cell = next(
        cell
        for table in reopened.tables()
        for cell in table.cells()
        if cell.text == "TABLE_CELL_LAYOUT_REFRESH"
    )
    assert not reopened_cell.element.xpath(".//hp:p/hp:linesegarray", namespaces=NS)


def test_replace_text_invalidates_linesegarray_in_table_cells(valid_hwpx_files: list[Path], tmp_path: Path) -> None:
    source = find_sample_with_section_xpath(valid_hwpx_files, ".//hp:tc//hp:p[hp:linesegarray and .//hp:t]")
    document = HwpxDocument.open(source)

    target_paragraph = next(
        paragraph
        for section in document.sections
        for paragraph in section.root_element.xpath(".//hp:tc//hp:p[hp:linesegarray and .//hp:t]", namespaces=NS)
        if paragraph.xpath(".//hp:t/text()", namespaces=NS)
    )
    original_text = "".join(target_paragraph.xpath(".//hp:t/text()", namespaces=NS))
    replacement = f"{original_text}_UPDATED"

    changed = document.replace_text(original_text, replacement, count=1)

    assert changed == 1
    assert not target_paragraph.xpath("./hp:linesegarray", namespaces=NS)

    output_path = tmp_path / "replace_text_layout_refresh.hwpx"
    document.save(output_path)

    reopened = HwpxDocument.open(output_path)
    reopened_paragraph = next(
        paragraph
        for section in reopened.sections
        for paragraph in section.root_element.xpath(".//hp:tc//hp:p", namespaces=NS)
        if "".join(paragraph.xpath(".//hp:t/text()", namespaces=NS)) == replacement
    )
    assert not reopened_paragraph.xpath("./hp:linesegarray", namespaces=NS)


def test_append_bookmark_invalidates_existing_paragraph_linesegarray() -> None:
    document = HwpxDocument.blank()
    paragraph = document.sections[0].root_element.xpath("./hp:p", namespaces=NS)[0]

    text_run = etree.SubElement(paragraph, "{http://www.hancom.co.kr/hwpml/2011/paragraph}run")
    text_run.set("charPrIDRef", "0")
    etree.SubElement(text_run, "{http://www.hancom.co.kr/hwpml/2011/paragraph}t").text = "Anchor paragraph"
    etree.SubElement(paragraph, "{http://www.hancom.co.kr/hwpml/2011/paragraph}linesegarray")

    document.append_bookmark("bookmark-anchor", section_index=0, paragraph_index=0)

    assert paragraph.xpath("./hp:run/hp:ctrl/hp:bookmark[@name='bookmark-anchor']", namespaces=NS)
    assert not paragraph.xpath("./hp:linesegarray", namespaces=NS)


def test_field_set_display_text_insertion_invalidates_linesegarray() -> None:
    document = HwpxDocument.blank()
    field = document.append_hyperlink("https://example.com", section_index=0, paragraph_index=0)
    paragraph = field.element.xpath("ancestor::hp:p[1]", namespaces=NS)[0]

    etree.SubElement(paragraph, "{http://www.hancom.co.kr/hwpml/2011/paragraph}linesegarray")

    field.set_display_text("Inserted display text")

    assert field.display_text == "Inserted display text"
    assert not paragraph.xpath("./hp:linesegarray", namespaces=NS)


def test_set_paragraph_text_preserves_section_properties(valid_hwpx_files: list[Path], tmp_path: Path) -> None:
    source = find_sample_with_section_xpath(valid_hwpx_files, "./hp:p[hp:run/hp:secPr]")
    document = HwpxDocument.open(source)

    paragraphs = document.sections[0].root_element.xpath("./hp:p", namespaces=NS)
    paragraph_index = next(
        index for index, paragraph in enumerate(paragraphs) if paragraph.xpath("./hp:run/hp:secPr", namespaces=NS)
    )

    document.set_paragraph_text(0, paragraph_index, "SECTION_SETTINGS_PRESERVED")

    paragraph = document.sections[0].root_element.xpath("./hp:p", namespaces=NS)[paragraph_index]
    assert paragraph.xpath("./hp:run/hp:secPr", namespaces=NS)
    assert not paragraph.xpath("./hp:linesegarray", namespaces=NS)

    output_path = tmp_path / "section_settings_preserved.hwpx"
    document.save(output_path)

    reopened = HwpxDocument.open(output_path)
    reopened_paragraph = reopened.sections[0].root_element.xpath("./hp:p", namespaces=NS)[paragraph_index]
    assert reopened_paragraph.xpath("./hp:run/hp:secPr", namespaces=NS)
    assert "SECTION_SETTINGS_PRESERVED" in reopened.get_document_text()


def test_set_paragraph_text_preserves_controls() -> None:
    root = etree.fromstring(
        """
        <hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
                xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
          <hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
            <hp:run charPrIDRef="0"><hp:secPr /></hp:run>
            <hp:run charPrIDRef="0"><hp:ctrl><hp:tbl /></hp:ctrl></hp:run>
            <hp:run charPrIDRef="0"><hp:t>VISIBLE</hp:t></hp:run>
            <hp:linesegarray />
          </hp:p>
        </hs:sec>
        """
    )

    section = HwpxDocument.blank().sections[0]
    section._root = root

    section.set_paragraph_text(0, "UPDATED")

    paragraph = section.root_element.xpath("./hp:p", namespaces=NS)[0]
    assert paragraph.xpath("./hp:run/hp:secPr", namespaces=NS)
    assert paragraph.xpath("./hp:run/hp:ctrl/hp:tbl", namespaces=NS)
    assert "UPDATED" in "".join(paragraph.xpath(".//hp:t/text()", namespaces=NS))
    assert not paragraph.xpath("./hp:linesegarray", namespaces=NS)


def test_header_footer_set_text_preserves_controls() -> None:
    document = HwpxDocument.blank()
    section = document.sections[0]
    block = etree.fromstring(
        """
        <hp:header xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" applyPageType="BOTH">
          <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">
            <hp:p id="10" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
              <hp:run charPrIDRef="0"><hp:ctrl><hp:newNum num="1" numType="PAGE" /></hp:ctrl></hp:run>
              <hp:run charPrIDRef="0"><hp:t>Header</hp:t></hp:run>
            </hp:p>
          </hp:subList>
        </hp:header>
        """
    )

    from jakal_hwpx.elements import HeaderFooterXml

    HeaderFooterXml(document, section, block).set_text("Edited header")

    assert block.xpath(".//hp:newNum", namespaces=NS)
    assert document.control_preservation_validation_errors() == []


def test_table_cell_set_text_preserves_controls() -> None:
    document = HwpxDocument.blank()
    section = document.sections[0]
    table_root = etree.fromstring(
        """
        <hp:tbl xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" rowCnt="1" colCnt="1">
          <hp:tr>
            <hp:tc>
              <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">
                <hp:p id="20" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
                  <hp:run charPrIDRef="0"><hp:ctrl><hp:fieldBegin id="11" type="HYPERLINK" fieldid="22" /></hp:ctrl></hp:run>
                  <hp:run charPrIDRef="0"><hp:t>Cell</hp:t></hp:run>
                  <hp:run charPrIDRef="0"><hp:ctrl><hp:fieldEnd beginIDRef="11" fieldid="22" /></hp:ctrl></hp:run>
                </hp:p>
              </hp:subList>
              <hp:cellAddr colAddr="0" rowAddr="0" />
              <hp:cellSpan colSpan="1" rowSpan="1" />
            </hp:tc>
          </hp:tr>
        </hp:tbl>
        """
    )

    from jakal_hwpx.elements import TableXml

    table = TableXml(document, section, table_root)
    table.cell(0, 0).set_text("Edited cell")

    assert table_root.xpath(".//hp:fieldBegin", namespaces=NS)
    assert table_root.xpath(".//hp:fieldEnd", namespaces=NS)
    assert document.control_preservation_validation_errors() == []


def test_shape_set_text_preserves_drawtext_controls() -> None:
    document = HwpxDocument.blank()
    section = document.sections[0]
    shape_root = etree.fromstring(
        """
        <hp:rect xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
          <hp:drawText>
            <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">
              <hp:p id="30" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
                <hp:run charPrIDRef="0"><hp:ctrl><hp:bookmark name="anchor" /></hp:ctrl></hp:run>
                <hp:run charPrIDRef="0"><hp:t>Shape text</hp:t></hp:run>
              </hp:p>
            </hp:subList>
          </hp:drawText>
        </hp:rect>
        """
    )

    from jakal_hwpx.elements import ShapeXml

    ShapeXml(document, section, shape_root).set_text("Edited shape")

    assert shape_root.xpath(".//hp:bookmark", namespaces=NS)
    assert document.control_preservation_validation_errors() == []


def test_validate_reports_control_preservation_failures(tmp_path: Path) -> None:
    document = HwpxDocument.blank()
    field = document.append_hyperlink("https://example.com", display_text="Example", section_index=0)
    paragraph = field.element.xpath("ancestor::hp:p[1]", namespaces=NS)[0]
    paragraph_index = next(
        index for index, node in enumerate(document.sections[0].root_element.xpath("./hp:p", namespaces=NS)) if node is paragraph
    )

    document.set_paragraph_text(0, paragraph_index, "Still safe")

    run_with_field_end = paragraph.xpath("./hp:run[hp:ctrl/hp:fieldEnd]", namespaces=NS)[0]
    paragraph.remove(run_with_field_end)

    errors = document.control_preservation_validation_errors()
    assert all(isinstance(error, ValidationIssue) for error in errors)
    assert any(error.part_path == "Contents/section0.xml" and error.paragraph_index is not None for error in errors)
    assert any("ctrl/fieldEnd" in error.message for error in errors)

    with pytest.raises(HwpxValidationError) as exc_info:
        document.save(tmp_path / "should_fail.hwpx", validate=True)

    assert all(isinstance(error, ValidationIssue) for error in exc_info.value.errors)
    assert all(error.code for error in exc_info.value.errors)
    assert exc_info.value.to_dicts()[0]["code"] == exc_info.value.errors[0].code
    assert "Contents/section0.xml paragraph" in str(exc_info.value)
    assert "ctrl/fieldEnd" in str(exc_info.value)


def test_append_row_rejects_template_rows_with_controls() -> None:
    document = HwpxDocument.blank()
    section = document.sections[0]
    table_root = etree.fromstring(
        """
        <hp:tbl xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" rowCnt="1" colCnt="1">
          <hp:tr>
            <hp:tc>
              <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">
                <hp:p id="40" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
                  <hp:run charPrIDRef="0"><hp:ctrl><hp:bookmark name="row-anchor" /></hp:ctrl></hp:run>
                  <hp:run charPrIDRef="0"><hp:t>Cell</hp:t></hp:run>
                </hp:p>
              </hp:subList>
              <hp:cellAddr colAddr="0" rowAddr="0" />
              <hp:cellSpan colSpan="1" rowSpan="1" />
            </hp:tc>
          </hp:tr>
        </hp:tbl>
        """
    )

    from jakal_hwpx.elements import TableXml

    table = TableXml(document, section, table_root)
    with pytest.raises(ValueError) as exc_info:
        table.append_row()

    assert "preserved controls" in str(exc_info.value)
    assert "ctrl/bookmark" in str(exc_info.value)


def test_append_paragraph_rejects_controlled_template_index() -> None:
    document = HwpxDocument.blank()
    section = document.sections[0]
    paragraph = section.root_element.xpath("./hp:p", namespaces=NS)[0]

    bookmark_run = etree.SubElement(paragraph, "{http://www.hancom.co.kr/hwpml/2011/paragraph}run")
    bookmark_run.set("charPrIDRef", "0")
    ctrl = etree.SubElement(bookmark_run, "{http://www.hancom.co.kr/hwpml/2011/paragraph}ctrl")
    bookmark = etree.SubElement(ctrl, "{http://www.hancom.co.kr/hwpml/2011/paragraph}bookmark")
    bookmark.set("name", "template-anchor")
    text_run = etree.SubElement(paragraph, "{http://www.hancom.co.kr/hwpml/2011/paragraph}run")
    text_run.set("charPrIDRef", "0")
    etree.SubElement(text_run, "{http://www.hancom.co.kr/hwpml/2011/paragraph}t").text = "Template"
    section.mark_modified()

    with pytest.raises(ValueError) as exc_info:
        document.append_paragraph("New paragraph", section_index=0, template_index=0)

    assert "preserved controls" in str(exc_info.value)
    assert "ctrl/bookmark" in str(exc_info.value)


def test_set_text_honors_newline_paragraphs() -> None:
    element = etree.fromstring(
        """
        <hp:tc xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
          <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">
            <hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
              <hp:run charPrIDRef="0"><hp:t>one</hp:t></hp:run>
              <hp:linesegarray />
            </hp:p>
            <hp:p id="1" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
              <hp:run charPrIDRef="0"><hp:t>two</hp:t></hp:run>
              <hp:linesegarray />
            </hp:p>
          </hp:subList>
        </hp:tc>
        """
    )

    _set_text(element, "alpha\nbeta\ngamma")

    paragraphs = element.xpath(".//hp:p", namespaces=NS)
    assert ["".join(paragraph.xpath(".//hp:t/text()", namespaces=NS)) for paragraph in paragraphs] == ["alpha", "beta", "gamma"]
    assert not element.xpath(".//hp:linesegarray", namespaces=NS)


def test_set_text_redistributes_single_line_text_across_existing_paragraphs() -> None:
    element = etree.fromstring(
        """
        <hp:tc xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
          <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">
            <hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
              <hp:run charPrIDRef="0"><hp:t>[ ] Parent benefit (name:     )</hp:t></hp:run>
              <hp:linesegarray />
            </hp:p>
            <hp:p id="1" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
              <hp:run charPrIDRef="0"><hp:t>[ ] Childcare allowance (name:     ), ([ ] Home [ ] Disabled [ ] Rural)</hp:t></hp:run>
              <hp:linesegarray />
            </hp:p>
            <hp:p id="2" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
              <hp:run charPrIDRef="0"><hp:t>[ ] Full-day service (name:   ), [ ] Part-time service (name:    )</hp:t></hp:run>
              <hp:linesegarray />
            </hp:p>
          </hp:subList>
        </hp:tc>
        """
    )

    _set_text(
        element,
        "[ ] Parent benefit (name: )[X] Childcare allowance (name: KIM), ([X] Home [ ] Disabled [ ] Rural)[ ] Full-day service (name: ), [ ] Part-time service (name: )",
    )

    paragraphs = element.xpath(".//hp:p", namespaces=NS)
    assert ["".join(paragraph.xpath(".//hp:t/text()", namespaces=NS)) for paragraph in paragraphs] == [
        "[ ] Parent benefit (name: )",
        "[X] Childcare allowance (name: KIM), ([X] Home [ ] Disabled [ ] Rural)",
        "[ ] Full-day service (name: ), [ ] Part-time service (name: )",
    ]
    assert not element.xpath(".//hp:linesegarray", namespaces=NS)


def test_replace_text_invalidates_linesegarray() -> None:
    paragraph = etree.fromstring(
        """
        <hp:p xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
          <hp:run charPrIDRef="0"><hp:t>alpha beta gamma</hp:t></hp:run>
          <hp:linesegarray />
        </hp:p>
        """
    )

    replaced = _replace_text(paragraph, "beta", "beta_LAYOUT_REFRESH", count=1)

    assert replaced == 1
    assert "".join(paragraph.xpath(".//hp:t/text()", namespaces=NS)) == "alpha beta_LAYOUT_REFRESH gamma"
    assert not paragraph.xpath("./hp:linesegarray", namespaces=NS)
