from __future__ import annotations

from pathlib import Path

from lxml import etree

from jakal_hwpx import HwpxDocument
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
