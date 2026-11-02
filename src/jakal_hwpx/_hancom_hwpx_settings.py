"""HWPX page, section, numbering, border, and note settings."""
from __future__ import annotations

from typing import Any

from lxml import etree

from ._hancom_models import (
    HeaderFooter,
    SectionSettings,
)
from ._hancom_styles import _coerce_bool
from .document import HwpxDocument
from .namespaces import (
    NS,
    qname,
)


def _extract_page_border_fills(section) -> list[dict[str, str | int]]:
    fills: list[dict[str, str | int]] = []
    for node in section.findall(".//hp:pageBorderFill"):
        offset = node.find("./hp:offset")
        fills.append(
            {
                "type": node.get_attr("type") or "BOTH",
                "borderFillIDRef": node.get_attr("borderFillIDRef") or "0",
                "textBorder": node.get_attr("textBorder") or "PAPER",
                "headerInside": node.get_attr("headerInside") or "0",
                "footerInside": node.get_attr("footerInside") or "0",
                "fillArea": node.get_attr("fillArea") or "PAPER",
                "left": int((offset.get_attr("left") if offset is not None else "0") or 0),
                "right": int((offset.get_attr("right") if offset is not None else "0") or 0),
                "top": int((offset.get_attr("top") if offset is not None else "0") or 0),
                "bottom": int((offset.get_attr("bottom") if offset is not None else "0") or 0),
            }
        )
    return fills


def _extract_page_numbers_with_positions(section) -> tuple[list[dict[str, str]], list[tuple[int | None, int | None]]]:
    from ._hancom_hwpx_read import _source_paragraph_position_for_element

    page_numbers: list[dict[str, str]] = []
    positions: list[tuple[int | None, int | None]] = []
    for node in section.findall(".//hp:pageNum"):
        if not _is_direct_section_paragraph_control(node.element):
            continue
        page_numbers.append(
            {
                "pos": node.get_attr("pos") or "BOTTOM_CENTER",
                "formatType": node.get_attr("formatType") or "DIGIT",
                "sideChar": node.get_attr("sideChar") or "-",
            }
        )
        positions.append(_source_paragraph_position_for_element(section, node.element))
    return page_numbers, positions


def _extract_page_numbers(section) -> list[dict[str, str]]:
    page_numbers, _positions = _extract_page_numbers_with_positions(section)
    return page_numbers


def _is_direct_section_paragraph_control(element: etree._Element) -> bool:
    parent = element.getparent()
    if parent is None or etree.QName(parent).localname != "ctrl":
        return False
    run = parent.getparent()
    if run is None or etree.QName(run).localname != "run":
        return False
    paragraph = run.getparent()
    if paragraph is None or etree.QName(paragraph).localname != "p":
        return False
    section_root = paragraph.getparent()
    return section_root is not None and etree.QName(section_root).localname == "sec"


def _extract_note_pr(section, kind: str) -> dict[str, object]:
    node = section.find(f".//hp:{kind}")
    if node is None:
        return {}
    auto_num = node.find("./hp:autoNumFormat")
    note_line = node.find("./hp:noteLine")
    note_spacing = node.find("./hp:noteSpacing")
    numbering = node.find("./hp:numbering")
    placement = node.find("./hp:placement")
    return {
        "autoNumFormat": {
            "type": auto_num.get_attr("type") if auto_num is not None else "DIGIT",
            "userChar": auto_num.get_attr("userChar") if auto_num is not None else "",
            "prefixChar": auto_num.get_attr("prefixChar") if auto_num is not None else "",
            "suffixChar": auto_num.get_attr("suffixChar") if auto_num is not None else "",
            "supscript": auto_num.get_attr("supscript") if auto_num is not None else "0",
        },
        "noteLine": {
            "length": note_line.get_attr("length") if note_line is not None else "-1",
            "type": note_line.get_attr("type") if note_line is not None else "SOLID",
            "width": note_line.get_attr("width") if note_line is not None else "0.12 mm",
            "color": note_line.get_attr("color") if note_line is not None else "#000000",
        },
        "noteSpacing": {
            "betweenNotes": note_spacing.get_attr("betweenNotes") if note_spacing is not None else "283",
            "belowLine": note_spacing.get_attr("belowLine") if note_spacing is not None else "567",
            "aboveLine": note_spacing.get_attr("aboveLine") if note_spacing is not None else "850",
        },
        "numbering": {
            "type": numbering.get_attr("type") if numbering is not None else "CONTINUOUS",
            "newNum": numbering.get_attr("newNum") if numbering is not None else "1",
        },
        "placement": {
            "place": placement.get_attr("place") if placement is not None else ("END_OF_DOCUMENT" if kind == "endNotePr" else "EACH_COLUMN"),
            "beneathText": placement.get_attr("beneathText") if placement is not None else "0",
        },
    }


def _extract_line_number_shape(section) -> dict[str, str]:
    node = section.find(".//hp:lineNumberShape")
    if node is None:
        return {}
    return {
        key: node.get_attr(key) or "0"
        for key in ("restartType", "countBy", "distance", "startNumber")
    }


def _extract_section_settings(section) -> SectionSettings:
    settings = section.document.section_settings(section.section_index())
    memo_shape_ids = {
        memo_shape.memo_shape_id
        for memo_shape in section.document.memo_shapes()
        if memo_shape.memo_shape_id is not None
    }
    memo_shape_id = settings.memo_shape_id if settings.memo_shape_id in memo_shape_ids else None
    page_numbers, page_number_positions = _extract_page_numbers_with_positions(section)
    return SectionSettings(
        page_width=settings.page_width,
        page_height=settings.page_height,
        landscape=settings.landscape,
        margins=settings.margins(),
        page_border_fills=_extract_page_border_fills(section),
        visibility=settings.visibility(),
        grid=settings.grid(),
        start_numbers=settings.start_numbers(),
        page_numbers=page_numbers,
        page_number_positions=page_number_positions,
        footnote_pr=_extract_note_pr(section, "footNotePr"),
        endnote_pr=_extract_note_pr(section, "endNotePr"),
        line_number_shape=_extract_line_number_shape(section),
        memo_shape_id=memo_shape_id,
    )


def _apply_section_settings(
    document: HwpxDocument,
    section_index: int,
    settings: SectionSettings,
    *,
    apply_page_numbers: bool = True,
) -> None:
    target = document.section_settings(section_index)
    # Hancom needs a column definition alongside section properties; otherwise
    # it can reset the first section's custom page setup to A4 on open.
    section = document.sections[section_index]
    if not section.root_element.xpath("./hp:p[1]//hp:colPr", namespaces=NS):
        column_control = etree.Element(qname("hp", "ctrl"))
        column_properties = etree.SubElement(column_control, qname("hp", "colPr"))
        for name, value in {
            "id": "", "type": "NEWSPAPER", "layout": "LEFT",
            "colCount": "1", "sameSz": "1", "sameGap": "0",
        }.items():
            column_properties.set(name, value)
        target.element.addnext(column_control)
        section.mark_modified()
    if settings.page_width is not None or settings.page_height is not None or settings.landscape is not None:
        width, height, landscape = settings.page_width, settings.page_height, settings.landscape
        # The shared model's LANDSCAPE alias uses oriented dimensions (as HWP
        # and DOCX do). HWPX stores paper dimensions before rotating the page.
        if landscape == "LANDSCAPE":
            width, height, landscape = height, width, "NARROWLY"
        elif landscape == "PORTRAIT":
            landscape = "WIDELY"
        target.set_page_size(
            width=width,
            height=height,
            landscape=landscape,
        )
    if settings.margins:
        target.set_margins(**settings.margins)
    if settings.page_border_fills:
        _apply_page_border_fills(document, section_index, settings.page_border_fills)
    if settings.visibility:
        target.set_visibility(
            hide_first_header=_coerce_bool(settings.visibility.get("hideFirstHeader")),
            hide_first_footer=_coerce_bool(settings.visibility.get("hideFirstFooter")),
            hide_first_master_page=_coerce_bool(settings.visibility.get("hideFirstMasterPage")),
            border=settings.visibility.get("border"),
            fill=settings.visibility.get("fill"),
            hide_first_page_num=_coerce_bool(settings.visibility.get("hideFirstPageNum")),
            hide_first_empty_line=_coerce_bool(settings.visibility.get("hideFirstEmptyLine")),
            show_line_number=_coerce_bool(settings.visibility.get("showLineNumber")),
        )
    if settings.grid:
        target.set_grid(
            line_grid=settings.grid.get("lineGrid"),
            char_grid=settings.grid.get("charGrid"),
            wonggoji_format=_coerce_bool(settings.grid.get("wonggojiFormat")),
        )
    if settings.start_numbers:
        target.set_start_numbers(
            page_starts_on=settings.start_numbers.get("pageStartsOn"),
            page=settings.start_numbers.get("page"),
            pic=settings.start_numbers.get("pic"),
            tbl=settings.start_numbers.get("tbl"),
            equation=settings.start_numbers.get("equation"),
        )
    if settings.memo_shape_id is not None:
        target.set_memo_shape_id(settings.memo_shape_id)
    if settings.line_number_shape:
        _apply_line_number_shape(document, section_index, settings.line_number_shape)
    if settings.footnote_pr or settings.endnote_pr:
        _apply_note_settings(
            document,
            section_index,
            footnote_pr=settings.footnote_pr or None,
            endnote_pr=settings.endnote_pr or None,
        )
    if apply_page_numbers and settings.page_numbers:
        _apply_page_numbers(document, section_index, settings.page_numbers)


def _apply_page_border_fills(
    document: HwpxDocument,
    section_index: int,
    page_border_fills: list[dict[str, str | int]],
) -> None:
    sec_pr = document.section_xml(section_index).find(".//hp:secPr")
    if sec_pr is None:
        return
    for node in list(sec_pr.element.xpath("./hp:pageBorderFill", namespaces=NS)):
        node.getparent().remove(node)
    for border_fill in page_border_fills:
        border_type = str(border_fill.get("type", "BOTH"))
        node = etree.SubElement(sec_pr.element, f"{{{NS['hp']}}}pageBorderFill")
        node.set("type", border_type)
        node.set("borderFillIDRef", str(border_fill.get("borderFillIDRef", "0")))
        node.set("textBorder", str(border_fill.get("textBorder", "PAPER")))
        node.set("headerInside", str(border_fill.get("headerInside", "0")))
        node.set("footerInside", str(border_fill.get("footerInside", "0")))
        node.set("fillArea", str(border_fill.get("fillArea", "PAPER")))
        offset = etree.SubElement(node, f"{{{NS['hp']}}}offset")
        offset.set("left", str(border_fill.get("left", 0)))
        offset.set("right", str(border_fill.get("right", 0)))
        offset.set("top", str(border_fill.get("top", 0)))
        offset.set("bottom", str(border_fill.get("bottom", 0)))
    document.sections[section_index].mark_modified()


def _apply_line_number_shape(
    document: HwpxDocument,
    section_index: int,
    line_number_shape: dict[str, str],
) -> None:
    node = document.section_xml(section_index).find(".//hp:lineNumberShape")
    if node is None:
        return
    for key in ("restartType", "countBy", "distance", "startNumber"):
        value = line_number_shape.get(key)
        if value is not None:
            node.set_attr(key, value)


def _apply_note_settings(
    document: HwpxDocument,
    section_index: int,
    *,
    footnote_pr: dict[str, object] | None = None,
    endnote_pr: dict[str, object] | None = None,
) -> None:
    sec_pr = document.section_xml(section_index).find(".//hp:secPr")
    if sec_pr is None:
        return
    for kind in ("footNotePr", "endNotePr"):
        for node in list(sec_pr.element.xpath(f"./hp:{kind}", namespaces=NS)):
            node.getparent().remove(node)
    if footnote_pr:
        _append_note_pr(sec_pr.element, "footNotePr", footnote_pr)
    if endnote_pr:
        _append_note_pr(sec_pr.element, "endNotePr", endnote_pr)
    document.sections[section_index].mark_modified()


def _append_note_pr(sec_pr: etree._Element, kind: str, note_pr: dict[str, Any]) -> None:
    from ._hancom_conversion_common import _xml_safe_scalar

    note_node = etree.SubElement(sec_pr, f"{{{NS['hp']}}}{kind}")
    auto_num = dict(note_pr.get("autoNumFormat", {}))
    auto_num_node = etree.SubElement(note_node, f"{{{NS['hp']}}}autoNumFormat")
    auto_num_node.set("type", _xml_safe_scalar(auto_num.get("type", "DIGIT")))
    auto_num_node.set("userChar", _xml_safe_scalar(auto_num.get("userChar", "")))
    auto_num_node.set("prefixChar", _xml_safe_scalar(auto_num.get("prefixChar", "")))
    auto_num_node.set("suffixChar", _xml_safe_scalar(auto_num.get("suffixChar", "")))
    auto_num_node.set("supscript", _xml_safe_scalar(auto_num.get("supscript", "0")))
    note_line = dict(note_pr.get("noteLine", {}))
    note_line_node = etree.SubElement(note_node, f"{{{NS['hp']}}}noteLine")
    note_line_node.set("length", str(note_line.get("length", "-1")))
    note_line_node.set("type", str(note_line.get("type", "SOLID")))
    note_line_node.set("width", str(note_line.get("width", "0.12 mm")))
    note_line_node.set("color", str(note_line.get("color", "#000000")))
    note_spacing = dict(note_pr.get("noteSpacing", {}))
    note_spacing_node = etree.SubElement(note_node, f"{{{NS['hp']}}}noteSpacing")
    note_spacing_node.set("betweenNotes", str(note_spacing.get("betweenNotes", "283")))
    note_spacing_node.set("belowLine", str(note_spacing.get("belowLine", "567")))
    note_spacing_node.set("aboveLine", str(note_spacing.get("aboveLine", "850")))
    numbering = dict(note_pr.get("numbering", {}))
    numbering_node = etree.SubElement(note_node, f"{{{NS['hp']}}}numbering")
    numbering_node.set("type", str(numbering.get("type", "CONTINUOUS")))
    numbering_node.set("newNum", str(numbering.get("newNum", "1")))
    placement = dict(note_pr.get("placement", {}))
    placement_node = etree.SubElement(note_node, f"{{{NS['hp']}}}placement")
    placement_node.set("place", str(placement.get("place", "END_OF_DOCUMENT" if kind == "endNotePr" else "EACH_COLUMN")))
    placement_node.set("beneathText", str(placement.get("beneathText", "0")))


def _apply_page_numbers(
    document: HwpxDocument,
    section_index: int,
    page_numbers: list[dict[str, str]],
) -> None:
    from ._hancom_hwpx_write import _append_control_host_paragraph

    section_xml = document.section_xml(section_index)
    for node in section_xml.findall(".//hp:pageNum"):
        parent = node.element.getparent()
        if parent is not None:
            parent.remove(node.element)
    for page_number in page_numbers:
        paragraph_index = _append_control_host_paragraph(document, section_index)
        document.append_control_xml(_page_number_xml(page_number), section_index=section_index, paragraph_index=paragraph_index)


def _apply_page_numbers_to_source_paragraphs(
    document: HwpxDocument,
    section_index: int,
    settings: SectionSettings,
) -> None:
    from ._hancom_hwpx_write import _append_control_host_paragraph

    section_xml = document.section_xml(section_index)
    for node in section_xml.findall(".//hp:pageNum"):
        parent = node.element.getparent()
        if parent is not None:
            parent.remove(node.element)

    paragraph_count = len(document.sections[section_index].paragraphs())
    for index, page_number in enumerate(settings.page_numbers):
        paragraph_index: int | None = None
        if index < len(settings.page_number_positions):
            paragraph_index = settings.page_number_positions[index][0]
        if paragraph_index is None or paragraph_index >= paragraph_count:
            paragraph_index = _append_control_host_paragraph(document, section_index)
            paragraph_count += 1
        document.append_control_xml(_page_number_xml(page_number), section_index=section_index, paragraph_index=paragraph_index)


def _page_number_xml(page_number: dict[str, str]) -> str:
    return (
        '<hp:pageNum xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" '
        f'pos="{page_number.get("pos", "BOTTOM_CENTER")}" '
        f'formatType="{page_number.get("formatType", "DIGIT")}" '
        f'sideChar="{page_number.get("sideChar", "-")}"/>'
    )


def _apply_header_footer_blocks(document: HwpxDocument, section_index: int, blocks: list[HeaderFooter]) -> None:
    from ._hancom_hwpx_write import _append_control_host_paragraph, _append_header_footer_block_to_hwpx

    for block in blocks:
        paragraph_index = _append_control_host_paragraph(document, section_index)
        _append_header_footer_block_to_hwpx(document, section_index, paragraph_index, block)
