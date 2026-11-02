"""HWPX paragraph grouping, nested block serialization, and field output."""
from __future__ import annotations

import json

from lxml import etree

from ._hancom_models import (
    _HWPX_CHART_CARRIER_PREFIX,
    AutoNumber,
    Bookmark,
    Chart,
    Equation,
    Field,
    Form,
    HancomBlock,
    HancomSection,
    HeaderFooter,
    Hyperlink,
    Memo,
    Note,
    Ole,
    Paragraph,
    Picture,
    SectionSettings,
    Shape,
    Table,
)
from ._hwpx_element_helpers import _set_text_node_content
from .document import HwpxDocument
from .namespaces import (
    NS,
    qname,
)


def _build_hwpx_form_parameters(block: Form) -> dict[str, str]:
    parameters: dict[str, str] = {
        "FormType": block.form_type,
        "Checked": "1" if block.checked else "0",
        "Editable": "1" if block.editable else "0",
        "Locked": "1" if block.locked else "0",
    }
    if block.value is not None:
        parameters["Value"] = block.value
    if block.items:
        parameters["Items"] = json.dumps([str(value) for value in block.items], ensure_ascii=True, separators=(",", ":"))
    if block.placeholder:
        parameters["Placeholder"] = block.placeholder
    return parameters


def _build_hwpx_memo_parameters(block: Memo) -> dict[str, str]:
    parameters: dict[str, str] = {
        "Visible": "1" if block.visible else "0",
    }
    if block.text:
        parameters["Text"] = block.text
    if block.author:
        parameters["Author"] = block.author
    if block.memo_id:
        parameters["MemoId"] = block.memo_id
    if block.anchor_id:
        parameters["AnchorId"] = block.anchor_id
    if block.order is not None:
        parameters["Order"] = str(block.order)
    return parameters


def _build_hwpx_chart_carrier_comment(block: Chart) -> str:
    payload = {
        "title": block.title,
        "chartType": block.chart_type,
        "categories": [str(value) for value in block.categories],
        "series": [dict(value) for value in block.series],
        "dataRef": block.data_ref,
        "legendVisible": bool(block.legend_visible),
        "shapeComment": block.shape_comment,
    }
    return _HWPX_CHART_CARRIER_PREFIX + json.dumps(payload, ensure_ascii=True, separators=(",", ":"))


def _reset_hwpx_section(document: HwpxDocument, section_index: int) -> None:
    paragraphs = document.sections[section_index].paragraphs()
    while len(paragraphs) > 1:
        document.delete_paragraph(section_index, 1)
        paragraphs = document.sections[section_index].paragraphs()
    document.set_paragraph_text(section_index, 0, "")


def _section_uses_source_paragraphs(section: HancomSection) -> bool:
    return any(
        getattr(block, "source_paragraph_index", None) is not None
        for block in [*section.header_footer_blocks, *section.blocks]
    )


def _write_grouped_section_to_hwpx(document: HwpxDocument, section_index: int, section: HancomSection) -> None:
    from ._hancom_conversion_common import _is_owned_text_shadow_paragraph
    from ._hancom_to_hwpx import _append_non_paragraph_block_to_hwpx

    paragraph_map: dict[int, int] = {}
    styled_paragraphs: set[int] = set()
    ordered_items = [
        block
        for _, block in sorted(
            enumerate([*section.header_footer_blocks, *section.blocks]),
            key=lambda item: (
                getattr(item[1], "source_paragraph_index", 10**9),
                getattr(item[1], "source_order", item[0]),
                item[0],
            ),
        )
    ]

    index = 0
    while index < len(ordered_items):
        block = ordered_items[index]
        next_block = ordered_items[index + 1] if index + 1 < len(ordered_items) else None
        if _is_owned_text_shadow_paragraph(block, next_block):
            index += 1
            continue

        source_paragraph_index = getattr(block, "source_paragraph_index", None)
        if source_paragraph_index is None:
            paragraph_index = _append_control_host_paragraph(document, section_index)
        else:
            paragraph_index = _ensure_hwpx_source_paragraph(
                document,
                section_index,
                paragraph_map,
                int(source_paragraph_index),
            )

        if isinstance(block, HeaderFooter):
            _append_header_footer_block_to_hwpx(document, section_index, paragraph_index, block)
        elif isinstance(block, Paragraph):
            _apply_hwpx_paragraph_style_block(
                document, section_index, paragraph_index, block,
                apply_character_style=paragraph_index not in styled_paragraphs,
            )
            styled_paragraphs.add(paragraph_index)
            if block.text:
                _append_text_to_hwpx_paragraph(
                    document,
                    section_index,
                    paragraph_index,
                    block.text,
                    char_pr_id=_resolve_hwpx_char_pr_id(document, block),
            )
            index += 1
            continue
        else:
            _append_non_paragraph_block_to_hwpx(document, section_index, paragraph_index, block)
        index += 1


def _ensure_hwpx_source_paragraph(
    document: HwpxDocument,
    section_index: int,
    paragraph_map: dict[int, int],
    source_paragraph_index: int,
) -> int:
    if source_paragraph_index in paragraph_map:
        return paragraph_map[source_paragraph_index]
    if not paragraph_map:
        paragraph_map[source_paragraph_index] = 0
        return 0
    target_index = _append_control_host_paragraph(document, section_index)
    paragraph_map[source_paragraph_index] = target_index
    return target_index


def _append_text_to_hwpx_paragraph(
    document: HwpxDocument,
    section_index: int,
    paragraph_index: int,
    text: str,
    *,
    char_pr_id: str | None = None,
) -> None:
    paragraph = document._paragraph_element(section_index, paragraph_index)
    run = document._append_run(paragraph, char_pr_id=char_pr_id)
    text_node = etree.SubElement(run, qname("hp", "t"))
    text_node.text = text
    document.sections[section_index].mark_modified()


def _is_supported_nested_hwpx_block(value: object) -> bool:
    return isinstance(value, (Paragraph, Bookmark, Hyperlink, Field, AutoNumber, Table, Picture, Equation, Note, Shape, Ole, Chart, Form, Memo))


def _build_hwpx_nested_sublist(document: HwpxDocument, blocks: list[HancomBlock], *, vertical_align: str = "CENTER") -> etree._Element:
    from ._hancom_hwpx_read import _sort_hwpx_nested_blocks_by_source_order
    from ._hancom_to_hwpx import _append_non_paragraph_block_to_hwpx

    sublist = etree.Element(qname("hp", "subList"))
    sublist.set("id", "")
    sublist.set("textDirection", "HORIZONTAL")
    sublist.set("lineWrap", "BREAK")
    sublist.set("vertAlign", vertical_align)
    sublist.set("linkListIDRef", "0")
    sublist.set("linkListNextIDRef", "0")
    sublist.set("textWidth", "0")
    sublist.set("textHeight", "0")
    sublist.set("hasTextRef", "0")
    sublist.set("hasNumRef", "0")

    grouped: dict[int, list[HancomBlock]] = {}
    for block in blocks:
        grouped.setdefault(int(getattr(block, "nested_paragraph_index", 0)), []).append(block)
    if not grouped:
        grouped[0] = []

    # Build on attached staging paragraphs so normal object writers can create
    # package relationships and globally unique control IDs, including nested
    # tables and pictures. Detach only after the complete sublist exists.
    section = document.sections[0]
    staged = []
    try:
        for _paragraph_index, paragraph_blocks in sorted(grouped.items()):
            paragraph_index = _append_control_host_paragraph(document, 0)
            paragraph = document._paragraph_element(0, paragraph_index)
            staged.append(paragraph)
            first = next((block for block in paragraph_blocks if isinstance(block, Paragraph)), None)
            if first is not None:
                _apply_hwpx_paragraph_style_block(document, 0, paragraph_index, first)
            for block in _sort_hwpx_nested_blocks_by_source_order(paragraph_blocks):
                if isinstance(block, Paragraph):
                    if block.text:
                        _append_text_to_hwpx_paragraph(document, 0, paragraph_index, block.text, char_pr_id=block.char_pr_id)
                elif not getattr(block, "source_text_owned_by_container", False):
                    _append_non_paragraph_block_to_hwpx(document, 0, paragraph_index, block)
        for paragraph in staged:
            sublist.append(paragraph)
    finally:
        for paragraph in staged:
            if paragraph.getparent() is section.root_element:
                section.root_element.remove(paragraph)
        section.mark_modified()
    return sublist


def _append_nested_block_to_hwpx_paragraph(document: HwpxDocument, paragraph: etree._Element, block: HancomBlock) -> None:
    if isinstance(block, Paragraph):
        if block.text:
            run = etree.SubElement(paragraph, qname("hp", "run"))
            run.set("charPrIDRef", block.char_pr_id or "0")
            etree.SubElement(run, qname("hp", "t")).text = block.text
        return
    if isinstance(block, Bookmark):
        run = etree.SubElement(paragraph, qname("hp", "run"))
        run.set("charPrIDRef", "0")
        ctrl = etree.SubElement(run, qname("hp", "ctrl"))
        etree.SubElement(ctrl, qname("hp", "bookmark")).set("name", block.name)
        return
    if isinstance(block, Hyperlink):
        _append_nested_field_to_hwpx_paragraph(
            document,
            paragraph,
            field_type="HYPERLINK",
            display_text=block.display_text,
            parameters=_hwpx_hyperlink_parameters(block.target),
        )
        return
    if isinstance(block, Field):
        _append_nested_field_to_hwpx_paragraph(
            document,
            paragraph,
            field_type=block.effective_native_field_type,
            display_text=block.display_text,
            name=block.name,
            parameters=block.parameters,
            editable=block.editable,
            dirty=block.dirty,
        )
        return
    if isinstance(block, AutoNumber):
        run = etree.SubElement(paragraph, qname("hp", "run"))
        run.set("charPrIDRef", "0")
        ctrl = etree.SubElement(run, qname("hp", "ctrl"))
        node = etree.SubElement(ctrl, qname("hp", block.kind))
        node.set("num", str(block.number))
        node.set("numType", str(block.number_type))


def _append_nested_field_to_hwpx_paragraph(
    document: HwpxDocument,
    paragraph: etree._Element,
    *,
    field_type: str,
    display_text: str | None = None,
    name: str | None = None,
    parameters: dict[str, int | str] | None = None,
    editable: bool = False,
    dirty: bool = False,
) -> None:
    begin_id = document._next_control_number(".//hp:fieldBegin/@id")
    field_id = document._next_control_number(".//hp:fieldBegin/@fieldid | .//hp:fieldEnd/@fieldid")

    begin_run = etree.SubElement(paragraph, qname("hp", "run"))
    begin_run.set("charPrIDRef", "0")
    begin_ctrl = etree.SubElement(begin_run, qname("hp", "ctrl"))
    field_begin = etree.SubElement(begin_ctrl, qname("hp", "fieldBegin"))
    field_begin.set("id", begin_id)
    field_begin.set("type", field_type)
    field_begin.set("name", name or "")
    field_begin.set("editable", "1" if editable else "0")
    field_begin.set("dirty", "1" if dirty else "0")
    field_begin.set("zorder", "0")
    field_begin.set("fieldid", field_id)

    if parameters:
        params = etree.SubElement(field_begin, qname("hp", "parameters"))
        params.set("cnt", str(len(parameters)))
        params.set("name", "")
        for key, value in parameters.items():
            tag_name = "integerParam" if isinstance(value, int) else "stringParam"
            parameter = etree.SubElement(params, qname("hp", tag_name))
            parameter.set("name", str(key))
            parameter.text = str(value)

    if display_text is not None:
        display_run = etree.SubElement(paragraph, qname("hp", "run"))
        display_run.set("charPrIDRef", "0")
        etree.SubElement(display_run, qname("hp", "t")).text = display_text

    end_run = etree.SubElement(paragraph, qname("hp", "run"))
    end_run.set("charPrIDRef", "0")
    end_ctrl = etree.SubElement(end_run, qname("hp", "ctrl"))
    field_end = etree.SubElement(end_ctrl, qname("hp", "fieldEnd"))
    field_end.set("beginIDRef", begin_id)
    field_end.set("fieldid", field_id)


def _hwpx_hyperlink_parameters(target: str) -> dict[str, int | str]:
    parameters: dict[str, int | str] = {"Command": target, "Path": target}
    if target.startswith("mailto:"):
        parameters["Category"] = "HWPHYPERLINK_TYPE_EMAIL"
    elif target.startswith(("http://", "https://")):
        parameters["Category"] = "HWPHYPERLINK_TYPE_WEB"
    else:
        parameters["Category"] = "HWPHYPERLINK_TYPE_PATH"
    return parameters


def _apply_hwpx_hyperlink_runs(field, block: Hyperlink | Field) -> None:
    char_id = getattr(block, "_docx_char_pr_id", None)
    if char_id is not None:
        field.element.xpath("ancestor::hp:run[1]", namespaces=NS)[0].set("charPrIDRef", str(char_id))
        field.section.mark_modified()
    rich = getattr(block, "_docx_hyperlink_runs", None)
    if not rich or rich[0] != block.display_text:
        return
    begin = field.element.xpath("ancestor::hp:run[1]", namespaces=NS)[0]
    paragraph = begin.getparent()
    end = field.element.xpath("ancestor::hp:p[1]//hp:fieldEnd[@beginIDRef=$id]/ancestor::hp:run[1]", namespaces=NS, id=field.control_id)[0]
    for node in list(paragraph)[paragraph.index(begin) + 1:paragraph.index(end)]:
        paragraph.remove(node)
    for segment in rich[1]:
        run = etree.Element(qname("hp", "run"), charPrIDRef=segment.char_pr_id or "0")
        text = etree.SubElement(run, qname("hp", "t"))
        _set_text_node_content(text, segment.text)
        end.addprevious(run)
    field.section.mark_modified()


def _hwpx_field_kwargs(block: Field) -> dict[str, object]:
    if "WordInstruction" in block.parameters or block.field_type in {"JAKAL_LIST_LABEL", "JAKAL_NOTE_STYLE", "JAKAL_WORD_FEATURE", "JAKAL_WORD_FIELD_BOUNDARY"}:
        import base64
        import zlib

        details = json.dumps({"kind": block.field_type, "name": block.name, "parameters": block.parameters,
                              "editable": block.editable, "dirty": block.dirty},
                             ensure_ascii=True, separators=(",", ":")).encode("utf-8")
        name = _WORD_EFFECT_FIELD_PREFIX + base64.urlsafe_b64encode(zlib.compress(details)).decode("ascii")
        if len(name) - len(_WORD_EFFECT_FIELD_PREFIX) > 131072:
            raise ValueError("The Word feature definition is too large for a reversible HWPX field fallback.")
        # Hancom rejects arbitrary field type names. Use its editable text
        # field, retaining our semantic kind in ordinary field parameters.
        result = {"field_type": "CLICK_HERE", "name": name, "editable": True, "dirty": True, "parameters": {
            **block.parameters, "JakalFieldKind": block.field_type, "Prop": 9,
            "Command": "Clickhere:set:42:Direction:wstring:0: HelpState:wstring:0:  ",
        }}
        if "WordInstruction" in block.parameters:
            from .addons.docx.native_effects import native_field

            native = native_field(str(block.parameters["WordInstruction"]), block.display_text or "")
            if native is not None:
                result.update(native)
        return result
    return {"field_type": block.effective_native_field_type, "name": block.name, "parameters": block.parameters,
            "editable": block.editable, "dirty": block.dirty}


_WORD_EFFECT_FIELD_PREFIX = "jakal:word-effect:v1:"


def _word_effect_from_field_name(name: str | None) -> dict | None:
    # Hancom rewrites CLICK_HERE to its native Prop/Command parameter set on
    # save. Names survive that normalization, so they carry the small feature
    # definition as well. No original document/package is embedded here.
    if not name or not name.startswith(_WORD_EFFECT_FIELD_PREFIX):
        return None
    import base64
    import zlib

    try:
        encoded = name.removeprefix(_WORD_EFFECT_FIELD_PREFIX)
        if len(encoded) > 131072:
            return None
        decoder = zlib.decompressobj()
        data = decoder.decompress(base64.urlsafe_b64decode(encoded), max_length=1048576)
        if not decoder.eof or decoder.unconsumed_tail:
            return None
        details = json.loads(data)
        if not isinstance(details, dict) or not isinstance(details.get("kind"), str) or not isinstance(details.get("parameters"), dict):
            return None
        details.setdefault("name", None)
        if details.get("name") is not None and not isinstance(details["name"], str):
            return None
        if not isinstance(details.get("editable", False), bool) or not isinstance(details.get("dirty", False), bool):
            return None
        if any(not isinstance(key, str) or not isinstance(value, (str, int)) for key, value in details["parameters"].items()):
            return None
        return details
    except (ValueError, TypeError, zlib.error):
        return None


def _apply_hwpx_nested_text_blocks(document: HwpxDocument, wrapper, block: object, *, vertical_align: str = "CENTER") -> None:
    from ._hancom_conversion_common import _nested_blocks

    textbox = getattr(block, "_word_extra", {}).get("textbox", {})
    if textbox:
        vertical_align = textbox.get("vertical_align", vertical_align)
    nested = _nested_blocks(block)
    if isinstance(block, Note) and (block.reference_run_properties is not None or block.marker_run_properties is not None):
        parameters: dict[str, str | int] = {}
        for key, value in (("ReferenceRPr", block.reference_run_properties), ("MarkerRPr", block.marker_run_properties)):
            if value is not None:
                parameters[key] = value
        # An empty native text field retains the small formatting definition
        # through Hancom's parameter normalization without printing any text.
        carrier = Field("JAKAL_NOTE_STYLE", display_text="", parameters=parameters)
        setattr(carrier, "nested_paragraph_index", getattr(nested[0], "nested_paragraph_index", 0) if nested else 0)
        setattr(carrier, "nested_source_order", -1000)
        if not nested:
            nested = [Paragraph(block.text)]
        nested = [carrier, *nested]
    if not nested:
        return
    sublist = _build_hwpx_nested_sublist(document, nested, vertical_align=vertical_align)
    element = wrapper.element
    local_name = etree.QName(element).localname
    if local_name in {"header", "footer", "footNote", "endNote", "hiddenComment"}:
        old_nodes = element.xpath("./hp:subList", namespaces=NS)
        if old_nodes:
            element.replace(old_nodes[0], sublist)
        else:
            element.append(sublist)
        wrapper.section.mark_modified()
        return

    draw_text_nodes = element.xpath("./hp:drawText", namespaces=NS)
    draw_text = draw_text_nodes[0] if draw_text_nodes else etree.SubElement(element, qname("hp", "drawText"))
    if not draw_text_nodes:
        draw_text.set("lastWidth", str(getattr(block, "width", 12000)))
        draw_text.set("name", "")
        draw_text.set("editable", "0")
    old_nodes = draw_text.xpath("./hp:subList", namespaces=NS)
    if old_nodes:
        draw_text.replace(old_nodes[0], sublist)
    else:
        draw_text.insert(0, sublist)
    wrapper.section.mark_modified()


def _append_header_footer_block_to_hwpx(
    document: HwpxDocument,
    section_index: int,
    paragraph_index: int,
    block: HeaderFooter,
) -> None:
    from ._hancom_conversion_common import _has_nested_blocks
    from ._hancom_hwpx_graphics import _apply_hwpx_shape_block
    from ._hancom_hwpx_settings import _extract_section_settings

    if block.apply_page_type == "FIRST":
        settings = _extract_section_settings(document.sections[section_index])
        fallback = _first_page_story_shape(block, settings)
        appended = document.append_shape(
            kind="rect", text="" if block.nested_blocks else block.text,
            width=fallback.width, height=fallback.height,
            section_index=section_index, paragraph_index=paragraph_index,
        )
        _apply_hwpx_shape_block(appended, fallback)
        _apply_hwpx_nested_text_blocks(document, appended, fallback, vertical_align="BOTTOM" if block.kind == "footer" else "TOP")
        # FIRST has no native HWPX header variant. The ordinary story is hidden
        # on page one; a page-relative, non-wrapping text box displays it there.
        document.section_settings(section_index).set_visibility(**{"hide_first_" + block.kind: True})
        return
    if block.kind == "header":
        appended = document.append_header(
            "" if _has_nested_blocks(block) else block.text,
            apply_page_type=block.apply_page_type,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )
        _apply_hwpx_nested_text_blocks(document, appended, block, vertical_align="TOP")
    elif block.kind == "footer":
        appended = document.append_footer(
            "" if _has_nested_blocks(block) else block.text,
            apply_page_type=block.apply_page_type,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )
        _apply_hwpx_nested_text_blocks(document, appended, block, vertical_align="TOP")


_FIRST_PAGE_STORY_MARKER = "jakal:first-page-story:"


def _first_page_story_shape(block: HeaderFooter, settings: SectionSettings) -> Shape:
    margins = settings.margins
    page_width = settings.page_height if settings.landscape == "NARROWLY" else settings.page_width
    height = max(2000, int(margins.get("header" if block.kind == "header" else "footer", 3600)))
    # Let multiline text occupy its natural height, including paragraphs larger
    # than the reserved story area. This does not consume body layout space.
    height = max(height, 1800 * max(1, block.text.count("\n") + 1))
    return Shape(
        "rect", text=block.text,
        width=int(page_width or 59528) - int(margins.get("left", 8504)) - int(margins.get("right", 8504)),
        height=height, shape_comment=_FIRST_PAGE_STORY_MARKER + block.kind,
        line_width=0, has_line_node=False, has_fill_node=False,
        text_margins={side: 0 for side in ("left", "right", "top", "bottom")},
        layout={"treatAsChar": "0", "textWrap": "IN_FRONT_OF_TEXT", "flowWithText": "0",
                "allowOverlap": "1", "holdAnchorAndSO": "1", "vertRelTo": "PAPER", "horzRelTo": "PAPER",
                "vertAlign": "BOTTOM" if block.kind == "footer" else "TOP", "horzAlign": "LEFT",
                "horzOffset": str(margins.get("left", 8504)),
                "vertOffset": str(margins.get("bottom" if block.kind == "footer" else "top", 4252))},
        nested_blocks=block.nested_blocks,
    )


def _restore_first_page_stories(blocks: list[HancomBlock], stories: list[HeaderFooter]) -> None:
    from ._hancom_hwpx_read import _mark_hwpx_source_paragraph

    for block in list(blocks):
        if not isinstance(block, Shape) or block.shape_comment not in {
            _FIRST_PAGE_STORY_MARKER + "header", _FIRST_PAGE_STORY_MARKER + "footer",
        }:
            continue
        story = HeaderFooter(block.shape_comment.removeprefix(_FIRST_PAGE_STORY_MARKER), block.text,
                             apply_page_type="FIRST", nested_blocks=block.nested_blocks)
        _mark_hwpx_source_paragraph(story, getattr(block, "source_paragraph_index", 0), order=getattr(block, "source_order", 0))
        stories.append(story)
        blocks.remove(block)


def _append_control_host_paragraph(document: HwpxDocument, section_index: int) -> int:
    document.append_paragraph("", section_index=section_index)
    index = len(document.sections[section_index].paragraphs()) - 1
    document._paragraph_element(section_index, index).attrib.update({"pageBreak": "0", "columnBreak": "0"})
    return index


def _apply_hwpx_paragraph_style_block(
    document: HwpxDocument, section_index: int, paragraph_index: int, block: Paragraph,
    *, apply_character_style: bool = True,
) -> None:
    attributes = getattr(block, "_hwpx_paragraph_attributes", None)
    if attributes:
        document._paragraph_element(section_index, paragraph_index).attrib.update(attributes)
    style_id = _resolve_hwpx_style_id(document, block)
    para_pr_id = _resolve_hwpx_para_pr_id(document, block)
    char_pr_id = _resolve_hwpx_char_pr_id(document, block) if apply_character_style else None
    if style_id is None and para_pr_id is None and char_pr_id is None:
        return
    paragraph = document._paragraph_element(section_index, paragraph_index)
    protected = [(run, run.get("charPrIDRef")) for run in paragraph.findall(qname("hp", "run"))
                 if run.xpath(".//hp:fieldBegin/hp:parameters/*[@name='JakalFieldKind' and text()='JAKAL_WORD_FEATURE']", namespaces=NS)]
    document.apply_style_to_paragraph(
        section_index,
        paragraph_index,
        style_id=style_id,
        para_pr_id=para_pr_id,
        char_pr_id=char_pr_id,
    )
    for run, reference in protected:
        if reference is not None:
            run.set("charPrIDRef", reference)


def _resolve_hwpx_style_id(document: HwpxDocument, block: Paragraph) -> str | None:
    if block.style_id is None:
        return None
    known = {style.style_id for style in document.styles() if style.style_id is not None}
    return block.style_id if block.style_id in known else None


def _resolve_hwpx_para_pr_id(document: HwpxDocument, block: Paragraph) -> str | None:
    candidate = block.para_pr_id
    if candidate is None and block.hwp_para_shape_id is not None:
        candidate = str(block.hwp_para_shape_id)
    if candidate is None:
        return None
    known = {style.style_id for style in document.paragraph_styles() if style.style_id is not None}
    return candidate if candidate in known else None


def _resolve_hwpx_char_pr_id(document: HwpxDocument, block: Paragraph) -> str | None:
    if block.char_pr_id is None:
        return None
    known = {style.style_id for style in document.character_styles() if style.style_id is not None}
    return block.char_pr_id if block.char_pr_id in known else None
