"""HWPX block extraction, nested text containers, and field carriers."""
from __future__ import annotations

import json
from typing import cast

from lxml import etree

from ._hancom_models import (
    _HWPX_CHART_CARRIER_PREFIX,
    _HWPX_FORM_FIELD_TYPE,
    _HWPX_MEMO_FIELD_TYPE,
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
    Paragraph,
    Picture,
    Shape,
)
from ._hancom_styles import _coerce_bool
from ._hwpx_element_helpers import _text_node_plain_text
from .namespaces import NS


def _extract_hwpx_section(section) -> HancomSection:
    from ._hancom_from_hwpx import _extract_hwpx_paragraph_blocks
    from ._hancom_hwpx_settings import _extract_section_settings
    from ._hancom_hwpx_write import _restore_first_page_stories

    blocks: list[HancomBlock] = []
    section_settings = _extract_section_settings(section)
    header_footer_blocks = _extract_header_footer_blocks(section)
    paragraphs = section.readonly_root_element.xpath("./hp:p", namespaces=NS)
    for paragraph_index, paragraph in enumerate(paragraphs):
        paragraph_blocks = _extract_hwpx_paragraph_blocks(section, paragraph, paragraph_index)
        if not paragraph_blocks and not _hwpx_paragraph_has_header_footer(paragraph):
            paragraph_blocks.append(_mark_hwpx_source_paragraph(_build_hwpx_paragraph_block(paragraph, ""), paragraph_index))
        blocks.extend(paragraph_blocks)
    _restore_first_page_stories(blocks, header_footer_blocks)
    return HancomSection(settings=section_settings, header_footer_blocks=header_footer_blocks, blocks=blocks)


def _hwpx_paragraph_has_header_footer(paragraph: etree._Element) -> bool:
    return bool(paragraph.xpath("./hp:run/hp:ctrl/hp:header | ./hp:run/hp:ctrl/hp:footer", namespaces=NS))


def _sort_hwpx_blocks_by_source_order(blocks: list[HancomBlock]) -> list[HancomBlock]:
    return [
        block
        for _index, block in sorted(
            enumerate(blocks),
            key=lambda item: (
                getattr(item[1], "source_order", 10**9),
                item[0],
            ),
        )
    ]


def _sort_hwpx_nested_blocks_by_source_order(blocks: list[HancomBlock]) -> list[HancomBlock]:
    return [
        block
        for _index, block in sorted(
            enumerate(blocks),
            key=lambda item: (
                getattr(item[1], "nested_paragraph_index", 10**9),
                getattr(item[1], "nested_source_order", 10**9),
                item[0],
            ),
        )
    ]


def _mark_hwpx_nested_paragraph(block, paragraph_index: int, *, order: int | None = None):
    setattr(block, "nested_paragraph_index", paragraph_index)
    if order is not None:
        setattr(block, "nested_source_order", order)
    return block


def _extract_hwpx_text_container_blocks(section, element: etree._Element) -> list[HancomBlock]:
    local_name = etree.QName(element).localname
    if local_name in {"rect", "line", "ellipse", "arc", "polygon", "curve", "connectLine", "container"}:
        paragraphs = element.xpath("./hp:drawText/hp:subList/hp:p", namespaces=NS)
    elif local_name in {"header", "footer", "footNote", "endNote", "hiddenComment"}:
        paragraphs = element.xpath("./hp:subList/hp:p", namespaces=NS)
    else:
        return []

    blocks: list[HancomBlock] = []
    for paragraph_index, paragraph in enumerate(paragraphs):
        blocks.extend(_extract_hwpx_nested_paragraph_blocks(section, paragraph, paragraph_index))
    if local_name not in {"header", "footer", "footNote", "endNote", "hiddenComment"} and all(
        isinstance(block, Paragraph) and not block.text for block in blocks
    ):
        return []
    return _sort_hwpx_nested_blocks_by_source_order(blocks)


def _extract_hwpx_nested_paragraph_blocks(section, paragraph: etree._Element, paragraph_index: int) -> list[HancomBlock]:
    blocks: list[HancomBlock] = []
    skipped_field_id: str | None = None
    for run in paragraph.xpath("./hp:run", namespaces=NS):
        for child in list(run):
            local_name = etree.QName(child).localname
            if local_name == "t":
                if skipped_field_id is None:
                    text = _text_node_plain_text(child).replace("\r", "")
                    if text:
                        blocks.append(
                            _mark_hwpx_nested_paragraph(
                                _build_hwpx_paragraph_block(paragraph, text, char_pr_id=run.get("charPrIDRef")),
                                paragraph_index,
                                order=_source_order_for_element(paragraph, child),
                            )
                        )
                continue

            candidates = list(child) if local_name == "ctrl" else [child]
            for node in candidates:
                node_local_name = etree.QName(node).localname
                if skipped_field_id is not None:
                    if node_local_name == "fieldEnd" and node.get("beginIDRef") == skipped_field_id:
                        skipped_field_id = None
                    continue
                if node_local_name not in _HWPX_NESTED_INLINE_BLOCK_NAMES:
                    continue
                extracted = _extract_hwpx_run_node(section, node)
                if extracted is None:
                    continue
                extracted_blocks = extracted if isinstance(extracted, list) else [extracted]
                for block in extracted_blocks:
                    blocks.append(
                        _mark_hwpx_nested_paragraph(
                            block,
                            paragraph_index,
                            order=_source_order_for_element(paragraph, node),
                        )
                    )
                if node_local_name == "fieldBegin":
                    skipped_field_id = node.get("id") or node.get("fieldid")
    if not blocks:
        blocks.append(_mark_hwpx_nested_paragraph(_build_hwpx_paragraph_block(paragraph, ""), paragraph_index))
    elif isinstance(blocks[0], (Picture, Equation, Hyperlink)):
        blocks.insert(0, _mark_hwpx_nested_paragraph(_build_hwpx_paragraph_block(paragraph, ""), paragraph_index, order=-1))
    return _sort_hwpx_nested_blocks_by_source_order(blocks)


def _extract_missing_hwpx_descendant_blocks(
    section,
    paragraph: etree._Element,
    paragraph_index: int,
    seen_nodes: set[str],
) -> list[HancomBlock]:
    expression = (
        ".//hp:bookmark | .//hp:fieldBegin | .//hp:checkBtn | .//hp:radioBtn | .//hp:btn | .//hp:edit | "
        ".//hp:comboBox | .//hp:listBox | .//hp:scrollBar | .//hp:autoNum | .//hp:newNum | .//hp:tbl | "
        ".//hp:pic | .//hp:footNote | .//hp:endNote | .//hp:hiddenComment | .//hp:equation | .//hp:chart | "
        ".//hp:rect | .//hp:line | .//hp:ellipse | .//hp:arc | .//hp:polygon | .//hp:curve | "
        ".//hp:connectLine | .//hp:textart | .//hp:container | .//hp:ole"
    )
    blocks: list[HancomBlock] = []
    memo_blocks: list[Memo] = []
    memo_carriers: list[Memo] = []
    for node in paragraph.xpath(expression, namespaces=NS):
        if _hwpx_node_is_inside_table_cell(node):
            continue  # Cell sublists own their controls; do not lift them into the body.
        if _hwpx_node_is_inside_text_container(node) and etree.QName(node).localname in _HWPX_NESTED_INLINE_BLOCK_NAMES:
            continue
        node_key = _hwpx_node_key(node)
        if node_key in seen_nodes:
            continue
        extracted = _extract_hwpx_run_node(section, node)
        if extracted is None:
            continue
        seen_nodes.add(node_key)
        if not _hwpx_node_is_inside_table_cell(node):
            caption_text = _extract_hwpx_caption_text(node)
            if caption_text:
                blocks.append(
                    _mark_hwpx_source_paragraph(
                        _build_hwpx_paragraph_block(paragraph, caption_text),
                        paragraph_index,
                        order=_source_order_for_element(paragraph, node),
                    )
                )
        extracted_blocks = extracted if isinstance(extracted, list) else [extracted]
        for block in extracted_blocks:
            if _hwpx_node_is_inside_table_cell(node):
                _clear_nested_owned_text(block)
            elif _hwpx_node_is_inside_text_container(node):
                _mark_hwpx_text_owned_by_container(block)
            owned_text = _hwpx_owned_inline_text(block)
            if owned_text:
                blocks.append(
                    _mark_hwpx_source_paragraph(
                        _build_hwpx_paragraph_block(paragraph, owned_text),
                        paragraph_index,
                        order=_source_order_for_element(paragraph, node),
                    )
                )
            _mark_hwpx_source_paragraph(block, paragraph_index, order=_source_order_for_element(paragraph, node))
            if isinstance(block, Memo) and _memo_requires_hwpx_carrier(block):
                memo_carriers.append(block)
            elif isinstance(block, Memo):
                memo_blocks.append(block)
            else:
                blocks.append(block)
    if memo_blocks or memo_carriers:
        for memo in _merge_hwpx_memos_with_carriers(memo_blocks, memo_carriers):
            blocks.append(_mark_hwpx_source_paragraph(memo, paragraph_index))
    return blocks


def _extract_hwpx_caption_text(node: etree._Element) -> str:
    if etree.QName(node).localname not in _CAPTION_TEXT_BLOCK_NAMES:
        return ""
    caption_nodes = node.xpath("./hp:caption", namespaces=NS)
    if not caption_nodes:
        return ""
    return "".join(text_node.text or "" for text_node in caption_nodes[0].xpath(".//hp:t", namespaces=NS))


_CAPTION_TEXT_BLOCK_NAMES = {
    "tbl",
    "pic",
    "ole",
    "rect",
    "line",
    "ellipse",
    "arc",
    "polygon",
    "curve",
    "connectLine",
    "textart",
    "container",
}


def _hwpx_node_is_inside_table_cell(node: etree._Element) -> bool:
    return any(etree.QName(ancestor).localname == "tc" for ancestor in node.iterancestors())


_HWPX_TEXT_CONTAINER_NAMES = {
    "header",
    "footer",
    "footNote",
    "endNote",
    "hiddenComment",
    "rect",
    "line",
    "ellipse",
    "arc",
    "polygon",
    "curve",
    "connectLine",
    "textart",
    "container",
}


_HWPX_NESTED_INLINE_BLOCK_NAMES = {
    "bookmark",
    "fieldBegin",
    "autoNum",
    "newNum",
    "tbl", "pic", "equation", "rect", "line", "ellipse", "arc", "polygon", "curve",
    "connectLine", "textart", "container", "ole", "chart", "footNote", "endNote", "hiddenComment",
    "checkBtn", "radioBtn", "btn", "edit", "comboBox", "listBox", "scrollBar",
}


def _hwpx_node_is_inside_text_container(node: etree._Element) -> bool:
    return any(etree.QName(ancestor).localname in _HWPX_TEXT_CONTAINER_NAMES for ancestor in node.iterancestors())


def _clear_nested_owned_text(block: HancomBlock) -> None:
    if isinstance(block, Shape):
        block.text = ""
    elif isinstance(block, Chart):
        block.title = ""
    elif isinstance(block, Hyperlink):
        block.display_text = None
    elif isinstance(block, Field):
        block.display_text = None
    elif isinstance(block, Form):
        block.label = ""
    elif isinstance(block, Memo):
        block.text = ""


def _mark_hwpx_text_owned_by_container(block: HancomBlock) -> None:
    if isinstance(block, (Hyperlink, Field)):
        setattr(block, "source_text_owned_by_container", True)


def _hwpx_owned_inline_text(block: HancomBlock) -> str | None:
    if getattr(block, "source_text_owned_by_container", False):
        return None
    if isinstance(block, (Hyperlink, Field)):
        return block.display_text
    return None


def _hwpx_node_key(node: etree._Element) -> str:
    return node.getroottree().getpath(node)


def _extract_hwpx_run_node(section, node: etree._Element) -> HancomBlock | list[HancomBlock] | None:
    from ._hancom_hwpx_graphics import (
        _extract_hwpx_chart,
        _extract_hwpx_equation,
        _extract_hwpx_memo,
        _extract_hwpx_note,
        _extract_hwpx_ole,
        _extract_hwpx_picture,
        _extract_hwpx_shape,
        _extract_hwpx_table,
    )

    local_name = etree.QName(node).localname
    if local_name == "bookmark":
        return _extract_bookmark(_build_bookmark_wrapper(section, node))
    if local_name == "fieldBegin":
        field = _build_field_wrapper(section, node)
        result: Hyperlink | Field | Form | Memo
        if (field.field_type or "").upper() == "HYPERLINK":
            result = Hyperlink(target=field.hyperlink_target or "", display_text=field.display_text or None)
        else:
            result = _extract_field(field)
        if isinstance(result, (Hyperlink, Field)):
            styled = [Paragraph(_text_node_plain_text(text), char_pr_id=text.getparent().get("charPrIDRef")) for text in field._display_text_nodes()]
            setattr(result, "_docx_hyperlink_runs", (result.display_text, styled))
            runs = node.xpath("ancestor::hp:run[1]", namespaces=NS)
            if runs:
                setattr(result, "_docx_char_pr_id", runs[0].get("charPrIDRef"))
            paragraphs = node.xpath("ancestor::hp:p[1]", namespaces=NS)
            if paragraphs:
                paragraph = paragraphs[0]
                setattr(result, "_word_native_para_pr_id", paragraph.get("paraPrIDRef"))
                setattr(result, "_word_native_style_id", paragraph.get("styleIDRef"))
                text_characters = paragraph.xpath("./hp:run[hp:t/text()]/@charPrIDRef", namespaces=NS)
                if text_characters:
                    setattr(result, "_word_native_text_char_id", text_characters[0])
        return result
    if local_name in {"checkBtn", "radioBtn", "btn", "edit", "comboBox", "listBox", "scrollBar"}:
        return _extract_hwpx_form(_build_form_wrapper(section, node))
    if local_name in {"autoNum", "newNum"}:
        number = _extract_auto_number(_build_auto_number_wrapper(section, node))
        runs = node.xpath("ancestor::hp:run[1]", namespaces=NS)
        if runs:
            setattr(number, "_docx_char_pr_id", runs[0].get("charPrIDRef"))
        formats = node.xpath("./hp:autoNumFormat/@type", namespaces=NS)
        if formats:
            setattr(number, "_docx_number_format", formats[0])
        superscript = node.xpath("./hp:autoNumFormat/@supscript", namespaces=NS)
        if superscript:
            setattr(number, "_docx_superscript", superscript[0] == "1")
        return number
    if local_name == "tbl":
        table = _extract_hwpx_table(_build_table_wrapper(section, node))
        paragraphs = node.xpath("ancestor::hp:p[1]", namespaces=NS)
        if paragraphs:
            setattr(table, "_hwpx_paragraph_attributes", {key: paragraphs[0].get(key, "0") for key in ("pageBreak", "columnBreak")})
            run = node.getparent()
            setattr(table, "_hwpx_anchor_paragraph", _build_hwpx_paragraph_block(paragraphs[0], "", char_pr_id=run.get("charPrIDRef")))
        return table
    if local_name == "pic":
        return _extract_hwpx_picture(_build_picture_wrapper(section, node))
    if local_name in {"footNote", "endNote"}:
        return _extract_hwpx_note(_build_note_wrapper(section, node))
    if local_name == "hiddenComment":
        return _extract_hwpx_memo(_build_memo_wrapper(section, node))
    if local_name == "equation":
        return _extract_hwpx_equation(_build_equation_wrapper(section, node))
    if local_name == "chart":
        if _hwpx_chart_is_covered_by_carrier_shape(node):
            return None
        return _extract_hwpx_chart(_build_chart_wrapper(section, node))
    if local_name in {"rect", "line", "ellipse", "arc", "polygon", "curve", "connectLine", "textart", "container"}:
        return _extract_hwpx_shape(_build_shape_wrapper(section, node))
    if local_name == "ole":
        if node.xpath("ancestor::hp:default[parent::hp:switch[hp:case/hp:chart]]", namespaces=NS):
            return None
        return _extract_hwpx_ole(_build_ole_wrapper(section, node))
    return None


def _hwpx_chart_is_covered_by_carrier_shape(node: etree._Element) -> bool:
    for ancestor in node.iterancestors():
        if etree.QName(ancestor).localname not in {
            "rect",
            "line",
            "ellipse",
            "arc",
            "polygon",
            "curve",
            "connectLine",
            "textart",
            "container",
        }:
            continue
        comments = ancestor.xpath("./hp:shapeComment/text()", namespaces=NS)
        if comments and str(comments[0]).startswith(_HWPX_CHART_CARRIER_PREFIX):
            return True
    return False


def _source_order_for_element(paragraph: etree._Element, element: etree._Element) -> int:
    for index, candidate in enumerate(paragraph.iter()):
        if candidate is element:
            return index
    return 10**9


def _mark_hwpx_source_paragraph(block, paragraph_index: int, *, order: int | None = None):
    setattr(block, "source_paragraph_index", paragraph_index)
    if order is not None:
        setattr(block, "source_order", order)
    return block


def _build_hwpx_paragraph_block(paragraph, text: str, *, char_pr_id: str | None = None) -> Paragraph:
    result = Paragraph(
        text=text,
        style_id=paragraph.get("styleIDRef"),
        para_pr_id=paragraph.get("paraPrIDRef"),
        char_pr_id=char_pr_id or _extract_hwpx_paragraph_char_pr_id(paragraph),
    )
    setattr(result, "_hwpx_paragraph_attributes", {key: paragraph.get(key, "0") for key in ("pageBreak", "columnBreak", "merged")})
    from ._hancom_text_layout import capture_empty_runs
    empty = capture_empty_runs(paragraph)
    if empty is not None:
        setattr(result, "_hwpx_empty_runs", empty)
    return result


def _extract_hwpx_paragraph_char_pr_id(paragraph) -> str | None:
    run_ids = [
        value
        for value in paragraph.xpath("./hp:run/@charPrIDRef", namespaces=NS)
        if isinstance(value, str) and value
    ]
    if not run_ids:
        return None
    first = run_ids[0]
    return first if all(value == first for value in run_ids) else None


def _extract_bookmark(bookmark) -> Bookmark:
    return Bookmark(name=bookmark.name or "")


def _extract_field(field) -> Field | Form | Memo:
    from ._hancom_hwpx_write import _word_effect_from_field_name

    if (field.field_type or "").upper() == _HWPX_FORM_FIELD_TYPE:
        return _extract_hwpx_form_field(field)
    if (field.field_type or "").upper() == _HWPX_MEMO_FIELD_TYPE:
        return _extract_hwpx_memo_field(field)
    parameters = field.parameter_map()
    restored = _word_effect_from_field_name(field.name)
    if restored is not None:
        display = field.display_text or None
        if restored["parameters"].get("NativeList") == "1" and not display:
            display = str(restored["parameters"].get("RenderedLabel", ""))
        return Field(field_type=restored["kind"], display_text=display,
                     name=restored["name"], parameters=restored["parameters"],
                     editable=restored.get("editable", False), dirty=restored.get("dirty", False))
    return Field(
        field_type=parameters.get("JakalFieldKind", field.field_type or ""),
        display_text=field.display_text or None,
        name=field.name,
        parameters={key: value for key, value in parameters.items() if key not in {"JakalFieldKind", "Prop", "Command"}}
        if "JakalFieldKind" in parameters else parameters,
        editable=field.element.get("editable") == "1",
        dirty=field.element.get("dirty") == "1",
    )


def _extract_hwpx_form_field(field) -> Form:
    from ._hancom_conversion_common import _json_list_parameter

    parameters = field.parameter_map()
    items = _json_list_parameter(parameters.get("Items"))
    return Form(
        label=field.display_text or "",
        form_type=parameters.get("FormType", "INPUT") or "INPUT",
        name=field.name or None,
        value=parameters.get("Value"),
        checked=_coerce_bool(parameters.get("Checked")) or False,
        items=items,
        editable=_coerce_bool(parameters.get("Editable")) is not False,
        locked=_coerce_bool(parameters.get("Locked")) or False,
        placeholder=parameters.get("Placeholder") or None,
    )


def _extract_hwpx_form(form) -> Form:
    return Form(
        label=form.label,
        form_type=form.form_type,
        name=form.name,
        value=form.value,
        checked=form.checked,
        items=form.items,
        editable=form.editable,
        locked=form.locked,
        placeholder=form.placeholder,
    )


def _extract_hwpx_memo_field(field) -> Memo:
    parameters = field.parameter_map()
    order_value = parameters.get("Order")
    try:
        order = int(order_value) if order_value not in (None, "") else None
    except (TypeError, ValueError):
        order = None
    return Memo(
        text=field.display_text or parameters.get("Text") or "",
        author=parameters.get("Author") or None,
        memo_id=parameters.get("MemoId") or field.name or None,
        anchor_id=parameters.get("AnchorId") or None,
        order=order,
        visible=_coerce_bool(parameters.get("Visible")) is not False,
    )


def _memo_requires_hwpx_carrier(block: Memo) -> bool:
    return any(
        value is not None
        for value in (block.author, block.memo_id, block.anchor_id, block.order)
    ) or block.visible is False


def _merge_hwpx_memo_with_carrier(memo: Memo, carrier: Memo) -> Memo:
    return Memo(
        text=memo.text or carrier.text,
        author=carrier.author,
        memo_id=carrier.memo_id,
        anchor_id=carrier.anchor_id,
        order=carrier.order,
        visible=carrier.visible,
    )


def _merge_hwpx_memos_with_carriers(memos: list[Memo], carriers: list[Memo]) -> list[Memo]:
    if not carriers:
        return list(memos)
    unused_carriers = list(range(len(carriers)))
    merged: list[Memo] = []

    for memo in memos:
        carrier_index = _find_matching_hwpx_memo_carrier(memo, carriers, unused_carriers)
        if carrier_index is None:
            merged.append(memo)
            continue
        unused_carriers.remove(carrier_index)
        merged.append(_merge_hwpx_memo_with_carrier(memo, carriers[carrier_index]))

    for carrier_index in unused_carriers:
        merged.append(carriers[carrier_index])
    return merged


def _find_matching_hwpx_memo_carrier(memo: Memo, carriers: list[Memo], unused_carriers: list[int]) -> int | None:
    memo_text = _normalize_hwpx_memo_pairing_key(memo.text)
    if memo_text:
        for carrier_index in unused_carriers:
            if _normalize_hwpx_memo_pairing_key(carriers[carrier_index].text) == memo_text:
                return carrier_index
    return unused_carriers[0] if unused_carriers else None


def _normalize_hwpx_memo_pairing_key(value: str | None) -> str:
    return " ".join((value or "").split())


def _extract_hwpx_chart_carrier(shape) -> Chart | None:
    from ._hancom_conversion_common import _normalize_int_map, _normalize_str_map

    comment = shape.shape_comment or ""
    if not comment.startswith(_HWPX_CHART_CARRIER_PREFIX):
        return None
    try:
        payload = json.loads(comment[len(_HWPX_CHART_CARRIER_PREFIX) :])
    except json.JSONDecodeError:
        return None
    if not isinstance(payload, dict):
        return None
    size = shape.size()
    return Chart(
        title=shape.text or str(payload.get("title", "")),
        chart_type=str(payload.get("chartType", "BAR") or "BAR"),
        categories=[str(value) for value in payload.get("categories", []) if value is not None],
        series=[dict(value) for value in payload.get("series", []) if isinstance(value, dict)],
        data_ref=str(payload["dataRef"]) if payload.get("dataRef") not in (None, "") else None,
        legend_visible=_coerce_bool(payload.get("legendVisible")) is not False,
        width=size.get("width", 12000),
        height=size.get("height", 3200),
        shape_comment=str(payload["shapeComment"]) if payload.get("shapeComment") not in (None, "") else None,
        layout=_normalize_str_map(shape.layout()),
        out_margins=_normalize_int_map(shape.out_margins()),
        rotation=_normalize_str_map(shape.rotation()),
    )


def _extract_auto_number(auto_number) -> AutoNumber:
    return AutoNumber(
        kind=auto_number.kind,
        number=auto_number.number or 1,
        number_type=auto_number.number_type or "PAGE",
    )


def _extract_header_footer_blocks(section) -> list[HeaderFooter]:
    from ._hancom_hwpx_graphics import _extract_hwpx_direct_text

    section_index = section.section_index()
    blocks: list[HeaderFooter] = []
    for header in section.document.headers(section_index=section_index):
        paragraph_index, source_order = _source_paragraph_position_for_element(section, header.element)
        blocks.append(
            _mark_hwpx_source_paragraph(
                HeaderFooter(
                    kind=header.kind,
                    text=_extract_hwpx_direct_text(header.element),
                    apply_page_type=header.apply_page_type or "BOTH",
                    nested_blocks=cast(list[object], _extract_hwpx_text_container_blocks(header.section, header.element)),
                ),
                paragraph_index or 0,
                order=source_order,
            )
        )
    for footer in section.document.footers(section_index=section_index):
        paragraph_index, source_order = _source_paragraph_position_for_element(section, footer.element)
        blocks.append(
            _mark_hwpx_source_paragraph(
                HeaderFooter(
                    kind=footer.kind,
                    text=_extract_hwpx_direct_text(footer.element),
                    apply_page_type=footer.apply_page_type or "BOTH",
                    nested_blocks=cast(list[object], _extract_hwpx_text_container_blocks(footer.section, footer.element)),
                ),
                paragraph_index or 0,
                order=source_order,
            )
        )
    return blocks


def _source_paragraph_position_for_element(section, element: etree._Element) -> tuple[int | None, int | None]:
    paragraphs = section.readonly_root_element.xpath("./hp:p", namespaces=NS)
    for index, paragraph in enumerate(paragraphs):
        if paragraph is element or element in paragraph.iterdescendants():
            return index, _source_order_for_element(paragraph, element)
    return None, None


def _source_paragraph_index_for_element(section, element: etree._Element) -> int | None:
    paragraphs = section.readonly_root_element.xpath("./hp:p", namespaces=NS)
    for index, paragraph in enumerate(paragraphs):
        if paragraph is element or element in paragraph.iterdescendants():
            return index
    return None


def _build_table_wrapper(section, node):
    from .elements import TableXml

    return TableXml(section.document, section, node)


def _build_picture_wrapper(section, node):
    from .elements import PictureXml

    return PictureXml(section.document, section, node)


def _build_field_wrapper(section, node):
    from .elements import FieldXml

    return FieldXml(section.document, section, node)


def _build_bookmark_wrapper(section, node):
    from .elements import BookmarkXml

    return BookmarkXml(section.document, section, node)


def _build_auto_number_wrapper(section, node):
    from .elements import AutoNumberXml

    return AutoNumberXml(section.document, section, node)


def _build_note_wrapper(section, node):
    from .elements import NoteXml

    return NoteXml(section.document, section, node)


def _build_memo_wrapper(section, node):
    from .elements import MemoXml

    return MemoXml(section.document, section, node)


def _build_form_wrapper(section, node):
    from .elements import FormXml

    return FormXml(section.document, section, node)


def _build_equation_wrapper(section, node):
    from .elements import EquationXml

    return EquationXml(section.document, section, node)


def _build_shape_wrapper(section, node):
    from .elements import ShapeXml

    return ShapeXml(section.document, section, node)


def _build_chart_wrapper(section, node):
    from .elements import ChartXml

    return ChartXml(section.document, section, node)


def _build_ole_wrapper(section, node):
    from .elements import OleXml

    return OleXml(section.document, section, node)
