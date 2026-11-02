"""Native HWP block extraction and paragraph/control interpretation."""
from __future__ import annotations

import struct
from typing import cast

from ._hancom_models import (
    AutoNumber,
    Bookmark,
    Field,
    HancomBlock,
    HancomSection,
    HeaderFooter,
    Note,
    Paragraph,
    SectionSettings,
    Shape,
    Table,
)
from ._hancom_styles import _parse_optional_int
from .hwp_binary import (
    TAG_CTRL_DATA,
    TAG_CTRL_HEADER,
    TAG_LIST_HEADER,
    TAG_PAGE_DEF,
    TAG_PARA_CHAR_SHAPE,
    TAG_PARA_HEADER,
    ParagraphHeaderRecord,
    SectionParagraphModel,
)
from .hwp_document import (
    HwpArcShapeObject,
    HwpConnectLineShapeObject,
    HwpContainerShapeObject,
    HwpCurveShapeObject,
    HwpDocument,
    HwpEllipseShapeObject,
    HwpLineShapeObject,
    HwpPolygonShapeObject,
    HwpRectangleShapeObject,
    HwpTextArtShapeObject,
)


def _extract_hwp_section(section) -> HancomSection:
    from ._hancom_hwpx_write import _restore_first_page_stories

    controls_by_paragraph: dict[int, list[object]] = {}
    header_footer_blocks: list[HeaderFooter] = []
    for control in section.controls():
        if getattr(control, "control_id", "") in {"head", "foot"}:
            block = _extract_hwp_header_footer(control)
            if block is not None:
                header_footer_blocks.append(block)
            continue
        controls_by_paragraph.setdefault(control.paragraph_index, []).append(control)

    blocks: list[HancomBlock] = []
    for paragraph in section.model().paragraphs():
        paragraph_controls = controls_by_paragraph.get(paragraph.index, [])
        if paragraph.header.level != 0:
            blocks.extend(_extract_supported_nested_hwp_blocks(paragraph_controls))
            continue
        blocks.extend(_extract_hwp_paragraph_blocks(paragraph, paragraph_controls))
    _restore_first_page_stories(blocks, header_footer_blocks)
    return HancomSection(
        settings=_extract_hwp_section_settings(section),
        header_footer_blocks=header_footer_blocks,
        blocks=blocks,
    )


def _hwp_effect_blocks(blocks: list[HancomBlock]):
    """The binary HWP writer uses displayed text for portable Word fields."""
    from copy import deepcopy

    from ._hancom_conversion_common import _is_owned_text_shadow_paragraph

    label = ""
    for position, block in enumerate(blocks):
        following = blocks[position + 1] if position + 1 < len(blocks) else None
        if _is_owned_text_shadow_paragraph(block, following):
            continue
        if isinstance(block, Field) and block.field_type == "JAKAL_LIST_LABEL":
            label = block.display_text or ""
            continue
        if isinstance(block, Field) and "WordInstruction" in block.parameters:
            yield Paragraph(block.display_text or "")
            continue
        if label and isinstance(block, Paragraph):
            if not block.text:
                continue
            copied = deepcopy(block)
            copied.text = label + copied.text
            label = ""
            yield copied
        else:
            if label:
                yield Paragraph(label)
                label = ""
            yield block
    if label:
        yield Paragraph(label)


def _extract_hwp_paragraph_blocks(paragraph, controls: list[object]) -> list[HancomBlock]:
    from ._hancom_conversion_common import _block_owned_text, _normalize_hwp_text
    from ._hancom_from_hwp import _extract_hwp_control
    from ._hancom_hwpx_read import _mark_hwpx_source_paragraph

    if not controls:
        text = _normalize_hwp_text(paragraph.text)
        if not text:
            return []
        return [_mark_hwpx_source_paragraph(_build_hwp_paragraph_block(paragraph, text), paragraph.index, order=0)]

    blocks: list[HancomBlock] = []
    pending_text_parts: list[str] = []
    control_index = 0
    tokens = _tokenize_hwp_paragraph_raw_text(paragraph.raw_text)

    for token_kind, token_value in tokens:
        if token_kind == "text":
            pending_text_parts.append(token_value)
            continue

        match_index = _find_matching_control_index(controls, control_index, token_value)
        if match_index is None:
            continue
        control = controls[match_index]
        control_index = match_index + 1
        block = _extract_hwp_control(control)
        pending_text = "".join(pending_text_parts)
        pending_text_parts.clear()
        normalized_pending = _normalize_hwp_text(pending_text)
        owned_text = _block_owned_text(block)
        if normalized_pending and (owned_text is None or normalized_pending != _normalize_hwp_text(owned_text)):
            blocks.append(_build_hwp_paragraph_block(paragraph, normalized_pending))
        if block is not None:
            blocks.append(block)

    trailing_text = _normalize_hwp_text("".join(pending_text_parts))
    if trailing_text:
        blocks.append(_build_hwp_paragraph_block(paragraph, trailing_text))

    for control in controls[control_index:]:
        block = _extract_hwp_control(control)
        if block is not None:
            blocks.append(block)
    return [
        _mark_hwpx_source_paragraph(block, paragraph.index, order=index)
        for index, block in enumerate(blocks)
    ]


def _extract_supported_nested_hwp_blocks(controls: list[object]) -> list[HancomBlock]:
    from ._hancom_from_hwp import _extract_hwp_control

    blocks: list[HancomBlock] = []
    for control in controls:
        block = _extract_hwp_control(control)
        if isinstance(block, (AutoNumber, Bookmark, Note)):
            blocks.append(block)
    return blocks


def _extract_hwp_nested_paragraph_blocks(paragraph, controls: list[object], paragraph_index: int) -> list[HancomBlock]:
    from ._hancom_hwpx_read import _mark_hwpx_nested_paragraph

    blocks = _extract_hwp_paragraph_blocks(paragraph, controls)
    return [
        _mark_hwpx_nested_paragraph(block, paragraph_index, order=index)
        for index, block in enumerate(blocks)
    ]


def _extract_hwp_section_settings(section) -> SectionSettings:
    from ._hancom_conversion_common import _normalize_hwp_setting_scalar

    settings = section.document.binary_document().section_page_settings(section.index)
    section_definition = section.document.binary_document().section_definition_settings(section.index)
    page_border_fills = section.document.binary_document().section_page_border_fills(section.index)
    page_numbers = section.document.binary_document().section_page_numbers(section.index)
    note_settings = section.document.binary_document().section_note_settings(section.index)
    page_width = settings.get("page_width")
    page_height = settings.get("page_height")
    margins = {
        key: value
        for key, value in settings.items()
        if key in {"left", "right", "top", "bottom", "header", "footer", "gutter"}
    }
    result = SectionSettings(
        page_width=page_width or None,
        page_height=page_height or None,
        landscape="LANDSCAPE" if (page_width or 0) > (page_height or 0) else "WIDELY",
        margins=margins,
        page_border_fills=page_border_fills,
        visibility=dict(section_definition.get("visibility", {})),
        grid=dict(section_definition.get("grid", {})),
        start_numbers=dict(section_definition.get("start_numbers", {})),
        page_numbers=page_numbers,
        footnote_pr=dict(note_settings.get("footNotePr", {})),
        endnote_pr=dict(note_settings.get("endNotePr", {})),
        numbering_shape_id=_normalize_hwp_setting_scalar(section_definition.get("numbering_shape_id")),
        memo_shape_id=_normalize_hwp_setting_scalar(section_definition.get("memo_shape_id")),
    )
    page = next((node for node in section.model().iter_nodes() if node.tag_id == TAG_PAGE_DEF), None)
    if page is not None and len(page.payload) >= 36:
        native_margins = {key: int.from_bytes(page.payload[offset:offset + 4], "little")
                          for key, offset in (("left", 8), ("right", 12), ("top", 16), ("bottom", 20),
                                              ("header", 24), ("footer", 28), ("gutter", 32))}
        setattr(result, "_docx_native_margins", (dict(result.margins), native_margins))
    return result


def _build_hwp_paragraph_block(paragraph, text: str) -> Paragraph:
    from ._hancom_conversion_common import _xml_safe_scalar

    header = getattr(paragraph, "header", None)
    style_id = getattr(header, "style_id", None) if header is not None else None
    para_shape_id = getattr(header, "para_shape_id", None) if header is not None else None
    shapes = next((node for node in header.children if node.tag_id == TAG_PARA_CHAR_SHAPE), None) if header else None
    runs = []
    if shapes is not None:
        entries = list(struct.iter_unpack("<II", shapes.payload)) if len(shapes.payload) % 8 == 0 else []
        raw = paragraph.raw_text.encode("utf-16-le", errors="surrogatepass")
        for index, (start, char_id) in enumerate(entries):
            end = entries[index + 1][0] if index + 1 < len(entries) else len(raw) // 2
            value = raw[start * 2:end * 2].decode("utf-16-le", errors="replace")
            visible = "".join(text for kind, text in _tokenize_hwp_paragraph_raw_text(value) if kind == "text")
            visible = visible.encode("utf-16-le", errors="surrogatepass").decode("utf-16-le", errors="replace")
            runs.append((_xml_safe_scalar(visible.replace("\r", "")), str(char_id)))
        if runs:
            runs[0] = (runs[0][0].lstrip(), runs[0][1])
            runs[-1] = (runs[-1][0].rstrip(), runs[-1][1])
    result = Paragraph(
        text=text,
        style_id=str(style_id) if style_id is not None else None,
        para_pr_id=str(para_shape_id) if para_shape_id is not None else None,
        hwp_para_shape_id=para_shape_id,
        hwp_style_id=style_id,
        hwp_split_flags=getattr(header, "split_flags", None),
        hwp_control_mask=getattr(header, "control_mask", None),
    )
    if runs and "".join(value for value, _ in runs) == text:
        setattr(result, "_docx_runs", runs)
        if len(runs) == 1:
            result.char_pr_id = runs[0][1]
    return result


def _tokenize_hwp_paragraph_raw_text(raw_text: str) -> list[tuple[str, str]]:
    from ._hancom_conversion_common import _drop_surrogate_codepoints

    encoded = _drop_surrogate_codepoints(raw_text).encode("utf-16-le")
    if not encoded:
        return []
    units = list(struct.unpack("<" + "H" * (len(encoded) // 2), encoded))
    tokens: list[tuple[str, str]] = []
    text_buffer: list[str] = []
    index = 0
    while index < len(units):
        value = units[index]
        if value >= 32:
            text_buffer.append(chr(value))
            index += 1
            continue

        if value in (9, 10, 13):
            text_buffer.append(chr(value))

        if (
            index + 7 < len(units)
            and units[index + 3] == 0
            and units[index + 4] == 0
            and units[index + 5] == 0
            and units[index + 6] == 0
        ):
            if text_buffer:
                tokens.append(("text", "".join(text_buffer)))
                text_buffer = []
            token_bytes = struct.pack("<8H", *units[index : index + 8])
            control_id = token_bytes[2:6][::-1].decode("latin1", errors="replace")
            tokens.append(("control", control_id))
            index += 8
            continue

        index += 1

    if text_buffer:
        tokens.append(("text", "".join(text_buffer)))
    return tokens


def _find_matching_control_index(controls: list[object], start_index: int, control_id: str) -> int | None:
    for index in range(start_index, len(controls)):
        if getattr(controls[index], "control_id", "") == control_id:
            return index
    return None


def _attach_hwp_docx_table(control, block: Table) -> None:
    """Retain native cell formatting as hints for the portable Word writer."""
    from .hwp_document import _build_control_wrapper

    definitions = control.document.docinfo_model().border_fill_records()
    fills, borders = {}, {}
    rich: dict[tuple[int, int], tuple[str, list[HancomBlock]]] = {}
    widths = (0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0)
    for cell in control.cells():
        if not 0 < cell.border_fill_id <= len(definitions):
            continue
        payload = definitions[cell.border_fill_id - 1].payload
        if len(payload) >= 40 and int.from_bytes(payload[32:36], "little") & 1:
            fills[(cell.row, cell.column)] = "#" + payload[36:39].hex().upper()
        values = {}
        for side, offset in (("left", 2), ("right", 8), ("top", 14), ("bottom", 20)):
            if len(payload) >= offset + 6:
                values[side] = {
                    "type": {0: "NONE", 1: "SOLID", 2: "DASH", 3: "DOT", 8: "DOUBLE_SLIM"}.get(payload[offset], "SOLID"),
                    "width": str(widths[min(payload[offset + 1], len(widths) - 1)]),
                    "color": "#" + payload[offset + 2:offset + 5].hex().upper(),
                }
        borders[(cell.row, cell.column)] = values
    address = None
    nested: list[HancomBlock] = []
    section_model = control.document.section_model(control.section_index)

    def finish() -> None:
        if address is not None:
            row, col = address
            if row < block.rows and col < block.cols:
                rich[address] = (block.cell_texts[row][col], list(nested))

    for child in control.control_node.children:
        if child.tag_id == TAG_LIST_HEADER and len(child.payload) == 47:
            finish()
            address = (int.from_bytes(child.payload[10:12], "little"), int.from_bytes(child.payload[8:10], "little"))
            nested = []
        elif child.tag_id == TAG_PARA_HEADER and address is not None:
            paragraph = SectionParagraphModel(section_index=control.section_index, index=len(nested), header=child)
            controls: list[object] = []
            for node in child.children:
                if node.tag_id == TAG_CTRL_HEADER:
                    wrapper = _build_control_wrapper(control.document, section_model, paragraph, node, len(controls))
                    if wrapper is not None:
                        controls.append(wrapper)
            nested.extend(_extract_hwp_nested_paragraph_blocks(paragraph, controls, paragraph.index))
    finish()
    setattr(block, "_docx_cell_content", rich)
    setattr(block, "_docx_table_details", {"cell_fill_colors": fills, "cell_border_styles": borders})


def _extract_hwp_table_column_widths(table) -> list[int]:
    widths = [0] * table.column_count
    for cell in table.cells():
        if cell.col_span == 1 and widths[cell.column] == 0:
            widths[cell.column] = cell.width
    fallback = max((value for value in widths if value > 0), default=1000)
    for cell in table.cells():
        if cell.col_span <= 1:
            continue
        missing = [index for index in range(cell.column, cell.column + cell.col_span) if widths[index] == 0]
        if not missing:
            continue
        distributed = max(cell.width // cell.col_span, 1)
        for index in missing:
            widths[index] = distributed
    return [value or fallback for value in widths]


def _apply_hwp_shape_specific_fields(shape, block: Shape) -> None:
    fields = dict(block.specific_fields)
    if not fields:
        return
    if isinstance(shape, (HwpLineShapeObject, HwpConnectLineShapeObject)):
        start = fields.get("start")
        end = fields.get("end")
        attributes = fields.get("attributes")
        if start is not None or end is not None or attributes is not None:
            shape.set_endpoints(start=start, end=end, attributes=int(cast(int, attributes)) if attributes is not None else None)
        return
    if isinstance(shape, HwpRectangleShapeObject) and "corner_radius" in fields:
        shape.set_corner_radius(int(cast(int, fields["corner_radius"])))
        return
    if isinstance(shape, HwpEllipseShapeObject):
        kwargs = {key: fields[key] for key in ("arc_flags", "center", "axis1", "axis2", "start", "end", "start_2", "end_2") if key in fields}
        if kwargs:
            shape.set_geometry(**kwargs)
        return
    if isinstance(shape, HwpArcShapeObject):
        kwargs = {key: fields[key] for key in ("arc_type", "center", "axis1", "axis2") if key in fields}
        if kwargs:
            shape.set_geometry(**kwargs)
        return
    if isinstance(shape, HwpPolygonShapeObject) and isinstance(fields.get("points"), list):
        shape.set_points(fields["points"], trailing_payload=fields.get("trailing_payload") if isinstance(fields.get("trailing_payload"), bytes) else None)
        return
    raw_payload = fields.get("raw_payload")
    if isinstance(shape, HwpCurveShapeObject) and isinstance(raw_payload, (bytes, bytearray)):
        shape.set_raw_payload(bytes(raw_payload))
        return
    if isinstance(shape, HwpContainerShapeObject) and isinstance(raw_payload, (bytes, bytearray)):
        shape.set_raw_payload(bytes(raw_payload))
        return
    if isinstance(shape, HwpTextArtShapeObject):
        if "text" in fields:
            shape.set_text(str(fields["text"]))
        if "font_name" in fields or "font_style" in fields:
            shape.set_style(
                font_name=str(fields["font_name"]) if "font_name" in fields else None,
                font_style=str(fields["font_style"]) if "font_style" in fields else None,
            )


def _parse_hwp_field_parameters(control) -> dict[str, str | int]:
    data_node = next((child for child in control.control_node.children if child.tag_id == TAG_CTRL_DATA), None)
    if data_node is None:
        return {}
    decoded = data_node.payload.decode("utf-16-le", errors="ignore")
    parameters: dict[str, str | int] = {}
    for entry in decoded.split(";"):
        if not entry or "=" not in entry:
            continue
        key, value = entry.split("=", 1)
        if value.isdigit():
            parameters[key] = int(value)
        else:
            parameters[key] = value
    return parameters


def _resolve_hwp_paragraph_attributes(
    document: HwpDocument,
    block: Paragraph,
    *,
    paragraph_attribute_limits: tuple[int, int] | None = None,
) -> tuple[int, int, int, int]:
    if paragraph_attribute_limits is None:
        docinfo = document.docinfo_model()
        para_shape_count = len(docinfo.para_shape_records())
        style_count = len(docinfo.style_records())
    else:
        para_shape_count, style_count = paragraph_attribute_limits
    para_shape_id = block.hwp_para_shape_id
    style_id = block.hwp_style_id
    if para_shape_id is None:
        para_shape_id = _parse_optional_int(block.para_pr_id)
    if style_id is None:
        style_id = _parse_optional_int(block.style_id)
    resolved_para_shape_id = (
        para_shape_id
        if para_shape_id is not None and 0 <= para_shape_id < para_shape_count and para_shape_id <= 0xFFFF
        else 0
    )
    resolved_style_id = (
        style_id
        if style_id is not None and 0 <= style_id < style_count and style_id <= 0xFF
        else 0
    )
    return (
        resolved_para_shape_id,
        resolved_style_id,
        block.hwp_split_flags or 0,
        block.hwp_control_mask or 0,
    )


def _extract_hwp_header_footer(control) -> HeaderFooter | None:
    text = getattr(control, "text", "")
    if not text:
        text = ""
    kind = getattr(control, "kind", "")
    if kind not in {"header", "footer"}:
        return None
    nested_blocks = _extract_hwp_header_footer_nested_blocks(control)
    if not text and not nested_blocks:
        return None
    return HeaderFooter(
        kind=kind,
        text=text,
        apply_page_type=getattr(control, "apply_page_type", "BOTH") or "BOTH",
        nested_blocks=cast(list[object], nested_blocks),
    )


def _extract_hwp_header_footer_nested_blocks(control) -> list[HancomBlock]:
    from .hwp_document import _build_control_wrapper

    blocks: list[HancomBlock] = []
    paragraph_index = 0
    section_model = control.document.section_model(control.section_index)
    for child in control.control_node.children:
        if not isinstance(child, ParagraphHeaderRecord):
            continue
        paragraph = SectionParagraphModel(
            section_index=control.section_index,
            index=paragraph_index,
            header=child,
        )
        control_ordinal = 0
        controls: list[object] = []
        for nested_child in child.children:
            if nested_child.tag_id != TAG_CTRL_HEADER:
                continue
            wrapper = _build_control_wrapper(control.document, section_model, paragraph, nested_child, control_ordinal)
            if wrapper is not None:
                controls.append(wrapper)
                control_ordinal += 1
        blocks.extend(_extract_hwp_nested_paragraph_blocks(paragraph, controls, paragraph_index))
        paragraph_index += 1
    return blocks
