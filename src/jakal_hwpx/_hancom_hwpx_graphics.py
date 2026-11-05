"""HWPX object geometry, table cells, pictures, equations, and shapes."""
from __future__ import annotations

from collections.abc import Mapping
from pathlib import Path
from typing import cast

from lxml import etree

from ._hancom_models import (
    Chart,
    Equation,
    Field,
    Memo,
    MemoShapeDefinition,
    Note,
    Ole,
    Picture,
    Shape,
    Table,
)
from ._hancom_styles import _coerce_bool
from .document import HwpxDocument
from .namespaces import (
    NS,
    qname,
)


def _optional_hwpx_int_attr(node: etree._Element | None, attr_name: str) -> int | None:
    if node is None:
        return None
    value = node.get(attr_name)
    if value in (None, ""):
        return None
    try:
        return int(value)
    except ValueError:
        return None


def _extract_hwpx_graphic_size_fields(element: etree._Element, *, width: int, height: int) -> dict[str, int | None]:
    original_size = element.xpath("./hp:orgSz", namespaces=NS)
    current_size = element.xpath("./hp:curSz", namespaces=NS)
    original = original_size[0] if original_size else None
    current = current_size[0] if current_size else None
    original_width = _optional_hwpx_int_attr(original, "width")
    original_height = _optional_hwpx_int_attr(original, "height")
    current_width = _optional_hwpx_int_attr(current, "width")
    current_height = _optional_hwpx_int_attr(current, "height")
    return {
        "original_width": original_width if original_width != width else None,
        "original_height": original_height if original_height != height else None,
        "current_width": current_width if current_width != width else None,
        "current_height": current_height if current_height != height else None,
    }


def _extract_hwpx_graphic_size_attributes(element: etree._Element) -> dict[str, str]:
    nodes = element.xpath("./hp:sz", namespaces=NS)
    if not nodes:
        return {}
    return {str(key): str(value) for key, value in nodes[0].attrib.items()}


def _apply_hwpx_graphic_size_node(wrapper, *, has_size_node: bool, size_attributes: dict[str, str]) -> None:
    nodes = wrapper.element.xpath("./hp:sz", namespaces=NS)
    if not has_size_node:
        for node in nodes:
            parent = node.getparent()
            if parent is not None:
                parent.remove(node)
        if nodes:
            wrapper.section.mark_modified()
        return
    if not size_attributes:
        return
    node = nodes[0] if nodes else etree.SubElement(wrapper.element, qname("hp", "sz"))
    for key, value in size_attributes.items():
        node.set(str(key), str(value))
    wrapper.section.mark_modified()


def _remove_hwpx_child_nodes(wrapper, xpath: str) -> None:
    nodes = wrapper.element.xpath(xpath, namespaces=NS)
    for node in nodes:
        parent = node.getparent()
        if parent is not None:
            parent.remove(node)
    if nodes:
        wrapper.section.mark_modified()


def _apply_hwpx_optional_graphic_nodes(
    wrapper,
    *,
    has_position_node: bool,
    has_out_margin_node: bool,
) -> None:
    if not has_position_node:
        _remove_hwpx_child_nodes(wrapper, "./hp:pos")
    if not has_out_margin_node:
        _remove_hwpx_child_nodes(wrapper, "./hp:outMargin")


def _hwpx_layout_kwargs(values: Mapping[str, object] | None) -> dict[str, object]:
    if not values:
        return {}
    mapping = {
        "textWrap": "text_wrap",
        "textFlow": "text_flow",
        "treatAsChar": "treat_as_char",
        "affectLSpacing": "affect_line_spacing",
        "flowWithText": "flow_with_text",
        "allowOverlap": "allow_overlap",
        "holdAnchorAndSO": "hold_anchor_and_so",
        "vertRelTo": "vert_rel_to",
        "horzRelTo": "horz_rel_to",
        "vertAlign": "vert_align",
        "horzAlign": "horz_align",
        "vertOffset": "vert_offset",
        "horzOffset": "horz_offset",
    }
    bool_fields = {
        "treatAsChar",
        "affectLSpacing",
        "flowWithText",
        "allowOverlap",
        "holdAnchorAndSO",
    }
    result: dict[str, object] = {}
    for source, target in mapping.items():
        value = values.get(source)
        if value in (None, ""):
            continue
        if source in bool_fields:
            coerced = _coerce_bool(value)
            if coerced is None:
                continue
            result[target] = coerced
            continue
        result[target] = value
    return result


def _hwpx_rotation_kwargs(values: Mapping[str, object] | None) -> dict[str, object]:
    if not values:
        return {}
    mapping = {
        "angle": "angle",
        "centerX": "center_x",
        "centerY": "center_y",
        "rotateimage": "rotate_image",
    }
    return {
        target: value
        for source, target in mapping.items()
        if (value := values.get(source)) not in (None, "")
    }


def _hwpx_line_style_kwargs(values: Mapping[str, object] | None) -> dict[str, object]:
    if not values:
        return {}
    mapping = {
        "color": "color",
        "width": "width",
        "style": "style",
        "endCap": "end_cap",
        "headStyle": "head_style",
        "tailStyle": "tail_style",
        "headfill": "head_fill",
        "tailfill": "tail_fill",
        "headSz": "head_size",
        "tailSz": "tail_size",
        "outlineStyle": "outline_style",
        "alpha": "alpha",
    }
    bool_fields = {"headfill", "tailfill"}
    result: dict[str, object] = {}
    for source, target in mapping.items():
        value = values.get(source)
        if value in (None, ""):
            continue
        if source in bool_fields:
            coerced = _coerce_bool(value)
            if coerced is None:
                continue
            result[target] = coerced
            continue
        result[target] = value
    return result


def _extract_hwpx_shape_points(node: etree._Element) -> list[dict[str, int]]:
    points: list[dict[str, int]] = []
    point_names = {"pt", "startPt", "endPt", "center", "ax1", "ax2", "start1", "start2", "end1", "end2"}
    for child in node:
        local_name = etree.QName(child).localname
        if local_name.startswith("pt") or local_name in point_names:
            try:
                points.append(
                    {
                        "x": int(child.get("x", "0")),
                        "y": int(child.get("y", "0")),
                    }
                )
            except ValueError:
                continue
    return points


def _extract_hwpx_shape_specific_fields(shape) -> dict[str, object]:
    points = _extract_hwpx_shape_points(shape.element)
    if shape.kind in {"line", "connectLine"} and len(points) >= 2:
        return {
            "start": points[0],
            "end": points[1],
        }
    if shape.kind == "polygon" and points:
        return {
            "point_count": len(points),
            "points": points,
        }
    return {}


_NESTED_TEXT_CONTROL_NAMES = {
    "tbl",
    "pic",
    "ole",
    "equation",
    "chart",
    "rect",
    "line",
    "ellipse",
    "arc",
    "polygon",
    "curve",
    "connectLine",
    "textart",
    "container",
    "fieldBegin",
    "fieldEnd",
    "header",
    "footer",
    "footNote",
    "endNote",
    "hiddenComment",
    "autoNum",
    "newNum",
    "bookmark",
}


def _extract_hwpx_direct_text(scope: etree._Element) -> str:
    chunks: list[str] = []
    for text_node in scope.xpath(".//hp:t", namespaces=NS):
        parent = text_node.getparent()
        skip = False
        while parent is not None and parent is not scope:
            if etree.QName(parent).localname in _NESTED_TEXT_CONTROL_NAMES:
                skip = True
                break
            parent = parent.getparent()
        if not skip:
            chunks.append(text_node.text or "")
    return "".join(chunks)


def _extract_hwpx_shape_direct_text(shape) -> str:
    direct_text = shape.element.get("text")
    if direct_text is not None:
        return direct_text
    draw_text_nodes = shape.element.xpath("./hp:drawText", namespaces=NS)
    if not draw_text_nodes:
        return ""
    return _extract_hwpx_direct_text(draw_text_nodes[0])


def _set_hwpx_shape_points(node: etree._Element, points: list[dict[str, int]]) -> None:
    kind = etree.QName(node).localname
    point_names = {"pt", "startPt", "endPt", "center", "ax1", "ax2", "start1", "start2", "end1", "end2"}
    for child in list(node):
        local_name = etree.QName(child).localname
        if local_name.startswith("pt") or local_name in point_names:
            node.remove(child)
    for index, point in enumerate(points):
        if kind in {"line", "connectLine"}:
            point_name = "startPt" if index == 0 else "endPt"
        elif kind == "polygon":
            point_name = "pt"
        else:
            point_name = f"pt{index}"
        child = etree.SubElement(node, qname("hc", point_name))
        child.set("x", str(int(point.get("x", 0))))
        child.set("y", str(int(point.get("y", 0))))
        if kind == "connectLine":
            child.set("subjectIDRef", str(int(point.get("subjectIDRef", 0))))
            child.set("subjectIdx", str(int(point.get("subjectIdx", 0))))


def _apply_hwpx_shape_specific_fields(shape, block: Shape) -> None:
    fields = dict(block.specific_fields)
    if not fields:
        return
    if block.kind in {"line", "connectLine"}:
        start = fields.get("start", {"x": 0, "y": 0})
        end = fields.get("end", {"x": 0, "y": 0})
        if isinstance(start, dict) and isinstance(end, dict):
            _set_hwpx_shape_points(shape.element, [start, end])
            shape.section.mark_modified()
        return
    raw_points = fields.get("points")
    if block.kind == "polygon" and isinstance(raw_points, list):
        points = [point for point in raw_points if isinstance(point, dict)]
        if points:
            _set_hwpx_shape_points(shape.element, points)
            shape.section.mark_modified()


def _apply_hwpx_table_block(table, block: Table) -> None:
    from ._hancom_conversion_common import _normalize_int_map
    from ._hancom_hwpx_write import _build_hwpx_nested_sublist

    anchor = getattr(block, "_hwpx_anchor_paragraph", None)
    if anchor is not None:
        paragraph = next(table.element.iterancestors(qname("hp", "p")), None)
        if paragraph is not None:
            if anchor.para_pr_id is not None:
                paragraph.set("paraPrIDRef", anchor.para_pr_id)
            if anchor.style_id is not None:
                paragraph.set("styleIDRef", anchor.style_id)
            if anchor.char_pr_id is not None:
                table.element.getparent().set("charPrIDRef", anchor.char_pr_id)

    for (row, column), (row_span, col_span) in (block.cell_spans or {}).items():
        table.merge_cells(row, column, row + row_span - 1, column + col_span - 1)
    native = getattr(block, "_hwpx_table_properties", {})
    if native and native.get("rows", block.rows) == block.rows and native.get("cols", block.cols) == block.cols and native.get("spans", block.cell_spans or {}) == (block.cell_spans or {}):
        from copy import deepcopy
        table.element.attrib.update(native["attributes"])
        for raw in native["children"]:
            node = etree.fromstring(raw)
            existing = table.element.find(node.tag)
            if existing is not None:
                table.element.replace(existing, node)
        for cell in table.cells():
            record = native["cells"].get((cell.row, cell.column))
            if record is None:
                continue
            cell.element.attrib.update(record["attributes"])
            sublist = cell.element.find(qname("hp", "subList"))
            if sublist is not None:
                sublist.attrib.update(record["sublist"])
            for raw in record["children"]:
                node = etree.fromstring(raw)
                existing = cell.element.find(node.tag)
                if existing is not None:
                    cell.element.replace(existing, deepcopy(node))
        table.section.mark_modified()
    for cell in table.cells():
        size = cell.element.find(qname("hp", "cellSz"))
        if size is None:
            continue
        if block.col_widths:
            size.set("width", str(sum(block.col_widths[cell.column:cell.column + cell.col_span])))
        if block.row_heights:
            heights = block.row_heights[cell.row:cell.row + cell.row_span]
            if all(height > 0 for height in heights):
                size.set("height", str(sum(heights)))
    size = table.element.find(qname("hp", "sz"))
    if size is not None:
        if block.col_widths:
            size.set("width", str(sum(block.col_widths)))
        if block.row_heights and all(height > 0 for height in block.row_heights):
            size.set("height", str(sum(block.row_heights)))
    table.section.mark_modified()
    if block.layout:
        table.set_layout(**_hwpx_layout_kwargs(block.layout))
    if block.out_margins:
        table.set_out_margins(**block.out_margins)
    if block.table_margins:
        table.set_in_margins(**block.table_margins)
    if block.cell_spacing is not None:
        table.set_cell_spacing(block.cell_spacing)
    table.set_table_border_fill_id(block.table_border_fill_id)
    if getattr(block, "_docx_new_table", False) and not getattr(block, "_hwpx_table_properties", None):
        # New HWPX tables reference borderFill 1 before any definitions exist.
        # If the first definition is a shaded header, transparent body cells
        # inherit that shade from the whole table. Reserve a neutral base.
        neutral = _ensure_hwpx_table_border_fill(table.document, None, {})
        table.set_table_border_fill_id(neutral)
        for cell in table.cells():
            cell.set_border_fill_id(neutral)
    if block.page_break:
        table.set_page_break(block.page_break)
    if block.repeat_header is not None:
        table.set_repeat_header(block.repeat_header)
    coordinates = {(cell.row, cell.column) for cell in table.cells()}
    for (row, column), border_fill_id in (block.cell_border_fill_ids or {}).items():
        if (row, column) in coordinates:
            table.cell(row, column).set_border_fill_id(border_fill_id)
    pattern_coordinates = {tuple(map(int, key.split(","))) for key in getattr(block, "_word_extra", {}).get("patterns", {})}
    for row, column in set(block.cell_fill_colors or {}) | set(block.cell_border_styles or {}) | pattern_coordinates:
        if (row, column) not in coordinates:
            continue
        color = (block.cell_fill_colors or {}).get((row, column))
        borders = (block.cell_border_styles or {}).get((row, column), {})
        pattern = getattr(block, "_word_extra", {}).get("patterns", {}).get(f"{row},{column}")
        border_fill_id = _ensure_hwpx_table_border_fill(table.document, color, borders, pattern=pattern)
        table.cell(row, column).set_border_fill_id(border_fill_id)
    for (row, column), margins in (block.cell_margins or {}).items():
        if (row, column) in coordinates:
            table.cell(row, column).set_margins(**_normalize_int_map(margins))
    for (row, column), vertical_align in (block.cell_vertical_aligns or {}).items():
        if (row, column) in coordinates:
            table.cell(row, column).set_vertical_align(vertical_align)
    for (row, column), (baseline, blocks) in getattr(block, "_docx_cell_content", {}).items():
        if (row, column) not in coordinates or block.cell_texts[row][column] != baseline:
            continue
        cell = table.cell(row, column)
        sublist = cell.element.find(qname("hp", "subList"))
        if sublist is None:
            continue
        rich = _build_hwpx_nested_sublist(table.document, blocks, vertical_align=cell.vertical_align)
        for child in list(sublist):
            sublist.remove(child)
        for child in list(rich):
            sublist.append(child)
        table.section.mark_modified()


def _ensure_hwpx_table_border_fill(
    document: HwpxDocument,
    fill_color: str | None,
    borders: dict[str, dict[str, str | int]],
    *, pattern: dict[str, str] | None = None,
) -> int:
    key = (fill_color, tuple((side, tuple(sorted(spec.items()))) for side, spec in sorted(borders.items())),
           tuple(sorted((pattern or {}).items())))
    cache = getattr(document, "_docx_table_border_fill_cache", None)
    if cache is None:
        cache = {}
        setattr(document, "_docx_table_border_fill_cache", cache)
    if key in cache:
        return cache[key]
    ref_list = document.header.root_element.find(qname("hh", "refList"))
    border_fills = ref_list.find(qname("hh", "borderFills"))
    if border_fills is None:
        border_fills = etree.Element(qname("hh", "borderFills"), itemCnt="0")
        char_properties = ref_list.find(qname("hh", "charProperties"))
        ref_list.insert(ref_list.index(char_properties) if char_properties is not None else 0, border_fills)
    existing = [int(node.get("id", "0")) for node in border_fills]
    style_id = max(existing, default=0) + 1
    node = etree.SubElement(
        border_fills, qname("hh", "borderFill"), id=str(style_id),
        threeD="0", shadow="0", centerLine="NONE", breakCellSeparateLine="0",
    )
    etree.SubElement(node, qname("hh", "slash"), type="NONE", Crooked="0", isCounter="0")
    etree.SubElement(node, qname("hh", "backSlash"), type="NONE", Crooked="0", isCounter="0")
    line_types = {"single": "SOLID", "dashed": "DASH", "dotted": "DOT", "double": "DOUBLE_SLIM", "none": "NONE", "nil": "NONE"}
    for side in ("left", "right", "top", "bottom"):
        spec = borders.get(side, {})
        line_type = line_types.get(str(spec.get("style", "none")), "SOLID")
        width_mm = max(0.1, int(spec.get("size", 4)) * 0.3527778 / 8)
        color = str(spec.get("color", "000000")).lstrip("#")
        etree.SubElement(
            node, qname("hh", f"{side}Border"), type=line_type,
            width=f"{width_mm:.2f} mm", color=f"#{color.upper()}",
        )
    etree.SubElement(node, qname("hh", "diagonal"), type="NONE", width="0.1 mm", color="#000000")
    if fill_color:
        brush = etree.SubElement(node, qname("hc", "fillBrush"))
        etree.SubElement(brush, qname("hc", "winBrush"), faceColor=fill_color, hatchColor="#000000", alpha="0")
    if pattern:
        from .addons.docx.formatting import apply_hwpx_pattern

        apply_hwpx_pattern(document, node, pattern)
    border_fills.set("itemCnt", str(len(border_fills)))
    document.header.mark_modified()
    cache[key] = style_id
    return style_id


def _apply_hwpx_picture_block(picture, block: Picture) -> None:
    drawing = getattr(block, "_word_extra", {}).get("drawing")
    if drawing:
        source = etree.fromstring(drawing.encode("utf-8"))
        transform = source.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm")
        flip = picture.element.find(qname("hp", "flip"))
        if transform is not None and flip is not None:
            flip.set("horizontal", "1" if transform.get("flipH", "0") in {"1", "true"} else "0")
            flip.set("vertical", "1" if transform.get("flipV", "0") in {"1", "true"} else "0")
    if any(
        value is not None
        for value in (block.original_width, block.original_height, block.current_width, block.current_height)
    ):
        picture.set_size(
            width=block.width,
            height=block.height,
            original_width=block.original_width,
            original_height=block.original_height,
            current_width=block.current_width,
            current_height=block.current_height,
        )
    _apply_hwpx_graphic_size_node(
        picture,
        has_size_node=block.has_size_node,
        size_attributes=block.size_attributes,
    )
    if block.shape_comment:
        picture.shape_comment = block.shape_comment
    if block.layout:
        picture.set_layout(**_hwpx_layout_kwargs(block.layout))
    if block.out_margins:
        picture.set_out_margins(**block.out_margins)
    if block.rotation:
        picture.set_rotation(**_hwpx_rotation_kwargs(block.rotation))
    if block.image_adjustment:
        picture.set_image_adjustment(**block.image_adjustment)
    if block.crop:
        picture.set_crop(**block.crop)
    line_style_kwargs = _hwpx_line_style_kwargs(block.line_style)
    if line_style_kwargs:
        picture.set_line_style(**line_style_kwargs)
    elif block.has_line_node or block.line_width > 0 or block.line_color != "#000000":
        picture.set_line_style(color=block.line_color, width=block.line_width, style="SOLID" if block.line_width > 0 else "NONE")
    _apply_hwpx_optional_graphic_nodes(
        picture,
        has_position_node=block.has_position_node,
        has_out_margin_node=block.has_out_margin_node,
    )


def _apply_hwpx_equation_block(equation, block: Equation) -> None:
    unsupported = getattr(block, "_unsupported_word_math", None)
    if unsupported:
        raise ValueError(unsupported)
    if hasattr(block, "_hwpx_equation_baseline"):
        baseline = block._hwpx_equation_baseline
        if baseline is None:
            equation.element.attrib.pop("baseLine", None)
        else:
            equation.element.set("baseLine", baseline)
        equation.section.mark_modified()
    character = getattr(block, "_hwpx_char_pr_id", None)
    if character is not None:
        equation.element.getparent().set("charPrIDRef", character)
        equation.section.mark_modified()
    if block.layout:
        equation.set_layout(**_hwpx_layout_kwargs(block.layout))
    if block.out_margins:
        equation.set_out_margins(**block.out_margins)
    if block.rotation:
        equation.set_rotation(**_hwpx_rotation_kwargs(block.rotation))


def _apply_hwpx_shape_block(shape, block: Shape) -> None:
    if any(
        value is not None
        for value in (block.original_width, block.original_height, block.current_width, block.current_height)
    ):
        shape.set_size(
            width=block.width,
            height=block.height,
            original_width=block.original_width,
            original_height=block.original_height,
            current_width=block.current_width,
            current_height=block.current_height,
        )
    _apply_hwpx_graphic_size_node(
        shape,
        has_size_node=block.has_size_node,
        size_attributes=block.size_attributes,
    )
    if block.kind == "textart" and block.has_text_margin_node and block.text:
        _ensure_hwpx_shape_draw_text(shape, block.text)
    if block.shape_comment:
        shape.shape_comment = block.shape_comment
    line_style_kwargs = _hwpx_line_style_kwargs(block.line_style)
    if line_style_kwargs:
        shape.set_line_style(**line_style_kwargs)
    else:
        shape.set_line_style(color=block.line_color, width=block.line_width, style="SOLID" if block.line_width > 0 else "NONE")
    if block.layout:
        shape.set_layout(**_hwpx_layout_kwargs(block.layout))
    if block.out_margins:
        shape.set_out_margins(**block.out_margins)
    if block.rotation:
        shape.set_rotation(**_hwpx_rotation_kwargs(block.rotation))
    if block.text_margins:
        shape.set_text_margins(**block.text_margins)
    _apply_hwpx_shape_specific_fields(shape, block)
    if not block.has_fill_node:
        _remove_hwpx_child_nodes(shape, "./hc:fillBrush")
    if not block.has_line_node:
        _remove_hwpx_child_nodes(shape, "./hp:lineShape")
    if not block.has_text_margin_node:
        _remove_hwpx_child_nodes(shape, "./hp:drawText/hp:textMargin")
    _apply_hwpx_optional_graphic_nodes(
        shape,
        has_position_node=block.has_position_node,
        has_out_margin_node=block.has_out_margin_node,
    )


def _ensure_hwpx_shape_draw_text(shape, text: str) -> None:
    if shape.element.xpath("./hp:drawText", namespaces=NS):
        return
    draw_text = etree.SubElement(shape.element, qname("hp", "drawText"))
    draw_text.set("lastWidth", str(shape.size().get("width", 12000)))
    draw_text.set("name", "")
    draw_text.set("editable", "0")

    sub_list = etree.SubElement(draw_text, qname("hp", "subList"))
    sub_list.set("id", "")
    sub_list.set("textDirection", "HORIZONTAL")
    sub_list.set("lineWrap", "BREAK")
    sub_list.set("vertAlign", "CENTER")
    sub_list.set("linkListIDRef", "0")
    sub_list.set("linkListNextIDRef", "0")
    sub_list.set("textWidth", "0")
    sub_list.set("textHeight", "0")
    sub_list.set("hasTextRef", "0")
    sub_list.set("hasNumRef", "0")

    paragraph = etree.SubElement(sub_list, qname("hp", "p"))
    paragraph.set("id", "0")
    paragraph.set("paraPrIDRef", "0")
    paragraph.set("styleIDRef", "0")
    paragraph.set("pageBreak", "0")
    paragraph.set("columnBreak", "0")
    paragraph.set("merged", "0")
    run = etree.SubElement(paragraph, qname("hp", "run"))
    run.set("charPrIDRef", "0")
    text_node = etree.SubElement(run, qname("hp", "t"))
    text_node.text = text

    text_margin = etree.SubElement(draw_text, qname("hp", "textMargin"))
    text_margin.set("left", "283")
    text_margin.set("right", "283")
    text_margin.set("top", "283")
    text_margin.set("bottom", "283")
    shape.section.mark_modified()


def _apply_hwpx_ole_block(ole, block: Ole) -> None:
    if any(
        value is not None
        for value in (block.original_width, block.original_height, block.current_width, block.current_height)
    ):
        ole.set_size(
            width=block.width,
            height=block.height,
            original_width=block.original_width,
            original_height=block.original_height,
            current_width=block.current_width,
            current_height=block.current_height,
        )
    _apply_hwpx_graphic_size_node(
        ole,
        has_size_node=block.has_size_node,
        size_attributes=block.size_attributes,
    )
    if block.shape_comment:
        ole.shape_comment = block.shape_comment
    if block.layout:
        ole.set_layout(**_hwpx_layout_kwargs(block.layout))
    if block.out_margins:
        ole.set_out_margins(**block.out_margins)
    if block.rotation:
        ole.set_rotation(**_hwpx_rotation_kwargs(block.rotation))
    if block.extent:
        ole.set_extent(**block.extent)
    line_style_kwargs = _hwpx_line_style_kwargs(block.line_style)
    if line_style_kwargs:
        ole.set_line_style(**line_style_kwargs)
    else:
        ole.set_line_style(color=block.line_color, width=block.line_width, style="SOLID" if block.line_width > 0 else "NONE")
    _apply_hwpx_optional_graphic_nodes(
        ole,
        has_position_node=block.has_position_node,
        has_out_margin_node=block.has_out_margin_node,
    )


def _extract_hwpx_table(table) -> Table:
    from ._hancom_conversion_common import _blank_cell_texts, _normalize_int_map, _normalize_str_map
    from ._hancom_hwpx_read import _extract_hwpx_nested_paragraph_blocks

    cell_texts = _blank_cell_texts(table.row_count, table.column_count)
    cell_spans: dict[tuple[int, int], tuple[int, int]] = {}
    cells = table.cells()
    for cell in cells:
        if cell.row < table.row_count and cell.column < table.column_count:
            cell_texts[cell.row][cell.column] = cell.text
            if cell.row_span != 1 or cell.col_span != 1:
                cell_spans[(cell.row, cell.column)] = (cell.row_span, cell.col_span)
    result = Table(
        rows=table.row_count,
        cols=table.column_count,
        cell_texts=cell_texts,
        cell_border_fill_ids={(cell.row, cell.column): cell.border_fill_id for cell in cells} or None,
        cell_margins={(cell.row, cell.column): cell.margins for cell in cells if cell.element.get("hasMargin") == "1"} or None,
        cell_vertical_aligns={
            (cell.row, cell.column): cell.vertical_align for cell in cells if cell.vertical_align != "CENTER"
        }
        or None,
        cell_spans=cell_spans or None,
        table_border_fill_id=table.table_border_fill_id,
        cell_spacing=table.cell_spacing,
        table_margins=table.in_margins() or None,
        layout=_normalize_str_map(table.layout()),
        out_margins=_normalize_int_map(table.out_margins()),
        page_break=table.page_break,
        repeat_header=table.repeat_header,
    )
    # Keep rich cell content for Word output, while cell_texts remains the public
    # editable view. Export only uses this content while the cell text matches.
    rich = {}
    widths = [0] * result.cols
    heights = [0] * result.rows
    fills = {}
    borders = {}
    definitions = {node.get("id"): node for node in table.document.header.readonly_root_element.xpath(".//hh:borderFill", namespaces=NS)}
    for cell in cells:
        content = []
        for index, paragraph in enumerate(cell.element.xpath("./hp:subList/hp:p", namespaces=NS)):
            content.extend(_extract_hwpx_nested_paragraph_blocks(table.section, paragraph, index))
        rich[(cell.row, cell.column)] = (cell.text, content)
        size = cell.element.find(qname("hp", "cellSz"))
        if size is not None:
            if cell.col_span == 1 and cell.column < len(widths):
                widths[cell.column] = int(size.get("width", "0"))
            if cell.row_span == 1 and cell.row < len(heights):
                heights[cell.row] = max(heights[cell.row], int(size.get("height", "0")))
        definition = definitions.get(str(cell.border_fill_id))
        if definition is not None:
            colors = definition.xpath("./hc:fillBrush/hc:winBrush/@faceColor", namespaces=NS)
            if colors and colors[0] not in {"none", "NONE"}:
                fills[(cell.row, cell.column)] = colors[0]
            values = {}
            for side in ("left", "right", "top", "bottom"):
                edge = definition.find(qname("hh", side + "Border"))
                if edge is not None:
                    values[side] = dict(edge.attrib)
            if values:
                borders[(cell.row, cell.column)] = values
    # These details are Word-export hints, not changes to the existing shared
    # authoring DSL or its native HWP/HWPX serialization behavior.
    setattr(result, "_docx_table_details", {
        "col_widths": widths if all(widths) else None,
        "row_heights": heights if any(heights) else None,
        "cell_fill_colors": fills, "cell_border_styles": borders,
    })
    setattr(result, "_docx_cell_content", rich)
    setattr(result, "_hwpx_table_properties", {
        "rows": result.rows, "cols": result.cols, "spans": result.cell_spans or {},
        "attributes": dict(table.element.attrib),
        "children": [etree.tostring(child) for child in table.element if etree.QName(child).localname in {"sz", "pos", "outMargin", "inMargin"}],
        "cells": {(cell.row, cell.column): {
            "attributes": dict(cell.element.attrib),
            "sublist": dict(cell.element.find(qname("hp", "subList")).attrib),
            "children": [etree.tostring(child) for child in cell.element if etree.QName(child).localname in {"cellAddr", "cellSpan", "cellSz", "cellMargin"}],
        } for cell in cells},
    })
    return result


def _extract_hwpx_picture(picture) -> Picture:
    from ._hancom_conversion_common import _normalize_int_map, _normalize_str_map

    path = picture.binary_part_path()
    extension = Path(path).suffix.lstrip(".") or None
    size = picture.size()
    line_style = picture.line_style()
    size_nodes = picture.element.xpath("./hp:sz", namespaces=NS)
    width = size.get("width", 7200)
    height = size.get("height", 7200)
    result = Picture(
        name=Path(path).name,
        data=picture.binary_data(),
        extension=extension,
        width=width,
        height=height,
        shape_comment=picture.shape_comment or None,
        layout=_normalize_str_map(picture.layout()),
        out_margins=_normalize_int_map(picture.out_margins()),
        rotation=_normalize_str_map(picture.rotation()),
        image_adjustment=_normalize_str_map(picture.image_adjustment()),
        crop=_normalize_int_map(picture.crop()) if hasattr(picture, "crop") else {},
        line_color=line_style.get("color", "#000000"),
        line_width=int(line_style.get("width") or "0"),
        line_style=_normalize_str_map(line_style),
        has_size_node=bool(size_nodes),
        size_attributes=_extract_hwpx_graphic_size_attributes(picture.element),
        has_line_node=bool(picture.element.xpath("./hp:lineShape", namespaces=NS)),
        has_position_node=bool(picture.element.xpath("./hp:pos", namespaces=NS)),
        has_out_margin_node=bool(picture.element.xpath("./hp:outMargin", namespaces=NS)),
        **_extract_hwpx_graphic_size_fields(picture.element, width=width, height=height),
    )
    flips = picture.element.find(qname("hp", "flip"))
    if flips is not None:
        setattr(result, "_native_picture_flip", dict(flips.attrib))
    image_dimension = picture.element.find(qname("hp", "imgDim"))
    if image_dimension is not None:
        setattr(result, "_native_picture_image_dim", dict(image_dimension.attrib))
    return result


def _extract_hwpx_note(note) -> Note:
    from ._hancom_hwpx_read import _extract_hwpx_text_container_blocks

    number = int(note.number) if note.number and str(note.number).isdigit() else None
    blocks = _extract_hwpx_text_container_blocks(note.section, note.element)
    formatting = next((block for block in blocks if isinstance(block, Field) and block.field_type == "JAKAL_NOTE_STYLE"), None)
    result = Note(
        kind=note.kind,
        text=note.text,
        number=number,
        nested_blocks=[block for block in blocks if block is not formatting],
    )
    runs = note.element.xpath("ancestor::hp:run[1]", namespaces=NS)
    if runs:
        setattr(result, "_docx_char_pr_id", runs[0].get("charPrIDRef"))
    if formatting is not None:
        reference = formatting.parameters.get("ReferenceRPr")
        marker = formatting.parameters.get("MarkerRPr")
        result.reference_run_properties = str(reference) if reference is not None else None
        result.marker_run_properties = str(marker) if marker is not None else None
    return result


def _extract_hwpx_memo(memo) -> Memo:
    return Memo(text=memo.text)


def _extract_hwpx_memo_shape(memo_shape) -> MemoShapeDefinition:
    return MemoShapeDefinition(
        memo_shape_id=memo_shape.memo_shape_id,
        width=memo_shape.width,
        line_width=memo_shape.line_width,
        line_type=memo_shape.line_type,
        line_color=memo_shape.line_color,
        fill_color=memo_shape.fill_color,
        active_color=memo_shape.active_color,
        memo_type=memo_shape.memo_type,
    )


def _extract_hwpx_equation(equation) -> Equation:
    from ._hancom_conversion_common import _normalize_int_map, _normalize_str_map

    size = equation.size()
    result = Equation(
        script=equation.script,
        width=size.get("width", 4800),
        height=size.get("height", 2300),
        shape_comment=equation.shape_comment or None,
        text_color=equation.element.get("textColor", "#000000"),
        base_unit=int(equation.element.get("baseUnit", "1100")),
        font=equation.element.get("font", "HYhwpEQ"),
        layout=_normalize_str_map(equation.layout()),
        out_margins=_normalize_int_map(equation.out_margins()),
        rotation=_normalize_str_map(equation.rotation()),
    )
    setattr(result, "_hwpx_equation_baseline", equation.element.get("baseLine"))
    setattr(result, "_hwpx_char_pr_id", equation.element.getparent().get("charPrIDRef", "0"))
    return result


def _extract_hwpx_chart(chart) -> Chart:
    from ._hancom_conversion_common import _normalize_int_map, _normalize_str_map

    size = chart.size()
    return Chart(
        title=chart.title,
        chart_type=chart.chart_type,
        categories=chart.categories,
        series=chart.series,
        data_ref=chart.data_ref,
        legend_visible=chart.legend_visible,
        width=size.get("width", 12000),
        height=size.get("height", 3200),
        shape_comment=chart.shape_comment or None,
        layout=_normalize_str_map(chart.layout()),
        out_margins=_normalize_int_map(chart.out_margins()),
        rotation=_normalize_str_map(chart.rotation()),
    )


def _extract_hwpx_shape(shape) -> Shape | Chart:
    from ._hancom_conversion_common import _normalize_int_map, _normalize_str_map
    from ._hancom_hwpx_read import _extract_hwpx_chart_carrier, _extract_hwpx_text_container_blocks

    carrier = _extract_hwpx_chart_carrier(shape)
    if carrier is not None:
        return carrier
    size = shape.size()
    fill_style = shape.fill_style()
    line_style = shape.line_style()
    size_nodes = shape.element.xpath("./hp:sz", namespaces=NS)
    width = size.get("width", 12000)
    height = size.get("height", 3200)
    return Shape(
        kind=shape.kind,
        text=_extract_hwpx_shape_direct_text(shape),
        width=width,
        height=height,
        fill_color=fill_style.get("faceColor", "#FFFFFF"),
        line_color=line_style.get("color", "#000000"),
        shape_comment=shape.shape_comment or None,
        layout=_normalize_str_map(shape.layout()),
        out_margins=_normalize_int_map(shape.out_margins()),
        rotation=_normalize_str_map(shape.rotation()),
        text_margins=_normalize_int_map(shape.text_margins()),
        specific_fields=_extract_hwpx_shape_specific_fields(shape),
        line_width=int(line_style.get("width", "33") or "33"),
        line_style=_normalize_str_map(line_style),
        has_line_node=bool(shape.element.xpath("./hp:lineShape", namespaces=NS)),
        has_size_node=bool(size_nodes),
        size_attributes=_extract_hwpx_graphic_size_attributes(shape.element),
        has_position_node=bool(shape.element.xpath("./hp:pos", namespaces=NS)),
        has_out_margin_node=bool(shape.element.xpath("./hp:outMargin", namespaces=NS)),
        has_fill_node=bool(shape.element.xpath("./hc:fillBrush", namespaces=NS)),
        has_text_margin_node=bool(shape.element.xpath("./hp:drawText/hp:textMargin", namespaces=NS)),
        nested_blocks=cast(list[object], _extract_hwpx_text_container_blocks(shape.section, shape.element)),
        **_extract_hwpx_graphic_size_fields(shape.element, width=width, height=height),
    )


def _extract_hwpx_ole(ole) -> Ole:
    from ._hancom_conversion_common import _normalize_int_map, _normalize_str_map

    size = ole.size()
    line_style = ole.line_style()
    size_nodes = ole.element.xpath("./hp:sz", namespaces=NS)
    width = size.get("width", 42001)
    height = size.get("height", 13501)
    return Ole(
        name=Path(ole.binary_part_path()).name,
        data=ole.binary_data(),
        width=width,
        height=height,
        shape_comment=ole.shape_comment or None,
        object_type=ole.object_type or "EMBEDDED",
        draw_aspect=ole.draw_aspect or "CONTENT",
        has_moniker=ole.has_moniker,
        eq_baseline=int(ole.element.get("eqBaseLine", "0")),
        line_color=line_style.get("color", "#000000"),
        line_width=int(line_style.get("width") or "0"),
        line_style=_normalize_str_map(line_style),
        layout=_normalize_str_map(ole.layout()),
        out_margins=_normalize_int_map(ole.out_margins()),
        rotation=_normalize_str_map(ole.rotation()),
        extent=_normalize_int_map(ole.extent()),
        has_size_node=bool(size_nodes),
        size_attributes=_extract_hwpx_graphic_size_attributes(ole.element),
        has_position_node=bool(ole.element.xpath("./hp:pos", namespaces=NS)),
        has_out_margin_node=bool(ole.element.xpath("./hp:outMargin", namespaces=NS)),
        **_extract_hwpx_graphic_size_fields(ole.element, width=width, height=height),
    )
