from __future__ import annotations

from collections import Counter
from copy import deepcopy
import json

from lxml import etree

from .namespaces import NS, qname


_FORM_METADATA_COMMAND_PREFIX = "JAKAL_FORM_META:"


def _text_nodes(element: etree._Element) -> list[etree._Element]:
    return list(element.xpath(".//hp:t", namespaces=NS))


def _first_node(element: etree._Element, expression: str) -> etree._Element | None:
    nodes = element.xpath(expression, namespaces=NS)
    return nodes[0] if nodes else None


def _bool_attr(value: bool) -> str:
    return "1" if value else "0"


def _set_optional_attributes(element: etree._Element | None, **attrs: object) -> None:
    if element is None:
        return
    for key, value in attrs.items():
        if value is None:
            continue
        if isinstance(value, bool):
            element.set(key, _bool_attr(value))
        else:
            element.set(key, str(value))


def _optional_int_attr(node: etree._Element | None, attr_name: str) -> int | None:
    if node is None:
        return None
    value = node.get(attr_name)
    if value in (None, ""):
        return None
    try:
        return int(value)
    except ValueError:
        return None


def _graphic_layout(element: etree._Element) -> dict[str, str]:
    pos = _first_node(element, "./hp:pos")
    layout = {
        "textWrap": element.get("textWrap", ""),
        "textFlow": element.get("textFlow", ""),
    }
    if pos is None:
        return layout
    for key in (
        "treatAsChar",
        "affectLSpacing",
        "flowWithText",
        "allowOverlap",
        "holdAnchorAndSO",
        "vertRelTo",
        "horzRelTo",
        "vertAlign",
        "horzAlign",
        "vertOffset",
        "horzOffset",
    ):
        layout[key] = pos.get(key, "")
    return layout


def _set_graphic_layout(
    element: etree._Element,
    *,
    text_wrap: str | None = None,
    text_flow: str | None = None,
    treat_as_char: bool | None = None,
    affect_line_spacing: bool | None = None,
    flow_with_text: bool | None = None,
    allow_overlap: bool | None = None,
    hold_anchor_and_so: bool | None = None,
    vert_rel_to: str | None = None,
    horz_rel_to: str | None = None,
    vert_align: str | None = None,
    horz_align: str | None = None,
    vert_offset: int | str | None = None,
    horz_offset: int | str | None = None,
) -> None:
    _set_optional_attributes(
        element,
        textWrap=text_wrap,
        textFlow=text_flow,
    )
    pos = _first_node(element, "./hp:pos")
    _set_optional_attributes(
        pos,
        treatAsChar=treat_as_char,
        affectLSpacing=affect_line_spacing,
        flowWithText=flow_with_text,
        allowOverlap=allow_overlap,
        holdAnchorAndSO=hold_anchor_and_so,
        vertRelTo=vert_rel_to,
        horzRelTo=horz_rel_to,
        vertAlign=vert_align,
        horzAlign=horz_align,
        vertOffset=vert_offset,
        horzOffset=horz_offset,
    )


def _margin_values(element: etree._Element, expression: str) -> dict[str, int]:
    margin = _first_node(element, expression)
    if margin is None:
        return {}
    return {key: int(margin.get(key, "0")) for key in ("left", "right", "top", "bottom")}


def _set_margin_values(
    element: etree._Element,
    expression: str,
    *,
    left: int | str | None = None,
    right: int | str | None = None,
    top: int | str | None = None,
    bottom: int | str | None = None,
) -> None:
    margin = _first_node(element, expression)
    _set_optional_attributes(margin, left=left, right=right, top=top, bottom=bottom)


def _graphic_size(element: etree._Element) -> dict[str, int]:
    size = _first_node(element, "./hp:sz")
    if size is None:
        return {}
    return {
        "width": int(size.get("width", "0")),
        "height": int(size.get("height", "0")),
    }


def _set_graphic_size(
    element: etree._Element,
    *,
    width: int | str | None = None,
    height: int | str | None = None,
    original_width: int | str | None = None,
    original_height: int | str | None = None,
    current_width: int | str | None = None,
    current_height: int | str | None = None,
    extent_x: int | str | None = None,
    extent_y: int | str | None = None,
) -> None:
    size_node = _first_node(element, "./hp:sz")
    original_node = _first_node(element, "./hp:orgSz")
    current_node = _first_node(element, "./hp:curSz")
    previous_width = _optional_int_attr(size_node, "width")
    previous_height = _optional_int_attr(size_node, "height")
    previous_original_width = _optional_int_attr(original_node, "width")
    previous_original_height = _optional_int_attr(original_node, "height")

    _set_optional_attributes(size_node, width=width, height=height)
    _set_optional_attributes(original_node, width=original_width, height=original_height)
    _set_optional_attributes(current_node, width=current_width, height=current_height)
    _set_optional_attributes(_first_node(element, "./hc:extent"), x=extent_x, y=extent_y)

    resolved_width = _optional_int_attr(size_node, "width")
    resolved_height = _optional_int_attr(size_node, "height")
    resolved_original_width = _optional_int_attr(original_node, "width")
    resolved_original_height = _optional_int_attr(original_node, "height")

    _sync_rotation_center_for_size(
        element,
        width=resolved_width,
        height=resolved_height,
        previous_width=previous_width,
        previous_height=previous_height,
    )
    _sync_picture_size_geometry(
        element,
        width=resolved_width,
        height=resolved_height,
        original_width=resolved_original_width,
        original_height=resolved_original_height,
        previous_original_width=previous_original_width,
        previous_original_height=previous_original_height,
    )
    _sync_shape_point_geometry(
        element,
        width=resolved_width,
        height=resolved_height,
        previous_width=previous_width,
        previous_height=previous_height,
    )
    _sync_extent_for_size(
        element,
        width=resolved_width,
        height=resolved_height,
        previous_width=previous_width,
        previous_height=previous_height,
        extent_x=extent_x,
        extent_y=extent_y,
    )


def _graphic_rotation(element: etree._Element) -> dict[str, str]:
    node = _first_node(element, "./hp:rotationInfo")
    if node is None:
        return {}
    return {
        key: node.get(key, "")
        for key in ("angle", "centerX", "centerY", "rotateimage")
    }


def _set_graphic_rotation(
    element: etree._Element,
    *,
    angle: int | str | None = None,
    center_x: int | str | None = None,
    center_y: int | str | None = None,
    rotate_image: bool | None = None,
) -> None:
    node = _first_node(element, "./hp:rotationInfo")
    if node is None:
        node = etree.SubElement(element, qname("hp", "rotationInfo"))
    _set_optional_attributes(
        node,
        angle=angle,
        centerX=center_x,
        centerY=center_y,
        rotateimage=rotate_image,
    )


def _sync_rotation_center_for_size(
    element: etree._Element,
    *,
    width: int | None,
    height: int | None,
    previous_width: int | None,
    previous_height: int | None,
) -> None:
    if width is None and height is None:
        return
    node = _first_node(element, "./hp:rotationInfo")
    if node is None:
        return

    angle = (node.get("angle") or "").strip()
    current_center_x = _optional_int_attr(node, "centerX")
    current_center_y = _optional_int_attr(node, "centerY")

    if width is not None:
        previous_default_x = None if previous_width is None else previous_width // 2
        if (
            current_center_x is None
            or (previous_default_x is not None and current_center_x == previous_default_x)
            or (current_center_x == 0 and angle in {"", "0"})
        ):
            node.set("centerX", str(width // 2))

    if height is not None:
        previous_default_y = None if previous_height is None else previous_height // 2
        if (
            current_center_y is None
            or (previous_default_y is not None and current_center_y == previous_default_y)
            or (current_center_y == 0 and angle in {"", "0"})
        ):
            node.set("centerY", str(height // 2))


def _sync_picture_size_geometry(
    element: etree._Element,
    *,
    width: int | None,
    height: int | None,
    original_width: int | None,
    original_height: int | None,
    previous_original_width: int | None,
    previous_original_height: int | None,
) -> None:
    if width is not None and height is not None and original_width not in (None, 0) and original_height not in (None, 0):
        scale = _first_node(element, "./hp:renderingInfo/hc:scaMatrix")
        if scale is not None:
            scale.set("e1", f"{width / (original_width or 1):.6f}")
            scale.set("e5", f"{height / (original_height or 1):.6f}")

    if original_width is None or original_height is None:
        return

    image_rect = _first_node(element, "./hp:imgRect")
    if image_rect is not None:
        for index, (x, y) in enumerate(((0, 0), (original_width, 0), (original_width, original_height), (0, original_height))):
            point = _first_node(image_rect, f"./hc:pt{index}")
            if point is not None:
                point.set("x", str(x))
                point.set("y", str(y))

    image_dim = _first_node(element, "./hp:imgDim")
    _set_optional_attributes(image_dim, dimwidth=original_width, dimheight=original_height)

    image_clip = _first_node(element, "./hp:imgClip")
    if image_clip is None:
        return

    current_left = _optional_int_attr(image_clip, "left")
    current_right = _optional_int_attr(image_clip, "right")
    current_top = _optional_int_attr(image_clip, "top")
    current_bottom = _optional_int_attr(image_clip, "bottom")
    if current_left is None or current_right is None or current_top is None or current_bottom is None:
        return

    if (
        current_left == 0
        and current_top == 0
        and (
            (current_right == 0 and current_bottom == 0)
            or (
                previous_original_width is not None
                and previous_original_height is not None
                and current_right == previous_original_width
                and current_bottom == previous_original_height
            )
        )
    ):
        image_clip.set("left", "0")
        image_clip.set("right", str(original_width))
        image_clip.set("top", "0")
        image_clip.set("bottom", str(original_height))


def _sync_shape_point_geometry(
    element: etree._Element,
    *,
    width: int | None,
    height: int | None,
    previous_width: int | None,
    previous_height: int | None,
) -> None:
    if width is None or height is None or previous_width is None or previous_height is None:
        return

    draw_text = _first_node(element, "./hp:drawText")
    if draw_text is not None:
        draw_text.set("lastWidth", str(width))

    points = [_first_node(element, f"./hc:pt{index}") for index in range(4)]
    present_points = [point for point in points if point is not None]
    if len(present_points) not in {2, 4}:
        return

    current_pairs = [(_optional_int_attr(point, "x"), _optional_int_attr(point, "y")) for point in present_points]
    if len(present_points) == 2:
        expected_pairs = [(0, 0), (previous_width, previous_height)]
        replacement_pairs = [(0, 0), (width, height)]
    else:
        expected_pairs = [(0, 0), (previous_width, 0), (previous_width, previous_height), (0, previous_height)]
        replacement_pairs = [(0, 0), (width, 0), (width, height), (0, height)]

    if current_pairs != expected_pairs:
        return

    for point, (x, y) in zip(present_points, replacement_pairs, strict=True):
        point.set("x", str(x))
        point.set("y", str(y))


def _sync_extent_for_size(
    element: etree._Element,
    *,
    width: int | None,
    height: int | None,
    previous_width: int | None,
    previous_height: int | None,
    extent_x: int | str | None,
    extent_y: int | str | None,
) -> None:
    node = _first_node(element, "./hc:extent")
    if node is None:
        return
    if extent_x is not None or extent_y is not None:
        _set_optional_attributes(node, x=extent_x, y=extent_y)
        return

    current_x = _optional_int_attr(node, "x")
    current_y = _optional_int_attr(node, "y")
    if width is not None and previous_width is not None and current_x == previous_width:
        node.set("x", str(width))
    if height is not None and previous_height is not None and current_y == previous_height:
        node.set("y", str(height))


def _line_style(element: etree._Element) -> dict[str, str]:
    node = _first_node(element, "./hp:lineShape")
    if node is None:
        return {}
    return {
        key: node.get(key, "")
        for key in (
            "color",
            "width",
            "style",
            "endCap",
            "headStyle",
            "tailStyle",
            "headfill",
            "tailfill",
            "headSz",
            "tailSz",
            "outlineStyle",
            "alpha",
        )
    }


def _set_line_style(
    element: etree._Element,
    *,
    color: str | None = None,
    width: int | str | None = None,
    style: str | None = None,
    end_cap: str | None = None,
    head_style: str | None = None,
    tail_style: str | None = None,
    head_fill: bool | None = None,
    tail_fill: bool | None = None,
    head_size: str | None = None,
    tail_size: str | None = None,
    outline_style: str | None = None,
    alpha: int | str | None = None,
) -> None:
    node = _first_node(element, "./hp:lineShape")
    if node is None:
        node = etree.SubElement(element, qname("hp", "lineShape"))
    _set_optional_attributes(
        node,
        color=color,
        width=width,
        style=style,
        endCap=end_cap,
        headStyle=head_style,
        tailStyle=tail_style,
        headfill=head_fill,
        tailfill=tail_fill,
        headSz=head_size,
        tailSz=tail_size,
        outlineStyle=outline_style,
        alpha=alpha,
    )





def _fill_style(element: etree._Element) -> dict[str, str]:
    node = _first_node(element, "./hc:fillBrush/hc:winBrush")
    if node is None:
        return {}
    return {
        key: node.get(key, "")
        for key in ("faceColor", "hatchColor", "alpha")
    }


def _set_fill_style(
    element: etree._Element,
    *,
    face_color: str | None = None,
    hatch_color: str | None = None,
    alpha: int | str | None = None,
) -> None:
    _set_optional_attributes(
        _first_node(element, "./hc:fillBrush/hc:winBrush"),
        faceColor=face_color,
        hatchColor=hatch_color,
        alpha=alpha,
    )


def _text_margin(element: etree._Element) -> dict[str, int]:
    return _margin_values(element, "./hp:drawText/hp:textMargin")


def _set_text_margin(
    element: etree._Element,
    *,
    left: int | str | None = None,
    right: int | str | None = None,
    top: int | str | None = None,
    bottom: int | str | None = None,
) -> None:
    _set_margin_values(element, "./hp:drawText/hp:textMargin", left=left, right=right, top=top, bottom=bottom)


def _image_adjustment(element: etree._Element) -> dict[str, str]:
    node = _first_node(element, "./hc:img")
    if node is None:
        return {}
    return {
        key: node.get(key, "")
        for key in ("bright", "contrast", "effect", "alpha")
    }


def _set_image_adjustment(
    element: etree._Element,
    *,
    bright: int | str | None = None,
    contrast: int | str | None = None,
    effect: str | None = None,
    alpha: int | str | None = None,
) -> None:
    _set_optional_attributes(_first_node(element, "./hc:img"), bright=bright, contrast=contrast, effect=effect, alpha=alpha)


def _crop_values(element: etree._Element) -> dict[str, int]:
    node = _first_node(element, "./hp:imgClip")
    if node is None:
        return {}
    return {key: int(node.get(key, "0")) for key in ("left", "right", "top", "bottom")}


def _set_crop_values(
    element: etree._Element,
    *,
    left: int | str | None = None,
    right: int | str | None = None,
    top: int | str | None = None,
    bottom: int | str | None = None,
) -> None:
    node = _first_node(element, "./hp:imgClip")
    if node is None:
        node = etree.SubElement(element, qname("hp", "imgClip"))
    _set_optional_attributes(node, left=left, right=right, top=top, bottom=bottom)


def _extract_text(element: etree._Element) -> str:
    return "".join(node.text or "" for node in _text_nodes(element))


def _replace_text(element: etree._Element, old: str, new: str, count: int = -1) -> int:
    if not old:
        raise ValueError("old must be non-empty.")
    remaining = count
    replaced = 0
    for node in _text_nodes(element):
        current = node.text or ""
        if old not in current:
            continue
        if remaining < 0:
            changed = current.count(old)
            node.text = current.replace(old, new)
        else:
            changed = 0
            updated = current
            while changed < remaining and old in updated:
                updated = updated.replace(old, new, 1)
                changed += 1
            node.text = updated
            remaining -= changed
        replaced += changed
        if changed and remaining == 0:
            break
    if replaced:
        _invalidate_paragraph_layout(element)
    return replaced


def _paragraphs_affected_by_text_edit(element: etree._Element) -> list[etree._Element]:
    paragraphs: list[etree._Element] = []
    seen: set[int] = set()

    def add(node: etree._Element) -> None:
        marker = id(node)
        if marker in seen:
            return
        seen.add(marker)
        paragraphs.append(node)

    if etree.QName(element).localname == "p":
        add(element)

    ancestors = element.xpath("ancestor::hp:p[1]", namespaces=NS)
    if ancestors:
        add(ancestors[0])

    for paragraph in _paragraph_nodes(element):
        add(paragraph)

    return paragraphs


def _invalidate_paragraph_layout(element: etree._Element) -> None:
    for paragraph in _paragraphs_affected_by_text_edit(element):
        for line_seg_array in paragraph.xpath("./hp:linesegarray", namespaces=NS):
            paragraph.remove(line_seg_array)


def _paragraph_nodes(element: etree._Element) -> list[etree._Element]:
    if etree.QName(element).localname == "p":
        return [element]

    paragraphs = list(element.xpath("./hp:p", namespaces=NS))
    for sublist in element.xpath("./hp:subList", namespaces=NS):
        paragraphs.extend(sublist.xpath("./hp:p", namespaces=NS))
    return paragraphs


def _ensure_first_paragraph(element: etree._Element) -> etree._Element:
    paragraph = element.xpath(".//hp:p[1]", namespaces=NS)
    if paragraph:
        return paragraph[0]

    sublists = element.xpath(".//hp:subList[1]", namespaces=NS)
    if sublists:
        sublist = sublists[0]
    else:
        sublist = etree.SubElement(element, qname("hp", "subList"))
        sublist.set("id", "")
        sublist.set("textDirection", "HORIZONTAL")
        sublist.set("lineWrap", "BREAK")
        sublist.set("vertAlign", "TOP")
        sublist.set("linkListIDRef", "0")
        sublist.set("linkListNextIDRef", "0")
        sublist.set("textWidth", "0")
        sublist.set("textHeight", "0")
        sublist.set("hasTextRef", "0")
        sublist.set("hasNumRef", "0")

    paragraph = etree.SubElement(sublist, qname("hp", "p"))
    paragraph.set("id", "0")
    paragraph.set("paraPrIDRef", "0")
    paragraph.set("styleIDRef", "0")
    paragraph.set("pageBreak", "0")
    paragraph.set("columnBreak", "0")
    paragraph.set("merged", "0")
    return paragraph


def _default_char_pr_id(paragraph: etree._Element) -> str:
    for run in paragraph.xpath("./hp:run", namespaces=NS):
        value = run.get("charPrIDRef")
        if value is not None:
            return value
    return "0"


def _ensure_run_with_text(paragraph: etree._Element, text: str) -> etree._Element:
    run = etree.SubElement(paragraph, qname("hp", "run"))
    run.set("charPrIDRef", _default_char_pr_id(paragraph))
    text_node = etree.SubElement(run, qname("hp", "t"))
    text_node.text = text
    return run


def _reset_paragraph_text(paragraph: etree._Element, text: str) -> None:
    for child in list(paragraph):
        paragraph.remove(child)
    _ensure_run_with_text(paragraph, text)


def _replace_paragraph_text_preserving_controls(
    paragraph: etree._Element,
    text: str,
    *,
    char_pr_id: str | None = None,
) -> None:
    resolved_char_pr_id = char_pr_id or _default_char_pr_id(paragraph)
    preserved_runs = [
        deepcopy(child)
        for child in paragraph.xpath("./hp:run[hp:secPr or hp:ctrl]", namespaces=NS)
    ]
    for preserved_run in preserved_runs:
        for text_node in preserved_run.xpath("./hp:t", namespaces=NS):
            preserved_run.remove(text_node)
    for child in list(paragraph):
        paragraph.remove(child)
    for preserved in preserved_runs:
        paragraph.append(preserved)
    run = etree.SubElement(paragraph, qname("hp", "run"))
    run.set("charPrIDRef", resolved_char_pr_id)
    text_node = etree.SubElement(run, qname("hp", "t"))
    text_node.text = text


def _preserved_structure_signature(paragraph: etree._Element) -> Counter[str]:
    signature: Counter[str] = Counter()
    for node in paragraph.xpath("./hp:run/hp:secPr | ./hp:run/hp:ctrl/*", namespaces=NS):
        local_name = etree.QName(node).localname
        parent_local_name = etree.QName(node.getparent()).localname if node.getparent() is not None else ""
        label = f"{parent_local_name}/{local_name}" if parent_local_name == "ctrl" else local_name
        attributes = []
        for key in ("id", "fieldid", "beginIDRef", "name", "instid", "num", "numType", "type"):
            value = node.get(key)
            if value:
                attributes.append(f"{key}={value}")
        if attributes:
            label = f"{label}[{', '.join(attributes)}]"
        signature[label] += 1
    return signature


def _capture_protected_paragraph_signatures(element: etree._Element) -> tuple[list[etree._Element], list[Counter[str]]]:
    paragraphs = _paragraph_nodes(element)
    return paragraphs, [_preserved_structure_signature(paragraph) for paragraph in paragraphs]


def _missing_preserved_tokens(expected: Counter[str], actual: Counter[str]) -> list[str]:
    missing: list[str] = []
    for token, count in expected.items():
        deficit = count - actual.get(token, 0)
        if deficit > 0:
            missing.extend([token] * deficit)
    return missing


def _clone_paragraph(paragraph: etree._Element) -> etree._Element:
    clone = deepcopy(paragraph)
    _reset_paragraph_text(clone, "")
    return clone


def _split_score(text: str, index: int, target: int) -> tuple[int, int]:
    penalty = abs(index - target) * 10
    prev_char = text[index - 1] if index > 0 else ""
    next_char = text[index] if index < len(text) else ""

    if prev_char in ")]":
        penalty -= 40
    elif prev_char == ",":
        penalty -= 30
    elif prev_char == " ":
        penalty -= 10

    if next_char == "[":
        penalty -= 25
    elif next_char == "(":
        penalty -= 10
    elif next_char == " ":
        penalty -= 5

    return penalty, index


def _snap_split_index(text: str, target: int, *, minimum: int, maximum: int) -> int:
    if minimum >= maximum:
        return minimum

    target = min(max(target, minimum), maximum)
    window = max(8, len(text) // 12)
    search_start = max(minimum, target - window)
    search_end = min(maximum, target + window)
    candidates = range(search_start, search_end + 1)
    return min(candidates, key=lambda index: _split_score(text, index, target))


def _distribute_text_across_paragraphs(text: str, paragraphs: list[etree._Element]) -> list[str]:
    if not paragraphs:
        return [text]

    if "\n" in text:
        return text.split("\n")

    if len(paragraphs) == 1:
        return [text]

    original_lengths = [len("".join(node.text or "" for node in paragraph.xpath(".//hp:t", namespaces=NS))) for paragraph in paragraphs]
    total_length = sum(original_lengths)
    if total_length <= 0:
        return [text] + [""] * (len(paragraphs) - 1)

    segments: list[str] = []
    start = 0

    for index in range(1, len(original_lengths)):
        remaining_paragraphs = len(original_lengths) - index
        minimum = start
        maximum = len(text) - remaining_paragraphs
        target = round(len(text) * (sum(original_lengths[:index]) / total_length))
        split_at = _snap_split_index(text, target, minimum=minimum, maximum=maximum)
        segments.append(text[start:split_at])
        start = split_at

    segments.append(text[start:])
    return segments


def _set_text(element: etree._Element, text: str) -> None:
    paragraphs = _paragraph_nodes(element)
    if paragraphs:
        parts = _distribute_text_across_paragraphs(text, paragraphs) if (len(paragraphs) > 1 or "\n" in text) else [text]
        while len(paragraphs) < len(parts):
            template = paragraphs[-1] if paragraphs else _ensure_first_paragraph(element)
            clone = _clone_paragraph(template)
            template.addnext(clone)
            paragraphs = _paragraph_nodes(element)
        for paragraph, value in zip(paragraphs, parts):
            _replace_paragraph_text_preserving_controls(paragraph, value)
        for paragraph in paragraphs[len(parts) :]:
            _replace_paragraph_text_preserving_controls(paragraph, "")
        _invalidate_paragraph_layout(element)
        return

    text_nodes = _text_nodes(element)
    if text_nodes:
        text_nodes[0].text = text
        for extra in text_nodes[1:]:
            extra.text = ""
        _invalidate_paragraph_layout(element)
        return

    paragraph = _ensure_first_paragraph(element)
    _replace_paragraph_text_preserving_controls(paragraph, text)
    _invalidate_paragraph_layout(element)


def _form_command_metadata(element: etree._Element) -> dict[str, object]:
    command = element.get("command", "")
    if not command.startswith(_FORM_METADATA_COMMAND_PREFIX):
        return {}
    try:
        payload = json.loads(command[len(_FORM_METADATA_COMMAND_PREFIX) :])
    except json.JSONDecodeError:
        return {}
    return payload if isinstance(payload, dict) else {}


def _write_form_command_metadata(
    element: etree._Element,
    *,
    updates: dict[str, object | None] | None = None,
) -> None:
    command = element.get("command", "")
    if command and not command.startswith(_FORM_METADATA_COMMAND_PREFIX):
        return
    metadata = dict(_form_command_metadata(element))
    for key, value in (updates or {}).items():
        if value is None or value == "":
            metadata.pop(key, None)
        else:
            metadata[str(key)] = value
    if metadata:
        element.set(
            "command",
            _FORM_METADATA_COMMAND_PREFIX + json.dumps(metadata, ensure_ascii=False, separators=(",", ":"), sort_keys=True),
        )
    else:
        element.set("command", "")


def _form_list_item_values(element: etree._Element) -> list[str]:
    values: list[str] = []
    for item in element.xpath("./hp:listItem", namespaces=NS):
        values.append(item.get("displayText") or item.get("value") or "")
    return values


def _set_form_list_items(element: etree._Element, values: list[str]) -> None:
    for item in list(element.xpath("./hp:listItem", namespaces=NS)):
        element.remove(item)
    for value in values:
        item = etree.SubElement(element, qname("hp", "listItem"))
        item.set("displayText", "")
        item.set("value", value)


def _form_caption_text(element: etree._Element) -> str | None:
    caption = _first_node(element, "./hp:caption")
    if caption is None:
        return None
    text = _extract_text(caption)
    return text if text else None


def _set_form_caption_text(
    element: etree._Element,
    value: str | None,
    *,
    char_pr_id: str | None = None,
) -> None:
    caption = _first_node(element, "./hp:caption")
    if not value:
        if caption is not None:
            element.remove(caption)
        return

    if caption is None:
        caption = etree.Element(qname("hp", "caption"))
        element.insert(0, caption)
    _set_optional_attributes(caption, side="LEFT", fullSz=False, width=-1, gap=850, lastWidth=0)

    sub_list = _first_node(caption, "./hp:subList")
    if sub_list is None:
        sub_list = etree.SubElement(caption, qname("hp", "subList"))
    _set_optional_attributes(
        sub_list,
        id="",
        textDirection="HORIZONTAL",
        lineWrap="BREAK",
        vertAlign="TOP",
        linkListIDRef="0",
        linkListNextIDRef="0",
        textWidth="0",
        textHeight="0",
        hasTextRef=False,
        hasNumRef=False,
    )

    paragraph = _first_node(sub_list, "./hp:p")
    if paragraph is None:
        paragraph = etree.SubElement(sub_list, qname("hp", "p"))
    _set_optional_attributes(
        paragraph,
        id="0",
        paraPrIDRef="0",
        styleIDRef="0",
        pageBreak="0",
        columnBreak="0",
        merged="0",
    )

    run = _first_node(paragraph, "./hp:run")
    if run is None:
        run = etree.SubElement(paragraph, qname("hp", "run"))
    run.set("charPrIDRef", char_pr_id or run.get("charPrIDRef") or "0")

    text_node = _first_form_text_node(run)
    if text_node is None:
        text_node = _first_node(run, "./hp:t")
    if text_node is None:
        text_node = etree.SubElement(run, qname("hp", "t"))
    text_node.text = value

    for extra_paragraph in list(sub_list.xpath("./hp:p[position()>1]", namespaces=NS)):
        sub_list.remove(extra_paragraph)
    for extra_run in list(paragraph.xpath("./hp:run[position()>1]", namespaces=NS)):
        paragraph.remove(extra_run)


def _first_form_text_node(element: etree._Element) -> etree._Element | None:
    nodes = element.xpath("./hp:text", namespaces=NS)
    return nodes[0] if nodes else None

for _compat_name in ["_text_nodes","_first_node","_bool_attr","_set_optional_attributes","_optional_int_attr","_graphic_layout","_set_graphic_layout","_margin_values","_set_margin_values","_graphic_size","_set_graphic_size","_graphic_rotation","_set_graphic_rotation","_sync_rotation_center_for_size","_sync_picture_size_geometry","_sync_shape_point_geometry","_sync_extent_for_size","_line_style","_set_line_style","_fill_style","_set_fill_style","_text_margin","_set_text_margin","_image_adjustment","_set_image_adjustment","_crop_values","_set_crop_values","_extract_text","_replace_text","_paragraphs_affected_by_text_edit","_invalidate_paragraph_layout","_paragraph_nodes","_ensure_first_paragraph","_default_char_pr_id","_ensure_run_with_text","_reset_paragraph_text","_replace_paragraph_text_preserving_controls","_preserved_structure_signature","_capture_protected_paragraph_signatures","_missing_preserved_tokens","_clone_paragraph","_split_score","_snap_split_index","_distribute_text_across_paragraphs","_set_text","_form_command_metadata","_write_form_command_metadata","_form_list_item_values","_set_form_list_items","_form_caption_text","_set_form_caption_text","_first_form_text_node"]:
    globals()[_compat_name].__module__ = "jakal_hwpx.elements"

