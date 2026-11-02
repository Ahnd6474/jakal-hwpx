"""Editable Word text boxes mapped to native Hancom text shapes."""
from __future__ import annotations

from copy import deepcopy
import re
from typing import Any

from lxml import etree

from ..._hancom_models import Shape
from .features import parse, xml

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
      "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
      "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
      "v": "urn:schemas-microsoft-com:vml"}


def find(node: Any, expression: str) -> list:
    return etree.XPath(expression, namespaces=NS)(node)


def select_alternatives(root: Any) -> None:
    """Read one supported text-box representation, never its duplicate fallback."""
    for alternate in list(root.iter("{" + NS["mc"] + "}AlternateContent")):
        selected = next((branch for branch in alternate
                         if find(branch, ".//wps:txbx/w:txbxContent")), None)
        if selected is None:
            selected = next((branch for branch in alternate if find(branch, ".//v:textbox/w:txbxContent")), None)
        if selected is not None:
            for child in list(selected):
                alternate.addprevious(child)
            alternate.getparent().remove(alternate)


def geometry(block: Shape) -> dict:
    return {name: deepcopy(getattr(block, name)) for name in
            ("width", "height", "layout", "out_margins", "text_margins", "rotation",
             "fill_color", "line_color", "line_width", "has_fill_node")}


def read_textbox(node: Any, source: Any, reader: Any) -> Shape | None:
    from docx.oxml import OxmlElement
    from docx.blkcntnr import BlockItemContainer

    contents = find(node, ".//wps:txbx/w:txbxContent | .//v:textbox/w:txbxContent")
    if not contents:
        return None
    if len(contents) != 1:
        raise ValueError("A grouped Word drawing contains multiple text boxes; individual mapping is required.")
    content = contents[0]
    template = deepcopy(node)
    for box in find(template, ".//w:txbxContent"):
        for child in list(box):
            box.remove(child)
    # VML's auxiliary Office drawing cache may contain stale original text.
    for item in template.iter():
        for attribute in list(item.attrib):
            if etree.QName(attribute).localname == "gfxdata":
                item.attrib.pop(attribute)
    shape = Shape("rect", has_fill_node=False, line_width=0,
                  layout={"treatAsChar": "1"}, shape_comment="jakal:word-textbox")
    extent = find(node, ".//wp:extent")
    if extent:
        shape.width, shape.height = (round(int(extent[0].get(key)) / 127) for key in ("cx", "cy"))
    anchor = find(node, ".//wp:anchor")
    if anchor:
        anchor = anchor[0]
        shape.layout.update(treatAsChar="0", flowWithText="1", allowOverlap=anchor.get("allowOverlap", "1"),
                            holdAnchorAndSO=anchor.get("locked", "0"))
        for direction, key in (("H", "horz"), ("V", "vert")):
            position = anchor.find("{" + NS["wp"] + "}position" + direction)
            if position is None:
                continue
            shape.layout[key + "RelTo"] = {"page": "PAPER", "margin": "PAGE", "column": "COLUMN",
                "paragraph": "PARA", "character": "PARA", "line": "LINE"}.get(position.get("relativeFrom"), "PAGE")
            offset = position.find("{" + NS["wp"] + "}posOffset")
            alignment = position.find("{" + NS["wp"] + "}align")
            shape.layout[key + "Offset"] = str(round(int(offset.text or 0) / 127)) if offset is not None else "0"
            shape.layout[key + "Align"] = (alignment.text or "LEFT").upper() if alignment is not None else ("LEFT" if direction == "H" else "TOP")
        wrapping = next((etree.QName(child).localname for child in anchor if etree.QName(child).localname.startswith("wrap")), "wrapNone")
        shape.layout["textWrap"] = {"wrapSquare": "SQUARE", "wrapTight": "SQUARE", "wrapThrough": "SQUARE",
            "wrapTopAndBottom": "TOP_AND_BOTTOM"}.get(wrapping, "BEHIND_TEXT" if anchor.get("behindDoc") == "1" else "IN_FRONT_OF_TEXT")
        shape.out_margins = {edge: round(int(anchor.get("dist" + suffix, "0")) / 127)
                             for edge, suffix in (("left", "L"), ("right", "R"), ("top", "T"), ("bottom", "B"))}
    properties = find(node, ".//wps:spPr")
    if properties:
        properties = properties[0]
        shape.has_fill_node = not bool(find(properties, "./a:noFill"))
        fill = find(properties, "./a:solidFill/a:srgbClr/@val")
        if fill:
            shape.fill_color = "#" + fill[0]
        lines = find(properties, "./a:ln")
        if lines:
            shape.has_line_node = not bool(find(lines[0], "./a:noFill"))
            shape.line_width = round(int(lines[0].get("w", "12700")) / 127) if shape.has_line_node else 0
            colors = find(lines[0], "./a:solidFill/a:srgbClr/@val")
            if colors:
                shape.line_color = "#" + colors[0]
        rotations = find(properties, "./a:xfrm/@rot")
        if rotations:
            shape.rotation = {"angle": str(int(rotations[0]) / 60000)}
    body = find(node, ".//wps:bodyPr")
    vertical = "TOP"
    if body:
        body = body[0]
        vertical = {"t": "TOP", "ctr": "CENTER", "b": "BOTTOM"}.get(body.get("anchor", "t"), "TOP")
        shape.text_margins = {edge: round(int(body.get(key, str(default))) / 127)
                              for edge, key, default in (("left", "lIns", 91440), ("right", "rIns", 91440),
                                                         ("top", "tIns", 45720), ("bottom", "bIns", 45720))}
    legacy = find(node, ".//v:shape")
    if legacy:
        values = dict(re.findall(r"([\w-]+)\s*:\s*([^;]+)", legacy[0].get("style", "")))
        for name in ("width", "height"):
            value = values.get(name, "")
            if value.endswith("pt"):
                setattr(shape, name, round(float(value[:-2]) * 100))
        shape.layout.update(treatAsChar="0", textWrap="IN_FRONT_OF_TEXT", vertRelTo="PARA", horzRelTo="COLUMN")
        for name, key in (("margin-left", "horzOffset"), ("margin-top", "vertOffset")):
            shape.layout[key] = str(round(float(values.get(name, "0pt").removesuffix("pt")) * 100))
        shape.has_fill_node = legacy[0].get("filled", "t") != "f"
    container = OxmlElement("w:hdr")
    container.extend(list(content))
    shape.nested_blocks = reader.container_blocks(BlockItemContainer(container, source))
    shape.text = "\n".join(paragraph.text for paragraph in BlockItemContainer(container, source).paragraphs)
    setattr(shape, "_word_extra", {"textbox": {"template": xml(template), "geometry": geometry(shape),
                                                "vertical_align": vertical}})
    return shape


def append_textbox(writer: Any, paragraph: Any, block: Shape) -> None:
    from docx.oxml import OxmlElement
    from docx.blkcntnr import BlockItemContainer
    from .pictures import _position

    extra = getattr(block, "_word_extra")["textbox"]
    template = parse(extra["template"])
    content = find(template, ".//w:txbxContent")[0]
    for child in list(content):
        content.remove(child)
    body = OxmlElement("w:hdr")
    container = BlockItemContainer(body, paragraph)
    if block.nested_blocks:
        writer.blocks(container, block.nested_blocks, nested=True)
    else:
        container.add_paragraph(block.text)
    content.extend(list(body))
    if not len(content) or content[-1].tag != "{" + NS["w"] + "}p":
        content.append(OxmlElement("w:p"))
    for properties in find(template, ".//wp:docPr"):
        properties.set("id", str(paragraph.part.next_id))
    baseline = extra["geometry"]
    for name, key in (("width", "cx"), ("height", "cy")):
        if getattr(block, name) != baseline.get(name):
            for extent in find(template, ".//wp:extent | .//wps:spPr/a:xfrm/a:ext"):
                extent.set(key, str(getattr(block, name) * 127))
    if find(template, ".//wp:anchor | .//wp:inline") and (
        any(block.layout.get(key) != value for key, value in baseline["layout"].items())
        or any(int(block.out_margins.get(key, 0)) != int(baseline["out_margins"].get(key, 0)) for key in ("left", "right", "top", "bottom"))
    ):
        _position(template, block)
    for properties in find(template, ".//wps:bodyPr"):
        for edge, key in (("left", "lIns"), ("right", "rIns"), ("top", "tIns"), ("bottom", "bIns")):
            if block.text_margins.get(edge) != baseline["text_margins"].get(edge):
                properties.set(key, str(block.text_margins.get(edge, 0) * 127))
    paragraph.add_run()._r.append(template)
