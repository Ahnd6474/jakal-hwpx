"""DrawingML picture geometry mapped to native Hancom picture properties."""
from __future__ import annotations

from copy import deepcopy
from io import BytesIO
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlparse
from urllib.request import urlopen

from ..._hancom_models import Picture
from .features import parse, xml


def _units(value: Any) -> int:
    return round(int(value) / 127)


def _linked_bytes(target: str, source_path: Path) -> bytes:
    uri = urlparse(target)
    limit = 32 * 1024 * 1024
    if uri.scheme in {"http", "https"}:
        with urlopen(target, timeout=15) as response:
            data = response.read(limit + 1)
    elif uri.scheme in {"", "file"}:
        path = Path(unquote(uri.path if uri.scheme else target))
        if uri.scheme == "file" and len(path.parts) and str(path).startswith("\\"):
            path = Path(str(path).lstrip("\\"))
        if not path.is_absolute():
            path = source_path.parent / path
        if path.stat().st_size > limit:
            raise ValueError("Linked Word picture exceeds 32 MiB.")
        data = path.read_bytes()
    else:
        raise ValueError(f"Unsupported linked picture URI scheme: {uri.scheme}")
    if len(data) > limit:
        raise ValueError("Linked Word picture exceeds 32 MiB.")
    return data


def read_picture(drawing: Any, source_path: Path) -> Picture | None:
    from docx.image.image import Image

    root = drawing._drawing
    blips = root.xpath(".//pic:blipFill/a:blip")
    if not blips:
        return None  # Charts/SmartArt are intentionally outside this mapping.
    from docx.oxml.ns import qn

    blip = blips[0]
    embedded, linked = blip.get(qn("r:embed")), blip.get(qn("r:link"))
    target = None
    if embedded:
        part = drawing.part.related_parts[embedded]
        data, name = part.blob, Path(str(part.partname)).name
    elif linked:
        target = drawing.part.rels[linked].target_ref
        data = _linked_bytes(target, source_path)
        name = "linked-picture." + Image.from_blob(data).ext
    else:
        raise ValueError("Word picture has neither embedded nor linked image data.")
    extent = root.xpath("./wp:inline/wp:extent | ./wp:anchor/wp:extent")
    width = _units(extent[0].get("cx")) if extent else 7200
    height = _units(extent[0].get("cy")) if extent else 7200
    result = Picture(name, data, extension=Path(name).suffix.lstrip("."), width=width, height=height,
                     layout={"treatAsChar": "1"})
    descriptions = root.xpath("./wp:inline/wp:docPr/@descr | ./wp:anchor/wp:docPr/@descr")
    result.shape_comment = descriptions[0] if descriptions else None
    crop = root.xpath(".//pic:blipFill/a:srcRect")
    if crop:
        # imgClip is a rectangle in the un-cropped original picture coordinates.
        result.original_width, result.original_height = width, height
        result.crop = {edge: round(int(crop[0].get(attribute, "0")) * size / 100000)
                       for edge, attribute, size in (("left", "l", width), ("top", "t", height))}
        result.crop.update(right=width - round(int(crop[0].get("r", "0")) * width / 100000),
                           bottom=height - round(int(crop[0].get("b", "0")) * height / 100000))
    transforms = root.xpath(".//pic:spPr/a:xfrm")
    if transforms and transforms[0].get("rot"):
        result.rotation = {"angle": str(int(transforms[0].get("rot")) / 60000),
                           "centerX": str(width // 2), "centerY": str(height // 2), "rotateimage": "1"}
    anchors = root.xpath("./wp:anchor")
    if anchors:
        anchor = anchors[0]
        result.layout.update(treatAsChar="0", flowWithText="1", allowOverlap=anchor.get("allowOverlap", "1"),
                             holdAnchorAndSO=anchor.get("locked", "0"))
        for direction, key in (("H", "horz"), ("V", "vert")):
            position = anchor.find(qn("wp:position" + direction))
            if position is not None:
                relative = position.get("relativeFrom", "margin")
                result.layout[key + "RelTo"] = {"page": "PAPER", "margin": "PAGE", "column": "COLUMN",
                                                    "paragraph": "PARA", "character": "PARA", "line": "LINE"}.get(relative, "PAGE")
                offset = position.find(qn("wp:posOffset"))
                align = position.find(qn("wp:align"))
                result.layout[key + "Offset"] = str(_units(offset.text or "0")) if offset is not None else "0"
                result.layout[key + "Align"] = (align.text or "LEFT").upper() if align is not None else ("LEFT" if key == "horz" else "TOP")
        wrap = next((child for child in anchor if child.tag.rsplit("}", 1)[-1].startswith("wrap")), None)
        name = wrap.tag.rsplit("}", 1)[-1] if wrap is not None else "wrapNone"
        result.layout["textWrap"] = {"wrapSquare": "SQUARE", "wrapTight": "SQUARE", "wrapThrough": "SQUARE",
                                     "wrapTopAndBottom": "TOP_AND_BOTTOM", "wrapNone": "BEHIND_TEXT" if anchor.get("behindDoc") == "1" else "IN_FRONT_OF_TEXT"}[name]
        result.layout["textFlow"] = {"left": "LEFT_ONLY", "right": "RIGHT_ONLY", "largest": "LARGEST_ONLY"}.get(wrap.get("wrapText", "bothSides") if wrap is not None else "bothSides", "BOTH_SIDES")
        result.out_margins = {edge: _units(anchor.get("dist" + suffix, "0"))
                              for edge, suffix in (("left", "L"), ("right", "R"), ("top", "T"), ("bottom", "B"))}
    hyperlinks = {}
    for node in root.xpath(".//a:hlinkClick | .//a:hlinkHover"):
        rid = node.get(qn("r:id"))
        if rid and rid in drawing.part.rels:
            hyperlinks[rid] = drawing.part.rels[rid].target_ref
    setattr(result, "_word_extra", {"drawing": xml(root), "linked": target, "hyperlinks": hyperlinks,
                                  "geometry": _geometry(result)})
    return result


def _geometry(block: Picture) -> dict:
    return {key: deepcopy(getattr(block, key)) for key in ("width", "height", "layout", "out_margins", "crop", "rotation")}


def append_picture(paragraph: Any, block: Picture) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.shared import Emu

    run = paragraph.add_run()
    try:
        shape = run.add_picture(BytesIO(block.data), width=Emu(block.width * 127), height=Emu(block.height * 127))
    except Exception as error:
        raise ValueError(f"Cannot export picture {block.name!r} to Word.") from error
    extra = getattr(block, "_word_extra", {})
    drawing = run._r.find(qn("w:drawing"))
    original = parse(extra["drawing"]) if extra.get("drawing") else deepcopy(drawing)
    blip = original.find(".//" + qn("a:blip"))
    blip.set(qn("r:embed"), shape._inline.graphic.graphicData.pic.blipFill.blip.embed)
    blip.attrib.pop(qn("r:link"), None)
    if extra.get("linked"):
        blip.set(qn("r:link"), paragraph.part.relate_to(extra["linked"], RT.IMAGE, is_external=True))
    for node in original.iter():
        rid = node.get(qn("r:id"))
        if rid in extra.get("hyperlinks", {}):
            node.set(qn("r:id"), paragraph.part.relate_to(extra["hyperlinks"][rid], RT.HYPERLINK, is_external=True))
    doc_pr = original.find(".//" + qn("wp:docPr"))
    doc_pr.set("id", str(shape._inline.docPr.id))
    doc_pr.set("descr", block.shape_comment or block.name)
    baseline = extra.get("geometry", {})
    for extent in original.findall(".//" + qn("wp:extent")):
        if baseline.get("width") != block.width:
            extent.set("cx", str(block.width * 127))
        if baseline.get("height") != block.height:
            extent.set("cy", str(block.height * 127))
    if baseline.get("crop") != block.crop:
        fill = original.find(".//" + qn("pic:blipFill"))
        crop = fill.find(qn("a:srcRect"))
        if crop is not None:
            fill.remove(crop)
        if block.crop:
            crop = OxmlElement("a:srcRect")
            image_dimension = getattr(block, "_native_picture_image_dim", {})
            width = int(image_dimension.get("dimwidth", block.original_width or block.width))
            height = int(image_dimension.get("dimheight", block.original_height or block.height))
            if width <= 0 or height <= 0:
                raise ValueError("A cropped picture must have positive original image dimensions.")
            for edge, attribute, size, inverse in (("left", "l", width, False), ("top", "t", height, False),
                                                    ("right", "r", width, True), ("bottom", "b", height, True)):
                value = block.crop.get(edge, size if inverse else 0)
                crop.set(attribute, str(round((size - value if inverse else value) * 100000 / size)))
            fill.insert(1, crop)
    if baseline.get("rotation") != block.rotation:
        transform = original.find(".//" + qn("a:xfrm"))
        if transform is not None:
            transform.set("rot", str(round(float(block.rotation.get("angle", "0")) * 60000)))
    flip = getattr(block, "_native_picture_flip", None)
    if flip is not None:
        transform = original.find(".//" + qn("a:xfrm"))
        if transform is not None:
            for native, word in (("horizontal", "flipH"), ("vertical", "flipV")):
                flipped = "1" if flip.get(native, "0") in {"1", "true"} else "0"
                if flipped != "0" or word in transform.attrib:
                    transform.set(word, flipped)
    old_layout = baseline.get("layout", {})
    layout_changed = any(block.layout.get(key) != value for key, value in old_layout.items())
    if not baseline:
        layout_changed = True
    margins_changed = any(int(block.out_margins.get(edge, 0)) != int(baseline.get("out_margins", {}).get(edge, 0))
                          for edge in ("left", "right", "top", "bottom"))
    if layout_changed or margins_changed:
        _position(original, block)
    drawing.getparent().replace(drawing, original)
    if extra.get("hyperlink"):
        link = OxmlElement("w:hyperlink")
        link.set(qn("r:id"), paragraph.part.relate_to(extra["hyperlink"], RT.HYPERLINK, is_external=True))
        run._r.addprevious(link)
        link.append(run._r)


def _position(drawing: Any, block: Any) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    container = drawing[0]
    floating = block.layout.get("treatAsChar", "1") == "0"
    container.tag = qn("wp:anchor" if floating else "wp:inline")
    for edge, suffix in (("left", "L"), ("right", "R"), ("top", "T"), ("bottom", "B")):
        container.set("dist" + suffix, str(block.out_margins.get(edge, 0) * 127))
    if not floating:
        for node in list(container):
            if node.tag in {qn("wp:simplePos"), qn("wp:positionH"), qn("wp:positionV")} or "}wrap" in node.tag:
                container.remove(node)
        for name in ("simplePos", "relativeHeight", "behindDoc", "locked", "layoutInCell", "allowOverlap"):
            container.attrib.pop(name, None)
        return
    wrap = block.layout.get("textWrap", "SQUARE")
    container.attrib.update({"simplePos": "0", "relativeHeight": "0", "behindDoc": "1" if wrap == "BEHIND_TEXT" else "0",
                             "locked": block.layout.get("holdAnchorAndSO", "0"), "layoutInCell": "1",
                             "allowOverlap": block.layout.get("allowOverlap", "1")})
    for node in list(container):
        if node.tag in {qn("wp:simplePos"), qn("wp:positionH"), qn("wp:positionV")} or "}wrap" in node.tag:
            container.remove(node)
    simple = OxmlElement("wp:simplePos")
    simple.attrib.update({"x": "0", "y": "0"})
    container.insert(0, simple)
    for index, direction, key in ((1, "H", "horz"), (2, "V", "vert")):
        pos = OxmlElement("wp:position" + direction)
        pos.set("relativeFrom", {"PAPER": "page", "PAGE": "margin", "COLUMN": "column", "PARA": "paragraph", "LINE": "line"}.get(block.layout.get(key + "RelTo", "PAGE"), "margin"))
        alignment = block.layout.get(key + "Align", "LEFT" if direction == "H" else "TOP")
        offset = int(block.layout.get(key + "Offset", "0"))
        value = OxmlElement("wp:align" if alignment not in {"LEFT", "TOP"} and not offset else "wp:posOffset")
        value.text = alignment.lower() if value.tag == qn("wp:align") else str(offset * 127)
        pos.append(value)
        container.insert(index, pos)
    name = {"SQUARE": "wrapSquare", "TIGHT": "wrapTight", "THROUGH": "wrapThrough", "TOP_AND_BOTTOM": "wrapTopAndBottom"}.get(wrap, "wrapNone")
    node = OxmlElement("wp:" + name)
    if name in {"wrapSquare", "wrapTight", "wrapThrough"}:
        node.set("wrapText", {"LEFT_ONLY": "left", "RIGHT_ONLY": "right", "LARGEST_ONLY": "largest"}.get(block.layout.get("textFlow", ""), "bothSides"))
    if name in {"wrapTight", "wrapThrough"}:
        polygon = OxmlElement("wp:wrapPolygon")
        polygon.set("edited", "0")
        for tag, x, y in (("start", 0, 0), ("lineTo", 0, 21600), ("lineTo", 21600, 21600), ("lineTo", 21600, 0), ("lineTo", 0, 0)):
            point = OxmlElement("wp:" + tag)
            point.attrib.update({"x": str(x), "y": str(y)})
            polygon.append(point)
        node.append(polygon)
    doc_pr = container.find(qn("wp:docPr"))
    container.insert(container.index(doc_pr), node)
