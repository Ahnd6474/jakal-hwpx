"""Validated OOXML formatting shared by the public Word detail handles."""
from __future__ import annotations

import math

from .editing import _color, _points, _tag


FONT_PROPERTIES = {"highlight", "character_spacing_pt", "scale_percent", "font_names", "languages"}
PARAGRAPH_BOOLEANS = {"word_wrap": "wordWrap", "auto_space_east_asian_latin": "autoSpaceDE",
                      "auto_space_east_asian_numbers": "autoSpaceDN", "snap_to_grid": "snapToGrid",
                      "suppress_line_numbers": "suppressLineNumbers"}
PARAGRAPH_PROPERTIES = {"line_spacing_pt", "line_spacing_rule", "tab_stops", *PARAGRAPH_BOOLEANS}


def element(parent, name, **attributes):
    from docx.oxml import OxmlElement

    node = OxmlElement("w:" + name)
    node.attrib.update({_tag(key): str(value) for key, value in attributes.items()})
    parent.append(node)
    return node


def replace(parent, name, attributes=None):
    """Replace one property, respecting python-docx's schema order where known."""
    for previous in list(parent.findall(_tag(name))):
        parent.remove(previous)
    if attributes is None:
        return None
    node = element(parent, name, **attributes)
    sequence = getattr(parent, "_tag_seq", ())
    # CT_* classes delete _tag_seq after constructing their descriptors. Use
    # insert_element_before with the ISO property order for the extras we add.
    orders = {
        "tblPr": ("tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription", "tblPrChange"),
        "rPr": ("rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dStrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange"),
        "pPr": ("pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"),
        "tcPr": ("cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark", "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange"),
        "sectPr": ("headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"),
    }
    order = orders.get(parent.tag.rsplit("}", 1)[-1], sequence)
    if name in order:
        successors = {_tag(tag) for tag in order[order.index(name) + 1:]}
        following = next((child for child in parent if child is not node and child.tag in successors), None)
        if following is not None:
            following.addprevious(node)
    return node


def font_extras(target, values):
    highlights = {"black", "blue", "cyan", "green", "magenta", "red", "yellow", "white", "darkBlue", "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow", "darkGray", "lightGray", "none"}
    if "highlight" in values and values["highlight"] is not None and values["highlight"] not in highlights:
        raise ValueError("Unknown Word highlight color.")
    spacing = values.get("character_spacing_pt")
    if spacing is not None and not math.isfinite(float(spacing)):
        raise ValueError("Character spacing must be finite.")
    scale = values.get("scale_percent")
    if scale is not None and (isinstance(scale, bool) or int(scale) != scale or not 1 <= scale <= 600):
        raise ValueError("Character scale must be an integer from 1 to 600 percent.")
    for key, allowed in (("font_names", {"ascii", "hAnsi", "eastAsia", "cs"}), ("languages", {"latin", "eastAsia", "bidi"})):
        if key in values and values[key] is not None:
            if not isinstance(values[key], dict) or values[key].keys() - allowed:
                raise ValueError("Invalid " + key + " script keys.")
            if any(value is not None and (not isinstance(value, str) or not value.strip()) for value in values[key].values()):
                raise ValueError("Script values must be non-empty strings or None.")
    if target is None:
        return
    pr = target._element.get_or_add_rPr()
    for key, tag, value in (("highlight", "highlight", values.get("highlight")), ("character_spacing_pt", "spacing", None if spacing is None else round(float(spacing) * 20)), ("scale_percent", "w", scale)):
        if key in values:
            replace(pr, tag, None if value is None else {"val": value})
    for key, tag in (("font_names", "rFonts"), ("languages", "lang")):
        if key not in values:
            continue
        items = values[key]
        node = pr.find(_tag(tag))
        if items is None:
            replace(pr, tag)
            continue
        if node is None:
            node = replace(pr, tag, {})
        for script, value in items.items():
            attr = "val" if script == "latin" else script
            if tag == "rFonts":
                node.attrib.pop(_tag("cstheme" if script == "cs" else script + "Theme"), None)
            if value is None:
                node.attrib.pop(_tag(attr), None)
            else:
                node.set(_tag(attr), value)


def paragraph_extras(target, values):
    for key in PARAGRAPH_BOOLEANS:
        if key in values and values[key] is not None and not isinstance(values[key], bool):
            raise ValueError(key + " must be boolean or None to inherit.")
    if "line_spacing_pt" in values and "line_spacing" in values:
        raise ValueError("Choose point or multiple line spacing.")
    if "line_spacing_rule" in values and "line_spacing_pt" not in values:
        raise ValueError("line_spacing_rule requires line_spacing_pt.")
    rule = values.get("line_spacing_rule", "exact")
    if rule not in {"exact", "at_least"}:
        raise ValueError("Point line spacing uses exact or at_least.")
    spacing = values.get("line_spacing_pt")
    if spacing is not None:
        _points(spacing, positive=True)
    stops = values.get("tab_stops")
    if stops is not None:
        for stop in stops:
            if stop.keys() - {"position_pt", "alignment", "leader"} or "position_pt" not in stop:
                raise ValueError("Tabs require position_pt and optional alignment/leader.")
            _points(stop["position_pt"])
            if stop.get("alignment", "left") not in {"left", "center", "right", "decimal", "bar", "clear"}:
                raise ValueError("Unknown tab alignment.")
            if stop.get("leader", "none") not in {"none", "dot", "hyphen", "underscore", "heavy", "middleDot"}:
                raise ValueError("Unknown tab leader.")
        positions = [round(float(stop["position_pt"]) * 20) for stop in stops]
        if len(positions) != len(set(positions)):
            raise ValueError("Tab positions must be unique.")
    if target is None:
        return
    for key, name in PARAGRAPH_BOOLEANS.items():
        if key in values:
            value = values[key]
            replace(target._element.get_or_add_pPr(), name, None if value is None else {"val": "1" if value else "0"})
    if "line_spacing_pt" in values:
        from docx.enum.text import WD_LINE_SPACING
        target.paragraph_format.line_spacing = None if spacing is None else _points(spacing, positive=True)
        target.paragraph_format.line_spacing_rule = None if spacing is None else WD_LINE_SPACING.EXACTLY if rule == "exact" else WD_LINE_SPACING.AT_LEAST
    if "tab_stops" in values:
        pr = target._element.get_or_add_pPr()
        tabs = replace(pr, "tabs", None if stops is None else {})
        for stop in sorted(stops or [], key=lambda stop: stop["position_pt"]):
            element(tabs, "tab", val=stop.get("alignment", "left"), leader=stop.get("leader", "none"), pos=round(float(stop["position_pt"]) * 20))


def cell_borders(cell, borders):
    if borders.keys() - {"top", "left", "bottom", "right", "insideH", "insideV", "start", "end", "tl2br", "tr2bl"}:
        raise ValueError("Unknown cell border edge.")
    converted = {}
    for edge, spec in borders.items():
        if spec is None:
            converted[edge] = None
            continue
        if spec.keys() - {"style", "width_pt", "color", "space_pt"}:
            raise ValueError("Unknown border property.")
        style = spec.get("style", "single")
        if style not in {"nil", "none", "single", "double", "dotted", "dashed", "thick", "dashSmallGap", "dotDash", "dotDotDash", "triple", "wave", "doubleWave"}:
            raise ValueError("Unsupported border style.")
        width = float(spec.get("width_pt", .5))
        if not math.isfinite(width) or not .25 <= width <= 12:
            raise ValueError("Border width must be between .25 and 12 pt.")
        space = float(spec.get("space_pt", 0))
        if not math.isfinite(space) or not 0 <= space <= 31 or int(space) != space:
            raise ValueError("Border spacing must be an integer from 0 to 31 pt.")
        converted[edge] = {"val": style, "sz": round(width * 8), "color": _color(spec.get("color", "000000")), "space": int(space)}
    pr = cell._tc.get_or_add_tcPr()
    node = pr.find(_tag("tcBorders"))
    if node is None:
        node = replace(pr, "tcBorders", {})
    for edge, attributes in converted.items():
        replace(node, edge, attributes)
