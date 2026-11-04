"""Floating placement and content-preserving table split/join operations."""
from __future__ import annotations

from copy import deepcopy
import math

from .editing import WordTable, _tag, _points, _balanced, _only_properties
from .editing_format import replace
from .editing_tables import geometry


def format_layout(table, *, alignment=None, indent_pt=None, autofit=None, wrap=None, x_pt=0, y_pt=0,
                  horizontal_from="margin", vertical_from="text", distance_pt=None):
    from docx.enum.table import WD_TABLE_ALIGNMENT
    alignments = {"left": WD_TABLE_ALIGNMENT.LEFT, "center": WD_TABLE_ALIGNMENT.CENTER, "right": WD_TABLE_ALIGNMENT.RIGHT}
    if alignment is not None and alignment not in alignments:
        raise ValueError("Table alignment is left, center, or right.")
    if wrap not in {None, "inline", "around"} or horizontal_from not in {"page", "margin", "text"} or vertical_from not in {"page", "margin", "text"}:
        raise ValueError("Unknown table wrap/reference frame.")
    for value in (x_pt, y_pt, indent_pt):
        if value is not None and not math.isfinite(float(value)):
            raise ValueError("Table offsets must be finite.")
    distances = distance_pt or {}
    if distances.keys() - {"top", "bottom", "left", "right"}:
        raise ValueError("Table text distances use top, bottom, left, right.")
    spaces = {key: round(_points(value) / 635) for key, value in distances.items()}
    native = table._native
    pr = native._tbl.tblPr
    if alignment is not None:
        native.alignment = alignments[alignment]
    if autofit is not None:
        native.autofit = bool(autofit)
    if indent_pt is not None:
        replace(pr, "tblInd", {"w": round(float(indent_pt) * 20), "type": "dxa"})
    if wrap == "inline":
        replace(pr, "tblpPr")
    elif wrap == "around":
        replace(pr, "tblpPr", {"horzAnchor": horizontal_from, "vertAnchor": vertical_from,
                               "tblpX": round(float(x_pt) * 20), "tblpY": round(float(y_pt) * 20),
                               **{key + "FromText": value for key, value in spaces.items()}})


def split_table(table, index):
    from docx.oxml import OxmlElement
    from .identity import ID
    rows, _, regions = geometry(table)
    if table._node.tblPr.find(_tag("tblpPr")) is not None:
        raise ValueError("Make a floating table inline before splitting it.")
    if not isinstance(index, int) or not 0 < index < len(rows):
        raise ValueError("Split needs a row strictly inside the table.")
    if any(region.top < index < region.bottom for region in regions):
        raise ValueError("A split cannot cut a vertical merge; split that cell first.")
    left, right = deepcopy(table._node), deepcopy(table._node)
    for row in left.tr_lst[index:]:
        left.remove(row)
    for row in right.tr_lst[:index]:
        right.remove(row)
    _balanced(left)
    _balanced(right)
    new = OxmlElement("w:tbl")
    new.append(deepcopy(table._node.tblPr))
    new.append(deepcopy(table._node.tblGrid))
    new.attrib.update({key: value for key, value in table._node.attrib.items() if key != ID})
    for row in rows[index:]:
        new.append(row)
    separator = OxmlElement("w:p")
    table._node.addnext(separator)
    separator.addnext(new)
    table._owner._ensure_tail()
    return WordTable(table._owner, new)


def join_tables(first, second, *, adopt_style=False):
    from lxml import etree
    first._check()
    second._check()
    if first._owner._editor is not second._owner._editor or first._owner.part is not second._owner.part or first._node.getparent() is not second._node.getparent():
        raise ValueError("Joined tables must be siblings in one document part.")
    if first._node is second._node:
        raise ValueError("A table cannot join itself.")
    _, a, _ = geometry(first)
    _, b, _ = geometry(second)
    if a != b:
        raise ValueError("Joined tables must have identical column grids and widths.")
    if not adopt_style and etree.tostring(first._node.tblPr) != etree.tostring(second._node.tblPr):
        raise ValueError("Table properties differ; explicitly adopt_style=True to keep the first table's style.")
    between, node = [], first._node.getnext()
    while node is not second._node:
        if node is None or node.tag != _tag("p") or not _only_properties(node) or node.find("./" + _tag("pPr") + "/" + _tag("sectPr")) is not None:
            raise ValueError("Tables must be adjacent, separated only by empty paragraphs.")
        between.append(node)
        node = node.getnext()
    _balanced(first._node)
    _balanced(second._node)
    for row in list(second._node.tr_lst):
        if row.trPr is not None:
            for header in list(row.trPr.findall(_tag("tblHeader"))):
                row.trPr.remove(header)
        first._node.append(row)
    for node in between + [second._node]:
        node.getparent().remove(node)
    first._owner._ensure_tail()
