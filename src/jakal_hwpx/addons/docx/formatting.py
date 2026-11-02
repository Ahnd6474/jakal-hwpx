"""Resolve Word theme and conditional table formatting before native mapping."""
from __future__ import annotations

from copy import deepcopy
import colorsys
from typing import Any

from .features import xml

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def color(node: Any, theme: Any, *, fill: bool = False) -> str | None:
    if node is None:
        return None
    key = "fill" if fill else "val" if node.tag == f"{{{W}}}color" else "color"
    value = node.get(f"{{{W}}}{key}", "")
    theme_key = "themeFill" if fill else "themeColor"
    theme_name = node.get(f"{{{W}}}{theme_key}")
    if theme_name and theme is not None:
        theme_name = {"background1": "lt1", "text1": "dk1", "background2": "lt2", "text2": "dk2",
                      "dark1": "dk1", "light1": "lt1", "dark2": "dk2", "light2": "lt2", "hyperlink": "hlink",
                      "followedHyperlink": "folHlink"}.get(theme_name, theme_name)
        nodes = theme.findall(f".//{{{A}}}clrScheme/{{{A}}}{theme_name}/*")
        if nodes:
            value = nodes[0].get("lastClr", nodes[0].get("val", value))
    if len(value) != 6 or any(char not in "0123456789abcdefABCDEF" for char in value):
        return None
    rgb = tuple(int(value[index:index + 2], 16) / 255 for index in (0, 2, 4))
    hue, luminance, saturation = colorsys.rgb_to_hls(*rgb)
    prefix = "themeFill" if fill else "theme"
    shade = node.get(f"{{{W}}}{prefix}Shade")
    tint = node.get(f"{{{W}}}{prefix}Tint")
    if shade:
        luminance *= int(shade, 16) / 255
    if tint:
        luminance = luminance * int(tint, 16) / 255 + 1 - int(tint, 16) / 255
    rgb = colorsys.hls_to_rgb(hue, luminance, saturation)
    return "#" + "".join(f"{min(255, max(0, round(channel * 255))):02X}" for channel in rgb)


def table_regions(table: Any, row: int, col: int, styles: list[Any]) -> list[str]:
    """Office precedence: rows, columns, edge columns, edge rows, corners.

    https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/f7f85a73-5f60-4f7e-a725-4f1f06198388
    """
    look = table._tbl.tblPr.find(f"{{{W}}}tblLook")
    mask = int(look.get(f"{{{W}}}val", "0000"), 16) if look is not None else 0

    def flag(name, bit):
        value = look.get(f"{{{W}}}{name}") if look is not None else None
        return value not in {"0", "false", "off"} if value is not None else bool(mask & bit)

    first_row, last_row = flag("firstRow", 0x20), flag("lastRow", 0x40)
    first_col, last_col = flag("firstColumn", 0x80), flag("lastColumn", 0x100)
    regions = ["wholeTable"]
    rows, cols = len(table.rows), len(table.columns)

    def band_size(tag):
        for style in reversed(styles):
            node = style.find(f"./{{{W}}}tblPr/{{{W}}}{tag}")
            if node is not None:
                return max(1, int(node.get(f"{{{W}}}val", "1")))
        return 1

    if not flag("noHBand", 0x200) and not (first_row and row == 0) and not (last_row and row == rows - 1):
        regions.append(("band1Horz", "band2Horz")[((row - int(first_row)) // band_size("tblStyleRowBandSize")) % 2])
    if not flag("noVBand", 0x400) and not (first_col and col == 0) and not (last_col and col == cols - 1):
        regions.append(("band1Vert", "band2Vert")[((col - int(first_col)) // band_size("tblStyleColBandSize")) % 2])
    for enabled, edge, value, expected in ((first_col, "firstCol", col, 0), (last_col, "lastCol", col, cols - 1),
                                           (first_row, "firstRow", row, 0), (last_row, "lastRow", row, rows - 1)):
        if enabled and value == expected:
            regions.append(edge)
    for enabled, name in ((first_row and first_col and row == col == 0, "nwCell"),
                          (first_row and last_col and row == 0 and col == cols - 1, "neCell"),
                          (last_row and first_col and row == rows - 1 and col == 0, "swCell"),
                          (last_row and last_col and row == rows - 1 and col == cols - 1, "seCell")):
        if enabled:
            regions.append(name)
    return regions


def table_properties(table: Any, row: int, col: int, styles: list[Any]) -> dict[str, Any]:
    from docx.oxml import OxmlElement

    result = {tag: OxmlElement("w:" + tag) for tag in ("tblPr", "tcPr", "pPr", "rPr")}

    def merge(source):
        if source is None:
            return
        target = result.get(source.tag.rsplit("}", 1)[-1])
        if target is None:
            return
        for node in source:
            previous = target.find(node.tag)
            if previous is not None:
                if node.tag in {f"{{{W}}}tcBorders", f"{{{W}}}tblBorders", f"{{{W}}}tcMar"}:
                    for side in node:
                        existing = previous.find(side.tag)
                        if existing is not None:
                            previous.remove(existing)
                        previous.append(deepcopy(side))
                    continue
                target.remove(previous)
            target.append(deepcopy(node))

    for style in styles:
        for node in style:
            merge(node)
    for region in table_regions(table, row, col, styles):
        for style in styles:
            for override in style.findall(f"{{{W}}}tblStylePr"):
                if override.get(f"{{{W}}}type") == region:
                    for node in override:
                        merge(node)
    merge(table._tbl.tblPr)
    merge(table.cell(row, col)._tc.tcPr)
    return result


def capture_table_properties(table: Any) -> dict:
    return {"tblPr": xml(table._tbl.tblPr),
            "rows": [xml(row._tr.trPr) for row in table.rows],
            "col_widths": [round(int(column.width) / 127) if column.width is not None else None for column in table.columns],
            "row_heights": [round(int(row.height) / 127) if row.height is not None else 0 for row in table.rows],
            "cells": {f"{row},{col}": xml(cell._tc.tcPr) for row, source in enumerate(table.rows)
                      for col, cell in enumerate(source.cells, start=source.grid_cols_before)}}


def apply_hwpx_character(element: Any, style: Any) -> None:
    """Map native spacing/baseline offsets in addition to retaining exact OOXML."""
    from ...namespaces import qname
    from lxml import etree

    raw = getattr(style, "_word_extra", {}).get("rPr")
    if not raw:
        return
    properties = etree.fromstring(raw.encode("utf-8"))
    height = float(style.height or 1100)
    for tag, native, factor in (("spacing", "spacing", 5), ("position", "offset", 50)):
        source = properties.find(f"{{{W}}}{tag}")
        if source is None:
            continue
        node = element.find(qname("hh", native))
        if node is None:
            node = etree.SubElement(element, qname("hh", native))
        value = round(int(source.get(f"{{{W}}}val", "0")) * factor / height * 100)
        for lang in ("hangul", "latin", "hanja", "japanese", "other", "symbol", "user"):
            node.set(lang, str(value))


def apply_hwpx_pattern(document: Any, border_fill: Any, pattern: dict[str, str]) -> None:
    """Use a native hatch, or a tiled bitmap for Word percentage patterns."""
    from lxml import etree
    from ...namespaces import qname

    brush = border_fill.find(qname("hc", "fillBrush"))
    if brush is None:
        brush = etree.SubElement(border_fill, qname("hc", "fillBrush"))
    for child in list(brush):
        brush.remove(child)
    name = pattern["type"]
    hatch = {"horzStripe": "HORIZONTAL", "thinHorzStripe": "HORIZONTAL", "vertStripe": "VERTICAL", "thinVertStripe": "VERTICAL",
             "reverseDiagStripe": "BACK_SLASH", "thinReverseDiagStripe": "BACK_SLASH", "diagStripe": "SLASH",
             "thinDiagStripe": "SLASH", "horzCross": "CROSS", "thinHorzCross": "CROSS", "diagCross": "CROSS_DIAGONAL",
             "thinDiagCross": "CROSS_DIAGONAL"}.get(name)
    if hatch or name == "solid":
        etree.SubElement(brush, qname("hc", "winBrush"), faceColor=pattern["foreground"] if name == "solid" else pattern["background"],
                         hatchColor=pattern["foreground"], alpha="0", **({"hatchStyle": hatch} if hatch else {}))
        return
    if name.startswith("pct"):
        import hashlib
        import struct
        import zlib

        percent = float(name.removeprefix("pct")) / 100
        if not 0 <= percent <= 1:
            raise ValueError("Invalid Word percentage shading pattern.")
        foreground = bytes.fromhex(pattern["foreground"].lstrip("#"))
        background = bytes.fromhex(pattern["background"].lstrip("#"))
        # Ordered 8x8 Bayer mask gives the requested foreground coverage.
        def rank(x, y):
            value = 0
            for bit in range(3):
                xb, yb = (x >> bit) & 1, (y >> bit) & 1
                value = (value << 2) | ((xb ^ yb) << 1) | yb
            return value

        pixels = b"".join(b"\0" + b"".join(foreground if rank(x, y) < round(percent * 64) else background for x in range(8)) for y in range(8))
        def chunk(tag, payload):
            return struct.pack(">I", len(payload)) + tag + payload + struct.pack(">I", zlib.crc32(tag + payload))
        data = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 8, 8, 8, 2, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(pixels)) + chunk(b"IEND", b"")
        identity = "word-pattern-" + hashlib.sha256(data).hexdigest()[:16]
        document.add_or_replace_binary(identity + ".png", data, media_type="image/png", manifest_id=identity)
        image = etree.SubElement(brush, qname("hc", "imgBrush"), mode="TILE")
        etree.SubElement(image, qname("hc", "img"), binaryItemIDRef=identity, bright="0", contrast="0", effect="REAL_PIC", alpha="0")
        return
    raise ValueError("Unsupported Word shading pattern: " + name)
