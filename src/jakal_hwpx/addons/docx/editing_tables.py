"""Grid-aware table edits, retaining physical cell XML and relationships."""
from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass, field

from .editing import _balanced, _points, _tag
from .editing_format import element, replace


@dataclass
class Region:
    top: int
    left: int
    bottom: int
    right: int
    cells: dict = field(default_factory=dict)


def geometry(handle) -> tuple[list, list[int], list[Region]]:
    table = handle._native
    rows = list(table._tbl.tr_lst)
    widths = [int(col.w or _points(72)) for col in table._tbl.tblGrid.gridCol_lst]
    regions: list[Region] = []
    above: dict[int, Region] = {}
    for r, row in enumerate(rows):
        if row.grid_before or row.grid_after:
            raise ValueError("Structural edits require a complete table grid; omitted edge cells must be filled first.")
        c, current = 0, {}
        for cell in row.tc_lst:
            if cell.tcPr is not None and cell.tcPr.find(_tag("hMerge")) is not None:
                raise ValueError("Legacy hMerge tables require normalization to gridSpan before structural editing.")
            end = c + cell.grid_span
            if end > len(widths):
                raise ValueError("Table cell exceeds its column grid.")
            if cell.vMerge == "continue":
                region = above.get(c)
                if region is None or region.left != c or region.right != end:
                    raise ValueError("Inconsistent vertical merge geometry.")
                region.bottom = r + 1
                region.cells[r] = cell
            else:
                region = Region(r, c, r + 1, end, {r: cell})
                regions.append(region)
            if cell.vMerge is not None:
                for col in range(c, end):
                    current[col] = region
            c = end
        if c != len(widths):
            raise ValueError("Table rows must cover the whole column grid.")
        above = current
    return rows, widths, regions


def empty_cell(template=None):
    from docx.oxml import OxmlElement
    cell = OxmlElement("w:tc")
    if template is not None and template.tcPr is not None:
        cell.append(deepcopy(template.tcPr))
        for tag in ("gridSpan", "vMerge", "hMerge"):
            replace(cell.tcPr, tag)
    element(cell, "p")
    return cell


def rebuild(handle, rows, widths, regions):
    """Reuse each original cell once; continuation cells retain their properties."""
    table = handle._node
    for row in rows:
        for cell in list(row.tc_lst):
            row.remove(cell)
    for region in regions:
        source = region.cells[region.top] if region.top in region.cells else next(iter(region.cells.values()))
        for r in range(region.top, region.bottom):
            cell = region.cells.get(r)
            if cell is None:
                cell = empty_cell(source)
            pr = cell.get_or_add_tcPr()
            replace(pr, "gridSpan", {"val": region.right - region.left} if region.right - region.left > 1 else None)
            replace(pr, "vMerge", {"val": "restart" if r == region.top else "continue"} if region.bottom - region.top > 1 else None)
            cell.width = sum(widths[region.left:region.right])
            region.cells[r] = cell
    for r, row in enumerate(rows):
        for region in sorted((item for item in regions if item.top <= r < item.bottom), key=lambda item: item.left):
            row.append(region.cells[r])
    for old in list(table.tr_lst):
        table.remove(old)
    for row in rows:
        table.append(row)
    grid = table.tblGrid
    for old in list(grid):
        grid.remove(old)
    for width in widths:
        column = grid.add_gridCol()
        column.w = width


def _values(regions, axis, index, values):
    """Values correspond to grid slots; merged continuations must be empty."""
    assignments = []
    for offset, value in enumerate(values or []):
        r, c = (index, offset) if axis == "row" else (offset, index)
        region = next(item for item in regions if item.top <= r < item.bottom and item.left <= c < item.right)
        if (r, c) != (region.top, region.left):
            if value is not None and str(value):
                raise ValueError("A value cannot overwrite a continued merged cell; use None for that grid slot.")
        elif value is not None:
            assignments.append((r, c, str(value)))
    return assignments


def insert_axis(handle, axis, index, values, width_pt=72):
    from docx.oxml import OxmlElement
    rows, widths, regions = geometry(handle)
    length = len(rows) if axis == "row" else len(widths)
    other = len(widths) if axis == "row" else len(rows)
    index = length if index is None else index
    if not isinstance(index, int) or not 0 <= index <= length or values is not None and len(values) > other:
        raise ValueError("Insertion index or values do not fit the table grid.")
    width = int(_points(width_pt, positive=True)) if axis == "column" else None
    # All work above rebuild is on descriptors/new XML only: input validation
    # never partially mutates the source table, even if the caller catches it.
    for region in regions:
        start, stop = (region.top, region.bottom) if axis == "row" else (region.left, region.right)
        if axis == "row":
            region.cells = {r + (r >= index): cell for r, cell in region.cells.items()}
            region.top += start >= index
            region.bottom += stop > index or start >= index
        else:
            region.left += start >= index
            region.right += stop > index or start >= index
    for slot in range(other):
        r, c = (index, slot) if axis == "row" else (slot, index)
        if not any(item.top <= r < item.bottom and item.left <= c < item.right for item in regions):
            regions.append(Region(r, c, r + 1, c + 1, {r: empty_cell()}))
    assignments = _values(regions, axis, index, values)
    if axis == "row":
        rows.insert(index, OxmlElement("w:tr"))
    else:
        widths.insert(index, width)
    rebuild(handle, rows, widths, regions)
    for r, c, value in assignments:
        handle.set_cell_text(r, c, value)


def delete_axis(handle, axis, index):
    rows, widths, regions = geometry(handle)
    length = len(rows) if axis == "row" else len(widths)
    if not isinstance(index, int) or not 0 <= index < length:
        raise IndexError("Table deletion index is out of range.")
    if length == 1:
        raise ValueError("Cannot delete the last table row or column; delete the table instead.")
    discarded = []
    retained = []
    for region in regions:
        start, stop = (region.top, region.bottom) if axis == "row" else (region.left, region.right)
        if start == index and stop == index + 1:
            discarded.extend(region.cells.values())
            continue
        if axis == "row":
            cells = region.cells
            # If a merge origin is removed, promote its complete content to
            # the first surviving row. No text/image/field content is lost.
            if region.top == index:
                old_continuation = cells.get(index + 1)
                if old_continuation is not None and old_continuation.xpath(".//w:t[string-length(.)>0] | .//w:tab | .//w:br | .//w:drawing | .//w:pict | .//w:fldChar | .//w:fldSimple | .//w:footnoteReference | .//w:endnoteReference | .//w:bookmarkStart | .//w:bookmarkEnd | .//m:oMath"):
                    raise ValueError("Cannot promote a vertical merge with non-empty continuation content; split it first.")
                cells[index + 1] = cells[index]
            region.cells = {r - (r > index): cell for r, cell in cells.items() if r != index}
            region.top -= start > index
            region.bottom -= stop > index
        else:
            region.left -= start > index
            region.right -= stop > index
        retained.append(region)
    for cell in discarded:
        _balanced(cell)
    if axis == "row":
        rows.pop(index)
    else:
        widths.pop(index)
    rebuild(handle, rows, widths, retained)


def split_cell(handle, row, col):
    handle._cell(row, col)  # Validate bounds before searching geometry.
    rows, widths, regions = geometry(handle)
    region = next(item for item in regions if item.top <= row < item.bottom and item.left <= col < item.right)
    if region.bottom - region.top == 1 and region.right - region.left == 1:
        return
    regions.remove(region)
    origin = region.cells[region.top]
    for r in range(region.top, region.bottom):
        for c in range(region.left, region.right):
            cell = region.cells.get(r) if c == region.left else None
            if cell is None:
                cell = empty_cell(origin)
            regions.append(Region(r, c, r + 1, c + 1, {r: cell}))
    rebuild(handle, rows, widths, regions)
