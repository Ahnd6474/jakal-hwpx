"""Bidirectional table mapping between Word and the shared document model."""
from __future__ import annotations

from typing import Any

from ..._hancom_models import Table


def read_table(reader: Any, source: Any) -> Table:
    """Project a python-docx table into the shared table model."""
    from docx.oxml.ns import qn

    from .formatting import capture_table_properties, color as resolve_color, table_properties
    from .reader import _border_styles, _cell_margins, _style_chain, _units

    rows, cols = len(source.rows), len(source.columns)
    if not rows or not cols:
        raise ValueError("DOCX contains a table with no rows or columns.")
    style_nodes = [style.element for style in reversed(list(_style_chain(source.style)))]
    texts = [["" for _ in range(cols)] for _ in range(rows)]
    coordinates: dict[Any, list[tuple[int, int]]] = {}
    vertical_aligns: dict[tuple[int, int], str] = {}
    cell_margins: dict[tuple[int, int], dict[str, int]] = {}
    cell_fill_colors: dict[tuple[int, int], str] = {}
    cell_patterns: dict[str, dict[str, str]] = {}
    cell_border_styles: dict[tuple[int, int], dict[str, dict[str, str | int]]] = {}
    # A custom table style without inherited tblCellMar has zero padding.
    # Do not let the HWPX factory's 141-unit cell defaults add new spacing.
    table_margins = {side: 0 for side in ("left", "right", "top", "bottom")}
    table_borders = {}
    for style in reversed(list(_style_chain(source.style))):
        properties = style.element.find(qn("w:tblPr"))
        if properties is not None:
            table_margins.update(_cell_margins(properties.find(qn("w:tblCellMar"))))
            table_borders.update(_border_styles(properties.find(qn("w:tblBorders"))))
    table_margins.update(_cell_margins(source._tbl.tblPr.find(qn("w:tblCellMar"))))
    table_borders.update(_border_styles(source._tbl.tblPr.find(qn("w:tblBorders"))))
    rich_cells = {}
    for row_index, row in enumerate(source.rows):
        for column, cell in enumerate(row.cells, start=row.grid_cols_before):
            if column >= cols:
                raise ValueError("DOCX table cells exceed the declared column grid.")
            positions = coordinates.setdefault(cell._tc, [])
            if not positions:
                texts[row_index][column] = cell.text
                resolved = table_properties(source, row_index, column, style_nodes)
                previous_context = reader.table_context
                reader.table_context = resolved
                try:
                    rich_cells[(row_index, column)] = (cell.text, reader.container_blocks(cell))
                finally:
                    reader.table_context = previous_context
                align = {0: "TOP", 1: "CENTER", 3: "BOTTOM"}.get(cell.vertical_alignment)
                vertical_aligns[(row_index, column)] = align or "TOP"
                properties = resolved["tcPr"]
                overrides = _cell_margins(properties.find(qn("w:tcMar")) if properties is not None else None)
                if overrides:
                    cell_margins[(row_index, column)] = {**table_margins, **overrides} if table_margins else overrides
                borders = _border_styles(properties.find(qn("w:tcBorders")) if properties is not None else None)
                conditional_borders = _border_styles(resolved["tblPr"].find(qn("w:tblBorders")))
                borders = {**conditional_borders, **borders}
                for side, inside, outer in (
                    ("top", "insideH", row_index == 0),
                    ("bottom", "insideH", row_index == rows - 1),
                    ("left", "insideV", column == 0),
                    ("right", "insideV", column == cols - 1),
                ):
                    if not outer and inside in borders:
                        borders[side] = borders[inside]
                if table_borders or borders:
                    cell_border_styles[(row_index, column)] = {**table_borders, **borders}
                    if any(
                        spec.get("style") not in {"single", "dashed", "dotted", "double", "none", "nil"}
                        for spec in cell_border_styles[(row_index, column)].values()
                    ):
                        reader.warn("DOCX table border style is approximated as a solid line.")
                shading = properties.find(qn("w:shd")) if properties is not None else None
                if shading is not None:
                    fill = resolve_color(shading, reader.theme, fill=True)
                    if fill:
                        cell_fill_colors[(row_index, column)] = fill
                    pattern = shading.get(qn("w:val"), "clear")
                    if pattern not in {"clear", "nil"}:
                        cell_patterns[f"{row_index},{column}"] = {
                            "type": pattern,
                            "foreground": resolve_color(shading, reader.theme) or "#000000",
                            "background": fill or "#FFFFFF",
                        }
            positions.append((row_index, column))
    spans = {}
    for positions in coordinates.values():
        first_row, first_col = positions[0]
        row_span = max(row for row, _ in positions) - first_row + 1
        col_span = max(col for _, col in positions) - first_col + 1
        if row_span > 1 or col_span > 1:
            spans[(first_row, first_col)] = (row_span, col_span)
    widths = [column.width for column in source.columns]
    heights = [row.height for row in source.rows]
    result = Table(
        rows=rows,
        cols=cols,
        cell_texts=texts,
        col_widths=[_units(width) for width in widths] if all(width is not None for width in widths) else None,
        row_heights=[_units(height) if height is not None else 0 for height in heights]
        if any(height is not None for height in heights)
        else None,
        cell_spans=spans or None,
        cell_vertical_aligns=vertical_aligns or None,
        table_margins=table_margins or None,
        cell_margins=cell_margins or None,
        cell_fill_colors=cell_fill_colors or None,
        cell_border_styles=cell_border_styles or None,
        repeat_header=bool(
            source.rows[0]._tr.xpath("./w:trPr/w:tblHeader[not(@w:val='0' or @w:val='false' or @w:val='off')]")
        ),
    )
    setattr(result, "_docx_cell_content", rich_cells)
    setattr(
        result,
        "_word_extra",
        {
            "table": capture_table_properties(source),
            "patterns": cell_patterns,
            "fills": {f"{row},{col}": color for (row, col), color in cell_fill_colors.items()},
            "borders": {f"{row},{col}": borders for (row, col), borders in cell_border_styles.items()},
        },
    )
    native = getattr(result, "_word_extra")["table"]
    native["row_controls"] = [reader.controls.get(row._tr, []) for row in source.rows]
    native["cell_controls"] = {
        f"{row},{col}": reader.controls.get(cell._tc, [])
        for row, item in enumerate(source.rows)
        for col, cell in enumerate(item.cells, start=item.grid_cols_before)
    }
    pending = reader.pending_breaks.pop(source._tbl.getparent(), None)
    if pending:
        setattr(result, "_hwpx_paragraph_attributes", {"pageBreak" if pending == "page" else "columnBreak": "1"})
    return result


def write_table(writer: Any, container: Any, block: Table) -> None:
    """Write a shared table model into a python-docx container."""
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT

    from .features import parse, set_property, wrap_controls
    from .writer import _color, _element, _length

    extra = getattr(block, "_word_extra", {})
    native = extra.get("table", {})

    details = getattr(block, "_docx_table_details", {})
    widths = block.col_widths or details.get("col_widths")
    heights = block.row_heights or details.get("row_heights")
    original_columns = bool(native and widths == native.get("col_widths"))
    if native and block.row_heights is None:
        heights = native.get("row_heights")
    fills = block.cell_fill_colors if block.cell_fill_colors is not None else details.get("cell_fill_colors", {})
    border_styles = (
        block.cell_border_styles if block.cell_border_styles is not None else details.get("cell_border_styles", {})
    )
    if container is writer.word or hasattr(container, "_tc"):
        table = container.add_table(rows=block.rows, cols=block.cols)
    else:
        table = container.add_table(rows=block.rows, cols=block.cols, width=_length(sum(widths or [36000])))
    if native.get("tblPr"):
        for node in parse(native["tblPr"]):
            set_property(table._tbl.tblPr, node)
    for row_index, properties in enumerate(native.get("rows", [])):
        if properties and row_index < block.rows:
            for node in parse(properties):
                set_property(table.rows[row_index]._tr.get_or_add_trPr(), node)
    if widths:
        if not original_columns:
            table.autofit = False
        for column, width in zip(table.columns, widths):
            column.width = _length(width)
        for row in table.rows:
            for cell, width in zip(row.cells, widths):
                cell.width = _length(width)
    for (row, col), (rows, cols) in (block.cell_spans or {}).items():
        table.cell(row, col).merge(table.cell(row + rows - 1, col + cols - 1))
    seen: set[Any] = set()
    for row in range(block.rows):
        if heights and row < len(heights) and heights[row] > 0:
            table.rows[row].height = _length(heights[row])
        if row == 0 and block.repeat_header:
            _element(table.rows[row]._tr.get_or_add_trPr(), "tblHeader")
        for col in range(block.cols):
            cell = table.cell(row, col)
            if cell._tc in seen:
                continue
            seen.add(cell._tc)
            value = (
                block.cell_texts[row][col]
                if row < len(block.cell_texts) and col < len(block.cell_texts[row])
                else ""
            )
            rich = getattr(block, "_docx_cell_content", {}).get((row, col))
            if rich and rich[0] == value:
                writer.blocks(cell, rich[1], first=cell.paragraphs[0], nested=True)
            else:
                cell.paragraphs[0].add_run(value)
            properties = cell._tc.get_or_add_tcPr()
            source_properties = native.get("cells", {}).get(f"{row},{col}")
            if source_properties:
                for node in parse(source_properties):
                    if node.tag.rsplit("}", 1)[-1] not in {"gridSpan", "vMerge", "tcW"}:
                        set_property(properties, node)
            fill = _color(fills.get((row, col)))
            if fill and fills.get((row, col)) != extra.get("fills", {}).get(f"{row},{col}"):
                _element(properties, "shd", val="clear", fill=fill)
            margins = dict(block.table_margins or {})
            margins.update((block.cell_margins or {}).get((row, col), {}))
            if margins:
                node = _element(properties, "tcMar")
                for side, margin_value in margins.items():
                    if side in {"left", "right", "top", "bottom"}:
                        _element(node, side, w=round(margin_value / 5), type="dxa")
            alignment = (block.cell_vertical_aligns or {}).get((row, col), "CENTER")
            cell.vertical_alignment = {
                "TOP": WD_CELL_VERTICAL_ALIGNMENT.TOP,
                "CENTER": WD_CELL_VERTICAL_ALIGNMENT.CENTER,
                "BOTTOM": WD_CELL_VERTICAL_ALIGNMENT.BOTTOM,
            }.get(alignment, WD_CELL_VERTICAL_ALIGNMENT.CENTER)
            borders = border_styles.get((row, col), {})
            if borders and borders != extra.get("borders", {}).get(f"{row},{col}"):
                node = _element(properties, "tcBorders")
                for side, values in borders.items():
                    if side not in {"left", "right", "top", "bottom"}:
                        continue
                    value = str(values.get("style", values.get("type", "SOLID")))
                    shape = {
                        "SOLID": "single",
                        "NONE": "nil",
                        "DASH": "dashed",
                        "DOT": "dotted",
                        "DOUBLE_SLIM": "double",
                        "single": "single",
                        "none": "nil",
                        "nil": "nil",
                        "dashed": "dashed",
                        "dotted": "dotted",
                        "double": "double",
                    }.get(value, "single")
                    width = float(str(values.get("width", "0.1")).replace("mm", ""))
                    size = int(values["size"]) if "size" in values else max(2, round(width * 72 / 25.4 * 8))
                    _element(
                        node,
                        side,
                        val=shape,
                        sz=size,
                        color=_color(str(values.get("color", "000000"))) or "000000",
                    )
    # Wrap rows/cells after all grid operations. python-docx intentionally
    # omits SDT-wrapped rows from its ordinary table.rows collection.
    native_rows = list(table.rows)
    cell_rows = [list(item.cells) for item in native_rows]
    wrapped_cells = set()
    for row, cells in enumerate(cell_rows):
        for col, cell in enumerate(cells):
            if cell._tc in wrapped_cells:
                continue
            wrapped_cells.add(cell._tc)
            wrap_controls(cell._tc, native.get("cell_controls", {}).get(f"{row},{col}", []), level="cell")
    for row, item in enumerate(native_rows):
        controls = native.get("row_controls", [])
        if row < len(controls):
            wrap_controls(item._tr, controls[row], level="row")
