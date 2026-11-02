"""Native HWP table cell normalization and binary payload construction."""
from __future__ import annotations

import struct
from collections.abc import (
    Sequence,
)
from dataclasses import (
    dataclass,
)
from typing import cast

from ._hwp_constants import (
    DEFAULT_HWP_TABLE_BORDER_FILL_ID,
    DEFAULT_HWP_TABLE_CELL_MARGIN,
)


def _normalize_table_cell_texts(
    rows: int,
    cols: int,
    cell_text: str | None,
    cell_texts: Sequence[str] | Sequence[Sequence[str]] | None,
) -> list[list[str]]:
    if rows < 1 or cols < 1:
        raise ValueError("rows and cols must both be at least 1.")

    matrix: list[list[str]]
    if cell_texts is None:
        matrix = [["" for _ in range(cols)] for _ in range(rows)]
        if cell_text is not None:
            matrix[0][0] = cell_text
        return matrix

    values = list(cell_texts)
    if values and isinstance(values[0], (list, tuple)):
        if len(values) != rows:
            raise ValueError(f"cell_texts row count must match rows={rows}.")
        matrix = []
        for row_index, row_values in enumerate(values):
            row_list = list(row_values)  # type: ignore[arg-type]
            if len(row_list) != cols:
                raise ValueError(f"cell_texts[{row_index}] column count must match cols={cols}.")
            matrix.append([str(value) for value in row_list])
        return matrix

    flat_values = [str(value) for value in values]  # type: ignore[arg-type]
    if len(flat_values) != rows * cols:
        raise ValueError(f"flat cell_texts length must be rows * cols ({rows * cols}).")
    return [flat_values[row_index * cols : (row_index + 1) * cols] for row_index in range(rows)]


def _normalize_table_measurements(count: int, values: Sequence[int] | None, *, default: int, name: str) -> list[int]:
    if values is None:
        return [default] * count
    result = [int(value) for value in values]
    if len(result) != count:
        raise ValueError(f"{name} length must match {count}.")
    if any(value < 1 for value in result):
        raise ValueError(f"{name} values must be positive integers.")
    return result


def _normalize_table_spans(
    rows: int,
    cols: int,
    spans: dict[tuple[int, int], tuple[int, int]] | None,
) -> dict[tuple[int, int], tuple[int, int]]:
    normalized: dict[tuple[int, int], tuple[int, int]] = {}
    for row in range(rows):
        for col in range(cols):
            normalized[(row, col)] = (1, 1)
    if spans is None:
        return normalized
    for key, value in spans.items():
        row, col = key
        row_span, col_span = value
        if not (0 <= row < rows and 0 <= col < cols):
            raise ValueError(f"cell_spans key {(row, col)} is out of range.")
        row_span = int(row_span)
        col_span = int(col_span)
        if row_span < 1 or col_span < 1:
            raise ValueError("cell span values must be positive integers.")
        if row + row_span > rows or col + col_span > cols:
            raise ValueError(f"cell span {(row_span, col_span)} at {(row, col)} exceeds table bounds.")
        normalized[(row, col)] = (row_span, col_span)
    return normalized


@dataclass(frozen=True)
class _TableCellSpec:
    row: int
    col: int
    row_span: int
    col_span: int
    width: int
    height: int
    text: str
    border_fill_id: int
    margins: tuple[int, int, int, int]


def _build_table_cell_specs(
    matrix: list[list[str]],
    *,
    row_heights: Sequence[int],
    col_widths: Sequence[int],
    cell_spans: dict[tuple[int, int], tuple[int, int]] | None,
    cell_border_fill_ids: dict[tuple[int, int], int] | None,
    default_border_fill_id: int,
    cell_margins: dict[tuple[int, int], tuple[int, int, int, int]] | None,
) -> list[_TableCellSpec]:
    rows = len(matrix)
    cols = len(matrix[0]) if matrix else 0
    span_map = _normalize_table_spans(rows, cols, cell_spans)
    covered = [[False for _ in range(cols)] for _ in range(rows)]
    specs: list[_TableCellSpec] = []
    for row in range(rows):
        for col in range(cols):
            if covered[row][col]:
                continue
            row_span, col_span = span_map[(row, col)]
            for covered_row in range(row, row + row_span):
                for covered_col in range(col, col + col_span):
                    if covered[covered_row][covered_col]:
                        raise ValueError(f"overlapping cell span at {(row, col)}.")
                    covered[covered_row][covered_col] = True
                    if (covered_row, covered_col) != (row, col) and matrix[covered_row][covered_col]:
                        raise ValueError(
                            f"cell_texts[{covered_row}][{covered_col}] must be empty because it is covered by span {(row_span, col_span)} at {(row, col)}."
                        )
            border_fill_id = int(cell_border_fill_ids.get((row, col), default_border_fill_id) if cell_border_fill_ids else default_border_fill_id)
            margins = cell_margins.get((row, col), DEFAULT_HWP_TABLE_CELL_MARGIN) if cell_margins else DEFAULT_HWP_TABLE_CELL_MARGIN
            specs.append(
                _TableCellSpec(
                    row=row,
                    col=col,
                    row_span=row_span,
                    col_span=col_span,
                    width=sum(col_widths[col : col + col_span]),
                    height=sum(row_heights[row : row + row_span]),
                    text=matrix[row][col],
                    border_fill_id=border_fill_id,
                    margins=cast(tuple[int, int, int, int], tuple(int(value) for value in margins)),
                )
            )
    return specs


def _build_table_record_payload(
    rows: int,
    cols: int,
    *,
    cell_spacing: int = 0,
    margin_left: int = 510,
    margin_right: int = 510,
    margin_top: int = 141,
    margin_bottom: int = 141,
    default_row_height: int = 1,
    border_fill_id: int = DEFAULT_HWP_TABLE_BORDER_FILL_ID,
    valid_zone_info_count: int = 0,
) -> bytes:
    attr = 0x04000006 if rows * cols > 1 else 0x00000006
    payload = bytearray()
    payload.extend(struct.pack("<IHHHHHHH", attr, rows, cols, cell_spacing, margin_left, margin_right, margin_top, margin_bottom))
    for _ in range(rows):
        payload.extend(int(default_row_height).to_bytes(2, "little", signed=False))
    payload.extend(int(border_fill_id).to_bytes(2, "little", signed=False))
    payload.extend(int(valid_zone_info_count).to_bytes(2, "little", signed=False))
    return bytes(payload)


def _build_table_cell_list_header_payload(
    cell: _TableCellSpec,
    *,
    list_flags: int | None = None,
) -> bytes:
    attr = 0x04000020 if list_flags is None else int(list_flags)
    margin_left, margin_right, margin_top, margin_bottom = cell.margins
    payload = bytearray()
    payload.extend((1).to_bytes(4, "little", signed=True))
    payload.extend(int(attr).to_bytes(4, "little", signed=False))
    payload.extend(int(cell.col).to_bytes(2, "little", signed=False))
    payload.extend(int(cell.row).to_bytes(2, "little", signed=False))
    payload.extend(int(cell.col_span).to_bytes(2, "little", signed=False))
    payload.extend(int(cell.row_span).to_bytes(2, "little", signed=False))
    payload.extend(int(cell.width).to_bytes(4, "little", signed=False))
    payload.extend(int(cell.height).to_bytes(4, "little", signed=False))
    payload.extend(int(margin_left).to_bytes(2, "little", signed=False))
    payload.extend(int(margin_right).to_bytes(2, "little", signed=False))
    payload.extend(int(margin_top).to_bytes(2, "little", signed=False))
    payload.extend(int(margin_bottom).to_bytes(2, "little", signed=False))
    payload.extend(int(cell.border_fill_id).to_bytes(2, "little", signed=False))
    payload.extend(int(cell.width).to_bytes(4, "little", signed=False))
    payload.extend(b"\x00" * 9)
    return bytes(payload)


# Preserve historical import and pickle identities.
for _public_type in (_TableCellSpec,):
    _public_type.__module__ = "jakal_hwpx.hwp_binary"
del _public_type
