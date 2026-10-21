from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
from typing import TYPE_CHECKING

from ._hwp_object_payloads import (  # noqa: F401 - compatibility re-exports
    _CONTROL_NATIVE_METADATA_PREFIX,
    _HANCOM_CONNECTLINE_COMPONENT_SIGNATURE,
    _HWP_OLE_DRAW_ASPECT_NAMES,
    _HWP_OLE_OBJECT_TYPE_NAMES,
    _HWP_SHAPE_ARC_SPECIFIC_SIZE,
    _HWP_SHAPE_ELLIPSE_SPECIFIC_SIZE,
    _HWP_SHAPE_LINE_SPECIFIC_SIZE,
    _HWP_SHAPE_RECTANGLE_SPECIFIC_SIZE,
    _HWP_TEXTART_PREFIX_PAYLOAD,
    _HWP_TEXTART_TRAILING_PAYLOAD,
    _SHAPE_NATIVE_METADATA_PREFIX,
    _build_arc_geometry_payload,
    _build_ellipse_geometry_payload,
    _build_line_specific_payload,
    _build_polygon_points_payload,
    _build_rectangle_specific_payload,
    _build_textart_payload_fields,
    _coerce_shape_point,
    _collect_control_native_metadata,
    _collect_shape_native_metadata,
    _control_metadata_node,
    _encode_length_prefixed_utf16,
    _graphic_layout_metadata_updates,
    _graphic_rotation_metadata_updates,
    _has_hancom_connectline_signature,
    _image_adjustment_metadata_updates,
    _metadata_bool_literal_is_valid,
    _metadata_bool_value,
    _metadata_json_dict_list,
    _metadata_json_list,
    _metadata_json_value,
    _metadata_prefixed_int_map,
    _metadata_prefixed_str_map,
    _parse_connectline_payload_fields,
    _parse_control_native_metadata,
    _parse_length_prefixed_utf16,
    _parse_object_description,
    _parse_raw_payload_fields,
    _parse_shape_fill_payload,
    _parse_shape_native_metadata,
    _parse_textart_payload_fields,
    _parse_textart_payload_text,
    _replace_object_description,
    _set_payload_size,
    _shape_metadata_node,
    _shape_specific_payload_size,
    _update_control_native_metadata,
    _update_shape_native_metadata,
)
from .hwp_binary import (
    _TABLE_CELL_CHAR_SHAPE,
    _TABLE_CELL_LINE_SEG,
    _TABLE_CELL_PARA_HEADER_TAIL,
    DEFAULT_HWP_TABLE_CELL_HEIGHT,
    DEFAULT_HWP_TABLE_CELL_WIDTH,
    TAG_LIST_HEADER,
    TAG_PARA_CHAR_SHAPE,
    TAG_PARA_HEADER,
    TAG_PARA_LINE_SEG,
    TAG_TABLE,
    ParagraphHeaderRecord,
    ParagraphTextRecord,
    RecordNode,
    _collect_text_from_node,
    SectionModel,
    _build_table_cell_list_header_payload,
    _build_table_record_payload,
    _TableCellSpec,
)

if TYPE_CHECKING:
    pass

from ._hwp_objects import HwpControlObject


@dataclass(frozen=True)
class HwpTableCellObject:
    row: int
    column: int
    col_span: int
    row_span: int
    width: int
    height: int
    border_fill_id: int
    margins: tuple[int, int, int, int]
    text: str


class HwpTableObject(HwpControlObject):
    @property
    def row_count(self) -> int:
        payload = self._table_record().payload
        return int.from_bytes(payload[4:6], "little")

    @property
    def column_count(self) -> int:
        payload = self._table_record().payload
        return int.from_bytes(payload[6:8], "little")

    @property
    def row_heights(self) -> list[int]:
        payload = self._table_record().payload
        row_count = self.row_count
        start = 18
        return [int.from_bytes(payload[start + index * 2 : start + index * 2 + 2], "little") for index in range(row_count)]

    @property
    def table_border_fill_id(self) -> int:
        payload = self._table_record().payload
        offset = 18 + self.row_count * 2
        return int.from_bytes(payload[offset : offset + 2], "little")

    @property
    def column_widths(self) -> list[int]:
        return self._column_widths()

    @property
    def cell_spacing(self) -> int:
        payload = self._table_record().payload
        return int.from_bytes(payload[8:10], "little", signed=False)

    @property
    def table_margins(self) -> tuple[int, int, int, int]:
        payload = self._table_record().payload
        return (
            int.from_bytes(payload[10:12], "little", signed=False),
            int.from_bytes(payload[12:14], "little", signed=False),
            int.from_bytes(payload[14:16], "little", signed=False),
            int.from_bytes(payload[16:18], "little", signed=False),
        )

    def cells(self) -> list[HwpTableCellObject]:
        result: list[HwpTableCellObject] = []
        children = self.control_node.children
        for index, node in enumerate(children):
            if node.tag_id != TAG_LIST_HEADER or len(node.payload) != 47:
                continue
            payload = node.payload
            text_node = children[index + 1] if index + 1 < len(children) and children[index + 1].tag_id == TAG_PARA_HEADER else None
            result.append(
                HwpTableCellObject(
                    row=int.from_bytes(payload[10:12], "little"),
                    column=int.from_bytes(payload[8:10], "little"),
                    col_span=int.from_bytes(payload[12:14], "little"),
                    row_span=int.from_bytes(payload[14:16], "little"),
                    width=int.from_bytes(payload[16:20], "little"),
                    height=int.from_bytes(payload[20:24], "little"),
                    margins=(
                        int.from_bytes(payload[24:26], "little"),
                        int.from_bytes(payload[26:28], "little"),
                        int.from_bytes(payload[28:30], "little"),
                        int.from_bytes(payload[30:32], "little"),
                    ),
                    border_fill_id=int.from_bytes(payload[32:34], "little"),
                    text=_collect_text_from_node(text_node) if text_node is not None else "",
                )
            )
        result.sort(key=lambda cell: (cell.row, cell.column))
        return result

    def cell(self, row: int, column: int) -> HwpTableCellObject:
        for cell in self.cells():
            if cell.row == row and cell.column == column:
                return cell
        raise IndexError(f"No HWP table cell at row={row}, column={column}.")

    def cell_text_matrix(self) -> list[list[str]]:
        matrix = [["" for _ in range(self.column_count)] for _ in range(self.row_count)]
        for cell in self.cells():
            if cell.row < self.row_count and cell.column < self.column_count:
                matrix[cell.row][cell.column] = cell.text
        return matrix

    def set_cell_text(self, row: int, column: int, text: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        children = control_node.children
        for index, node in enumerate(children):
            if node.tag_id != TAG_LIST_HEADER or len(node.payload) != 47:
                continue
            payload = node.payload
            current_column = int.from_bytes(payload[8:10], "little")
            current_row = int.from_bytes(payload[10:12], "little")
            if current_row != row or current_column != column:
                continue
            para_node = children[index + 1] if index + 1 < len(children) and children[index + 1].tag_id == TAG_PARA_HEADER else None
            if not isinstance(para_node, ParagraphHeaderRecord):
                raise ValueError("HWP table cell does not contain a writable paragraph header.")
            text_record = next((child for child in para_node.children if isinstance(child, ParagraphTextRecord)), None)
            if text_record is None:
                raise ValueError("HWP table cell does not contain a para text record.")
            text_record.set_raw_text(f"{text}\r")
            para_node.char_count = len(text) + 1
            self._commit(section_model)
            return
        raise IndexError(f"No HWP table cell at row={row}, column={column}.")

    def set_row_heights(self, heights: Sequence[int]) -> None:
        section_model, _paragraph, control_node = self._live_context()
        values = [int(value) for value in heights]
        if len(values) != self.row_count:
            raise ValueError(f"row heights length must match row_count={self.row_count}.")
        record = self._table_record(control_node)
        payload = bytearray(record.payload)
        start = 18
        for index, value in enumerate(values):
            payload[start + index * 2 : start + index * 2 + 2] = value.to_bytes(2, "little")
        record.payload = bytes(payload)
        self._commit(section_model)

    def set_table_border_fill_id(self, border_fill_id: int) -> None:
        section_model, _paragraph, control_node = self._live_context()
        record = self._table_record(control_node)
        payload = bytearray(record.payload)
        offset = 18 + self.row_count * 2
        payload[offset : offset + 2] = int(border_fill_id).to_bytes(2, "little")
        record.payload = bytes(payload)
        self._commit(section_model)

    def set_column_widths(self, widths: Sequence[int]) -> None:
        values = [int(value) for value in widths]
        if len(values) != self.column_count:
            raise ValueError(f"column widths length must match column_count={self.column_count}.")
        section_model, _paragraph, control_node = self._live_context()
        row_heights = list(self.row_heights)
        updated_specs = [
            _TableCellSpec(
                row=spec.row,
                col=spec.col,
                row_span=spec.row_span,
                col_span=spec.col_span,
                width=sum(values[spec.col : spec.col + spec.col_span]),
                height=sum(row_heights[spec.row : spec.row + spec.row_span]),
                text=spec.text,
                border_fill_id=spec.border_fill_id,
                margins=spec.margins,
            )
            for spec in self._cell_specs()
        ]
        self._replace_table_cells(
            section_model,
            control_node,
            row_count=self.row_count,
            column_count=self.column_count,
            row_heights=row_heights,
            cell_specs=updated_specs,
            table_border_fill_id=self.table_border_fill_id,
            cell_spacing=self.cell_spacing,
            table_margins=self.table_margins,
        )

    def set_table_margins(
        self,
        *,
        left: int | None = None,
        right: int | None = None,
        top: int | None = None,
        bottom: int | None = None,
    ) -> None:
        section_model, _paragraph, control_node = self._live_context()
        record = self._table_record(control_node)
        payload = bytearray(record.payload)
        current_left, current_right, current_top, current_bottom = self.table_margins
        payload[10:12] = int(current_left if left is None else left).to_bytes(2, "little", signed=False)
        payload[12:14] = int(current_right if right is None else right).to_bytes(2, "little", signed=False)
        payload[14:16] = int(current_top if top is None else top).to_bytes(2, "little", signed=False)
        payload[16:18] = int(current_bottom if bottom is None else bottom).to_bytes(2, "little", signed=False)
        record.payload = bytes(payload)
        self._commit(section_model)

    def set_cell_spacing(self, value: int) -> None:
        section_model, _paragraph, control_node = self._live_context()
        record = self._table_record(control_node)
        payload = bytearray(record.payload)
        payload[8:10] = int(value).to_bytes(2, "little", signed=False)
        record.payload = bytes(payload)
        self._commit(section_model)

    def set_cell_border_fill_id(self, row: int, column: int, border_fill_id: int) -> None:
        section_model, _paragraph, control_node = self._live_context()
        row_heights = list(self.row_heights)
        updated_specs: list[_TableCellSpec] = []
        found = False
        for spec in self._cell_specs():
            if spec.row == row and spec.col == column:
                updated_specs.append(
                    _TableCellSpec(
                        row=spec.row,
                        col=spec.col,
                        row_span=spec.row_span,
                        col_span=spec.col_span,
                        width=spec.width,
                        height=spec.height,
                        text=spec.text,
                        border_fill_id=int(border_fill_id),
                        margins=spec.margins,
                    )
                )
                found = True
            else:
                updated_specs.append(spec)
        if not found:
            raise IndexError(f"No HWP table cell at row={row}, column={column}.")
        self._replace_table_cells(
            section_model,
            control_node,
            row_count=self.row_count,
            column_count=self.column_count,
            row_heights=row_heights,
            cell_specs=updated_specs,
            table_border_fill_id=self.table_border_fill_id,
            cell_spacing=self.cell_spacing,
            table_margins=self.table_margins,
        )

    def set_cell_margins(
        self,
        row: int,
        column: int,
        *,
        left: int | None = None,
        right: int | None = None,
        top: int | None = None,
        bottom: int | None = None,
    ) -> None:
        section_model, _paragraph, control_node = self._live_context()
        row_heights = list(self.row_heights)
        updated_specs: list[_TableCellSpec] = []
        found = False
        for spec in self._cell_specs():
            if spec.row == row and spec.col == column:
                current_left, current_right, current_top, current_bottom = spec.margins
                updated_specs.append(
                    _TableCellSpec(
                        row=spec.row,
                        col=spec.col,
                        row_span=spec.row_span,
                        col_span=spec.col_span,
                        width=spec.width,
                        height=spec.height,
                        text=spec.text,
                        border_fill_id=spec.border_fill_id,
                        margins=(
                            int(current_left if left is None else left),
                            int(current_right if right is None else right),
                            int(current_top if top is None else top),
                            int(current_bottom if bottom is None else bottom),
                        ),
                    )
                )
                found = True
            else:
                updated_specs.append(spec)
        if not found:
            raise IndexError(f"No HWP table cell at row={row}, column={column}.")
        self._replace_table_cells(
            section_model,
            control_node,
            row_count=self.row_count,
            column_count=self.column_count,
            row_heights=row_heights,
            cell_specs=updated_specs,
            table_border_fill_id=self.table_border_fill_id,
            cell_spacing=self.cell_spacing,
            table_margins=self.table_margins,
        )

    def append_row(self) -> list[HwpTableCellObject]:
        section_model, _paragraph, control_node = self._live_context()
        row_heights = list(self.row_heights)
        new_row_index = self.row_count
        row_heights.append(row_heights[-1] if row_heights else DEFAULT_HWP_TABLE_CELL_HEIGHT)
        column_widths = self._column_widths()
        current_specs = self._cell_specs()
        for column_index, column_width in enumerate(column_widths):
            current_specs.append(
                _TableCellSpec(
                    row=new_row_index,
                    col=column_index,
                    row_span=1,
                    col_span=1,
                    width=column_width,
                    height=row_heights[-1],
                    text="",
                    border_fill_id=self.table_border_fill_id,
                    margins=self._default_cell_margins(),
                )
            )
        self._replace_table_cells(
            section_model,
            control_node,
            row_count=self.row_count + 1,
            column_count=self.column_count,
            row_heights=row_heights,
            cell_specs=current_specs,
            table_border_fill_id=self.table_border_fill_id,
            cell_spacing=self.cell_spacing,
            table_margins=self.table_margins,
        )
        refreshed = self.document.section(self.section_index).tables()[self._table_index_in_section()]
        new_row_index = refreshed.row_count - 1
        return [cell for cell in refreshed.cells() if cell.row == new_row_index]

    def merge_cells(self, start_row: int, start_column: int, end_row: int, end_column: int) -> HwpTableCellObject:
        if not (0 <= start_row <= end_row < self.row_count):
            raise IndexError("row merge range is out of bounds.")
        if not (0 <= start_column <= end_column < self.column_count):
            raise IndexError("column merge range is out of bounds.")

        section_model, _paragraph, control_node = self._live_context()
        row_heights = list(self.row_heights)
        column_widths = self._column_widths()
        merged_specs: list[_TableCellSpec] = []
        anchor: _TableCellSpec | None = None

        for spec in self._cell_specs():
            within_range = (
                start_row <= spec.row <= end_row
                and start_column <= spec.col <= end_column
            )
            if not within_range:
                merged_specs.append(spec)
                continue
            if spec.row == start_row and spec.col == start_column:
                anchor = _TableCellSpec(
                    row=start_row,
                    col=start_column,
                    row_span=end_row - start_row + 1,
                    col_span=end_column - start_column + 1,
                    width=sum(column_widths[start_column : end_column + 1]),
                    height=sum(row_heights[start_row : end_row + 1]),
                    text=spec.text,
                    border_fill_id=spec.border_fill_id,
                    margins=spec.margins,
                )
            elif spec.row + spec.row_span - 1 > end_row or spec.col + spec.col_span - 1 > end_column:
                raise ValueError("merge range intersects an existing spanned cell.")

        if anchor is None:
            raise IndexError("No anchor cell exists at the merge start coordinate.")

        merged_specs.append(anchor)
        merged_specs.sort(key=lambda item: (item.row, item.col))
        self._replace_table_cells(
            section_model,
            control_node,
            row_count=self.row_count,
            column_count=self.column_count,
            row_heights=row_heights,
            cell_specs=merged_specs,
            table_border_fill_id=self.table_border_fill_id,
            cell_spacing=self.cell_spacing,
            table_margins=self.table_margins,
        )
        refreshed = self.document.section(self.section_index).tables()[self._table_index_in_section()]
        return refreshed.cell(start_row, start_column)

    def _table_index_in_section(self) -> int:
        section = self.document.section(self.section_index)
        tables = section.tables()
        matches = [
            table
            for table in tables
            if table.paragraph_index == self.paragraph_index and table._control_ordinal == self._control_ordinal
        ]
        if matches:
            return tables.index(matches[0])
        for index, table in enumerate(tables):
            if table.paragraph_index == self.paragraph_index:
                return index
        raise ValueError("Could not locate the current HWP table within its section.")

    def _table_record(self, control_node: RecordNode | None = None) -> RecordNode:
        target = self.control_node if control_node is None else control_node
        for node in target.iter_descendants():
            if node.tag_id == TAG_TABLE:
                return node
        raise ValueError("HWP table control does not contain a table record.")

    def _cell_specs(self) -> list[_TableCellSpec]:
        return [
            _TableCellSpec(
                row=cell.row,
                col=cell.column,
                row_span=cell.row_span,
                col_span=cell.col_span,
                width=cell.width,
                height=cell.height,
                text=cell.text,
                border_fill_id=cell.border_fill_id,
                margins=cell.margins,
            )
            for cell in self.cells()
        ]

    def _column_widths(self) -> list[int]:
        widths = [0] * self.column_count
        for cell in self.cells():
            if cell.col_span == 1 and widths[cell.column] == 0:
                widths[cell.column] = cell.width
        fallback = max((value for value in widths if value > 0), default=DEFAULT_HWP_TABLE_CELL_WIDTH)
        for cell in self.cells():
            if cell.col_span <= 1:
                continue
            missing = [index for index in range(cell.column, cell.column + cell.col_span) if widths[index] == 0]
            if not missing:
                continue
            distributed = max(cell.width // cell.col_span, 1)
            for index in missing:
                widths[index] = distributed
        return [value or fallback for value in widths]

    def _default_cell_margins(self) -> tuple[int, int, int, int]:
        existing_cells = self.cells()
        if existing_cells:
            return existing_cells[0].margins
        return (510, 510, 141, 141)

    def _replace_table_cells(
        self,
        section_model: SectionModel,
        control_node: RecordNode,
        *,
        row_count: int,
        column_count: int,
        row_heights: Sequence[int],
        cell_specs: Sequence[_TableCellSpec],
        table_border_fill_id: int,
        cell_spacing: int,
        table_margins: tuple[int, int, int, int],
    ) -> None:
        table_record = self._table_record(control_node)
        margin_left, margin_right, margin_top, margin_bottom = table_margins
        table_record.payload = _build_table_record_payload(
            row_count,
            column_count,
            cell_spacing=cell_spacing,
            margin_left=margin_left,
            margin_right=margin_right,
            margin_top=margin_top,
            margin_bottom=margin_bottom,
            default_row_height=1,
            border_fill_id=table_border_fill_id,
        )
        payload = bytearray(table_record.payload)
        row_sizes_offset = 18
        for row_index, row_height in enumerate(row_heights):
            offset = row_sizes_offset + row_index * 2
            payload[offset : offset + 2] = int(row_height).to_bytes(2, "little", signed=False)
        table_record.payload = bytes(payload)

        new_children: list[RecordNode] = [table_record]
        for cell in sorted(cell_specs, key=lambda item: (item.row, item.col)):
            list_header = RecordNode(tag_id=TAG_LIST_HEADER, level=2, payload=_build_table_cell_list_header_payload(cell))
            raw_text = f"{cell.text}\r"
            char_count = len(raw_text)
            paragraph_header = ParagraphHeaderRecord(
                level=2,
                char_count=0x80000000 | char_count,
                control_mask=int.from_bytes(_TABLE_CELL_PARA_HEADER_TAIL[0:4], "little"),
                para_shape_id=int.from_bytes(_TABLE_CELL_PARA_HEADER_TAIL[4:6], "little"),
                style_id=_TABLE_CELL_PARA_HEADER_TAIL[6],
                split_flags=_TABLE_CELL_PARA_HEADER_TAIL[7],
                trailing_payload=_TABLE_CELL_PARA_HEADER_TAIL[8:],
            )
            paragraph_header.add_child(
                ParagraphTextRecord(
                    level=3,
                    raw_text=raw_text,
                )
            )
            paragraph_header.add_child(
                RecordNode(
                    tag_id=TAG_PARA_CHAR_SHAPE,
                    level=3,
                    payload=_TABLE_CELL_CHAR_SHAPE,
                )
            )
            paragraph_header.add_child(
                RecordNode(
                    tag_id=TAG_PARA_LINE_SEG,
                    level=3,
                    payload=_TABLE_CELL_LINE_SEG,
                )
            )
            new_children.extend((list_header, paragraph_header))

        control_node.children = []
        for child in new_children:
            control_node.add_child(child)
        self._commit(section_model)

