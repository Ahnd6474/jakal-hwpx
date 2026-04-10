from __future__ import annotations

import struct
import zlib
from collections.abc import Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterator

import olefile

from ._cfb_writer import write_compound_file
from .exceptions import HwpBinaryEditError, InvalidHwpFileError


FILE_HEADER_SIGNATURE = b"HWP Document File"
FILE_HEADER_SIZE = 256

TAG_DOCUMENT_PROPERTIES = 0x010
TAG_ID_MAPPINGS = 0x011
TAG_BIN_DATA = 0x012
TAG_FACE_NAME = 0x013
TAG_BORDER_FILL = 0x014
TAG_CHAR_SHAPE = 0x015
TAG_TAB_DEF = 0x016
TAG_NUMBERING = 0x017
TAG_BULLET = 0x018
TAG_PARA_SHAPE = 0x019
TAG_STYLE = 0x01A
TAG_DOC_DATA = 0x01B
TAG_COMPATIBLE_DOCUMENT = 0x01E
TAG_LAYOUT_COMPATIBILITY = 0x01F
TAG_PARA_HEADER = 0x042
TAG_PARA_TEXT = 0x043
TAG_PARA_CHAR_SHAPE = 0x044
TAG_PARA_LINE_SEG = 0x045
TAG_PARA_RANGE_TAG = 0x046
TAG_CTRL_HEADER = 0x047
TAG_LIST_HEADER = 0x048
TAG_PAGE_DEF = 0x049
TAG_FOOTNOTE_SHAPE = 0x04A
TAG_PAGE_BORDER_FILL = 0x04B
TAG_SHAPE_COMPONENT = 0x04C
TAG_TABLE = 0x04D
TAG_SHAPE_COMPONENT_LINE = 0x04E
TAG_SHAPE_COMPONENT_RECTANGLE = 0x04F
TAG_SHAPE_COMPONENT_ELLIPSE = 0x050
TAG_SHAPE_COMPONENT_ARC = 0x051
TAG_SHAPE_COMPONENT_POLYGON = 0x052
TAG_SHAPE_COMPONENT_CURVE = 0x053
TAG_SHAPE_COMPONENT_OLE = 0x054
TAG_SHAPE_COMPONENT_PICTURE = 0x055
TAG_SHAPE_COMPONENT_CONTAINER = 0x056
TAG_CTRL_DATA = 0x057
TAG_EQEDIT = 0x058
TAG_SHAPE_COMPONENT_TEXTART = 0x05A
TAG_FORM_OBJECT = 0x05B
TAG_MEMO_SHAPE = 0x05C
TAG_MEMO_LIST = 0x05D
TAG_CHART_DATA = 0x05E
TAG_TRACK_CHANGE = 0x05F
TAG_TRACK_CHANGE_AUTHOR = 0x060

FLAG_COMPRESSED = 1 << 0
FLAG_PASSWORD_PROTECTED = 1 << 1
FLAG_DISTRIBUTABLE = 1 << 2


TAG_NAMES = {
    TAG_DOCUMENT_PROPERTIES: "document_properties",
    TAG_ID_MAPPINGS: "id_mappings",
    TAG_BIN_DATA: "bin_data",
    TAG_FACE_NAME: "face_name",
    TAG_BORDER_FILL: "border_fill",
    TAG_CHAR_SHAPE: "char_shape",
    TAG_TAB_DEF: "tab_def",
    TAG_NUMBERING: "numbering",
    TAG_BULLET: "bullet",
    TAG_PARA_SHAPE: "para_shape",
    TAG_STYLE: "style",
    TAG_DOC_DATA: "doc_data",
    TAG_COMPATIBLE_DOCUMENT: "compatible_document",
    TAG_LAYOUT_COMPATIBILITY: "layout_compatibility",
    TAG_PARA_HEADER: "para_header",
    TAG_PARA_TEXT: "para_text",
    TAG_PARA_CHAR_SHAPE: "para_char_shape",
    TAG_PARA_LINE_SEG: "para_line_seg",
    TAG_PARA_RANGE_TAG: "para_range_tag",
    TAG_CTRL_HEADER: "ctrl_header",
    TAG_LIST_HEADER: "list_header",
    TAG_PAGE_DEF: "page_def",
    TAG_FOOTNOTE_SHAPE: "footnote_shape",
    TAG_PAGE_BORDER_FILL: "page_border_fill",
    TAG_SHAPE_COMPONENT: "shape_component",
    TAG_TABLE: "table",
    TAG_SHAPE_COMPONENT_LINE: "shape_line",
    TAG_SHAPE_COMPONENT_RECTANGLE: "shape_rectangle",
    TAG_SHAPE_COMPONENT_ELLIPSE: "shape_ellipse",
    TAG_SHAPE_COMPONENT_ARC: "shape_arc",
    TAG_SHAPE_COMPONENT_POLYGON: "shape_polygon",
    TAG_SHAPE_COMPONENT_CURVE: "shape_curve",
    TAG_SHAPE_COMPONENT_OLE: "shape_ole",
    TAG_SHAPE_COMPONENT_PICTURE: "shape_picture",
    TAG_SHAPE_COMPONENT_CONTAINER: "shape_container",
    TAG_CTRL_DATA: "ctrl_data",
    TAG_EQEDIT: "equation",
    TAG_SHAPE_COMPONENT_TEXTART: "shape_textart",
    TAG_FORM_OBJECT: "form_object",
    TAG_MEMO_SHAPE: "memo_shape",
    TAG_MEMO_LIST: "memo_list",
    TAG_CHART_DATA: "chart_data",
    TAG_TRACK_CHANGE: "track_change",
    TAG_TRACK_CHANGE_AUTHOR: "track_change_author",
}


DEFAULT_HWP_HYPERLINK_URL = "http://www.chungsa.go.kr/chungsa/cms/1/3/4.html"
DEFAULT_HWP_HYPERLINK_TEXT = "http://www.chungsa.go.kr/chungsa/cms/1/3/4.html"
DEFAULT_HWP_HYPERLINK_METADATA_FIELDS = ("1", "0", "0")

DEFAULT_HWP_TABLE_ROWS = 1
DEFAULT_HWP_TABLE_COLS = 1
DEFAULT_HWP_TABLE_CELL_WIDTH = 47341
DEFAULT_HWP_TABLE_CELL_HEIGHT = 282
DEFAULT_HWP_TABLE_BORDER_FILL_ID = 1
DEFAULT_HWP_TABLE_CELL_MARGIN = (510, 510, 141, 141)
DEFAULT_HWP_EQUATION_SCRIPT = "x+y"
DEFAULT_HWP_SHAPE_TEXT = ""
DEFAULT_HWP_SHAPE_WIDTH = 12000
DEFAULT_HWP_SHAPE_HEIGHT = 3200

_TABLE_TOP_PARA_HEADER_TAIL = bytes.fromhex("0008000019000000010000000100000000800000")
_TABLE_TOP_PARA_TEXT = bytes.fromhex("0b00206c627400000000000000000b000d00")
_TABLE_TOP_CHAR_SHAPE = bytes.fromhex("0000000004000000")
_TABLE_TOP_LINE_SEG = bytes.fromhex("00000000b87600008c0400008c040000dd030000f401000038ffffff04bd000000000600")
_TABLE_CTRL_HEADER = bytes.fromhex(
    "206c627411232a08000000000000000059bd000072030000020000008d008d008d008d002b945768000000000000"
)
_TABLE_CELL_PARA_HEADER_TAIL = bytes.fromhex("000000000b000000010000000100000000800000")
_TABLE_CELL_CHAR_SHAPE = bytes.fromhex("000000002f000000")
_TABLE_CELL_LINE_SEG = bytes.fromhex("00000000000000005802000058020000fe010000e0ffffff000000005cb9000000000600")

_HYPERLINK_PARA_HEADER_TAIL = bytes.fromhex("180000000000000003000000010000000080")
_HYPERLINK_CHAR_SHAPE = bytes.fromhex("000000000600000013000000070000004200000006000000")
_HYPERLINK_LINE_SEG = bytes.fromhex("0000000027e90000b0040000b0040000fc030000d00200000000000018a6000000000600")
_HYPERLINK_TEXT_START_MARKER = bytes.fromhex("03006b6c682500000000000000000300")
_HYPERLINK_TEXT_END_MARKER = bytes.fromhex("04006b6c680000000000000000000400")
_HYPERLINK_CTRL_PREFIX = bytes.fromhex("6b6c68250008000000")
_HYPERLINK_CTRL_SUFFIX = bytes.fromhex("93caed41")

_PICTURE_TOP_PARA_HEADER_TAIL = bytes.fromhex("000800000e00000001000000010000000080")
_PICTURE_TOP_PARA_TEXT = bytes.fromhex("0b00206f736700000000000000000b000d00")
_PICTURE_TOP_CHAR_SHAPE = bytes.fromhex("0000000006000000")
_PICTURE_TOP_LINE_SEG = bytes.fromhex("00000000381800009c8100009c8100002b6e00000c030000000000002cb1000000000600")
_PICTURE_CTRL_HEADER = bytes.fromhex(
    "206f736711232a040000000000000000b8b000009c8100000000000000000000000000009c52be4200000000"
)
_PICTURE_SHAPE_COMPONENT = bytes.fromhex(
    "6369702463697024000000000000000000000100b8b000009c810000b8b000009c8100000000082400005c580000ce4000000100000000000000"
    "f03f000000000000000000000000000000000000000000000000000000000000f03f0000000000000000000000000000f03f00000000000000000000"
    "000000000000000000000000000000000000f03f0000000000000000000000000000f03f000000000000000000000000000000000000000000000000"
    "0000000000f03f0000000000000000"
)
_PICTURE_SHAPE_PICTURE = bytes.fromhex(
    "0000000000000000010000c00000000000000000b8b0000000000000b8b000009c810000000000009c81000000000000000000007cb0000024810000"
    "00000000000000000000000100"
)


def _build_control_id_payload(control_id: str) -> bytes:
    normalized = control_id[:4].ljust(4)
    return normalized.encode("latin1", errors="replace")[::-1]


def _normalize_field_control_id(field_type: str) -> str:
    stripped = field_type.strip()
    if stripped.startswith("%"):
        return stripped[:4].ljust(4)
    alnum = "".join(character.lower() for character in stripped if character.isalnum())
    if not alnum:
        return "%fld"
    return ("%" + alnum[:3]).ljust(4)


def _build_equation_payload(script: str) -> bytes:
    return b"\x00\x00\x00\x00" + f"* {script}".encode("utf-16-le")


def _build_minimal_shape_component_payload(width: int, height: int) -> bytes:
    payload = bytearray(28)
    payload[20:24] = int(width).to_bytes(4, "little", signed=False)
    payload[24:28] = int(height).to_bytes(4, "little", signed=False)
    return bytes(payload)


def _normalize_stream_path(path: str | tuple[str, ...] | list[str]) -> str:
    if isinstance(path, str):
        return path.strip("/").replace("\\", "/")
    return "/".join(path)


def _stream_components(path: str) -> list[str]:
    normalized = _normalize_stream_path(path)
    return [part for part in normalized.split("/") if part]


def _replace_with_limit(source: str, old: str, new: str, count: int) -> tuple[str, int]:
    if count < 0:
        return source.replace(old, new), source.count(old)

    replaced = 0
    result = source
    while replaced < count and old in result:
        result = result.replace(old, new, 1)
        replaced += 1
    return result, replaced


@dataclass(frozen=True)
class HwpBinaryFileHeader:
    signature: str
    version_bytes: tuple[int, int, int, int]
    flags: int

    @classmethod
    def from_bytes(cls, data: bytes) -> "HwpBinaryFileHeader":
        if len(data) < 40:
            raise InvalidHwpFileError("HWP FileHeader stream is shorter than the minimum header size.")
        signature = data[:32].split(b"\x00", 1)[0].decode("ascii", errors="replace")
        return cls(
            signature=signature,
            version_bytes=tuple(data[32:36]),
            flags=int.from_bytes(data[36:40], "little"),
        )

    @property
    def version(self) -> str:
        return ".".join(str(part) for part in reversed(self.version_bytes))

    @property
    def compressed(self) -> bool:
        return bool(self.flags & FLAG_COMPRESSED)

    @property
    def password_protected(self) -> bool:
        return bool(self.flags & FLAG_PASSWORD_PROTECTED)

    @property
    def distributable(self) -> bool:
        return bool(self.flags & FLAG_DISTRIBUTABLE)


@dataclass(frozen=True)
class HwpRecord:
    tag_id: int
    level: int
    size: int
    header_size: int
    offset: int
    payload: bytes

    def to_bytes(self) -> bytes:
        if self.header_size == 8 or self.size >= 0xFFF:
            header = (self.tag_id & 0x3FF) | ((self.level & 0x3FF) << 10) | (0xFFF << 20)
            return header.to_bytes(4, "little") + self.size.to_bytes(4, "little") + self.payload
        header = (self.tag_id & 0x3FF) | ((self.level & 0x3FF) << 10) | ((self.size & 0xFFF) << 20)
        return header.to_bytes(4, "little") + self.payload


@dataclass(frozen=True)
class HwpDocumentProperties:
    section_count: int
    page_start_number: int
    footnote_start_number: int
    endnote_start_number: int
    picture_start_number: int
    table_start_number: int
    equation_start_number: int
    list_id: int
    paragraph_id: int
    character_unit_position: int


@dataclass(frozen=True)
class HwpParagraph:
    index: int
    section_index: int
    char_count: int
    control_mask: int
    para_shape_id: int
    style_id: int
    split_flags: int
    raw_text: str
    text: str
    text_record_offset: int
    text_record_size: int


@dataclass(frozen=True)
class HwpStreamCapacity:
    path: str
    compressed: bool
    original_size: int
    current_size: int
    free_bytes: int
    fits: bool


def hwp_tag_name(tag_id: int) -> str:
    return TAG_NAMES.get(tag_id, f"tag_{tag_id:03d}")


def _parse_document_properties_payload(payload: bytes) -> HwpDocumentProperties:
    if len(payload) < 26:
        raise InvalidHwpFileError("Document properties record is shorter than expected.")
    return HwpDocumentProperties(
        section_count=int.from_bytes(payload[0:2], "little"),
        page_start_number=int.from_bytes(payload[2:4], "little"),
        footnote_start_number=int.from_bytes(payload[4:6], "little"),
        endnote_start_number=int.from_bytes(payload[6:8], "little"),
        picture_start_number=int.from_bytes(payload[8:10], "little"),
        table_start_number=int.from_bytes(payload[10:12], "little"),
        equation_start_number=int.from_bytes(payload[12:14], "little"),
        list_id=int.from_bytes(payload[14:18], "little"),
        paragraph_id=int.from_bytes(payload[18:22], "little"),
        character_unit_position=int.from_bytes(payload[22:26], "little"),
    )


def _build_document_properties_payload(properties: HwpDocumentProperties) -> bytes:
    payload = bytearray(26)
    payload[0:2] = properties.section_count.to_bytes(2, "little")
    payload[2:4] = properties.page_start_number.to_bytes(2, "little")
    payload[4:6] = properties.footnote_start_number.to_bytes(2, "little")
    payload[6:8] = properties.endnote_start_number.to_bytes(2, "little")
    payload[8:10] = properties.picture_start_number.to_bytes(2, "little")
    payload[10:12] = properties.table_start_number.to_bytes(2, "little")
    payload[12:14] = properties.equation_start_number.to_bytes(2, "little")
    payload[14:18] = properties.list_id.to_bytes(4, "little")
    payload[18:22] = properties.paragraph_id.to_bytes(4, "little")
    payload[22:26] = properties.character_unit_position.to_bytes(4, "little")
    return bytes(payload)


def _build_hyperlink_para_text_payload(display_text: str) -> bytes:
    return _HYPERLINK_TEXT_START_MARKER + display_text.encode("utf-16-le") + _HYPERLINK_TEXT_END_MARKER + "\r".encode("utf-16-le")


def _build_hyperlink_command(url: str, metadata_fields: Sequence[str | int] | None = None) -> str:
    fields = DEFAULT_HWP_HYPERLINK_METADATA_FIELDS if metadata_fields is None else tuple(str(value) for value in metadata_fields)
    return url.replace(":", "\\:") + ";" + ";".join(fields) + ";"


def _build_hyperlink_ctrl_payload(url: str, metadata_fields: Sequence[str | int] | None = None) -> bytes:
    normalized = _build_hyperlink_command(url, metadata_fields=metadata_fields)
    return _HYPERLINK_CTRL_PREFIX + len(normalized).to_bytes(2, "little") + normalized.encode("utf-16-le") + _HYPERLINK_CTRL_SUFFIX


def _normalize_table_cell_texts(
    rows: int,
    cols: int,
    cell_text: str | None,
    cell_texts: Sequence[str] | Sequence[Sequence[str]] | None,
) -> list[list[str]]:
    if rows < 1 or cols < 1:
        raise ValueError("rows and cols must both be at least 1.")

    if cell_texts is None:
        matrix = [["" for _ in range(cols)] for _ in range(rows)]
        if cell_text is not None:
            matrix[0][0] = cell_text
        return matrix

    values = list(cell_texts)
    if values and isinstance(values[0], (list, tuple)):
        if len(values) != rows:
            raise ValueError(f"cell_texts row count must match rows={rows}.")
        matrix: list[list[str]] = []
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
                    margins=tuple(int(value) for value in margins),
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


@dataclass
class RecordNode:
    tag_id: int
    level: int
    payload: bytes
    header_size: int = 4
    offset: int = -1
    children: list["RecordNode"] = field(default_factory=list)

    @property
    def tag_name(self) -> str:
        return hwp_tag_name(self.tag_id)

    @property
    def size(self) -> int:
        return len(self.payload)

    def to_record(self) -> HwpRecord:
        return HwpRecord(
            tag_id=self.tag_id,
            level=self.level,
            size=self.size,
            header_size=self.header_size,
            offset=self.offset,
            payload=self.payload,
        )

    def add_child(self, child: "RecordNode") -> None:
        child._normalize_level(self.level + 1)
        self.children.append(child)

    def _normalize_level(self, level: int) -> None:
        delta = level - self.level
        self.level = level
        for child in self.children:
            child._normalize_level(child.level + delta)

    def iter_preorder(self) -> Iterator["RecordNode"]:
        yield self
        for child in self.children:
            yield from child.iter_preorder()

    def iter_descendants(self) -> Iterator["RecordNode"]:
        for child in self.children:
            yield child
            yield from child.iter_descendants()

    def to_records(self) -> list[HwpRecord]:
        return [node.to_record() for node in self.iter_preorder()]

    def clone(self) -> "RecordNode":
        node = record_node_from_record(self.to_record())
        node.children = [child.clone() for child in self.children]
        return node


class UnknownRecord(RecordNode):
    pass


class TypedRecord(RecordNode):
    pass


class DocumentPropertiesRecord(TypedRecord):
    def __init__(
        self,
        *,
        level: int,
        properties: HwpDocumentProperties,
        header_size: int = 4,
        offset: int = -1,
        children: list[RecordNode] | None = None,
    ) -> None:
        self.properties = properties
        super().__init__(
            tag_id=TAG_DOCUMENT_PROPERTIES,
            level=level,
            payload=_build_document_properties_payload(properties),
            header_size=header_size,
            offset=offset,
            children=list(children or []),
        )

    @classmethod
    def from_record(cls, record: HwpRecord) -> "DocumentPropertiesRecord":
        return cls(
            level=record.level,
            properties=_parse_document_properties_payload(record.payload),
            header_size=record.header_size,
            offset=record.offset,
        )

    def sync_payload(self) -> None:
        self.payload = _build_document_properties_payload(self.properties)

    def to_record(self) -> HwpRecord:
        self.sync_payload()
        return super().to_record()

    def to_properties(self) -> HwpDocumentProperties:
        return self.properties


ID_MAPPINGS_INDEX_NAMES = [
    "bin_data",
    "ko_fonts",
    "en_fonts",
    "cn_fonts",
    "jp_fonts",
    "other_fonts",
    "symbol_fonts",
    "user_fonts",
    "border_fills",
    "char_shapes",
    "tab_defs",
    "numberings",
    "bullets",
    "para_shapes",
    "styles",
    "memo_shapes",
    "track_changes",
    "track_change_authors",
]


class IdMappingsRecord(TypedRecord):
    def __init__(
        self,
        *,
        level: int,
        counts: list[int],
        trailing_payload: bytes = b"",
        header_size: int = 4,
        offset: int = -1,
        children: list[RecordNode] | None = None,
    ) -> None:
        self.counts = list(counts)
        self.trailing_payload = trailing_payload
        super().__init__(
            tag_id=TAG_ID_MAPPINGS,
            level=level,
            payload=self._build_payload(),
            header_size=header_size,
            offset=offset,
            children=list(children or []),
        )

    @classmethod
    def from_record(cls, record: HwpRecord) -> "IdMappingsRecord":
        count_bytes = len(record.payload) // 4
        counts = [int.from_bytes(record.payload[index * 4 : (index + 1) * 4], "little") for index in range(count_bytes)]
        trailing_offset = count_bytes * 4
        return cls(
            level=record.level,
            counts=counts,
            trailing_payload=record.payload[trailing_offset:],
            header_size=record.header_size,
            offset=record.offset,
        )

    def _build_payload(self) -> bytes:
        payload = bytearray()
        for count in self.counts:
            payload.extend(int(count).to_bytes(4, "little"))
        payload.extend(self.trailing_payload)
        return bytes(payload)

    def sync_payload(self) -> None:
        self.payload = self._build_payload()

    def to_record(self) -> HwpRecord:
        self.sync_payload()
        return super().to_record()

    def get_count(self, index: int) -> int:
        return self.counts[index] if index < len(self.counts) else 0

    def set_count(self, index: int, value: int) -> None:
        while len(self.counts) <= index:
            self.counts.append(0)
        self.counts[index] = int(value)

    @property
    def bin_data_count(self) -> int:
        return self.get_count(0)

    @bin_data_count.setter
    def bin_data_count(self, value: int) -> None:
        self.set_count(0, value)

    def named_counts(self) -> dict[str, int]:
        result: dict[str, int] = {}
        for index, count in enumerate(self.counts):
            key = ID_MAPPINGS_INDEX_NAMES[index] if index < len(ID_MAPPINGS_INDEX_NAMES) else f"index_{index}"
            result[key] = count
        return result


class BinDataRecord(TypedRecord):
    def __init__(
        self,
        *,
        level: int,
        flags: int,
        storage_id: int | None = None,
        extension: str | None = None,
        abs_path: str | None = None,
        rel_path: str | None = None,
        header_size: int = 4,
        offset: int = -1,
        children: list[RecordNode] | None = None,
    ) -> None:
        self.flags = flags
        self.storage_id = storage_id
        self.extension = extension
        self.abs_path = abs_path
        self.rel_path = rel_path
        super().__init__(
            tag_id=TAG_BIN_DATA,
            level=level,
            payload=self._build_payload(),
            header_size=header_size,
            offset=offset,
            children=list(children or []),
        )

    @property
    def kind(self) -> int:
        return self.flags & 0x000F

    @classmethod
    def from_record(cls, record: HwpRecord) -> "BinDataRecord":
        payload = record.payload
        flags = int.from_bytes(payload[0:2], "little") if len(payload) >= 2 else 0
        kind = flags & 0x000F
        if kind == 0:
            abs_len = int.from_bytes(payload[2:4], "little") if len(payload) >= 4 else 0
            cursor = 4
            abs_path = payload[cursor : cursor + abs_len * 2].decode("utf-16-le", errors="ignore")
            cursor += abs_len * 2
            rel_len = int.from_bytes(payload[cursor : cursor + 2], "little") if len(payload) >= cursor + 2 else 0
            cursor += 2
            rel_path = payload[cursor : cursor + rel_len * 2].decode("utf-16-le", errors="ignore")
            return cls(
                level=record.level,
                flags=flags,
                abs_path=abs_path,
                rel_path=rel_path,
                header_size=record.header_size,
                offset=record.offset,
            )

        storage_id = int.from_bytes(payload[2:4], "little") if len(payload) >= 4 else None
        ext_len = int.from_bytes(payload[4:6], "little") if len(payload) >= 6 else 0
        extension = payload[6 : 6 + ext_len * 2].decode("utf-16-le", errors="ignore") if ext_len else None
        return cls(
            level=record.level,
            flags=flags,
            storage_id=storage_id,
            extension=extension,
            header_size=record.header_size,
            offset=record.offset,
        )

    @classmethod
    def embedded(
        cls,
        *,
        level: int,
        storage_id: int,
        extension: str,
        flags: int = 0x0001,
    ) -> "BinDataRecord":
        return cls(level=level, flags=flags, storage_id=storage_id, extension=extension)

    def _build_payload(self) -> bytes:
        payload = bytearray()
        payload.extend(int(self.flags).to_bytes(2, "little"))
        if self.kind == 0:
            abs_path = self.abs_path or ""
            rel_path = self.rel_path or ""
            payload.extend(len(abs_path).to_bytes(2, "little"))
            payload.extend(abs_path.encode("utf-16-le"))
            payload.extend(len(rel_path).to_bytes(2, "little"))
            payload.extend(rel_path.encode("utf-16-le"))
            return bytes(payload)

        payload.extend(int(self.storage_id or 0).to_bytes(2, "little"))
        extension = (self.extension or "").lstrip(".")
        payload.extend(len(extension).to_bytes(2, "little"))
        payload.extend(extension.encode("utf-16-le"))
        return bytes(payload)

    def sync_payload(self) -> None:
        self.payload = self._build_payload()

    def to_record(self) -> HwpRecord:
        self.sync_payload()
        return super().to_record()


class ParagraphHeaderRecord(TypedRecord):
    def __init__(
        self,
        *,
        level: int,
        char_count: int,
        control_mask: int,
        para_shape_id: int,
        style_id: int,
        split_flags: int,
        header_size: int = 4,
        offset: int = -1,
        children: list[RecordNode] | None = None,
        trailing_payload: bytes = b"",
    ) -> None:
        self.char_count = char_count
        self.control_mask = control_mask
        self.para_shape_id = para_shape_id
        self.style_id = style_id
        self.split_flags = split_flags
        self.trailing_payload = trailing_payload
        payload = self._build_payload()
        super().__init__(
            tag_id=TAG_PARA_HEADER,
            level=level,
            payload=payload,
            header_size=header_size,
            offset=offset,
            children=list(children or []),
        )

    @classmethod
    def from_record(cls, record: HwpRecord) -> "ParagraphHeaderRecord":
        payload = record.payload
        return cls(
            level=record.level,
            char_count=int.from_bytes(payload[0:4], "little") if len(payload) >= 4 else 0,
            control_mask=int.from_bytes(payload[4:8], "little") if len(payload) >= 8 else 0,
            para_shape_id=int.from_bytes(payload[8:10], "little") if len(payload) >= 10 else 0,
            style_id=payload[10] if len(payload) >= 11 else 0,
            split_flags=payload[11] if len(payload) >= 12 else 0,
            header_size=record.header_size,
            offset=record.offset,
            trailing_payload=payload[12:] if len(payload) > 12 else b"",
        )

    def _build_payload(self) -> bytes:
        payload = bytearray(12)
        payload[0:4] = self.char_count.to_bytes(4, "little")
        payload[4:8] = self.control_mask.to_bytes(4, "little")
        payload[8:10] = self.para_shape_id.to_bytes(2, "little")
        payload[10] = self.style_id & 0xFF
        payload[11] = self.split_flags & 0xFF
        payload.extend(self.trailing_payload)
        return bytes(payload)

    def sync_payload(self) -> None:
        self.payload = self._build_payload()

    def to_record(self) -> HwpRecord:
        self.sync_payload()
        return super().to_record()


class ParagraphTextRecord(TypedRecord):
    def __init__(
        self,
        *,
        level: int,
        raw_text: str,
        header_size: int = 4,
        offset: int = -1,
        children: list[RecordNode] | None = None,
    ) -> None:
        self.raw_text = raw_text
        super().__init__(
            tag_id=TAG_PARA_TEXT,
            level=level,
            payload=raw_text.encode("utf-16-le"),
            header_size=header_size,
            offset=offset,
            children=list(children or []),
        )

    @classmethod
    def from_record(cls, record: HwpRecord) -> "ParagraphTextRecord":
        return cls(
            level=record.level,
            raw_text=record.payload.decode("utf-16-le", errors="ignore"),
            header_size=record.header_size,
            offset=record.offset,
        )

    @property
    def text(self) -> str:
        return _strip_para_text_controls(self.raw_text)

    def set_raw_text(self, value: str) -> None:
        self.raw_text = value
        self.payload = value.encode("utf-16-le")

    def set_text_same_length(self, value: str) -> None:
        if len(value.encode("utf-16-le")) != len(self.raw_text.encode("utf-16-le")):
            raise HwpBinaryEditError(
                "ParagraphTextRecord.set_text_same_length requires the new text to match the record's UTF-16 byte length."
            )
        self.set_raw_text(value)

    def replace_text_same_length(self, old: str, new: str, *, count: int = -1) -> int:
        if not old:
            raise ValueError("old must be a non-empty string.")
        if len(old.encode("utf-16-le")) != len(new.encode("utf-16-le")):
            raise HwpBinaryEditError(
                "ParagraphTextRecord.replace_text_same_length requires old and new to have the same UTF-16 byte length."
            )
        updated, replaced = _replace_with_limit(self.raw_text, old, new, count)
        self.set_raw_text(updated)
        return replaced

    def to_record(self) -> HwpRecord:
        self.payload = self.raw_text.encode("utf-16-le")
        return super().to_record()


def record_node_from_record(record: HwpRecord) -> RecordNode:
    if record.tag_id == TAG_DOCUMENT_PROPERTIES:
        return DocumentPropertiesRecord.from_record(record)
    if record.tag_id == TAG_ID_MAPPINGS:
        return IdMappingsRecord.from_record(record)
    if record.tag_id == TAG_BIN_DATA:
        return BinDataRecord.from_record(record)
    if record.tag_id == TAG_PARA_HEADER:
        return ParagraphHeaderRecord.from_record(record)
    if record.tag_id == TAG_PARA_TEXT:
        return ParagraphTextRecord.from_record(record)
    if record.tag_id in TAG_NAMES:
        return TypedRecord(
            tag_id=record.tag_id,
            level=record.level,
            payload=record.payload,
            header_size=record.header_size,
            offset=record.offset,
        )
    return UnknownRecord(
        tag_id=record.tag_id,
        level=record.level,
        payload=record.payload,
        header_size=record.header_size,
        offset=record.offset,
    )


def build_record_tree(records: list[HwpRecord]) -> list[RecordNode]:
    roots: list[RecordNode] = []
    stack: list[RecordNode] = []
    for record in records:
        node = record_node_from_record(record)
        while stack and stack[-1].level >= node.level:
            stack.pop()
        if stack:
            stack[-1].children.append(node)
        else:
            roots.append(node)
        stack.append(node)
    return roots


def flatten_record_tree(roots: list[RecordNode]) -> list[HwpRecord]:
    records: list[HwpRecord] = []
    for root in roots:
        records.extend(root.to_records())
    return records


@dataclass
class DocInfoModel:
    roots: list[RecordNode]

    @classmethod
    def from_records(cls, records: list[HwpRecord]) -> "DocInfoModel":
        return cls(roots=build_record_tree(records))

    def document_properties_record(self) -> DocumentPropertiesRecord:
        for root in self.roots:
            if isinstance(root, DocumentPropertiesRecord):
                return root
        raise InvalidHwpFileError("DocInfo does not start with a document properties record.")

    def document_properties(self) -> HwpDocumentProperties:
        return self.document_properties_record().to_properties()

    def id_mappings_record(self) -> IdMappingsRecord:
        for root in self.roots:
            if isinstance(root, IdMappingsRecord):
                return root
        raise InvalidHwpFileError("DocInfo does not contain an id mappings record.")

    def bin_data_records(self) -> list[BinDataRecord]:
        return [node for node in self.iter_nodes() if isinstance(node, BinDataRecord)]

    def face_name_records(self) -> list[RecordNode]:
        return self.records_by_tag_id(TAG_FACE_NAME)

    def border_fill_records(self) -> list[RecordNode]:
        return self.records_by_tag_id(TAG_BORDER_FILL)

    def char_shape_records(self) -> list[RecordNode]:
        return self.records_by_tag_id(TAG_CHAR_SHAPE)

    def para_shape_records(self) -> list[RecordNode]:
        return self.records_by_tag_id(TAG_PARA_SHAPE)

    def style_records(self) -> list[RecordNode]:
        return self.records_by_tag_id(TAG_STYLE)

    def next_bin_data_id(self) -> int:
        id_record = self.id_mappings_record()
        return id_record.bin_data_count + 1

    def add_embedded_bindata(self, extension: str, *, storage_id: int | None = None, flags: int = 0x0001) -> BinDataRecord:
        id_record = self.id_mappings_record()
        if storage_id is None:
            storage_id = self.next_bin_data_id()
        bindata = BinDataRecord.embedded(level=0, storage_id=storage_id, extension=extension, flags=flags)
        insert_at = 0
        for index, root in enumerate(self.roots):
            if isinstance(root, BinDataRecord):
                insert_at = index + 1
            elif isinstance(root, IdMappingsRecord):
                insert_at = index + 1
        self.roots.insert(insert_at, bindata)
        id_record.bin_data_count = max(id_record.bin_data_count, storage_id)
        return bindata

    def remove_bindata(self, storage_id: int) -> bool:
        for index, root in enumerate(self.roots):
            if isinstance(root, BinDataRecord) and root.storage_id == storage_id:
                del self.roots[index]
                id_record = self.id_mappings_record()
                remaining_ids = [record.storage_id or 0 for record in self.bin_data_records()]
                id_record.bin_data_count = max(remaining_ids, default=0)
                return True
        return False

    def tag_counts(self) -> dict[str, int]:
        counts: dict[str, int] = {}
        for node in self.iter_nodes():
            counts[node.tag_name] = counts.get(node.tag_name, 0) + 1
        return counts

    def iter_nodes(self) -> Iterator[RecordNode]:
        for root in self.roots:
            yield from root.iter_preorder()

    def records_by_tag_id(self, tag_id: int) -> list[RecordNode]:
        return [node for node in self.iter_nodes() if node.tag_id == tag_id]

    def to_records(self) -> list[HwpRecord]:
        return flatten_record_tree(self.roots)

    def to_bytes(self) -> bytes:
        return b"".join(record.to_bytes() for record in self.to_records())


@dataclass
class SectionParagraphModel:
    section_index: int
    index: int
    header: ParagraphHeaderRecord

    def text_record(self) -> ParagraphTextRecord | None:
        for node in self.header.iter_descendants():
            if isinstance(node, ParagraphTextRecord):
                return node
        return None

    @property
    def raw_text(self) -> str:
        text_record = self.text_record()
        return text_record.raw_text if text_record is not None else ""

    @property
    def text(self) -> str:
        text_record = self.text_record()
        return text_record.text if text_record is not None else ""

    def control_nodes(self) -> list[RecordNode]:
        return [node for node in self.header.iter_descendants() if node.tag_id == TAG_CTRL_HEADER]

    def descendant_nodes(self) -> list[RecordNode]:
        return list(self.header.iter_descendants())

    def set_text_same_length(self, value: str) -> None:
        text_record = self.text_record()
        if text_record is None:
            raise HwpBinaryEditError("Paragraph does not contain a para text record.")
        if text_record.raw_text != text_record.text:
            raise HwpBinaryEditError(
                "SectionParagraphModel.set_text_same_length is only supported for paragraphs without hidden control characters."
            )
        text_record.set_text_same_length(value)
        self.header.char_count = len(value)

    def replace_text_same_length(self, old: str, new: str, *, count: int = -1) -> int:
        text_record = self.text_record()
        if text_record is None:
            return 0
        replaced = text_record.replace_text_same_length(old, new, count=count)
        self.header.char_count = len(text_record.raw_text)
        return replaced


@dataclass
class SectionModel:
    section_index: int
    roots: list[RecordNode]

    @classmethod
    def from_records(cls, section_index: int, records: list[HwpRecord]) -> "SectionModel":
        return cls(section_index=section_index, roots=build_record_tree(records))

    def iter_nodes(self) -> Iterator[RecordNode]:
        for root in self.roots:
            yield from root.iter_preorder()

    def paragraphs(self) -> list[SectionParagraphModel]:
        paragraphs: list[SectionParagraphModel] = []
        paragraph_index = 0
        for root in self.iter_nodes():
            if isinstance(root, ParagraphHeaderRecord):
                paragraphs.append(
                    SectionParagraphModel(
                        section_index=self.section_index,
                        index=paragraph_index,
                        header=root,
                    )
                )
                paragraph_index += 1
        return paragraphs

    def control_nodes(self) -> list[RecordNode]:
        return [node for node in self.iter_nodes() if node.tag_id == TAG_CTRL_HEADER]

    def append_paragraph(
        self,
        text: str,
        *,
        para_shape_id: int = 0,
        style_id: int = 0,
        split_flags: int = 0,
        control_mask: int = 0,
    ) -> SectionParagraphModel:
        header = ParagraphHeaderRecord(
            level=0,
            char_count=len(text),
            control_mask=control_mask,
            para_shape_id=para_shape_id,
            style_id=style_id,
            split_flags=split_flags,
        )
        header.add_child(ParagraphTextRecord(level=1, raw_text=text))
        self.roots.append(header)
        return SectionParagraphModel(
            section_index=self.section_index,
            index=len(self.paragraphs()) - 1,
            header=header,
        )

    def insert_paragraph(
        self,
        paragraph_index: int,
        text: str,
        *,
        para_shape_id: int = 0,
        style_id: int = 0,
        split_flags: int = 0,
        control_mask: int = 0,
    ) -> SectionParagraphModel:
        current_paragraphs = self.paragraphs()
        if paragraph_index < 0 or paragraph_index > len(current_paragraphs):
            raise IndexError(f"paragraph_index {paragraph_index} is out of range for section {self.section_index}.")
        header = ParagraphHeaderRecord(
            level=0,
            char_count=len(text),
            control_mask=control_mask,
            para_shape_id=para_shape_id,
            style_id=style_id,
            split_flags=split_flags,
        )
        header.add_child(ParagraphTextRecord(level=1, raw_text=text))
        if paragraph_index == len(current_paragraphs):
            self.roots.append(header)
        else:
            self.roots.insert(self.roots.index(current_paragraphs[paragraph_index].header), header)
        return SectionParagraphModel(
            section_index=self.section_index,
            index=paragraph_index,
            header=header,
        )

    def to_records(self) -> list[HwpRecord]:
        return flatten_record_tree(self.roots)

    def to_bytes(self) -> bytes:
        return b"".join(record.to_bytes() for record in self.to_records())


def _iter_records(data: bytes) -> Iterator[HwpRecord]:
    offset = 0
    while offset + 4 <= len(data):
        header = int.from_bytes(data[offset : offset + 4], "little")
        tag_id = header & 0x3FF
        level = (header >> 10) & 0x3FF
        size = (header >> 20) & 0xFFF
        header_size = 4
        if size == 0xFFF:
            if offset + 8 > len(data):
                break
            size = int.from_bytes(data[offset + 4 : offset + 8], "little")
            header_size = 8
        payload_offset = offset + header_size
        payload_end = payload_offset + size
        if payload_end > len(data):
            break
        yield HwpRecord(
            tag_id=tag_id,
            level=level,
            size=size,
            header_size=header_size,
            offset=offset,
            payload=data[payload_offset:payload_end],
        )
        offset = payload_end


def _strip_para_text_controls(raw_text: str) -> str:
    units = list(struct.unpack("<" + "H" * (len(raw_text.encode("utf-16-le")) // 2), raw_text.encode("utf-16-le")))
    output: list[str] = []
    index = 0
    while index < len(units):
        value = units[index]
        if value >= 32:
            output.append(chr(value))
            index += 1
            continue

        if value in (9, 10, 13):
            output.append(chr(value))

        # HWP extended/inline controls occupy 8 UTF-16 code units total.
        if (
            index + 7 < len(units)
            and units[index + 3] == 0
            and units[index + 4] == 0
            and units[index + 5] == 0
            and units[index + 6] == 0
        ):
            index += 8
            continue

        index += 1
    return "".join(output)


class HwpBinaryDocument:
    def __init__(
        self,
        source_path: Path,
        streams: dict[str, bytes],
        file_header: HwpBinaryFileHeader,
    ) -> None:
        self.source_path = source_path
        self._streams = dict(streams)
        self._original_sizes = {path: len(data) for path, data in streams.items()}
        self._file_header = file_header

    @classmethod
    def open(cls, path: str | Path) -> "HwpBinaryDocument":
        input_path = Path(path).expanduser().resolve()
        if not input_path.exists():
            raise FileNotFoundError(input_path)
        if not olefile.isOleFile(str(input_path)):
            raise InvalidHwpFileError(f"{input_path} is not a valid OLE Compound File HWP document.")

        with olefile.OleFileIO(str(input_path)) as ole:
            streams: dict[str, bytes] = {}
            for entry in ole.listdir(streams=True, storages=False):
                stream_path = _normalize_stream_path(entry)
                streams[stream_path] = ole.openstream(entry).read()

        file_header_stream = streams.get("FileHeader")
        if file_header_stream is None:
            raise InvalidHwpFileError(f"{input_path} does not contain a FileHeader stream.")

        file_header = HwpBinaryFileHeader.from_bytes(file_header_stream)
        if file_header.signature != FILE_HEADER_SIGNATURE.decode("ascii"):
            raise InvalidHwpFileError(f"{input_path} does not have the expected HWP file signature.")

        return cls(source_path=input_path, streams=streams, file_header=file_header)

    def file_header(self) -> HwpBinaryFileHeader:
        return self._file_header

    def list_stream_paths(self) -> list[str]:
        return sorted(self._streams)

    def section_stream_paths(self) -> list[str]:
        return sorted(path for path in self._streams if path.startswith("BodyText/Section"))

    def bindata_stream_paths(self) -> list[str]:
        return sorted(path for path in self._streams if path.startswith("BinData/"))

    def has_stream(self, path: str | tuple[str, ...] | list[str]) -> bool:
        return _normalize_stream_path(path) in self._streams

    def stream_size(self, path: str | tuple[str, ...] | list[str]) -> int:
        normalized = _normalize_stream_path(path)
        return self._original_sizes[normalized]

    def stream_capacity(self, path: str | tuple[str, ...] | list[str], data: bytes, *, compress: bool | None = None) -> HwpStreamCapacity:
        normalized = _normalize_stream_path(path)
        if normalized not in self._streams:
            raise KeyError(normalized)
        if compress is None:
            compress = self._stream_is_compressed(normalized)
        original_size = self._original_sizes[normalized]
        if compress:
            compressor = zlib.compressobj(level=9, wbits=-15)
            current_size = len(compressor.compress(data) + compressor.flush())
        else:
            current_size = len(data)
        return HwpStreamCapacity(
            path=normalized,
            compressed=compress,
            original_size=original_size,
            current_size=current_size,
            free_bytes=original_size - current_size,
            fits=current_size <= original_size,
        )

    def section_capacity(self, section_index: int, records: list[HwpRecord] | None = None) -> HwpStreamCapacity:
        path = self.section_stream_paths()[section_index]
        if records is None:
            data = self.read_stream(path)
        else:
            data = b"".join(record.to_bytes() for record in records)
        return self.stream_capacity(path, data, compress=True)

    def docinfo_capacity(self, records: list[HwpRecord] | None = None) -> HwpStreamCapacity:
        data = self.read_stream("DocInfo") if records is None else b"".join(record.to_bytes() for record in records)
        return self.stream_capacity("DocInfo", data, compress=True)

    def read_stream(
        self,
        path: str | tuple[str, ...] | list[str],
        *,
        decompress: bool | None = None,
    ) -> bytes:
        normalized = _normalize_stream_path(path)
        if normalized not in self._streams:
            raise KeyError(normalized)

        data = self._streams[normalized]
        if decompress is None:
            decompress = self._stream_is_compressed(normalized)
        if not decompress:
            return data
        return zlib.decompress(data, -15)

    def add_stream(self, path: str | tuple[str, ...] | list[str], data: bytes) -> None:
        normalized = _normalize_stream_path(path)
        self._streams[normalized] = data
        self._original_sizes.setdefault(normalized, len(data))

    def remove_stream(self, path: str | tuple[str, ...] | list[str]) -> None:
        normalized = _normalize_stream_path(path)
        self._streams.pop(normalized, None)
        self._original_sizes.pop(normalized, None)

    def write_stream(
        self,
        path: str | tuple[str, ...] | list[str],
        data: bytes,
        *,
        compress: bool | None = None,
    ) -> None:
        normalized = _normalize_stream_path(path)
        if normalized not in self._streams:
            raise KeyError(normalized)

        original_size = self._original_sizes[normalized]
        if compress is None:
            compress = self._stream_is_compressed(normalized)

        if compress:
            compressor = zlib.compressobj(level=9, wbits=-15)
            payload = compressor.compress(data) + compressor.flush()
        else:
            payload = data

        self._streams[normalized] = payload

    def preview_text(self) -> str:
        data = self.read_stream("PrvText", decompress=False)
        return data.decode("utf-16-le", errors="ignore").rstrip("\x00")

    def set_preview_text(self, value: str) -> None:
        encoded = value.encode("utf-16-le")
        self.write_stream("PrvText", encoded, compress=False)

    def docinfo_records(self) -> list[HwpRecord]:
        return list(_iter_records(self.read_stream("DocInfo")))

    def document_properties(self) -> HwpDocumentProperties:
        records = self.docinfo_records()
        if not records or records[0].tag_id != TAG_DOCUMENT_PROPERTIES:
            raise InvalidHwpFileError("DocInfo does not start with a document properties record.")
        return _parse_document_properties_payload(records[0].payload)

    def docinfo_model(self) -> DocInfoModel:
        return DocInfoModel.from_records(self.docinfo_records())

    def replace_docinfo_model(self, model: DocInfoModel) -> None:
        self.write_stream("DocInfo", model.to_bytes(), compress=True)

    def section_records(self, section_index: int) -> list[HwpRecord]:
        path = self.section_stream_paths()[section_index]
        return list(_iter_records(self.read_stream(path)))

    def replace_section_records(self, section_index: int, records: list[HwpRecord]) -> None:
        path = self.section_stream_paths()[section_index]
        raw = b"".join(record.to_bytes() for record in records)
        self.write_stream(path, raw, compress=True)

    def section_model(self, section_index: int = 0) -> SectionModel:
        return SectionModel.from_records(section_index, self.section_records(section_index))

    def replace_section_model(self, section_index: int, model: SectionModel) -> None:
        self.replace_section_records(section_index, model.to_records())

    def append_paragraph(
        self,
        text: str,
        *,
        section_index: int = 0,
        para_shape_id: int = 0,
        style_id: int = 0,
        split_flags: int = 0,
        control_mask: int = 0,
    ) -> SectionParagraphModel:
        model = self.section_model(section_index)
        paragraph = model.append_paragraph(
            text,
            para_shape_id=para_shape_id,
            style_id=style_id,
            split_flags=split_flags,
            control_mask=control_mask,
        )
        self.replace_section_model(section_index, model)
        return paragraph

    def append_table(
        self,
        cell_text: str | None = None,
        *,
        rows: int = DEFAULT_HWP_TABLE_ROWS,
        cols: int = DEFAULT_HWP_TABLE_COLS,
        cell_texts: Sequence[str] | Sequence[Sequence[str]] | None = None,
        row_heights: Sequence[int] | None = None,
        col_widths: Sequence[int] | None = None,
        cell_spans: dict[tuple[int, int], tuple[int, int]] | None = None,
        cell_border_fill_ids: dict[tuple[int, int], int] | None = None,
        table_border_fill_id: int = DEFAULT_HWP_TABLE_BORDER_FILL_ID,
        profile_root: str | Path | None = None,
        section_index: int | None = None,
    ) -> None:
        self._append_table_with_text(
            cell_text,
            rows=rows,
            cols=cols,
            cell_texts=cell_texts,
            row_heights=row_heights,
            col_widths=col_widths,
            cell_spans=cell_spans,
            cell_border_fill_ids=cell_border_fill_ids,
            table_border_fill_id=table_border_fill_id,
            profile_root=profile_root,
            section_index=section_index,
        )

    def append_picture(
        self,
        image_bytes: bytes | None = None,
        extension: str | None = None,
        *,
        profile_root: str | Path | None = None,
        section_index: int | None = None,
    ) -> None:
        if image_bytes is None:
            self._append_picture_with_storage_id(
                self._resolve_default_picture_storage_id(),
                profile_root=profile_root,
                section_index=section_index,
            )
            return
        self._append_picture_with_bindata(image_bytes, extension=extension or "png", profile_root=profile_root, section_index=section_index)

    def append_hyperlink(
        self,
        url: str | None = None,
        *,
        text: str | None = None,
        metadata_fields: Sequence[str | int] | None = None,
        profile_root: str | Path | None = None,
        section_index: int | None = None,
    ) -> None:
        self._append_hyperlink_with_text(
            url or DEFAULT_HWP_HYPERLINK_URL,
            text=text or DEFAULT_HWP_HYPERLINK_TEXT,
            metadata_fields=metadata_fields,
            profile_root=profile_root,
            section_index=section_index,
        )

    def append_field(
        self,
        *,
        field_type: str,
        display_text: str | None = None,
        name: str | None = None,
        parameters: dict[str, str | int] | None = None,
        editable: bool = False,
        dirty: bool = False,
        section_index: int | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        normalized_field_type = field_type.strip()
        if normalized_field_type.upper() in {"HYPERLINK", "%HLK"}:
            target = None
            if parameters is not None:
                target = (
                    parameters.get("Path")
                    or parameters.get("Command")
                    or parameters.get("url")
                    or parameters.get("URL")
                )
            self.append_hyperlink(
                str(target or DEFAULT_HWP_HYPERLINK_URL),
                text=display_text,
                section_index=section_index,
            )
            return

        target_section_index = self._resolve_append_section_index(None, section_index)
        display_value = display_text or name or normalized_field_type
        control_payload = _build_control_id_payload(_normalize_field_control_id(normalized_field_type))
        control_node = RecordNode(tag_id=TAG_CTRL_HEADER, level=1, payload=control_payload)
        if parameters:
            parameter_text = ";".join(f"{key}={value}" for key, value in parameters.items())
            control_node.add_child(RecordNode(tag_id=TAG_CTRL_DATA, level=2, payload=parameter_text.encode("utf-16-le")))

        model = self.section_model(target_section_index)
        if paragraph_index is None or paragraph_index < 0:
            paragraph = model.append_paragraph(
                f"{display_value}\r",
                control_mask=1 if editable or dirty else 0,
            )
        else:
            paragraph = model.insert_paragraph(
                paragraph_index,
                f"{display_value}\r",
                control_mask=1 if editable or dirty else 0,
            )
        paragraph.header.add_child(control_node)
        self.replace_section_model(target_section_index, model)

    def append_equation(
        self,
        script: str = DEFAULT_HWP_EQUATION_SCRIPT,
        *,
        width: int = 4800,
        height: int = 2300,
        section_index: int | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        target_section_index = self._resolve_append_section_index(None, section_index)
        model = self.section_model(target_section_index)
        if paragraph_index is None or paragraph_index < 0:
            paragraph = model.append_paragraph("\r")
        else:
            paragraph = model.insert_paragraph(paragraph_index, "\r")
        control_node = RecordNode(tag_id=TAG_CTRL_HEADER, level=1, payload=_build_control_id_payload("eqed"))
        control_node.add_child(RecordNode(tag_id=TAG_EQEDIT, level=2, payload=_build_equation_payload(script)))
        paragraph.header.add_child(control_node)
        self.replace_section_model(target_section_index, model)

    def append_shape(
        self,
        *,
        kind: str = "rect",
        text: str = DEFAULT_HWP_SHAPE_TEXT,
        width: int = DEFAULT_HWP_SHAPE_WIDTH,
        height: int = DEFAULT_HWP_SHAPE_HEIGHT,
        section_index: int | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        kind_to_tag = {
            "line": TAG_SHAPE_COMPONENT_LINE,
            "rect": TAG_SHAPE_COMPONENT_RECTANGLE,
            "ellipse": TAG_SHAPE_COMPONENT_ELLIPSE,
            "arc": TAG_SHAPE_COMPONENT_ARC,
            "polygon": TAG_SHAPE_COMPONENT_POLYGON,
            "curve": TAG_SHAPE_COMPONENT_CURVE,
            "container": TAG_SHAPE_COMPONENT_CONTAINER,
            "textart": TAG_SHAPE_COMPONENT_TEXTART,
        }
        if kind not in kind_to_tag:
            supported = ", ".join(sorted(kind_to_tag))
            raise ValueError(f"append_shape() supports kind values: {supported}")

        target_section_index = self._resolve_append_section_index(None, section_index)
        model = self.section_model(target_section_index)
        paragraph_text = f"{text}\r" if text else "\r"
        if paragraph_index is None or paragraph_index < 0:
            paragraph = model.append_paragraph(paragraph_text)
        else:
            paragraph = model.insert_paragraph(paragraph_index, paragraph_text)
        control_node = RecordNode(tag_id=TAG_CTRL_HEADER, level=1, payload=_build_control_id_payload("gso "))
        shape_component = RecordNode(
            tag_id=TAG_SHAPE_COMPONENT,
            level=2,
            payload=_build_minimal_shape_component_payload(width, height),
        )
        shape_component.add_child(RecordNode(tag_id=kind_to_tag[kind], level=3, payload=b""))
        control_node.add_child(shape_component)
        paragraph.header.add_child(control_node)
        self.replace_section_model(target_section_index, model)

    def append_ole(
        self,
        name: str,
        data: bytes,
        *,
        width: int = DEFAULT_HWP_SHAPE_WIDTH,
        height: int = DEFAULT_HWP_SHAPE_HEIGHT,
        section_index: int | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        target_section_index = self._resolve_append_section_index(None, section_index)
        self.add_embedded_bindata(data, extension="ole")
        model = self.section_model(target_section_index)
        paragraph_text = f"{name}\r" if name else "\r"
        if paragraph_index is None or paragraph_index < 0:
            paragraph = model.append_paragraph(paragraph_text)
        else:
            paragraph = model.insert_paragraph(paragraph_index, paragraph_text)
        control_node = RecordNode(tag_id=TAG_CTRL_HEADER, level=1, payload=_build_control_id_payload("gso "))
        shape_component = RecordNode(
            tag_id=TAG_SHAPE_COMPONENT,
            level=2,
            payload=_build_minimal_shape_component_payload(width, height),
        )
        shape_component.add_child(RecordNode(tag_id=TAG_SHAPE_COMPONENT_OLE, level=3, payload=b""))
        control_node.add_child(shape_component)
        paragraph.header.add_child(control_node)
        self.replace_section_model(target_section_index, model)

    def paragraphs(self, section_index: int = 0) -> list[HwpParagraph]:
        paragraphs: list[HwpParagraph] = []
        current: dict[str, int | str] | None = None
        paragraph_index = -1
        for record in self.section_records(section_index):
            if record.tag_id == TAG_PARA_HEADER:
                if current is not None:
                    paragraphs.append(
                        HwpParagraph(
                            index=paragraph_index,
                            section_index=section_index,
                            char_count=int(current["char_count"]),
                            control_mask=int(current["control_mask"]),
                            para_shape_id=int(current["para_shape_id"]),
                            style_id=int(current["style_id"]),
                            split_flags=int(current["split_flags"]),
                            raw_text=str(current["raw_text"]),
                            text=str(current["text"]),
                            text_record_offset=int(current["text_record_offset"]),
                            text_record_size=int(current["text_record_size"]),
                        )
                    )
                paragraph_index += 1
                current = {
                    "char_count": int.from_bytes(record.payload[0:4], "little") if len(record.payload) >= 4 else 0,
                    "control_mask": int.from_bytes(record.payload[4:8], "little") if len(record.payload) >= 8 else 0,
                    "para_shape_id": int.from_bytes(record.payload[8:10], "little") if len(record.payload) >= 10 else 0,
                    "style_id": record.payload[10] if len(record.payload) >= 11 else 0,
                    "split_flags": record.payload[11] if len(record.payload) >= 12 else 0,
                    "raw_text": "",
                    "text": "",
                    "text_record_offset": -1,
                    "text_record_size": 0,
                }
                continue

            if record.tag_id == TAG_PARA_TEXT and current is not None:
                raw_text = record.payload.decode("utf-16-le", errors="ignore")
                current["raw_text"] = raw_text
                current["text"] = _strip_para_text_controls(raw_text)
                current["text_record_offset"] = record.offset
                current["text_record_size"] = record.size

        if current is not None:
            paragraphs.append(
                HwpParagraph(
                    index=paragraph_index,
                    section_index=section_index,
                    char_count=int(current["char_count"]),
                    control_mask=int(current["control_mask"]),
                    para_shape_id=int(current["para_shape_id"]),
                    style_id=int(current["style_id"]),
                    split_flags=int(current["split_flags"]),
                    raw_text=str(current["raw_text"]),
                    text=str(current["text"]),
                    text_record_offset=int(current["text_record_offset"]),
                    text_record_size=int(current["text_record_size"]),
                )
            )
        return paragraphs

    def get_document_text(self) -> str:
        fragments: list[str] = []
        for index, _path in enumerate(self.section_stream_paths()):
            for paragraph in self.paragraphs(index):
                cleaned = paragraph.text.replace("\r", "\n").strip("\n")
                if cleaned:
                    fragments.append(cleaned)
        return "\n".join(fragments)

    def set_paragraph_text_same_length(
        self,
        section_index: int,
        paragraph_index: int,
        new_text: str,
    ) -> None:
        paragraph = self.paragraphs(section_index)[paragraph_index]
        if paragraph.raw_text != paragraph.text:
            raise HwpBinaryEditError(
                "set_paragraph_text_same_length is only supported for paragraphs without hidden control characters."
            )
        if len(new_text.encode("utf-16-le")) != len(paragraph.raw_text.encode("utf-16-le")):
            raise HwpBinaryEditError(
                "set_paragraph_text_same_length requires the new text to match the paragraph's UTF-16 byte length."
            )
        self._rewrite_paragraph_text_record(section_index, paragraph_index, lambda _current: new_text)

    def replace_paragraph_text_same_length(
        self,
        section_index: int,
        paragraph_index: int,
        old: str,
        new: str,
        *,
        count: int = -1,
    ) -> int:
        if not old:
            raise ValueError("old must be a non-empty string.")
        if len(old.encode("utf-16-le")) != len(new.encode("utf-16-le")):
            raise HwpBinaryEditError(
                "replace_paragraph_text_same_length requires old and new to have the same UTF-16 byte length."
            )

        replaced = 0

        def updater(current: str) -> str:
            nonlocal replaced
            updated, changed = _replace_with_limit(current, old, new, count)
            replaced = changed
            return updated

        self._rewrite_paragraph_text_record(section_index, paragraph_index, updater)
        return replaced

    def replace_text_same_length(
        self,
        old: str,
        new: str,
        *,
        section_index: int | None = None,
        count: int = -1,
    ) -> int:
        if not old:
            raise ValueError("old must be a non-empty string.")
        if len(old.encode("utf-16-le")) != len(new.encode("utf-16-le")):
            raise HwpBinaryEditError("replace_text_same_length requires old and new to have the same UTF-16 byte length.")

        remaining = count
        replacements = 0
        target_paths = self.section_stream_paths()
        if section_index is not None:
            target_paths = [target_paths[section_index]]

        for path in target_paths:
            raw = self.read_stream(path)
            chunks: list[bytes] = []
            changed_in_stream = 0

            for record in _iter_records(raw):
                header_bytes = raw[record.offset : record.offset + record.header_size]
                payload = record.payload
                if record.tag_id == TAG_PARA_TEXT:
                    current_text = payload.decode("utf-16-le", errors="ignore")
                    limit = remaining if remaining >= 0 else -1
                    updated_text, changed = _replace_with_limit(current_text, old, new, limit)
                    if changed:
                        payload = updated_text.encode("utf-16-le")
                        changed_in_stream += changed
                        replacements += changed
                        if remaining >= 0:
                            remaining -= changed
                chunks.append(header_bytes)
                chunks.append(payload)
                if remaining == 0:
                    # Keep the rest of the stream untouched when the requested replacement count is exhausted.
                    tail_offset = record.offset + record.header_size + len(record.payload)
                    chunks.append(raw[tail_offset:])
                    break
            else:
                tail_offset = len(raw)

            if changed_in_stream:
                updated_raw = b"".join(chunks)
                if len(updated_raw) != len(raw):
                    raise HwpBinaryEditError("Section rewrite changed the decompressed stream length unexpectedly.")
                self.write_stream(path, updated_raw, compress=True)

            if remaining == 0:
                break

        return replacements

    def _rewrite_paragraph_text_record(
        self,
        section_index: int,
        paragraph_index: int,
        update_text: Callable[[str], str],
    ) -> None:
        path = self.section_stream_paths()[section_index]
        raw = self.read_stream(path)
        chunks: list[bytes] = []
        current_paragraph_index = -1
        changed = False

        for record in _iter_records(raw):
            header_bytes = raw[record.offset : record.offset + record.header_size]
            payload = record.payload

            if record.tag_id == TAG_PARA_HEADER:
                current_paragraph_index += 1
            elif record.tag_id == TAG_PARA_TEXT and current_paragraph_index == paragraph_index:
                current_text = payload.decode("utf-16-le", errors="ignore")
                updated_text = update_text(current_text)
                updated_payload = updated_text.encode("utf-16-le")
                if len(updated_payload) != len(payload):
                    raise HwpBinaryEditError("Paragraph text rewrite changed the paragraph text record size.")
                payload = updated_payload
                changed = True

            chunks.append(header_bytes)
            chunks.append(payload)

        if not changed:
            raise IndexError(
                f"Paragraph {paragraph_index} in section {section_index} does not contain a writable PARA_TEXT record."
            )

        updated_raw = b"".join(chunks)
        if len(updated_raw) != len(raw):
            raise HwpBinaryEditError("Section rewrite changed the decompressed stream length unexpectedly.")
        self.write_stream(path, updated_raw, compress=True)

    def save(self, path: str | Path | None = None) -> Path:
        target_path = self.source_path if path is None else Path(path).expanduser().resolve()
        self._write_to_path(target_path)
        if path is not None:
            self.source_path = target_path
        return target_path

    def save_copy(self, path: str | Path) -> Path:
        target_path = Path(path).expanduser().resolve()
        self._write_to_path(target_path)
        return target_path

    def _stream_is_compressed(self, path: str) -> bool:
        if not self._file_header.compressed:
            return False
        return path == "DocInfo" or path.startswith("BodyText/")

    def _compress_stream_preserving_capacity(self, path: str, data: bytes) -> bytes:
        capacity = self.stream_capacity(path, data, compress=True)
        original_size = capacity.original_size
        compressor = zlib.compressobj(level=9, wbits=-15)
        compressed = compressor.compress(data) + compressor.flush()
        if not capacity.fits:
            raise HwpBinaryEditError(
                f"{path} compressed to {capacity.current_size} bytes and no longer fits in its original {original_size}-byte stream."
            )
        return compressed + (b"\x00" * (original_size - len(compressed)))

    def _write_to_path(self, target_path: Path) -> None:
        target_path.parent.mkdir(parents=True, exist_ok=True)
        write_compound_file(target_path, self._streams)

    def _append_feature_via_profile(
        self,
        feature: str,
        *,
        profile_root: str | Path | None,
        section_index: int | None,
    ) -> None:
        from .hwp_pure_profile import HwpPureProfile, append_feature_from_profile

        profile = HwpPureProfile.load_bundled() if profile_root is None else HwpPureProfile.load(profile_root)
        if section_index is not None and section_index != profile.target_section_index:
            raise HwpBinaryEditError(
                f"Pure profile feature templates currently target section {profile.target_section_index}, "
                f"but section {section_index} was requested."
        )
        append_feature_from_profile(self, profile, feature)

    def _resolve_append_section_index(self, profile_root: str | Path | None, section_index: int | None) -> int:
        if section_index is not None:
            return section_index
        if profile_root is None:
            from .hwp_pure_profile import HwpPureProfile

            return HwpPureProfile.load_bundled().target_section_index
        profile_path = Path(profile_root).expanduser().resolve()
        if not profile_path.exists():
            return 0
        metadata_path = profile_path / "profile.json"
        if not metadata_path.exists():
            return 0
        from .hwp_pure_profile import HwpPureProfile

        return HwpPureProfile.load(profile_path).target_section_index

    def add_embedded_bindata(
        self,
        data: bytes,
        *,
        extension: str,
        storage_id: int | None = None,
        flags: int = 0x0001,
    ) -> tuple[int, str]:
        model = self.docinfo_model()
        bindata = model.add_embedded_bindata(extension, storage_id=storage_id, flags=flags)
        self.replace_docinfo_model(model)
        stream_path = f"BinData/BIN{int(bindata.storage_id or 0):04d}.{(bindata.extension or extension).lstrip('.').lower()}"
        self.add_stream(stream_path, data)
        return int(bindata.storage_id or 0), stream_path

    def remove_embedded_bindata(self, storage_id: int) -> bool:
        model = self.docinfo_model()
        removed = model.remove_bindata(storage_id)
        if not removed:
            return False
        self.replace_docinfo_model(model)
        prefix = f"BinData/BIN{storage_id:04d}."
        for stream_path in list(self._streams):
            if stream_path.startswith(prefix):
                self.remove_stream(stream_path)
        return True

    def _resolve_default_picture_storage_id(self) -> int:
        for stream_path in self.bindata_stream_paths():
            name = Path(stream_path).name
            if not name.startswith("BIN") or "." not in name:
                continue
            storage_token = name[3:].split(".", 1)[0]
            if storage_token.isdigit():
                return int(storage_token)
        raise HwpBinaryEditError(
            "append_picture() without image_bytes requires at least one existing BinData stream in the document."
        )

    def _build_picture_records(self, storage_id: int) -> list[HwpRecord]:
        top_char_count = len(_PICTURE_TOP_PARA_TEXT) // 2
        picture_payload = bytearray(_PICTURE_SHAPE_PICTURE)
        if len(picture_payload) <= 72:
            raise HwpBinaryEditError("Built picture payload is shorter than expected.")
        picture_payload[71:73] = int(storage_id).to_bytes(2, "little", signed=False)
        return [
            HwpRecord(
                tag_id=TAG_PARA_HEADER,
                level=0,
                size=22,
                header_size=4,
                offset=-1,
                payload=top_char_count.to_bytes(4, "little") + _PICTURE_TOP_PARA_HEADER_TAIL,
            ),
            HwpRecord(tag_id=TAG_PARA_TEXT, level=1, size=len(_PICTURE_TOP_PARA_TEXT), header_size=4, offset=-1, payload=_PICTURE_TOP_PARA_TEXT),
            HwpRecord(tag_id=TAG_PARA_CHAR_SHAPE, level=1, size=len(_PICTURE_TOP_CHAR_SHAPE), header_size=4, offset=-1, payload=_PICTURE_TOP_CHAR_SHAPE),
            HwpRecord(tag_id=TAG_PARA_LINE_SEG, level=1, size=len(_PICTURE_TOP_LINE_SEG), header_size=4, offset=-1, payload=_PICTURE_TOP_LINE_SEG),
            HwpRecord(tag_id=TAG_CTRL_HEADER, level=1, size=len(_PICTURE_CTRL_HEADER), header_size=4, offset=-1, payload=_PICTURE_CTRL_HEADER),
            HwpRecord(
                tag_id=TAG_SHAPE_COMPONENT,
                level=2,
                size=len(_PICTURE_SHAPE_COMPONENT),
                header_size=4,
                offset=-1,
                payload=_PICTURE_SHAPE_COMPONENT,
            ),
            HwpRecord(
                tag_id=TAG_SHAPE_COMPONENT_PICTURE,
                level=3,
                size=len(picture_payload),
                header_size=4,
                offset=-1,
                payload=bytes(picture_payload),
            ),
        ]

    def _append_picture_with_storage_id(
        self,
        storage_id: int,
        *,
        profile_root: str | Path | None,
        section_index: int | None,
    ) -> None:
        target_section_index = self._resolve_append_section_index(profile_root, section_index)
        section_records = self.section_records(target_section_index)
        section_records.extend(self._build_picture_records(storage_id))
        self.replace_section_records(target_section_index, section_records)

    def _append_picture_with_bindata(
        self,
        image_bytes: bytes,
        *,
        extension: str,
        profile_root: str | Path | None,
        section_index: int | None,
    ) -> None:
        storage_id, _stream_path = self.add_embedded_bindata(image_bytes, extension=extension)
        self._append_picture_with_storage_id(storage_id, profile_root=profile_root, section_index=section_index)

    def _append_table_with_text(
        self,
        cell_text: str | None,
        *,
        rows: int,
        cols: int,
        cell_texts: Sequence[str] | Sequence[Sequence[str]] | None,
        row_heights: Sequence[int] | None,
        col_widths: Sequence[int] | None,
        cell_spans: dict[tuple[int, int], tuple[int, int]] | None,
        cell_border_fill_ids: dict[tuple[int, int], int] | None,
        table_border_fill_id: int,
        profile_root: str | Path | None,
        section_index: int | None,
    ) -> None:
        target_section_index = self._resolve_append_section_index(profile_root, section_index)
        matrix = _normalize_table_cell_texts(rows, cols, cell_text, cell_texts)
        normalized_row_heights = _normalize_table_measurements(
            rows,
            row_heights,
            default=DEFAULT_HWP_TABLE_CELL_HEIGHT,
            name="row_heights",
        )
        normalized_col_widths = _normalize_table_measurements(
            cols,
            col_widths,
            default=DEFAULT_HWP_TABLE_CELL_WIDTH,
            name="col_widths",
        )
        cell_specs = _build_table_cell_specs(
            matrix,
            row_heights=normalized_row_heights,
            col_widths=normalized_col_widths,
            cell_spans=cell_spans,
            cell_border_fill_ids=cell_border_fill_ids,
            default_border_fill_id=table_border_fill_id,
            cell_margins=None,
        )
        top_char_count = len(_TABLE_TOP_PARA_TEXT) // 2
        records_to_append = [
            HwpRecord(
                tag_id=TAG_PARA_HEADER,
                level=0,
                size=24,
                header_size=4,
                offset=-1,
                payload=top_char_count.to_bytes(4, "little") + _TABLE_TOP_PARA_HEADER_TAIL,
            ),
            HwpRecord(tag_id=TAG_PARA_TEXT, level=1, size=len(_TABLE_TOP_PARA_TEXT), header_size=4, offset=-1, payload=_TABLE_TOP_PARA_TEXT),
            HwpRecord(tag_id=TAG_PARA_CHAR_SHAPE, level=1, size=len(_TABLE_TOP_CHAR_SHAPE), header_size=4, offset=-1, payload=_TABLE_TOP_CHAR_SHAPE),
            HwpRecord(tag_id=TAG_PARA_LINE_SEG, level=1, size=len(_TABLE_TOP_LINE_SEG), header_size=4, offset=-1, payload=_TABLE_TOP_LINE_SEG),
            HwpRecord(tag_id=TAG_CTRL_HEADER, level=1, size=len(_TABLE_CTRL_HEADER), header_size=4, offset=-1, payload=_TABLE_CTRL_HEADER),
        ]
        table_payload = _build_table_record_payload(
            rows,
            cols,
            default_row_height=1,
            border_fill_id=table_border_fill_id,
        )
        table_payload = bytearray(table_payload)
        row_sizes_offset = 18
        for row_index, row_height in enumerate(normalized_row_heights):
            offset = row_sizes_offset + row_index * 2
            table_payload[offset : offset + 2] = int(row_height).to_bytes(2, "little", signed=False)
        records_to_append.append(HwpRecord(tag_id=TAG_TABLE, level=2, size=len(table_payload), header_size=4, offset=-1, payload=bytes(table_payload)))
        for cell in cell_specs:
            cell_payload = _build_table_cell_list_header_payload(cell)
            records_to_append.append(
                HwpRecord(tag_id=TAG_LIST_HEADER, level=2, size=len(cell_payload), header_size=4, offset=-1, payload=cell_payload)
            )
            cell_raw_text = str(cell.text) + "\r"
            cell_char_count = len(cell_raw_text)
            records_to_append.append(
                HwpRecord(
                    tag_id=TAG_PARA_HEADER,
                    level=2,
                    size=24,
                    header_size=4,
                    offset=-1,
                    payload=(0x80000000 | cell_char_count).to_bytes(4, "little") + _TABLE_CELL_PARA_HEADER_TAIL,
                )
            )
            records_to_append.append(
                HwpRecord(
                    tag_id=TAG_PARA_TEXT,
                    level=3,
                    size=len(cell_raw_text.encode("utf-16-le")),
                    header_size=4,
                    offset=-1,
                    payload=cell_raw_text.encode("utf-16-le"),
                )
            )
            records_to_append.append(
                HwpRecord(
                    tag_id=TAG_PARA_CHAR_SHAPE,
                    level=3,
                    size=len(_TABLE_CELL_CHAR_SHAPE),
                    header_size=4,
                    offset=-1,
                    payload=_TABLE_CELL_CHAR_SHAPE,
                )
            )
            records_to_append.append(
                HwpRecord(
                    tag_id=TAG_PARA_LINE_SEG,
                    level=3,
                    size=len(_TABLE_CELL_LINE_SEG),
                    header_size=4,
                    offset=-1,
                    payload=_TABLE_CELL_LINE_SEG,
                )
            )
        section_records = self.section_records(target_section_index)
        section_records.extend(records_to_append)
        self.replace_section_records(target_section_index, section_records)

    def _append_hyperlink_with_text(
        self,
        url: str,
        *,
        text: str | None,
        metadata_fields: Sequence[str | int] | None,
        profile_root: str | Path | None,
        section_index: int | None,
    ) -> None:
        target_section_index = self._resolve_append_section_index(profile_root, section_index)
        display_text = text or url
        raw_payload = _build_hyperlink_para_text_payload(display_text)
        char_count = len(raw_payload) // 2
        control_payload = _build_hyperlink_ctrl_payload(url, metadata_fields=metadata_fields)

        records_to_append = [
            HwpRecord(
                tag_id=TAG_PARA_HEADER,
                level=0,
                size=22,
                header_size=4,
                offset=-1,
                payload=(0x80000000 | char_count).to_bytes(4, "little") + _HYPERLINK_PARA_HEADER_TAIL,
            ),
            HwpRecord(tag_id=TAG_PARA_TEXT, level=1, size=len(raw_payload), header_size=4, offset=-1, payload=raw_payload),
            HwpRecord(tag_id=TAG_PARA_CHAR_SHAPE, level=1, size=len(_HYPERLINK_CHAR_SHAPE), header_size=4, offset=-1, payload=_HYPERLINK_CHAR_SHAPE),
            HwpRecord(tag_id=TAG_PARA_LINE_SEG, level=1, size=len(_HYPERLINK_LINE_SEG), header_size=4, offset=-1, payload=_HYPERLINK_LINE_SEG),
            HwpRecord(tag_id=TAG_CTRL_HEADER, level=1, size=len(control_payload), header_size=4, offset=-1, payload=control_payload),
        ]
        section_records = self.section_records(target_section_index)
        section_records.extend(records_to_append)
        self.replace_section_records(target_section_index, section_records)
