from __future__ import annotations

import copy
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
DEFAULT_HWP_EQUATION_FONT = "HYhwpEQ"
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

_DEFAULT_EQUATION_CTRL_HEADER = bytes.fromhex(
    "6465716511222a0c000000000000000045120000ef0800001300000000000000000000006d58f07000000000030031bcb9c2b0c6"
)
_DEFAULT_SHAPE_CTRL_HEADER = bytes.fromhex(
    "206f736700206a048d2a0000fb1400006e2300002e0c00004902000000000000000000009158d143000000000000"
)
_DEFAULT_ELLIPSE_CTRL_HEADER = bytes.fromhex(
    "206f736711222a100000000000000000d80e0000d0070000000000000000000000000000e84d387f000000000600c0d0d0c685c7c8b2e4b22e00"
)
_DEFAULT_ARC_CTRL_HEADER = bytes.fromhex(
    "206f736711222a000000000000000000a00f0000980800000000000000000000000000008150387f00000000050038d685c7c8b2e4b22e00"
)
_DEFAULT_POLYGON_CTRL_HEADER = bytes.fromhex(
    "206f736711222a00000000000000000068100000600900000100000000000000000000008350387f000000000700e4b201ac15d685c7c8b2e4b22e00"
)
_DEFAULT_TEXTART_CTRL_HEADER = bytes.fromhex(
    "206f736711220a140000000000000000e457000088130000000000003800380000000000cc76387f00000000070000aef5b9dcc285c7c8b2e4b22e00"
)
_DEFAULT_SHAPE_COMPONENT = bytes.fromhex(
    "636572246365722400000000feffffff000001006c2300002c0c00006e2300002d0c0000000000010000b7110000170600000100000000000000"
    "f03f000000000000000000000000000000000000000000000000000000000000f03f00000000000000c0a435ff44e700f03f00000000000000000000"
    "000000000000000000000000001db2e606a102f03f0000000000000040000000000000f03f000000000000000000000000000000000000000000000000"
    "0000000000f03f00000000000000000000000038000000410000c0000000000000000000000000000000000000000000000000009258d1030000"
)
_DEFAULT_ELLIPSE_SHAPE_COMPONENT = bytes.fromhex(
    "6c6c65246c6c65240000000000000000000001000000000000000000000000000000000000000900000000000000000000000100000000000000f03f000000000000000000000000000000000000000000000000000000000000f03f0000000000000000000000000000f03f000000000000000000000000000000000000000000000000000000000000f03f0000000000000000000000000000f03f000000000000000000000000000000000000000000000000000000000000f03f000000000000000000000000000000000000000000000000000000000000000000b2b2b2000000000000000000e94d383f0000"
)
_DEFAULT_ARC_SHAPE_COMPONENT = bytes.fromhex(
    "63726124637261240000000000000000000001000000000000000000000000000000000000000800000000000000000000000100000000000000f03f000000000000000000000000000000000000000000000000000000000000f03f0000000000000000000000000000f03f000000000000000000000000000000000000000000000000000000000000f03f0000000000000000000000000000f03f000000000000000000000000000000000000000000000000000000000000f03f000000000000000000000000000000000000000000000000000000000000000000b2b2b20000000000000000008250383f0000"
)
_DEFAULT_POLYGON_SHAPE_COMPONENT = bytes.fromhex(
    "6c6f70246c6f70240000000000000000000001000000000000000000000000000000000000000800000000000000000000000100000000000000f03f000000000000000000000000000000000000000000000000000000000000f03f0000000000000000000000000000f03f000000000000000000000000000000000000000000000000000000000000f03f0000000000000000000000000000f03f000000000000000000000000000000000000000000000000000000000000f03f000000000000000000000000000000000000000000000000000000000000000000b2b2b20000000000000000008450383f0000"
)
_DEFAULT_TEXTART_SHAPE_COMPONENT = bytes.fromhex(
    "74617424746174240000000000000000000001005d3700005d3700005d3700005d37000000000b00000000000000000000000100000000000000f03f000000000000000000000000000000000000000000000000000000000000f03f0000000000000000000000000000f03f000000000000000000000000000000000000000000000000000000000000f03f0000000000000000000000000000f03f000000000000000000000000000000000000000000000000000000000000f03f00000000000000004455660021000000000000000201000000ddeeff000000000001000000000000000000000000b2b2b2000000000000000000cd76383f0000"
)
_DEFAULT_SHAPE_LIST_HEADER = bytes.fromhex("010000000000000000000000000000006e230000000000000000000000000000ff1b0201000000004001000000")
_DEFAULT_SHAPE_TEXT_PARA_CONTROL_MASK = int.from_bytes(bytes.fromhex("00000080"), "little", signed=False)
_DEFAULT_SHAPE_TEXT_PARA_SHAPE_ID = 0x17
_DEFAULT_SHAPE_TEXT_STYLE_ID = 0x16
_DEFAULT_SHAPE_TEXT_SPLIT_FLAGS = 0x00
_DEFAULT_SHAPE_TEXT_PARA_TRAILING = bytes.fromhex("070000000100000000800000")
_DEFAULT_SHAPE_TEXT_CHAR_SHAPE = bytes.fromhex(
    "000000001a000000010000001b000000020000001c000000030000001b000000040000001d000000050000001e000000"
)
_DEFAULT_SHAPE_TEXT_LINE_SEG = bytes.fromhex("000000006400000098080000980800004e07000028050000000000006c23000000000600")
_DEFAULT_SHAPE_RECTANGLE_PAYLOAD = bytes.fromhex("3200000000000000006c230000000000006c2300002c0c0000000000002c0c0000")
_DEFAULT_SHAPE_LINE_PAYLOAD = bytes.fromhex("0000000000000000640000006400000000000000")
_DEFAULT_ELLIPSE_PAYLOAD = bytes.fromhex(
    "00000000200020007d00200020007e00200020007f000d00000000004a01000000000000000000000000000083002000200084002000200085000d00"
)
_DEFAULT_ARC_PAYLOAD = bytes.fromhex("00200020008100200020008200200020008300200020008400")
_DEFAULT_POLYGON_PAYLOAD = bytes.fromhex("0000000000000000")
_EQUATION_VERSION_MARKER = "Equation Version 60"
_DEFAULT_TEXTART_FONT_NAME = "Malgun Gothic"
_DEFAULT_TEXTART_FONT_STYLE = "Bold"
_DEFAULT_SHAPE_LINE_WIDTH = 33
_DEFAULT_OLE_LINE_WIDTH = 0
_HWP_FILLED_SHAPE_COMMON_PAYLOAD_SIZE = 32
_SHAPE_NATIVE_METADATA_PREFIX = "JAKAL_SHAPE_META"
_HWP_COMPLEX_SHAPE_TEMPLATES = {
    "ellipse": {
        "ctrl": _DEFAULT_ELLIPSE_CTRL_HEADER,
        "component": _DEFAULT_ELLIPSE_SHAPE_COMPONENT,
    },
    "arc": {
        "ctrl": _DEFAULT_ARC_CTRL_HEADER,
        "component": _DEFAULT_ARC_SHAPE_COMPONENT,
    },
    "polygon": {
        "ctrl": _DEFAULT_POLYGON_CTRL_HEADER,
        "component": _DEFAULT_POLYGON_SHAPE_COMPONENT,
    },
    "textart": {
        "ctrl": _DEFAULT_TEXTART_CTRL_HEADER,
        "component": _DEFAULT_TEXTART_SHAPE_COMPONENT,
    },
}

_HWP_LINE_STYLE_CODES = {
    "SOLID": 0,
    "NONE": 0,
}
_HWP_LINE_END_CAP_CODES = {
    "ROUND": 0,
    "FLAT": 1,
}
_HWP_LINE_ARROW_CODES = {
    "NORMAL": 0,
}
_HWP_LINE_ARROW_SIZE_CODES = {
    "SMALL_SMALL": 0,
    "SMALL_MEDIUM": 1,
    "SMALL_LARGE": 2,
    "MEDIUM_SMALL": 3,
    "MEDIUM_MEDIUM": 4,
    "MEDIUM_LARGE": 5,
    "LARGE_SMALL": 6,
    "LARGE_MEDIUM": 7,
    "LARGE_LARGE": 8,
}
_HWP_LINE_OUTLINE_STYLE_CODES = {
    "NORMAL": 0,
    "OUTER": 1,
    "INNER": 2,
}
_HWP_DRAW_ASPECT_CODES = {
    "CONTENT": 1,
    "THUMBNAIL": 2,
    "ICON": 4,
    "DOCPRINT": 8,
}
_HWP_OLE_OBJECT_TYPE_CODES = {
    "UNKNOWN": 0,
    "EMBEDDED": 1,
    "LINK": 2,
    "STATIC": 3,
    "EQUATION": 4,
}

_CONTROL_MARKER_CODES = {
    "head": 0x10,
    "foot": 0x10,
    "atno": 0x12,
    "nwno": 0x15,
    "pgnp": 0x15,
    "gso ": 0x0B,
    "fn  ": 0x11,
    "en  ": 0x11,
    "bokm": 0x16,
}
_AUTO_NUMBER_TYPE_CODES = {
    "PAGE": 0x0C,
}

_HEADER_CTRL_HEADER = bytes.fromhex("646165680000000002000000")
_HEADER_LIST_HEADER = bytes.fromhex("01000000000000003ebc00009b100000000000000000000000000000000000000000")
_HEADER_PARA_HEADER = bytes.fromhex("01000080000000001b000a00010000000100000000800000")
_HEADER_PARA_CHAR_SHAPE = bytes.fromhex("0000000014000000")
_HEADER_PARA_LINE_SEG = bytes.fromhex("00000000000000006400000064000000550000003c000000000000003cbc000000000600")

_FOOTER_CTRL_HEADER = bytes.fromhex("746f6f6600000000")
_FOOTER_LIST_HEADER = bytes.fromhex("01000000400000003ebc00009c100000000000000000000000000000000000000000")
_FOOTER_PARA_HEADER = bytes.fromhex("01000080000000002f00000001000000010000000080")
_FOOTER_PARA_CHAR_SHAPE = bytes.fromhex("000000006a000000")
_FOOTER_PARA_LINE_SEG = bytes.fromhex("00000000000000008403000084030000fd0200001c020000000000003cbc000000000600")

_BOOKMARK_CTRL_DATA_PREFIX = bytes.fromhex("1b020100000000400100")

_FOOTNOTE_CTRL_HEADER = bytes.fromhex("20206e6601000000000029000000000000000000")
_ENDNOTE_CTRL_HEADER = bytes.fromhex("20206e6501000000000029000000000000000000")
_NOTE_LIST_HEADER = bytes.fromhex("01000000000000000000000000000000")
_NOTE_PARA_HEADER = bytes.fromhex("000000800000000000000000010000000000010000000000")
_NOTE_CHAR_SHAPE = bytes.fromhex("0000000000000000")
_NOTE_LINE_SEG = bytes.fromhex("0000000000000000e8030000e803000052030000580200000000000018a6000000000600")

_BLANK_DROP_CONTROL_IDS = {
    "atno",
    "bokm",
    "en  ",
    "eqed",
    "fn  ",
    "foot",
    "gso ",
    "head",
    "nwno",
    "tbl ",
}


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


def _build_equation_payload(script: str, *, font: str = DEFAULT_HWP_EQUATION_FONT) -> bytes:
    body = f"* {script}`{_EQUATION_VERSION_MARKER}\x07{font}"
    return b"\x00\x00\x00\x00" + body.encode("utf-16-le")


def _build_graphic_control_header_payload(
    template: bytes,
    control_id: str,
    *,
    width: int,
    height: int,
    description: str | None = None,
) -> bytes:
    payload = bytearray(template)
    payload[0:4] = _build_control_id_payload(control_id)
    if len(payload) >= 24:
        payload[16:20] = int(width).to_bytes(4, "little", signed=False)
        payload[20:24] = int(height).to_bytes(4, "little", signed=False)
    encoded_description = str(description or "").encode("utf-16-le")
    base = bytes(payload[:44]).ljust(44, b"\x00")
    return base + (len(encoded_description) // 2).to_bytes(2, "little", signed=False) + encoded_description


def _build_minimal_shape_component_payload(width: int, height: int) -> bytes:
    payload = bytearray(28)
    payload[20:24] = int(width).to_bytes(4, "little", signed=False)
    payload[24:28] = int(height).to_bytes(4, "little", signed=False)
    return bytes(payload)


def _shape_control_header_template(kind: str) -> bytes:
    return _HWP_COMPLEX_SHAPE_TEMPLATES.get(kind, {}).get("ctrl", _DEFAULT_SHAPE_CTRL_HEADER)


def _build_shape_component_payload(
    kind: str,
    width: int,
    height: int,
    *,
    fill_color: str = "#FFFFFF",
    line_color: str = "#000000",
    line_width: int = _DEFAULT_SHAPE_LINE_WIDTH,
) -> bytes:
    template = _HWP_COMPLEX_SHAPE_TEMPLATES.get(kind, {}).get("component")
    if template is not None:
        payload = bytearray(template)
        if len(payload) >= 200:
            payload[196:200] = _build_colorref(line_color)
        if len(payload) >= 202:
            payload[200:202] = max(0, min(int(line_width), 0xFFFF)).to_bytes(2, "little", signed=False)
        if kind == "textart" and len(payload) >= 217:
            payload[213:217] = _build_colorref(fill_color)
        return bytes(payload)
    payload = bytearray(_DEFAULT_SHAPE_COMPONENT)
    payload[20:24] = int(width).to_bytes(4, "little", signed=False)
    payload[24:28] = int(height).to_bytes(4, "little", signed=False)
    payload[28:32] = int(width).to_bytes(4, "little", signed=False)
    payload[32:36] = int(height).to_bytes(4, "little", signed=False)
    return bytes(payload)


def _build_shape_native_metadata_payload(**fields: str) -> bytes:
    serialized = ";".join(f"{name}={value}" for name, value in fields.items())
    text = f"{_SHAPE_NATIVE_METADATA_PREFIX};{serialized}" if serialized else _SHAPE_NATIVE_METADATA_PREFIX
    return text.encode("utf-16-le")


def _build_line_attribute_flags(
    *,
    line_style: str = "SOLID",
    end_cap: str = "FLAT",
    head_style: str = "NORMAL",
    tail_style: str = "NORMAL",
    head_size: str = "MEDIUM_MEDIUM",
    tail_size: str = "MEDIUM_MEDIUM",
    head_fill: bool = True,
    tail_fill: bool = True,
) -> int:
    value = _HWP_LINE_STYLE_CODES.get(str(line_style).upper(), 0) & 0x3F
    value |= (_HWP_LINE_END_CAP_CODES.get(str(end_cap).upper(), 1) & 0x0F) << 6
    value |= (_HWP_LINE_ARROW_CODES.get(str(head_style).upper(), 0) & 0x3F) << 10
    value |= (_HWP_LINE_ARROW_CODES.get(str(tail_style).upper(), 0) & 0x3F) << 16
    value |= (_HWP_LINE_ARROW_SIZE_CODES.get(str(head_size).upper(), 4) & 0x0F) << 22
    value |= (_HWP_LINE_ARROW_SIZE_CODES.get(str(tail_size).upper(), 4) & 0x0F) << 26
    if head_fill:
        value |= 1 << 30
    if tail_fill:
        value |= 1 << 31
    return value


def _build_shape_line_info_payload(
    *,
    color: str = "#000000",
    width: int = _DEFAULT_SHAPE_LINE_WIDTH,
    line_style: str = "SOLID",
    end_cap: str = "FLAT",
    head_style: str = "NORMAL",
    tail_style: str = "NORMAL",
    head_size: str = "MEDIUM_MEDIUM",
    tail_size: str = "MEDIUM_MEDIUM",
    head_fill: bool = True,
    tail_fill: bool = True,
    outline_style: str = "NORMAL",
) -> bytes:
    payload = bytearray(11)
    payload[0:4] = _build_colorref(color)
    payload[4:6] = max(0, min(int(width), 0xFFFF)).to_bytes(2, "little", signed=False)
    payload[6:10] = _build_line_attribute_flags(
        line_style=line_style,
        end_cap=end_cap,
        head_style=head_style,
        tail_style=tail_style,
        head_size=head_size,
        tail_size=tail_size,
        head_fill=head_fill,
        tail_fill=tail_fill,
    ).to_bytes(4, "little", signed=False)
    payload[10] = _HWP_LINE_OUTLINE_STYLE_CODES.get(str(outline_style).upper(), 0) & 0xFF
    return bytes(payload)


def _build_shape_fill_payload(*, face_color: str = "#FFFFFF", hatch_color: str = "#000000") -> bytes:
    payload = bytearray()
    payload.extend((0x00000001).to_bytes(4, "little", signed=False))
    payload.extend(_build_colorref(face_color))
    payload.extend(_build_colorref(hatch_color))
    payload.extend((0).to_bytes(4, "little", signed=True))
    payload.append(0)
    payload.extend((0).to_bytes(4, "little", signed=False))
    return bytes(payload)


def _build_filled_shape_common_payload(*, fill_color: str = "#FFFFFF", line_color: str = "#000000") -> bytes:
    return (
        _build_shape_line_info_payload(color=line_color, width=_DEFAULT_SHAPE_LINE_WIDTH)
        + _build_shape_fill_payload(face_color=fill_color, hatch_color=line_color)
    )


def _build_ellipse_specific_payload(width: int, height: int, *, arc_flags: int = 0) -> bytes:
    payload = bytearray()
    center_x = int(width) // 2
    center_y = int(height) // 2
    axis1_x = int(width) // 2
    axis1_y = 0
    axis2_x = 0
    axis2_y = int(height) // 2
    start_pos_x = int(width)
    start_pos_y = center_y
    end_pos_x = center_x
    end_pos_y = 0
    for value in (
        int(arc_flags),
        center_x,
        center_y,
        axis1_x,
        axis1_y,
        axis2_x,
        axis2_y,
        start_pos_x,
        start_pos_y,
        end_pos_x,
        end_pos_y,
        start_pos_x,
        start_pos_y,
        end_pos_x,
        end_pos_y,
    ):
        payload.extend(int(value).to_bytes(4, "little", signed=True))
    return bytes(payload)


def _build_arc_specific_payload(width: int, height: int) -> bytes:
    payload = bytearray()
    center_x = int(width) // 2
    center_y = int(height) // 2
    axis1_x = int(width) // 2
    axis1_y = 0
    axis2_x = 0
    axis2_y = int(height) // 2
    for value in (0x00000002, center_x, center_y, axis1_x, axis1_y, axis2_x, axis2_y):
        payload.extend(int(value).to_bytes(4, "little", signed=True))
    return bytes(payload)


def _build_polygon_specific_payload(width: int, height: int) -> bytes:
    x_points = (0, int(width), int(width), 0)
    y_points = (0, 0, int(height), int(height))
    payload = bytearray()
    payload.extend(len(x_points).to_bytes(2, "little", signed=False))
    for value in x_points:
        payload.extend(int(value).to_bytes(4, "little", signed=True))
    for value in y_points:
        payload.extend(int(value).to_bytes(4, "little", signed=True))
    return bytes(payload)


def _build_textart_specific_payload(
    text: str,
    *,
    font_name: str = _DEFAULT_TEXTART_FONT_NAME,
    font_style: str = _DEFAULT_TEXTART_FONT_STYLE,
) -> bytes:
    raw_text = f"{text}\r"
    payload = bytearray(bytes.fromhex("00000000000000005d370000000000005d3700005d370000000000005d370000"))
    payload.extend(len(raw_text).to_bytes(2, "little", signed=False))
    payload.extend(raw_text.encode("utf-16-le"))
    payload.extend(len(font_name).to_bytes(2, "little", signed=False))
    payload.extend(font_name.encode("utf-16-le"))
    payload.extend(len(font_style).to_bytes(2, "little", signed=False))
    payload.extend(font_style.encode("utf-16-le"))
    payload.extend(bytes.fromhex("01000000000000007800000064000000000000000000000000000000000000000000000000000000"))
    return bytes(payload)


def _build_ole_attribute_flags(
    *,
    object_type: str = "EMBEDDED",
    draw_aspect: str = "CONTENT",
    has_moniker: bool = False,
    eq_baseline: int = 0,
) -> int:
    value = _HWP_DRAW_ASPECT_CODES.get(str(draw_aspect).upper(), 1) & 0xFF
    if has_moniker:
        value |= 1 << 8
    baseline = max(0, min(int(eq_baseline), 0x7F))
    value |= (baseline & 0x7F) << 9
    value |= (_HWP_OLE_OBJECT_TYPE_CODES.get(str(object_type).upper(), 1) & 0x3F) << 16
    return value


def _build_ole_specific_payload(
    *,
    storage_id: int,
    width: int,
    height: int,
    object_type: str = "EMBEDDED",
    draw_aspect: str = "CONTENT",
    has_moniker: bool = False,
    eq_baseline: int = 0,
    line_color: str = "#000000",
    line_width: int = _DEFAULT_OLE_LINE_WIDTH,
) -> bytes:
    payload = bytearray(24)
    payload[0:4] = _build_ole_attribute_flags(
        object_type=object_type,
        draw_aspect=draw_aspect,
        has_moniker=has_moniker,
        eq_baseline=eq_baseline,
    ).to_bytes(4, "little", signed=False)
    payload[4:8] = int(width).to_bytes(4, "little", signed=True)
    payload[8:12] = int(height).to_bytes(4, "little", signed=True)
    payload[12:14] = max(0, min(int(storage_id), 0xFFFF)).to_bytes(2, "little", signed=False)
    payload[14:18] = _build_colorref(line_color)
    payload[18:20] = max(0, min(int(line_width), 0xFFFF)).to_bytes(2, "little", signed=False)
    payload[20:24] = _build_line_attribute_flags(
        line_style="NONE" if int(line_width) <= 0 else "SOLID",
        end_cap="ROUND",
        head_style="NORMAL",
        tail_style="NORMAL",
        head_size="SMALL_SMALL",
        tail_size="SMALL_SMALL",
        head_fill=False,
        tail_fill=False,
    ).to_bytes(4, "little", signed=False)
    return bytes(payload)


def _build_shape_specific_payload(
    kind: str,
    *,
    text: str = DEFAULT_HWP_SHAPE_TEXT,
    width: int,
    height: int,
    fill_color: str = "#FFFFFF",
    line_color: str = "#000000",
) -> bytes:
    if kind == "rect":
        return _build_filled_shape_common_payload(fill_color=fill_color, line_color=line_color) + _DEFAULT_SHAPE_RECTANGLE_PAYLOAD
    if kind == "ellipse":
        return _DEFAULT_ELLIPSE_PAYLOAD
    if kind == "arc":
        return _DEFAULT_ARC_PAYLOAD
    if kind == "polygon":
        return _DEFAULT_POLYGON_PAYLOAD
    if kind == "line":
        common = _build_shape_line_info_payload(color=line_color, width=_DEFAULT_SHAPE_LINE_WIDTH)
        payload = bytearray(_DEFAULT_SHAPE_LINE_PAYLOAD)
        payload[8:12] = int(width).to_bytes(4, "little", signed=False)
        payload[12:16] = int(height).to_bytes(4, "little", signed=False)
        return common + bytes(payload)
    if kind == "textart":
        return _build_textart_specific_payload(text)
    return b""


def _build_shape_text_records(text: str) -> tuple[RecordNode, ParagraphHeaderRecord]:
    raw_text = f"{text}\r"
    paragraph_header = ParagraphHeaderRecord(
        level=3,
        char_count=0x80000000 | len(raw_text),
        control_mask=_DEFAULT_SHAPE_TEXT_PARA_CONTROL_MASK,
        para_shape_id=_DEFAULT_SHAPE_TEXT_PARA_SHAPE_ID,
        style_id=_DEFAULT_SHAPE_TEXT_STYLE_ID,
        split_flags=_DEFAULT_SHAPE_TEXT_SPLIT_FLAGS,
        trailing_payload=_DEFAULT_SHAPE_TEXT_PARA_TRAILING,
    )
    paragraph_header.add_child(ParagraphTextRecord(level=4, raw_text=raw_text))
    paragraph_header.add_child(RecordNode(tag_id=TAG_PARA_CHAR_SHAPE, level=4, payload=_DEFAULT_SHAPE_TEXT_CHAR_SHAPE))
    paragraph_header.add_child(RecordNode(tag_id=TAG_PARA_LINE_SEG, level=4, payload=_DEFAULT_SHAPE_TEXT_LINE_SEG))
    return (
        RecordNode(tag_id=TAG_LIST_HEADER, level=3, payload=_DEFAULT_SHAPE_LIST_HEADER),
        paragraph_header,
    )


def _decode_control_id_payload(payload: bytes) -> str:
    if len(payload) < 4:
        return ""
    return payload[:4][::-1].decode("latin1", errors="replace")


def _iter_paragraph_raw_text_tokens(raw_text: str) -> Iterator[tuple[str, str, str]]:
    encoded = raw_text.encode("utf-16-le")
    if not encoded:
        return
    units = list(struct.unpack("<" + "H" * (len(encoded) // 2), encoded))
    text_buffer: list[str] = []
    index = 0
    while index < len(units):
        value = units[index]
        if value >= 32:
            text_buffer.append(chr(value))
            index += 1
            continue

        if value in (9, 10, 13):
            text_buffer.append(chr(value))

        if (
            index + 7 < len(units)
            and units[index + 3] == 0
            and units[index + 4] == 0
            and units[index + 5] == 0
            and units[index + 6] == 0
        ):
            if text_buffer:
                fragment = "".join(text_buffer)
                yield ("text", "", fragment)
                text_buffer = []
            token_bytes = struct.pack("<8H", *units[index : index + 8])
            control_id = token_bytes[2:6][::-1].decode("latin1", errors="replace")
            yield ("control", control_id, token_bytes.decode("utf-16-le", errors="ignore"))
            index += 8
            continue

        index += 1

    if text_buffer:
        yield ("text", "", "".join(text_buffer))


def _is_blank_preserved_control_id(control_id: str) -> bool:
    return not (control_id.startswith("%") or control_id in _BLANK_DROP_CONTROL_IDS)


def _build_blank_paragraph_raw_text(raw_text: str) -> str:
    kept_fragments = [
        fragment
        for token_kind, control_id, fragment in _iter_paragraph_raw_text_tokens(raw_text)
        if token_kind == "control" and _is_blank_preserved_control_id(control_id)
    ]
    return "".join(kept_fragments) + "\r"


def _build_blank_section_model(model: "SectionModel", *, section_index: int | None = None) -> "SectionModel":
    target_section_index = model.section_index if section_index is None else section_index
    paragraphs = model.paragraphs()
    if not paragraphs:
        blank_model = SectionModel(section_index=target_section_index, roots=[])
        blank_model.append_paragraph("\r")
        return blank_model

    source_header = paragraphs[0].header
    blank_raw_text = _build_blank_paragraph_raw_text(paragraphs[0].raw_text)
    blank_children: list[RecordNode] = []
    text_record_added = False

    for child in source_header.children:
        if isinstance(child, ParagraphTextRecord):
            blank_children.append(
                ParagraphTextRecord(
                    level=child.level,
                    raw_text=blank_raw_text,
                    header_size=child.header_size,
                    offset=child.offset,
                )
            )
            text_record_added = True
            continue
        if child.tag_id == TAG_CTRL_HEADER:
            control_id = _decode_control_id_payload(child.payload)
            if _is_blank_preserved_control_id(control_id):
                blank_children.append(child.clone())
            continue
        blank_children.append(child.clone())

    if not text_record_added:
        blank_children.insert(0, ParagraphTextRecord(level=source_header.level + 1, raw_text=blank_raw_text))

    blank_header = ParagraphHeaderRecord(
        level=source_header.level,
        char_count=len(blank_raw_text),
        control_mask=source_header.control_mask,
        para_shape_id=source_header.para_shape_id,
        style_id=source_header.style_id,
        split_flags=source_header.split_flags,
        header_size=source_header.header_size,
        offset=source_header.offset,
        trailing_payload=source_header.trailing_payload,
    )
    blank_header.children = blank_children
    blank_header.sync_payload()
    return SectionModel(section_index=target_section_index, roots=[blank_header])


def _build_blank_paragraph_header(source_paragraph: "SectionParagraphModel") -> ParagraphHeaderRecord:
    source_header = source_paragraph.header
    blank_raw_text = _build_blank_paragraph_raw_text(source_paragraph.raw_text)
    blank_children: list[RecordNode] = []
    text_record_added = False

    for child in source_header.children:
        if isinstance(child, ParagraphTextRecord):
            blank_children.append(
                ParagraphTextRecord(
                    level=child.level,
                    raw_text=blank_raw_text,
                    header_size=child.header_size,
                    offset=child.offset,
                )
            )
            text_record_added = True
            continue
        if child.tag_id == TAG_CTRL_HEADER:
            control_id = _decode_control_id_payload(child.payload)
            if _is_blank_preserved_control_id(control_id):
                blank_children.append(child.clone())
            continue
        blank_children.append(child.clone())

    if not text_record_added:
        blank_children.insert(0, ParagraphTextRecord(level=source_header.level + 1, raw_text=blank_raw_text))

    blank_header = ParagraphHeaderRecord(
        level=source_header.level,
        char_count=len(blank_raw_text),
        control_mask=source_header.control_mask,
        para_shape_id=source_header.para_shape_id,
        style_id=source_header.style_id,
        split_flags=source_header.split_flags,
        header_size=source_header.header_size,
        offset=source_header.offset,
        trailing_payload=source_header.trailing_payload,
    )
    blank_header.children = blank_children
    blank_header.sync_payload()
    return blank_header


def _insert_blank_like_paragraph(model: "SectionModel", paragraph_index: int | None = None) -> "SectionParagraphModel":
    current_paragraphs = model.paragraphs()
    target_index = len(current_paragraphs) if paragraph_index is None or paragraph_index < 0 else paragraph_index
    if target_index < 0 or target_index > len(current_paragraphs):
        raise IndexError(f"paragraph_index {target_index} is out of range for section {model.section_index}.")

    if current_paragraphs:
        header = _build_blank_paragraph_header(current_paragraphs[0])
    else:
        header = ParagraphHeaderRecord(level=0, char_count=1)
        header.add_child(ParagraphTextRecord(level=1, raw_text="\r"))

    if target_index == len(current_paragraphs):
        model.roots.append(header)
    else:
        root_index = model._root_insert_index_for_paragraph(target_index)
        model.roots.insert(root_index, header)

    return SectionParagraphModel(section_index=model.section_index, index=target_index, header=header)


def _find_section_definition_control_node(model: "SectionModel") -> RecordNode | None:
    paragraphs = model.paragraphs()
    if not paragraphs:
        return None
    for child in paragraphs[0].header.children:
        if child.tag_id == TAG_CTRL_HEADER and _decode_control_id_payload(child.payload) == "secd":
            return child
    return None


def _find_child_record(node: RecordNode, tag_id: int) -> RecordNode | None:
    return next((child for child in node.children if child.tag_id == tag_id), None)


def _find_paragraph_control_node(model: "SectionModel", control_id: str) -> RecordNode | None:
    paragraphs = model.paragraphs()
    if not paragraphs:
        return None
    encoded = _build_control_id_payload(control_id)
    for child in paragraphs[0].header.children:
        if child.tag_id == TAG_CTRL_HEADER and child.payload[:4] == encoded:
            return child
    return None


def _insert_paragraph_control_node(model: "SectionModel", control_node: RecordNode, *, after_control_id: str | None = None) -> None:
    paragraphs = model.paragraphs()
    if not paragraphs:
        raise HwpBinaryEditError("Section does not contain a writable paragraph for inserting controls.")
    children = paragraphs[0].header.children
    insert_at = len(children)
    if after_control_id is not None:
        encoded = _build_control_id_payload(after_control_id)
        for index, child in enumerate(children):
            if child.tag_id == TAG_CTRL_HEADER and child.payload[:4] == encoded:
                insert_at = index + 1
                break
    children.insert(insert_at, control_node)


_SECTION_DEF_HIDE_HEADER_FLAG = 1 << 0
_SECTION_DEF_HIDE_FOOTER_FLAG = 1 << 1
_SECTION_DEF_HIDE_MASTER_PAGE_FLAG = 1 << 2
_SECTION_DEF_HIDE_PAGE_NUM_FLAG = 1 << 5
_SECTION_DEF_HIDE_EMPTY_LINE_FLAG = 1 << 19
_SECTION_DEF_WONGGOJI_FORMAT_FLAG = 1 << 22
_SECTION_PAGE_STARTS_ON_MASK = 0x3
_SECTION_PAGE_STARTS_ON_CODES = {
    "BOTH": 0,
    "EVEN": 1,
    "ODD": 2,
}
_SECTION_PAGE_STARTS_ON_VALUES = {value: key for key, value in _SECTION_PAGE_STARTS_ON_CODES.items()}
_PAGE_HIDING_FLAGS = {
    "hideFirstHeader": 0x01,
    "hideFirstFooter": 0x02,
    "hideFirstMasterPage": 0x04,
    "hideBorder": 0x08,
    "hideFill": 0x10,
    "hideFirstPageNum": 0x20,
}
_SECTION_DEF_BORDER_FIRST_FLAG = 1 << 8
_SECTION_DEF_FILL_FIRST_FLAG = 1 << 9
_PAGE_NUM_POSITION_CODES = {
    "NONE": 0,
    "TOP_LEFT": 1,
    "TOP_CENTER": 2,
    "TOP_RIGHT": 3,
    "BOTTOM_LEFT": 4,
    "BOTTOM_CENTER": 5,
    "BOTTOM_RIGHT": 6,
    "OUTSIDE_TOP": 7,
    "OUTSIDE_BOTTOM": 8,
    "INSIDE_TOP": 9,
    "INSIDE_BOTTOM": 10,
}
_PAGE_NUM_POSITION_VALUES = {value: key for key, value in _PAGE_NUM_POSITION_CODES.items()}
_NUMBER_TYPE_CODES = {
    "DIGIT": 0,
    "CIRCLED_DIGIT": 1,
    "ROMAN_CAPITAL": 2,
    "ROMAN_SMALL": 3,
    "LATIN_CAPITAL": 4,
    "LATIN_SMALL": 5,
    "CIRCLED_LATIN_CAPITAL": 6,
    "CIRCLED_LATIN_SMALL": 7,
    "HANGUL_SYLLABLE": 8,
    "CIRCLED_HANGUL_SYLLABLE": 9,
    "HANGUL_JAMO": 10,
    "CIRCLED_HANGUL_JAMO": 11,
    "HANGUL_PHONETIC": 12,
    "IDEOGRAPH": 13,
    "CIRCLED_IDEOGRAPH": 14,
    "DECAGON_CIRCLE_HANGUL": 15,
    "DECAGON_CIRCLE_HANJA": 16,
}
_NUMBER_TYPE_VALUES = {value: key for key, value in _NUMBER_TYPE_CODES.items()}
_NOTE_PLACEMENT_CODES = {
    "EACH_COLUMN": 0,
    "MERGED_COLUMN": 1,
    "RIGHT_MOST_COLUMN": 2,
    "END_OF_DOCUMENT": 0,
    "END_OF_SECTION": 1,
}
_NOTE_NUMBERING_CODES = {
    "CONTINUOUS": 0,
    "ON_SECTION": 1,
    "ON_PAGE": 2,
}
_NOTE_NUMBERING_VALUES = {value: key for key, value in _NOTE_NUMBERING_CODES.items()}
_NOTE_LINE_TYPE_CODES = {
    "SOLID": 1,
    "DOT": 2,
    "DASH": 3,
}
_NOTE_LINE_TYPE_VALUES = {value: key for key, value in _NOTE_LINE_TYPE_CODES.items()}
_NOTE_LINE_WIDTH_CODES = {
    "0.1 mm": 0,
    "0.12 mm": 1,
    "0.15 mm": 2,
    "0.2 mm": 3,
    "0.25 mm": 4,
    "0.3 mm": 5,
    "0.4 mm": 6,
    "0.5 mm": 7,
    "0.6 mm": 8,
    "0.7 mm": 9,
    "1.0 mm": 10,
    "1.5 mm": 11,
    "2.0 mm": 12,
    "3.0 mm": 13,
    "4.0 mm": 14,
    "5.0 mm": 15,
}
_NOTE_LINE_WIDTH_VALUES = {value: key for key, value in _NOTE_LINE_WIDTH_CODES.items()}


def _parse_section_definition_payload(payload: bytes) -> dict[str, object]:
    body = payload[4:].ljust(26, b"\x00")
    attributes = int.from_bytes(body[0:4], "little")
    return {
        "attributes": attributes,
        "space_columns": int.from_bytes(body[4:6], "little"),
        "line_grid": int.from_bytes(body[6:8], "little"),
        "char_grid": int.from_bytes(body[8:10], "little"),
        "tab_stop": int.from_bytes(body[10:14], "little"),
        "numbering_shape_id": int.from_bytes(body[14:16], "little"),
        "page": int.from_bytes(body[16:18], "little"),
        "pic": int.from_bytes(body[18:20], "little"),
        "tbl": int.from_bytes(body[20:22], "little"),
        "equation": int.from_bytes(body[22:24], "little"),
        "language": int.from_bytes(body[24:26], "little"),
    }


def _build_section_definition_payload(
    payload: bytes,
    *,
    visibility: dict[str, str] | None = None,
    grid: dict[str, int] | None = None,
    start_numbers: dict[str, str] | None = None,
    numbering_shape_id: int | None = None,
) -> bytes:
    updated = bytearray(payload[:30].ljust(30, b"\x00"))
    attributes = int.from_bytes(updated[4:8], "little")
    if visibility is not None:
        visibility_flags = {
            "hideFirstHeader": _SECTION_DEF_HIDE_HEADER_FLAG,
            "hideFirstFooter": _SECTION_DEF_HIDE_FOOTER_FLAG,
            "hideFirstMasterPage": _SECTION_DEF_HIDE_MASTER_PAGE_FLAG,
            "hideFirstPageNum": _SECTION_DEF_HIDE_PAGE_NUM_FLAG,
            "hideFirstEmptyLine": _SECTION_DEF_HIDE_EMPTY_LINE_FLAG,
        }
        for key, flag in visibility_flags.items():
            if key not in visibility or visibility[key] is None:
                continue
            if str(visibility[key]) in {"1", "true", "True"}:
                attributes |= flag
            else:
                attributes &= ~flag
    if grid is not None:
        if "lineGrid" in grid and grid["lineGrid"] is not None:
            updated[10:12] = int(grid["lineGrid"]).to_bytes(2, "little", signed=False)
        if "charGrid" in grid and grid["charGrid"] is not None:
            updated[12:14] = int(grid["charGrid"]).to_bytes(2, "little", signed=False)
        if "wonggojiFormat" in grid and grid["wonggojiFormat"] is not None:
            if int(grid["wonggojiFormat"]):
                attributes |= _SECTION_DEF_WONGGOJI_FORMAT_FLAG
            else:
                attributes &= ~_SECTION_DEF_WONGGOJI_FORMAT_FLAG
    if numbering_shape_id is not None:
        updated[18:20] = int(numbering_shape_id).to_bytes(2, "little", signed=False)
    if start_numbers is not None:
        for key, offset in {"page": 20, "pic": 22, "tbl": 24, "equation": 26}.items():
            if key in start_numbers and start_numbers[key] is not None:
                updated[offset : offset + 2] = int(start_numbers[key]).to_bytes(2, "little", signed=False)
    updated[4:8] = attributes.to_bytes(4, "little", signed=False)
    return bytes(updated) + payload[30:]


def _parse_page_starts_on_payload(payload: bytes) -> str:
    attributes = int.from_bytes(payload[4:8].ljust(4, b"\x00"), "little")
    return _SECTION_PAGE_STARTS_ON_VALUES.get(attributes & _SECTION_PAGE_STARTS_ON_MASK, "BOTH")


def _build_page_starts_on_payload(page_starts_on: str) -> bytes:
    code = _SECTION_PAGE_STARTS_ON_CODES.get(str(page_starts_on).upper(), 0)
    return _build_control_id_payload("pgct") + code.to_bytes(4, "little", signed=False)


def _parse_page_hiding_payload(payload: bytes) -> dict[str, str]:
    attributes = int.from_bytes(payload[4:8].ljust(4, b"\x00"), "little")
    return {
        "hideFirstHeader": "1" if attributes & _PAGE_HIDING_FLAGS["hideFirstHeader"] else "0",
        "hideFirstFooter": "1" if attributes & _PAGE_HIDING_FLAGS["hideFirstFooter"] else "0",
        "hideFirstMasterPage": "1" if attributes & _PAGE_HIDING_FLAGS["hideFirstMasterPage"] else "0",
        "hideFirstPageNum": "1" if attributes & _PAGE_HIDING_FLAGS["hideFirstPageNum"] else "0",
        "hideBorder": "1" if attributes & _PAGE_HIDING_FLAGS["hideBorder"] else "0",
        "hideFill": "1" if attributes & _PAGE_HIDING_FLAGS["hideFill"] else "0",
    }


def _build_page_hiding_payload(visibility: dict[str, str]) -> bytes:
    attributes = 0
    for key, flag in _PAGE_HIDING_FLAGS.items():
        value = visibility.get(key)
        if value is not None and str(value) in {"1", "true", "True"}:
            attributes |= flag
    return _build_control_id_payload("pghd") + attributes.to_bytes(4, "little", signed=False)


def _parse_page_num_payload(payload: bytes) -> dict[str, str]:
    attributes = int.from_bytes(payload[4:8].ljust(4, b"\x00"), "little")
    side_char = payload[14:16].decode("utf-16-le", errors="ignore") if len(payload) >= 16 else "-"
    return {
        "pos": _PAGE_NUM_POSITION_VALUES.get((attributes >> 8) & 0x0F, "BOTTOM_CENTER"),
        "formatType": _NUMBER_TYPE_VALUES.get(attributes & 0xFF, "DIGIT"),
        "sideChar": side_char or "-",
    }


def _build_page_num_payload(page_num: dict[str, str]) -> bytes:
    format_code = _NUMBER_TYPE_CODES.get(str(page_num.get("formatType", "DIGIT")).upper(), 0)
    position_code = _PAGE_NUM_POSITION_CODES.get(str(page_num.get("pos", "BOTTOM_CENTER")).upper(), 5)
    attributes = format_code | (position_code << 8)
    side_char = str(page_num.get("sideChar", "-"))[:1]
    return (
        _build_control_id_payload("pgnp")
        + attributes.to_bytes(4, "little", signed=False)
        + ("\x00\x00\x00" + side_char).encode("utf-16-le")
    )


def _parse_note_shape_payload(payload: bytes, *, kind: str) -> dict[str, object]:
    values = payload[:26].ljust(26, b"\x00")
    attributes = int.from_bytes(values[0:4], "little")
    line_type_code = values[20] if len(values) > 20 else 1
    line_width_code = values[21] if len(values) > 21 else 1
    if kind == "endNotePr":
        place = {0: "END_OF_DOCUMENT", 1: "END_OF_SECTION"}.get((attributes >> 8) & 0x03, "END_OF_DOCUMENT")
        length = str(int.from_bytes(values[12:14], "little", signed=False))
    else:
        place = {0: "EACH_COLUMN", 1: "MERGED_COLUMN", 2: "RIGHT_MOST_COLUMN"}.get((attributes >> 8) & 0x03, "EACH_COLUMN")
        length = str(int.from_bytes(values[12:14], "little", signed=True))
    return {
        "autoNumFormat": {
            "type": _NUMBER_TYPE_VALUES.get(attributes & 0xFF, "DIGIT"),
            "userChar": values[4:6].decode("utf-16-le", errors="ignore").rstrip("\x00"),
            "prefixChar": values[6:8].decode("utf-16-le", errors="ignore").rstrip("\x00"),
            "suffixChar": values[8:10].decode("utf-16-le", errors="ignore").rstrip("\x00"),
            "supscript": "1" if attributes & (1 << 12) else "0",
        },
        "numbering": {
            "type": _NOTE_NUMBERING_VALUES.get((attributes >> 10) & 0x03, "CONTINUOUS"),
            "newNum": str(int.from_bytes(values[10:12], "little")),
        },
        "noteLine": {
            "length": length,
            "type": _NOTE_LINE_TYPE_VALUES.get(line_type_code, "SOLID"),
            "width": _NOTE_LINE_WIDTH_VALUES.get(line_width_code, "0.12 mm"),
            "color": _parse_colorref(values[22:26]),
        },
        "noteSpacing": {
            "aboveLine": str(int.from_bytes(values[14:16], "little")),
            "belowLine": str(int.from_bytes(values[16:18], "little")),
            "betweenNotes": str(int.from_bytes(values[18:20], "little")),
        },
        "placement": {
            "place": place,
            "beneathText": "1" if attributes & (1 << 13) else "0",
        },
    }


def _build_note_shape_payload(payload: bytes, *, note_pr: dict[str, object], kind: str) -> bytes:
    original = payload[:26].ljust(26, b"\x00")
    attributes = int.from_bytes(original[0:4], "little")
    auto_num = dict(note_pr.get("autoNumFormat", {}))
    numbering = dict(note_pr.get("numbering", {}))
    note_line = dict(note_pr.get("noteLine", {}))
    note_spacing = dict(note_pr.get("noteSpacing", {}))
    placement = dict(note_pr.get("placement", {}))
    attributes &= ~0xFF
    attributes |= _NUMBER_TYPE_CODES.get(str(auto_num.get("type", "DIGIT")).upper(), 0)
    attributes &= ~(0x03 << 8)
    if kind == "endNotePr":
        attributes |= ({ "END_OF_DOCUMENT": 0, "END_OF_SECTION": 1 }.get(str(placement.get("place", "END_OF_DOCUMENT")).upper(), 0) << 8)
    else:
        attributes |= ({ "EACH_COLUMN": 0, "MERGED_COLUMN": 1, "RIGHT_MOST_COLUMN": 2 }.get(str(placement.get("place", "EACH_COLUMN")).upper(), 0) << 8)
    attributes &= ~(0x03 << 10)
    attributes |= (_NOTE_NUMBERING_CODES.get(str(numbering.get("type", "CONTINUOUS")).upper(), 0) << 10)
    if str(auto_num.get("supscript", "0")) in {"1", "true", "True"}:
        attributes |= 1 << 12
    else:
        attributes &= ~(1 << 12)
    if str(placement.get("beneathText", "0")) in {"1", "true", "True"}:
        attributes |= 1 << 13
    else:
        attributes &= ~(1 << 13)
    updated = bytearray(26)
    updated[0:4] = attributes.to_bytes(4, "little", signed=False)
    updated[4:6] = str(auto_num.get("userChar", ""))[:1].encode("utf-16-le")
    updated[6:8] = str(auto_num.get("prefixChar", ""))[:1].encode("utf-16-le")
    updated[8:10] = str(auto_num.get("suffixChar", ""))[:1].encode("utf-16-le")
    updated[10:12] = int(numbering.get("newNum", 1)).to_bytes(2, "little", signed=False)
    length_value = int(str(note_line.get("length", -1)))
    if kind == "endNotePr":
        updated[12:14] = (length_value & 0xFFFF).to_bytes(2, "little", signed=False)
    else:
        updated[12:14] = length_value.to_bytes(2, "little", signed=True)
    updated[14:16] = int(note_spacing.get("aboveLine", 850)).to_bytes(2, "little", signed=False)
    updated[16:18] = int(note_spacing.get("belowLine", 567)).to_bytes(2, "little", signed=False)
    updated[18:20] = int(note_spacing.get("betweenNotes", 283)).to_bytes(2, "little", signed=False)
    updated[20] = _NOTE_LINE_TYPE_CODES.get(str(note_line.get("type", "SOLID")).upper(), 1)
    updated[21] = _NOTE_LINE_WIDTH_CODES.get(str(note_line.get("width", "0.12 mm")), 1)
    updated[22:26] = _build_colorref(str(note_line.get("color", "#000000")))
    return bytes(updated) + payload[26:]


def _parse_colorref(payload: bytes) -> str:
    values = payload[:4].ljust(4, b"\x00")
    blue = values[0]
    green = values[1]
    red = values[2]
    return f"#{red:02X}{green:02X}{blue:02X}"


def _build_colorref(value: str) -> bytes:
    normalized = str(value).strip().lstrip("#")
    if len(normalized) != 6:
        normalized = "000000"
    red = int(normalized[0:2], 16)
    green = int(normalized[2:4], 16)
    blue = int(normalized[4:6], 16)
    return bytes((blue, green, red, 0))


def _parse_page_def_payload(payload: bytes) -> dict[str, int]:
    values = list(struct.unpack("<10I", payload[:40].ljust(40, b"\x00")))
    return {
        "page_width": values[0],
        "page_height": values[1],
        "left": values[2],
        "right": values[3],
        "top": values[4],
        "header": values[5],
        "footer": values[6],
        "bottom": values[7],
        "gutter": values[8],
    }


def _build_page_def_payload(
    payload: bytes,
    *,
    page_width: int | None = None,
    page_height: int | None = None,
    landscape: str | None = None,
    margins: dict[str, int] | None = None,
) -> bytes:
    values = list(struct.unpack("<10I", payload[:40].ljust(40, b"\x00")))
    width = values[0] if page_width is None else int(page_width)
    height = values[1] if page_height is None else int(page_height)
    if landscape is not None and page_width is None and page_height is None:
        normalized_landscape = str(landscape).upper()
        if normalized_landscape == "LANDSCAPE" and width < height:
            width, height = height, width
        elif normalized_landscape in {"WIDELY", "PORTRAIT"} and width > height:
            width, height = height, width
    values[0] = width
    values[1] = height
    margin_mapping = {
        "left": 2,
        "right": 3,
        "top": 4,
        "header": 5,
        "footer": 6,
        "bottom": 7,
        "gutter": 8,
    }
    for key, index in margin_mapping.items():
        if margins is not None and key in margins and margins[key] is not None:
            values[index] = int(margins[key])
    updated = struct.pack("<10I", *values)
    return updated + payload[40:]


_SECTION_PAGE_BORDER_FILL_TYPES = ("BOTH", "EVEN", "ODD")
_PAGE_BORDER_FILL_TEXT_BORDER_FLAG = 0x01
_PAGE_BORDER_FILL_HEADER_INSIDE_FLAG = 0x02
_PAGE_BORDER_FILL_FOOTER_INSIDE_FLAG = 0x04
_PAGE_BORDER_FILL_BORDER_AREA_FLAG = 0x10


def _parse_page_border_fill_payload(payload: bytes, *, border_type: str) -> dict[str, str | int]:
    values = payload[:14].ljust(14, b"\x00")
    flags, left, right, top, bottom, border_fill_id = struct.unpack("<I5H", values)
    return {
        "type": border_type,
        "borderFillIDRef": str(border_fill_id),
        "textBorder": "PAPER" if flags & _PAGE_BORDER_FILL_TEXT_BORDER_FLAG else "BORDER",
        "headerInside": "1" if flags & _PAGE_BORDER_FILL_HEADER_INSIDE_FLAG else "0",
        "footerInside": "1" if flags & _PAGE_BORDER_FILL_FOOTER_INSIDE_FLAG else "0",
        "fillArea": "BORDER" if flags & _PAGE_BORDER_FILL_BORDER_AREA_FLAG else "PAPER",
        "left": left,
        "right": right,
        "top": top,
        "bottom": bottom,
    }


def _build_page_border_fill_payload(
    payload: bytes,
    *,
    border_fill: dict[str, str | int],
) -> bytes:
    values = payload[:14].ljust(14, b"\x00")
    flags, left, right, top, bottom, border_fill_id = struct.unpack("<I5H", values)
    if str(border_fill.get("textBorder", "PAPER")).upper() == "PAPER":
        flags |= _PAGE_BORDER_FILL_TEXT_BORDER_FLAG
    else:
        flags &= ~_PAGE_BORDER_FILL_TEXT_BORDER_FLAG
    if str(border_fill.get("headerInside", "0")) in {"1", "true", "True"}:
        flags |= _PAGE_BORDER_FILL_HEADER_INSIDE_FLAG
    else:
        flags &= ~_PAGE_BORDER_FILL_HEADER_INSIDE_FLAG
    if str(border_fill.get("footerInside", "0")) in {"1", "true", "True"}:
        flags |= _PAGE_BORDER_FILL_FOOTER_INSIDE_FLAG
    else:
        flags &= ~_PAGE_BORDER_FILL_FOOTER_INSIDE_FLAG
    if str(border_fill.get("fillArea", "PAPER")).upper() == "BORDER":
        flags |= _PAGE_BORDER_FILL_BORDER_AREA_FLAG
    else:
        flags &= ~_PAGE_BORDER_FILL_BORDER_AREA_FLAG
    left = int(border_fill.get("left", left))
    right = int(border_fill.get("right", right))
    top = int(border_fill.get("top", top))
    bottom = int(border_fill.get("bottom", bottom))
    border_fill_id = int(border_fill.get("borderFillIDRef", border_fill_id))
    updated = struct.pack("<I5H", flags, left, right, top, bottom, border_fill_id)
    return updated + payload[14:]


def _build_control_marker_text(control_id: str) -> str:
    marker_code = _CONTROL_MARKER_CODES.get(control_id)
    if marker_code is None:
        raise ValueError(f"Unsupported control marker id: {control_id}")
    control_payload = _build_control_id_payload(control_id)
    first_unit = int.from_bytes(control_payload[0:2], "little")
    second_unit = int.from_bytes(control_payload[2:4], "little")
    units = (marker_code, first_unit, second_unit, 0, 0, 0, 0, marker_code)
    return struct.pack("<8H", *units).decode("utf-16-le", errors="ignore")


def _build_auto_number_payload(kind: str, *, number: int | str, number_type: str) -> bytes:
    if kind == "newNum":
        payload = bytearray(_build_control_id_payload("nwno"))
        payload.extend((0).to_bytes(4, "little"))
        payload.extend(int(number).to_bytes(2, "little", signed=False))
        return bytes(payload)
    if kind == "autoNum":
        payload = bytearray(_build_control_id_payload("atno"))
        payload.extend((0).to_bytes(4, "little"))
        payload.extend(_AUTO_NUMBER_TYPE_CODES.get(str(number_type).upper(), 0x0C).to_bytes(4, "little"))
        payload.extend((0).to_bytes(4, "little"))
        return bytes(payload)
    raise ValueError("append_auto_number() kind must be 'newNum' or 'autoNum'.")


def _build_bookmark_ctrl_data_payload(name: str) -> bytes:
    encoded_name = name.encode("utf-16-le")
    payload = bytearray(_BOOKMARK_CTRL_DATA_PREFIX)
    payload.extend(len(name).to_bytes(2, "little", signed=False))
    payload.extend(encoded_name)
    return bytes(payload)


def _build_note_ctrl_payload(control_id: str) -> bytes:
    if control_id == "fn  ":
        return _FOOTNOTE_CTRL_HEADER
    if control_id == "en  ":
        return _ENDNOTE_CTRL_HEADER
    raise ValueError(f"Unsupported note control id: {control_id}")


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
            root_index = self._root_insert_index_for_paragraph(paragraph_index)
            self.roots.insert(root_index, header)
        return SectionParagraphModel(
            section_index=self.section_index,
            index=paragraph_index,
            header=header,
        )

    def _root_insert_index_for_paragraph(self, paragraph_index: int) -> int:
        seen = 0
        for root_index, root in enumerate(self.roots):
            paragraph_count = sum(1 for node in root.iter_preorder() if isinstance(node, ParagraphHeaderRecord))
            if seen + paragraph_count > paragraph_index:
                return root_index
            seen += paragraph_count
        return len(self.roots)

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

    def ensure_section_count(self, section_count: int) -> None:
        if section_count < 1:
            raise ValueError("section_count must be at least 1.")

        current_count = len(self.section_stream_paths())
        if current_count == 0:
            raise InvalidHwpFileError("HWP document does not contain any BodyText/Section streams.")

        while current_count < section_count:
            template_index = 1 if current_count > 1 else current_count - 1
            template_model = _build_blank_section_model(self.section_model(template_index), section_index=current_count)
            raw = template_model.to_bytes()
            if self._stream_is_compressed("BodyText/Section0"):
                compressor = zlib.compressobj(level=9, wbits=-15)
                payload = compressor.compress(raw) + compressor.flush()
            else:
                payload = raw
            self.add_stream(f"BodyText/Section{current_count}", payload)
            current_count += 1

        self._set_document_properties_section_count(current_count)

    def reset_body_sections_to_blank(self) -> None:
        section_paths = self.section_stream_paths()
        for section_index, _path in enumerate(section_paths):
            self.replace_section_model(section_index, _build_blank_section_model(self.section_model(section_index)))
        self._set_document_properties_section_count(len(section_paths))
        self.set_preview_text("")

    def section_page_settings(self, section_index: int) -> dict[str, int]:
        self.ensure_section_count(section_index + 1)
        section_definition = _find_section_definition_control_node(self.section_model(section_index))
        if section_definition is None:
            return {}
        page_def = _find_child_record(section_definition, TAG_PAGE_DEF)
        if page_def is None:
            return {}
        return _parse_page_def_payload(page_def.payload)

    def set_section_page_settings(
        self,
        section_index: int,
        *,
        page_width: int | None = None,
        page_height: int | None = None,
        landscape: str | None = None,
        margins: dict[str, int] | None = None,
    ) -> None:
        self.ensure_section_count(section_index + 1)
        model = self.section_model(section_index)
        section_definition = _find_section_definition_control_node(model)
        if section_definition is None:
            raise HwpBinaryEditError(f"Section {section_index} does not contain a writable secd control.")
        page_def = _find_child_record(section_definition, TAG_PAGE_DEF)
        if page_def is None:
            raise HwpBinaryEditError(f"Section {section_index} does not contain a writable page_def record.")
        page_def.payload = _build_page_def_payload(
            page_def.payload,
            page_width=page_width,
            page_height=page_height,
            landscape=landscape,
            margins=margins,
        )
        self.replace_section_model(section_index, model)

    def section_page_border_fills(self, section_index: int) -> list[dict[str, str | int]]:
        self.ensure_section_count(section_index + 1)
        section_definition = _find_section_definition_control_node(self.section_model(section_index))
        if section_definition is None:
            return []
        records = [child for child in section_definition.children if child.tag_id == TAG_PAGE_BORDER_FILL]
        fills: list[dict[str, str | int]] = []
        for index, record in enumerate(records):
            border_type = _SECTION_PAGE_BORDER_FILL_TYPES[index] if index < len(_SECTION_PAGE_BORDER_FILL_TYPES) else f"EXTRA_{index}"
            fills.append(_parse_page_border_fill_payload(record.payload, border_type=border_type))
        return fills

    def set_section_page_border_fills(
        self,
        section_index: int,
        page_border_fills: list[dict[str, str | int]],
    ) -> None:
        self.ensure_section_count(section_index + 1)
        model = self.section_model(section_index)
        section_definition = _find_section_definition_control_node(model)
        if section_definition is None:
            raise HwpBinaryEditError(f"Section {section_index} does not contain a writable secd control.")
        records = [child for child in section_definition.children if child.tag_id == TAG_PAGE_BORDER_FILL]
        if not records:
            raise HwpBinaryEditError(f"Section {section_index} does not contain writable page_border_fill records.")
        record_by_type = {
            _SECTION_PAGE_BORDER_FILL_TYPES[index]: record
            for index, record in enumerate(records[: len(_SECTION_PAGE_BORDER_FILL_TYPES)])
        }
        for border_fill in page_border_fills:
            border_type = str(border_fill.get("type", "BOTH")).upper()
            record = record_by_type.get(border_type)
            if record is None:
                continue
            record.payload = _build_page_border_fill_payload(record.payload, border_fill=border_fill)
        self.replace_section_model(section_index, model)

    def section_definition_settings(self, section_index: int) -> dict[str, object]:
        self.ensure_section_count(section_index + 1)
        model = self.section_model(section_index)
        section_definition = _find_section_definition_control_node(model)
        if section_definition is None:
            return {}
        values = _parse_section_definition_payload(section_definition.payload)
        page_hiding = _parse_page_hiding_payload(_find_paragraph_control_node(model, "pghd").payload) if _find_paragraph_control_node(model, "pghd") is not None else {}
        page_starts_on = "BOTH"
        page_starts_on_control = _find_paragraph_control_node(model, "pgct")
        if page_starts_on_control is not None:
            page_starts_on = _parse_page_starts_on_payload(page_starts_on_control.payload)
        attributes = int(values.get("attributes", 0))
        border_hidden = page_hiding.get("hideBorder") == "1"
        fill_hidden = page_hiding.get("hideFill") == "1"
        border_first = bool(attributes & _SECTION_DEF_BORDER_FIRST_FLAG)
        fill_first = bool(attributes & _SECTION_DEF_FILL_FIRST_FLAG)
        return {
            "visibility": {
                "hideFirstHeader": page_hiding.get("hideFirstHeader", "0"),
                "hideFirstFooter": page_hiding.get("hideFirstFooter", "0"),
                "hideFirstMasterPage": page_hiding.get("hideFirstMasterPage", "0"),
                "hideFirstPageNum": page_hiding.get("hideFirstPageNum", "0"),
                "hideFirstEmptyLine": "1" if attributes & _SECTION_DEF_HIDE_EMPTY_LINE_FLAG else "0",
                "border": "HIDE_FIRST" if border_hidden and border_first else "HIDE_ALL" if border_hidden else "SHOW_FIRST" if border_first else "SHOW_ALL",
                "fill": "HIDE_FIRST" if fill_hidden and fill_first else "HIDE_ALL" if fill_hidden else "SHOW_FIRST" if fill_first else "SHOW_ALL",
                "showLineNumber": "1" if _section_uses_synthetic_line_number_layout(model) else "0",
            },
            "grid": {
                "lineGrid": int(values.get("line_grid", 0)),
                "charGrid": int(values.get("char_grid", 0)),
                "wonggojiFormat": 1 if attributes & _SECTION_DEF_WONGGOJI_FORMAT_FLAG else 0,
            },
            "start_numbers": {
                "pageStartsOn": page_starts_on,
                "page": str(values.get("page", 0)),
                "pic": str(values.get("pic", 0)),
                "tbl": str(values.get("tbl", 0)),
                "equation": str(values.get("equation", 0)),
            },
            "numbering_shape_id": str(values.get("numbering_shape_id", 0)),
        }

    def set_section_definition_settings(
        self,
        section_index: int,
        *,
        visibility: dict[str, str] | None = None,
        grid: dict[str, int] | None = None,
        start_numbers: dict[str, str] | None = None,
        numbering_shape_id: int | None = None,
    ) -> None:
        self.ensure_section_count(section_index + 1)
        model = self.section_model(section_index)
        section_definition = _find_section_definition_control_node(model)
        if section_definition is None:
            raise HwpBinaryEditError(f"Section {section_index} does not contain a writable secd control.")
        section_definition.payload = _build_section_definition_payload(
            section_definition.payload,
            visibility=visibility,
            grid=grid,
            start_numbers=start_numbers,
            numbering_shape_id=numbering_shape_id,
        )
        if visibility is not None:
            current = self.section_definition_settings(section_index)["visibility"]
            current.update({key: value for key, value in visibility.items() if value is not None})
            border_value = str(current.get("border", "SHOW_ALL")).upper()
            fill_value = str(current.get("fill", "SHOW_ALL")).upper()
            payload = bytearray(section_definition.payload)
            attributes = int.from_bytes(payload[4:8], "little")
            if border_value in {"SHOW_FIRST", "HIDE_FIRST"}:
                attributes |= _SECTION_DEF_BORDER_FIRST_FLAG
            else:
                attributes &= ~_SECTION_DEF_BORDER_FIRST_FLAG
            if fill_value in {"SHOW_FIRST", "HIDE_FIRST"}:
                attributes |= _SECTION_DEF_FILL_FIRST_FLAG
            else:
                attributes &= ~_SECTION_DEF_FILL_FIRST_FLAG
            payload[4:8] = attributes.to_bytes(4, "little", signed=False)
            section_definition.payload = bytes(payload)
            page_hiding_visibility = {
                "hideFirstHeader": current.get("hideFirstHeader", "0"),
                "hideFirstFooter": current.get("hideFirstFooter", "0"),
                "hideFirstMasterPage": current.get("hideFirstMasterPage", "0"),
                "hideFirstPageNum": current.get("hideFirstPageNum", "0"),
                "hideBorder": "1" if border_value in {"HIDE_FIRST", "HIDE_ALL"} else "0",
                "hideFill": "1" if fill_value in {"HIDE_FIRST", "HIDE_ALL"} else "0",
            }
            page_hiding_control = _find_paragraph_control_node(model, "pghd")
            if any(value == "1" for value in page_hiding_visibility.values()):
                if page_hiding_control is None:
                    page_hiding_control = RecordNode(tag_id=TAG_CTRL_HEADER, level=1, payload=_build_page_hiding_payload(page_hiding_visibility))
                    _insert_paragraph_control_node(model, page_hiding_control, after_control_id="cold")
                else:
                    page_hiding_control.payload = _build_page_hiding_payload(page_hiding_visibility)
            elif page_hiding_control is not None:
                for paragraph in model.paragraphs():
                    if page_hiding_control in paragraph.header.children:
                        paragraph.header.children.remove(page_hiding_control)
                        break
            if visibility.get("showLineNumber") is not None:
                _apply_section_show_line_numbers(model, str(current.get("showLineNumber", "0")) in {"1", "true", "True"})
        if start_numbers is not None and start_numbers.get("pageStartsOn") is not None:
            page_starts_on = str(start_numbers.get("pageStartsOn", "BOTH")).upper()
            page_starts_on_control = _find_paragraph_control_node(model, "pgct")
            if page_starts_on == "BOTH":
                if page_starts_on_control is not None:
                    model.paragraphs()[0].header.children.remove(page_starts_on_control)
            else:
                if page_starts_on_control is None:
                    page_starts_on_control = RecordNode(
                        tag_id=TAG_CTRL_HEADER,
                        level=1,
                        payload=_build_page_starts_on_payload(page_starts_on),
                    )
                    _insert_paragraph_control_node(model, page_starts_on_control, after_control_id="secd")
                else:
                    page_starts_on_control.payload = _build_page_starts_on_payload(page_starts_on)
        self.replace_section_model(section_index, model)

    def section_page_numbers(self, section_index: int) -> list[dict[str, str]]:
        self.ensure_section_count(section_index + 1)
        model = self.section_model(section_index)
        page_numbers: list[dict[str, str]] = []
        for paragraph in model.paragraphs():
            for child in paragraph.header.children:
                if child.tag_id == TAG_CTRL_HEADER and child.payload[:4] == _build_control_id_payload("pgnp"):
                    page_numbers.append(_parse_page_num_payload(child.payload))
        return page_numbers

    def set_section_page_numbers(self, section_index: int, page_numbers: list[dict[str, str]]) -> None:
        self.ensure_section_count(section_index + 1)
        model = self.section_model(section_index)
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(model)
        for paragraph in model.paragraphs():
            paragraph.header.children = [
                child
                for child in paragraph.header.children
                if not (child.tag_id == TAG_CTRL_HEADER and child.payload[:4] == _build_control_id_payload("pgnp"))
            ]
        for page_num in page_numbers:
            raw_text = _build_control_marker_text("pgnp") + "\r"
            paragraph = model.append_paragraph(raw_text)
            if use_synthetic_line_numbers:
                _ensure_paragraph_line_seg(paragraph)
            paragraph.header.add_child(
                RecordNode(tag_id=TAG_CTRL_HEADER, level=1, payload=_build_page_num_payload(page_num))
            )
        self.replace_section_model(section_index, model)

    def section_note_settings(self, section_index: int) -> dict[str, dict[str, object]]:
        self.ensure_section_count(section_index + 1)
        section_definition = _find_section_definition_control_node(self.section_model(section_index))
        if section_definition is None:
            return {}
        records = [child for child in section_definition.children if child.tag_id == TAG_FOOTNOTE_SHAPE]
        settings: dict[str, dict[str, object]] = {}
        if len(records) >= 1:
            settings["footNotePr"] = _parse_note_shape_payload(records[0].payload, kind="footNotePr")
        if len(records) >= 2:
            settings["endNotePr"] = _parse_note_shape_payload(records[1].payload, kind="endNotePr")
        return settings

    def set_section_note_settings(
        self,
        section_index: int,
        *,
        footnote_pr: dict[str, object] | None = None,
        endnote_pr: dict[str, object] | None = None,
    ) -> None:
        self.ensure_section_count(section_index + 1)
        model = self.section_model(section_index)
        section_definition = _find_section_definition_control_node(model)
        if section_definition is None:
            raise HwpBinaryEditError(f"Section {section_index} does not contain a writable secd control.")
        records = [child for child in section_definition.children if child.tag_id == TAG_FOOTNOTE_SHAPE]
        if len(records) >= 1 and footnote_pr is not None:
            records[0].payload = _build_note_shape_payload(records[0].payload, note_pr=footnote_pr, kind="footNotePr")
        if len(records) >= 2 and endnote_pr is not None:
            records[1].payload = _build_note_shape_payload(records[1].payload, note_pr=endnote_pr, kind="endNotePr")
        self.replace_section_model(section_index, model)

    def _set_document_properties_section_count(self, section_count: int) -> None:
        model = self.docinfo_model()
        properties_record = model.document_properties_record()
        properties = properties_record.to_properties()
        properties_record.properties = HwpDocumentProperties(
            section_count=section_count,
            page_start_number=properties.page_start_number,
            footnote_start_number=properties.footnote_start_number,
            endnote_start_number=properties.endnote_start_number,
            picture_start_number=properties.picture_start_number,
            table_start_number=properties.table_start_number,
            equation_start_number=properties.equation_start_number,
            list_id=properties.list_id,
            paragraph_id=properties.paragraph_id,
            character_unit_position=properties.character_unit_position,
        )
        self.replace_docinfo_model(model)

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
        self.ensure_section_count(section_index + 1)
        model = self.section_model(section_index)
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(model)
        paragraph = model.append_paragraph(
            text,
            para_shape_id=para_shape_id,
            style_id=style_id,
            split_flags=split_flags,
            control_mask=control_mask,
        )
        if use_synthetic_line_numbers:
            _ensure_paragraph_line_seg(paragraph)
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
        width: int = DEFAULT_HWP_SHAPE_WIDTH,
        height: int = DEFAULT_HWP_SHAPE_HEIGHT,
        profile_root: str | Path | None = None,
        section_index: int | None = None,
    ) -> None:
        if image_bytes is None:
            self._append_picture_with_storage_id(
                self._resolve_default_picture_storage_id(),
                width=width,
                height=height,
                profile_root=profile_root,
                section_index=section_index,
            )
            return
        self._append_picture_with_bindata(
            image_bytes,
            extension=extension or "png",
            width=width,
            height=height,
            profile_root=profile_root,
            section_index=section_index,
        )

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
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(model)
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
        if use_synthetic_line_numbers:
            _ensure_paragraph_line_seg(paragraph)
        paragraph.header.add_child(control_node)
        self.replace_section_model(target_section_index, model)

    def append_auto_number(
        self,
        *,
        number: int | str = 1,
        number_type: str = "PAGE",
        kind: str = "newNum",
        section_index: int | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        normalized_kind = kind.strip()
        if normalized_kind not in {"newNum", "autoNum"}:
            raise ValueError("append_auto_number() kind must be 'newNum' or 'autoNum'.")

        target_section_index = self._resolve_append_section_index(None, section_index)
        model = self.section_model(target_section_index)
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(model)
        control_id = "nwno" if normalized_kind == "newNum" else "atno"
        raw_text = _build_control_marker_text(control_id) + "\r"
        if paragraph_index is None or paragraph_index < 0:
            paragraph = model.append_paragraph(raw_text)
        else:
            paragraph = model.insert_paragraph(paragraph_index, raw_text)
        if use_synthetic_line_numbers:
            _ensure_paragraph_line_seg(paragraph)
        control_node = RecordNode(
            tag_id=TAG_CTRL_HEADER,
            level=1,
            payload=_build_auto_number_payload(normalized_kind, number=number, number_type=number_type),
        )
        paragraph.header.add_child(control_node)
        self.replace_section_model(target_section_index, model)

    def append_header(
        self,
        text: str,
        *,
        section_index: int | None = None,
    ) -> None:
        self._upsert_header_footer_control("head", text=text, section_index=section_index)

    def append_footer(
        self,
        text: str,
        *,
        section_index: int | None = None,
    ) -> None:
        self._upsert_header_footer_control("foot", text=text, section_index=section_index)

    def append_bookmark(
        self,
        name: str,
        *,
        section_index: int | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        target_section_index = self._resolve_append_section_index(None, section_index)
        model = self.section_model(target_section_index)
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(model)
        raw_text = _build_control_marker_text("bokm") + "\r"
        if paragraph_index is None or paragraph_index < 0:
            paragraph = model.append_paragraph(raw_text)
        else:
            paragraph = model.insert_paragraph(paragraph_index, raw_text)
        if use_synthetic_line_numbers:
            _ensure_paragraph_line_seg(paragraph)
        control_node = RecordNode(tag_id=TAG_CTRL_HEADER, level=1, payload=_build_control_id_payload("bokm"))
        control_node.add_child(RecordNode(tag_id=TAG_CTRL_DATA, level=2, payload=_build_bookmark_ctrl_data_payload(name)))
        paragraph.header.add_child(control_node)
        self.replace_section_model(target_section_index, model)

    def append_note(
        self,
        text: str,
        *,
        kind: str = "footNote",
        number: int | str | None = None,
        section_index: int | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        normalized_kind = kind.strip()
        if normalized_kind not in {"footNote", "endNote"}:
            raise ValueError("append_note() kind must be 'footNote' or 'endNote'.")

        target_section_index = self._resolve_append_section_index(None, section_index)
        model = self.section_model(target_section_index)
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(model)
        control_id = "fn  " if normalized_kind == "footNote" else "en  "
        raw_text = _build_control_marker_text(control_id) + "\r"
        if paragraph_index is None or paragraph_index < 0:
            paragraph = model.append_paragraph(raw_text)
        else:
            paragraph = model.insert_paragraph(paragraph_index, raw_text)
        if use_synthetic_line_numbers:
            _ensure_paragraph_line_seg(paragraph)
        paragraph.header.add_child(_build_note_control_node(control_id, text))
        self.replace_section_model(target_section_index, model)

    def append_footnote(
        self,
        text: str,
        *,
        number: int | str | None = None,
        section_index: int | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        self.append_note(
            text,
            kind="footNote",
            number=number,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_endnote(
        self,
        text: str,
        *,
        number: int | str | None = None,
        section_index: int | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        self.append_note(
            text,
            kind="endNote",
            number=number,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_equation(
        self,
        script: str = DEFAULT_HWP_EQUATION_SCRIPT,
        *,
        width: int = 4800,
        height: int = 2300,
        font: str = DEFAULT_HWP_EQUATION_FONT,
        shape_comment: str | None = None,
        section_index: int | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        target_section_index = self._resolve_append_section_index(None, section_index)
        model = self.section_model(target_section_index)
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(model)
        if paragraph_index is None or paragraph_index < 0:
            paragraph = model.append_paragraph("\r")
        else:
            paragraph = model.insert_paragraph(paragraph_index, "\r")
        if use_synthetic_line_numbers:
            _ensure_paragraph_line_seg(paragraph)
        control_node = RecordNode(
            tag_id=TAG_CTRL_HEADER,
            level=1,
            payload=_build_graphic_control_header_payload(
                _DEFAULT_EQUATION_CTRL_HEADER,
                "eqed",
                width=width,
                height=height,
                description=shape_comment,
            ),
        )
        control_node.add_child(RecordNode(tag_id=TAG_EQEDIT, level=2, payload=_build_equation_payload(script, font=font)))
        paragraph.header.add_child(control_node)
        self.replace_section_model(target_section_index, model)

    def append_shape(
        self,
        *,
        kind: str = "rect",
        text: str = DEFAULT_HWP_SHAPE_TEXT,
        width: int = DEFAULT_HWP_SHAPE_WIDTH,
        height: int = DEFAULT_HWP_SHAPE_HEIGHT,
        fill_color: str = "#FFFFFF",
        line_color: str = "#000000",
        shape_comment: str | None = None,
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
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(model)
        paragraph = _insert_blank_like_paragraph(model, paragraph_index)
        if use_synthetic_line_numbers:
            _ensure_paragraph_line_seg(paragraph)
        _append_control_marker_to_paragraph(paragraph, "gso ")
        control_node = RecordNode(
            tag_id=TAG_CTRL_HEADER,
            level=1,
            payload=_build_graphic_control_header_payload(
                _shape_control_header_template(kind),
                "gso ",
                width=width,
                height=height,
                description=shape_comment,
            ),
        )
        metadata_fields: dict[str, str] = {}
        if kind in {"ellipse", "arc", "polygon"} and fill_color:
            metadata_fields["fill_color"] = fill_color
        if kind in {"ellipse", "arc", "polygon"} and text:
            metadata_fields["text"] = text
        if metadata_fields:
            control_node.add_child(
                RecordNode(
                    tag_id=TAG_CTRL_DATA,
                    level=2,
                    payload=_build_shape_native_metadata_payload(**metadata_fields),
                )
            )
        shape_component = RecordNode(
            tag_id=TAG_SHAPE_COMPONENT,
            level=2,
            payload=_build_shape_component_payload(
                kind,
                width,
                height,
                fill_color=fill_color,
                line_color=line_color,
            ),
        )
        if text and kind not in {"line", "ellipse", "arc", "polygon", "textart"}:
            list_header, paragraph_header = _build_shape_text_records(text)
            shape_component.add_child(list_header)
            shape_component.add_child(paragraph_header)
        shape_component.add_child(
            RecordNode(
                tag_id=kind_to_tag[kind],
                level=3,
                payload=_build_shape_specific_payload(
                    kind,
                    text=text,
                    width=width,
                    height=height,
                    fill_color=fill_color,
                    line_color=line_color,
                ),
            )
        )
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
        shape_comment: str | None = None,
        object_type: str = "EMBEDDED",
        draw_aspect: str = "CONTENT",
        has_moniker: bool = False,
        eq_baseline: int = 0,
        line_color: str = "#000000",
        line_width: int = _DEFAULT_OLE_LINE_WIDTH,
        section_index: int | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        target_section_index = self._resolve_append_section_index(None, section_index)
        storage_id, _ = self.add_embedded_bindata(data, extension="ole")
        model = self.section_model(target_section_index)
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(model)
        paragraph_text = f"{name}\r" if name else "\r"
        if paragraph_index is None or paragraph_index < 0:
            paragraph = model.append_paragraph(paragraph_text)
        else:
            paragraph = model.insert_paragraph(paragraph_index, paragraph_text)
        if use_synthetic_line_numbers:
            _ensure_paragraph_line_seg(paragraph)
        control_node = RecordNode(
            tag_id=TAG_CTRL_HEADER,
            level=1,
            payload=_build_graphic_control_header_payload(
                _shape_control_header_template("ole"),
                "gso ",
                width=width,
                height=height,
                description=shape_comment,
            ),
        )
        shape_component = RecordNode(
            tag_id=TAG_SHAPE_COMPONENT,
            level=2,
            payload=_build_shape_component_payload("ole", width, height),
        )
        if name:
            list_header, paragraph_header = _build_shape_text_records(name)
            shape_component.add_child(list_header)
            shape_component.add_child(paragraph_header)
        shape_component.add_child(
            RecordNode(
                tag_id=TAG_SHAPE_COMPONENT_OLE,
                level=3,
                payload=_build_ole_specific_payload(
                    storage_id=storage_id,
                    width=width,
                    height=height,
                    object_type=object_type,
                    draw_aspect=draw_aspect,
                    has_moniker=has_moniker,
                    eq_baseline=eq_baseline,
                    line_color=line_color,
                    line_width=line_width,
                ),
            )
        )
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
            self.ensure_section_count(section_index + 1)
            return section_index
        if profile_root is None:
            from .hwp_pure_profile import HwpPureProfile

            target_section_index = HwpPureProfile.load_bundled().target_section_index
            self.ensure_section_count(target_section_index + 1)
            return target_section_index
        profile_path = Path(profile_root).expanduser().resolve()
        if not profile_path.exists():
            self.ensure_section_count(1)
            return 0
        metadata_path = profile_path / "profile.json"
        if not metadata_path.exists():
            self.ensure_section_count(1)
            return 0
        from .hwp_pure_profile import HwpPureProfile

        target_section_index = HwpPureProfile.load(profile_path).target_section_index
        self.ensure_section_count(target_section_index + 1)
        return target_section_index

    def _upsert_header_footer_control(
        self,
        control_id: str,
        *,
        text: str,
        section_index: int | None,
    ) -> None:
        target_section_index = self._resolve_append_section_index(None, section_index)
        model = self.section_model(target_section_index)
        paragraphs = model.paragraphs()
        if not paragraphs:
            paragraph = model.append_paragraph("\r")
            paragraphs = model.paragraphs()
        target_paragraph = paragraphs[0]
        control_node = next(
            (
                child
                for child in target_paragraph.header.children
                if child.tag_id == TAG_CTRL_HEADER and _decode_control_id_payload(child.payload) == control_id
            ),
            None,
        )
        if control_node is None:
            control_node = _build_header_footer_control_node(control_id)
            target_paragraph.header.add_child(control_node)
            _append_control_marker_to_paragraph(target_paragraph, control_id)
        _set_header_footer_control_text(control_node, text)
        self.replace_section_model(target_section_index, model)

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

    def _build_picture_records(self, storage_id: int, *, width: int, height: int) -> list[HwpRecord]:
        top_char_count = len(_PICTURE_TOP_PARA_TEXT) // 2
        picture_payload = bytearray(_PICTURE_SHAPE_PICTURE)
        if len(picture_payload) <= 72:
            raise HwpBinaryEditError("Built picture payload is shorter than expected.")
        picture_payload[71:73] = int(storage_id).to_bytes(2, "little", signed=False)
        shape_component_payload = bytearray(_PICTURE_SHAPE_COMPONENT)
        shape_component_payload[20:24] = int(width).to_bytes(4, "little", signed=False)
        shape_component_payload[24:28] = int(height).to_bytes(4, "little", signed=False)
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
                payload=bytes(shape_component_payload),
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
        width: int,
        height: int,
        profile_root: str | Path | None,
        section_index: int | None,
    ) -> None:
        target_section_index = self._resolve_append_section_index(profile_root, section_index)
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(self.section_model(target_section_index))
        section_records = self.section_records(target_section_index)
        section_records.extend(self._build_picture_records(storage_id, width=width, height=height))
        self.replace_section_records(target_section_index, section_records)
        if use_synthetic_line_numbers:
            model = self.section_model(target_section_index)
            _apply_section_show_line_numbers(model, True)
            self.replace_section_model(target_section_index, model)

    def _append_picture_with_bindata(
        self,
        image_bytes: bytes,
        *,
        extension: str,
        width: int,
        height: int,
        profile_root: str | Path | None,
        section_index: int | None,
    ) -> None:
        storage_id, _stream_path = self.add_embedded_bindata(image_bytes, extension=extension)
        self._append_picture_with_storage_id(
            storage_id,
            width=width,
            height=height,
            profile_root=profile_root,
            section_index=section_index,
        )

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
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(self.section_model(target_section_index))
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
        if use_synthetic_line_numbers:
            model = self.section_model(target_section_index)
            _apply_section_show_line_numbers(model, True)
            self.replace_section_model(target_section_index, model)

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
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(self.section_model(target_section_index))
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
        if use_synthetic_line_numbers:
            model = self.section_model(target_section_index)
            _apply_section_show_line_numbers(model, True)
            self.replace_section_model(target_section_index, model)


def _build_header_footer_control_node(control_id: str) -> RecordNode:
    if control_id == "head":
        control_payload = _HEADER_CTRL_HEADER
        list_payload = _HEADER_LIST_HEADER
        paragraph_payload = _HEADER_PARA_HEADER
        char_shape_payload = _HEADER_PARA_CHAR_SHAPE
        line_seg_payload = _HEADER_PARA_LINE_SEG
    elif control_id == "foot":
        control_payload = _FOOTER_CTRL_HEADER
        list_payload = _FOOTER_LIST_HEADER
        paragraph_payload = _FOOTER_PARA_HEADER
        char_shape_payload = _FOOTER_PARA_CHAR_SHAPE
        line_seg_payload = _FOOTER_PARA_LINE_SEG
    else:
        raise ValueError(f"Unsupported header/footer control id: {control_id}")

    control_node = RecordNode(tag_id=TAG_CTRL_HEADER, level=1, payload=control_payload)
    control_node.add_child(RecordNode(tag_id=TAG_LIST_HEADER, level=2, payload=list_payload))
    paragraph_header = ParagraphHeaderRecord.from_record(
        HwpRecord(
            tag_id=TAG_PARA_HEADER,
            level=2,
            size=len(paragraph_payload),
            header_size=4,
            offset=-1,
            payload=paragraph_payload,
        )
    )
    paragraph_header.add_child(RecordNode(tag_id=TAG_PARA_CHAR_SHAPE, level=3, payload=char_shape_payload))
    paragraph_header.add_child(RecordNode(tag_id=TAG_PARA_LINE_SEG, level=3, payload=line_seg_payload))
    control_node.add_child(paragraph_header)
    return control_node


def _append_control_marker_to_paragraph(paragraph: SectionParagraphModel, control_id: str) -> None:
    text_record = paragraph.text_record()
    if text_record is None:
        raise HwpBinaryEditError("Paragraph does not contain a para text record.")
    marker = _build_control_marker_text(control_id)
    if marker in text_record.raw_text:
        return
    raw_text = text_record.raw_text or "\r"
    if not raw_text.endswith("\r"):
        raw_text += "\r"
    updated = raw_text[:-1] + marker + "\r"
    text_record.set_raw_text(updated)
    paragraph.header.char_count = len(updated)
    paragraph.header.sync_payload()


def _set_header_footer_control_text(control_node: RecordNode, text: str) -> None:
    paragraph_header = next(
        (child for child in control_node.children if isinstance(child, ParagraphHeaderRecord)),
        None,
    )
    if paragraph_header is None:
        raise HwpBinaryEditError("Header/footer control does not contain a paragraph header.")
    text_record = next((child for child in paragraph_header.children if isinstance(child, ParagraphTextRecord)), None)
    raw_text = f"{text}\r" if text else ""
    if text_record is None:
        if not raw_text:
            paragraph_header.char_count = 0
            paragraph_header.sync_payload()
            return
        text_record = ParagraphTextRecord(level=paragraph_header.level + 1, raw_text=raw_text)
        paragraph_header.children.insert(0, text_record)
    else:
        text_record.set_raw_text(raw_text)
    paragraph_header.char_count = len(raw_text)
    paragraph_header.sync_payload()


def _build_note_control_node(control_id: str, text: str) -> RecordNode:
    control_node = RecordNode(tag_id=TAG_CTRL_HEADER, level=1, payload=_build_note_ctrl_payload(control_id))
    control_node.add_child(RecordNode(tag_id=TAG_LIST_HEADER, level=2, payload=_NOTE_LIST_HEADER))
    paragraph_header_payload = bytearray(_NOTE_PARA_HEADER)
    raw_text = f"{text}\r" if text else ""
    paragraph_header_payload[0:4] = (0x80000000 | len(raw_text)).to_bytes(4, "little", signed=False)
    paragraph_header = ParagraphHeaderRecord.from_record(
        HwpRecord(
            tag_id=TAG_PARA_HEADER,
            level=2,
            size=len(paragraph_header_payload),
            header_size=4,
            offset=-1,
            payload=bytes(paragraph_header_payload),
        )
    )
    if raw_text:
        paragraph_header.add_child(ParagraphTextRecord(level=3, raw_text=raw_text))
    paragraph_header.add_child(RecordNode(tag_id=TAG_PARA_CHAR_SHAPE, level=3, payload=_NOTE_CHAR_SHAPE))
    paragraph_header.add_child(RecordNode(tag_id=TAG_PARA_LINE_SEG, level=3, payload=_NOTE_LINE_SEG))
    control_node.add_child(paragraph_header)
    return control_node


def _build_default_paragraph_line_seg_payload(paragraph_index: int) -> bytes:
    payload = bytearray(_NOTE_LINE_SEG)
    payload[4:8] = max(paragraph_index, 0).to_bytes(4, "little", signed=False)
    return bytes(payload)


def _paragraph_line_seg_node(paragraph: SectionParagraphModel) -> RecordNode | None:
    return next((child for child in paragraph.header.children if child.tag_id == TAG_PARA_LINE_SEG), None)


def _paragraph_uses_synthetic_line_number_layout(paragraph: SectionParagraphModel) -> bool:
    line_seg = _paragraph_line_seg_node(paragraph)
    if line_seg is None:
        return False
    if line_seg.payload != _build_default_paragraph_line_seg_payload(paragraph.index):
        return False
    trailing = paragraph.header.trailing_payload
    return len(trailing) > 4 and trailing[4] == 1


def _top_level_section_paragraphs(model: SectionModel) -> list[SectionParagraphModel]:
    return [paragraph for paragraph in model.paragraphs() if paragraph.header.level == 0]


def _section_uses_synthetic_line_number_layout(model: SectionModel) -> bool:
    paragraphs = _top_level_section_paragraphs(model)
    return bool(paragraphs) and all(_paragraph_uses_synthetic_line_number_layout(paragraph) for paragraph in paragraphs)


def _set_paragraph_line_number_flag(paragraph: SectionParagraphModel, enabled: bool) -> None:
    trailing = bytearray(paragraph.header.trailing_payload)
    while len(trailing) <= 4:
        trailing.append(0)
    trailing[4] = 1 if enabled else 0
    paragraph.header.trailing_payload = bytes(trailing)
    paragraph.header.sync_payload()


def _ensure_paragraph_line_seg(paragraph: SectionParagraphModel) -> None:
    line_seg = _paragraph_line_seg_node(paragraph)
    payload = _build_default_paragraph_line_seg_payload(paragraph.index)
    if line_seg is None:
        paragraph.header.add_child(RecordNode(tag_id=TAG_PARA_LINE_SEG, level=paragraph.header.level + 1, payload=payload))
    else:
        line_seg.payload = payload
    _set_paragraph_line_number_flag(paragraph, True)


def _remove_synthetic_paragraph_line_seg(paragraph: SectionParagraphModel) -> None:
    line_seg = _paragraph_line_seg_node(paragraph)
    if line_seg is not None and line_seg.payload == _build_default_paragraph_line_seg_payload(paragraph.index):
        paragraph.header.children.remove(line_seg)
    _set_paragraph_line_number_flag(paragraph, False)


def _apply_section_show_line_numbers(model: SectionModel, enabled: bool) -> None:
    for paragraph in _top_level_section_paragraphs(model):
        if enabled:
            _ensure_paragraph_line_seg(paragraph)
        else:
            _remove_synthetic_paragraph_line_seg(paragraph)
