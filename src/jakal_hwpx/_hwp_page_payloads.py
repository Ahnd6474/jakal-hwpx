from __future__ import annotations

import struct

from ._hwp_payload_codecs import _parse_colorref


def _build_control_id_payload(control_id: str) -> bytes:
    normalized = control_id[:4].ljust(4)
    return normalized.encode("latin1", errors="replace")[::-1]


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
_SECTION_DEFINITION_STARTS_ON_SHIFT = 20
_SECTION_DEFINITION_STARTS_ON_MASK = 0x3 << _SECTION_DEFINITION_STARTS_ON_SHIFT
_SECTION_DEFINITION_STARTS_ON_CODES = {
    "BOTH": 0,
    "ODD": 1,
    "EVEN": 2,
}
_SECTION_DEFINITION_STARTS_ON_VALUES = {
    value: key for key, value in _SECTION_DEFINITION_STARTS_ON_CODES.items()
}
_HEADER_FOOTER_APPLY_PAGE_TYPE_CODES = {
    "BOTH": 0,
    "EVEN": 1,
    "ODD": 2,
}
_HEADER_FOOTER_APPLY_PAGE_TYPE_VALUES = {
    value: key for key, value in _HEADER_FOOTER_APPLY_PAGE_TYPE_CODES.items()
}
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
        if "pageStartsOn" in start_numbers and start_numbers["pageStartsOn"] is not None:
            starts_on = str(start_numbers["pageStartsOn"]).upper()
            starts_on_code = _SECTION_DEFINITION_STARTS_ON_CODES.get(starts_on, 0)
            attributes &= ~_SECTION_DEFINITION_STARTS_ON_MASK
            attributes |= starts_on_code << _SECTION_DEFINITION_STARTS_ON_SHIFT
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
    auto_num_value = note_pr.get("autoNumFormat")
    numbering_value = note_pr.get("numbering")
    note_line_value = note_pr.get("noteLine")
    note_spacing_value = note_pr.get("noteSpacing")
    placement_value = note_pr.get("placement")
    auto_num = dict(auto_num_value) if isinstance(auto_num_value, dict) else {}
    numbering = dict(numbering_value) if isinstance(numbering_value, dict) else {}
    note_line = dict(note_line_value) if isinstance(note_line_value, dict) else {}
    note_spacing = dict(note_spacing_value) if isinstance(note_spacing_value, dict) else {}
    placement = dict(placement_value) if isinstance(placement_value, dict) else {}
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
    updated = bytearray(payload)
    if len(updated) < 26:
        updated.extend(b"\x00" * (26 - len(updated)))
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
    return bytes(updated)
def _build_colorref(value: str) -> bytes:
    normalized = str(value).strip().lstrip("#")
    if len(normalized) != 6:
        normalized = "000000"
    red = int(normalized[0:2], 16)
    green = int(normalized[2:4], 16)
    blue = int(normalized[4:6], 16)
    return bytes((blue, green, red, 0))


def _parse_rgb_color(payload: bytes) -> str:
    values = payload[:4].ljust(4, b"\x00")
    red = values[0]
    green = values[1]
    blue = values[2]
    return f"#{red:02X}{green:02X}{blue:02X}"


def _build_rgb_color(value: str) -> bytes:
    normalized = str(value).strip().lstrip("#")
    if len(normalized) != 6:
        normalized = "000000"
    red = int(normalized[0:2], 16)
    green = int(normalized[2:4], 16)
    blue = int(normalized[4:6], 16)
    return bytes((red, green, blue, 0))


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

for _compat_name in ["_parse_section_definition_payload","_build_section_definition_payload","_parse_page_starts_on_payload","_build_page_starts_on_payload","_parse_page_hiding_payload","_build_page_hiding_payload","_parse_page_num_payload","_build_page_num_payload","_parse_note_shape_payload","_build_note_shape_payload","_build_colorref","_parse_rgb_color","_build_rgb_color","_parse_page_def_payload","_build_page_def_payload","_parse_page_border_fill_payload","_build_page_border_fill_payload"]:
    globals()[_compat_name].__module__ = "jakal_hwpx.hwp_binary"

