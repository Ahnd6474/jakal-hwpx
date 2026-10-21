from __future__ import annotations

import struct

from ._hwp_constants import (
    _HWP_CHAR_FONT_ORDER,
    _HWP_LINE_ARROW_NAMES,
    _HWP_LINE_ARROW_SIZE_NAMES,
    _HWP_LINE_END_CAP_NAMES,
    _HWP_LINE_OUTLINE_STYLE_NAMES,
    _HWP_LINE_STYLE_NAMES,
    _HWP_TAB_LEADER_NAMES,
    _HWP_TAB_TYPE_NAMES,
    _MEMO_SHAPE_NATIVE_METADATA_PREFIX,
    TAG_SHAPE_COMPONENT_ARC,
    TAG_SHAPE_COMPONENT_CONTAINER,
    TAG_SHAPE_COMPONENT_CURVE,
    TAG_SHAPE_COMPONENT_ELLIPSE,
    TAG_SHAPE_COMPONENT_LINE,
    TAG_SHAPE_COMPONENT_POLYGON,
    TAG_SHAPE_COMPONENT_RECTANGLE,
    TAG_SHAPE_COMPONENT_TEXTART,
)


def _parse_colorref(payload: bytes) -> str:
    values = payload[:4].ljust(4, b"\x00")
    blue = values[0]
    green = values[1]
    red = values[2]
    return f"#{red:02X}{green:02X}{blue:02X}"

def _parse_shape_fill_payload(payload: bytes) -> dict[str, object]:
    if len(payload) < 4:
        return {}
    fill_type = int.from_bytes(payload[0:4], "little", signed=False)
    result: dict[str, object] = {
        "fill_type": fill_type,
        "fill_type_flags": {
            "color_pattern": bool(fill_type & 0x00000001),
            "image_fill": bool(fill_type & 0x00000002),
            "gradation_fill": bool(fill_type & 0x00000004),
        },
    }
    if fill_type & 0x00000001:
        if len(payload) >= 8:
            result["faceColor"] = _parse_colorref(payload[4:8])
        if len(payload) >= 12:
            result["hatchColor"] = _parse_colorref(payload[8:12])
        if len(payload) >= 16:
            result["pattern_flags"] = int.from_bytes(payload[12:16], "little", signed=False)
        if len(payload) >= 17:
            result["pattern_alpha"] = payload[16]
        if len(payload) > 17:
            result["pattern_tail"] = payload[17:]
    elif len(payload) > 4:
        result["raw_payload"] = payload[4:]
    result["u16_values"] = _payload_u16_values(payload)
    result["u32_values"] = _payload_u32_values(payload)
    return result

def _parse_memo_shape_native_metadata(payload: bytes) -> dict[str, str]:
    if not payload:
        return {}
    text = payload.decode("utf-16-le", errors="ignore")
    if not text.startswith(_MEMO_SHAPE_NATIVE_METADATA_PREFIX):
        return {}
    metadata: dict[str, str] = {}
    for item in text.split(";")[1:]:
        key, separator, value = item.partition("=")
        if separator and key:
            metadata[key] = value
    return metadata

def _parse_list_header_payload(payload: bytes) -> dict[str, object]:
    values = bytes(payload)
    result: dict[str, object] = {
        "u16_values": _payload_u16_values(values),
        "u32_values": _payload_u32_values(values),
    }
    if len(values) >= 4:
        result["paragraph_count"] = int.from_bytes(values[0:4], "little", signed=False)
    if len(values) >= 8:
        property_flags = int.from_bytes(values[4:8], "little", signed=False)
        result["property_flags"] = property_flags
        result["property_bits"] = _bit_flags(property_flags)
    return result

def _parse_memo_shape_payload(payload: bytes) -> dict[str, object]:
    values = bytes(payload)
    result: dict[str, object] = {
        "raw_payload": values,
        "u16_values": _payload_u16_values(values),
        "u32_values": _payload_u32_values(values),
    }
    metadata = _parse_memo_shape_native_metadata(values)
    if metadata:
        result.update(metadata)
        if "width" in metadata:
            try:
                result["width"] = int(metadata["width"])
            except ValueError:
                pass
        if "lineWidth" in metadata:
            try:
                result["line_width"] = int(metadata["lineWidth"])
            except ValueError:
                pass
        if "lineType" in metadata:
            result["line_type"] = metadata["lineType"]
        if "lineColor" in metadata:
            result["line_color"] = metadata["lineColor"]
        if "fillColor" in metadata:
            result["fill_color"] = metadata["fillColor"]
        if "activeColor" in metadata:
            result["active_color"] = metadata["activeColor"]
        if "memoType" in metadata:
            result["memo_type"] = metadata["memoType"]
    decoded = values.decode("utf-16-le", errors="ignore").replace("\x00", "")
    if decoded:
        result["utf16_text"] = decoded
    return result

def _parse_form_object_payload(payload: bytes) -> dict[str, object]:
    values = bytes(payload)
    result: dict[str, object] = {
        "raw_payload": values,
        "u16_values": _payload_u16_values(values),
        "u32_values": _payload_u32_values(values),
    }
    decoded = values.decode("utf-16-le", errors="ignore").replace("\x00", "")
    if decoded:
        result["utf16_text"] = decoded
        result["label"] = decoded
    return result

def _parse_memo_list_payload(payload: bytes) -> dict[str, object]:
    values = bytes(payload)
    result: dict[str, object] = {
        "raw_payload": values,
        "u16_values": _payload_u16_values(values),
        "u32_values": _payload_u32_values(values),
    }
    decoded = values.decode("utf-16-le", errors="ignore").replace("\x00", "")
    if decoded:
        result["utf16_text"] = decoded
        result["text"] = decoded
    return result

def _parse_chart_data_payload(payload: bytes) -> dict[str, object]:
    values = bytes(payload)
    result: dict[str, object] = {
        "raw_payload": values,
        "u16_values": _payload_u16_values(values),
        "u32_values": _payload_u32_values(values),
    }
    decoded = values.decode("utf-16-le", errors="ignore").replace("\x00", "")
    if decoded:
        result["utf16_text"] = decoded
        result["title"] = decoded
    return result

def _parse_track_change_payload(payload: bytes) -> dict[str, object]:
    values = bytes(payload)
    result: dict[str, object] = {
        "u16_values": _payload_u16_values(values),
        "u32_values": _payload_u32_values(values),
    }
    if len(values) >= 4:
        flags = int.from_bytes(values[0:4], "little", signed=False)
        result["flags"] = flags
        result["flag_bits"] = _bit_flags(flags)
    decoded = values.decode("utf-16-le", errors="ignore").replace("\x00", "")
    if decoded:
        result["utf16_text"] = decoded
    return result

def _parse_track_change_author_payload(payload: bytes) -> dict[str, object]:
    values = bytes(payload)
    name, cursor = _parse_length_prefixed_utf16(values, 0)
    result: dict[str, object] = {
        "name": name,
        "name_length": len(name),
        "u16_values": _payload_u16_values(values),
        "u32_values": _payload_u32_values(values),
    }
    if cursor < len(values):
        result["trailing_payload"] = values[cursor:]
    return result

def _parse_textart_payload_text(payload: bytes) -> str:
    if len(payload) < 34:
        return ""
    text_length = int.from_bytes(payload[32:34], "little", signed=False)
    end = 34 + text_length * 2
    if end > len(payload):
        return ""
    return payload[34:end].decode("utf-16-le", errors="ignore").replace("\r", "")

def _parse_length_prefixed_utf16(payload: bytes, cursor: int) -> tuple[str, int]:
    if cursor + 2 > len(payload):
        return ("", cursor)
    char_count = int.from_bytes(payload[cursor : cursor + 2], "little", signed=False)
    cursor += 2
    end = min(cursor + char_count * 2, len(payload))
    encoded = payload[cursor:end]
    if len(encoded) % 2:
        encoded = encoded[:-1]
    return (encoded.decode("utf-16-le", errors="ignore"), end)

def _parse_line_attribute_flags(value: int) -> dict[str, object]:
    return {
        "line_style_code": value & 0x3F,
        "line_style": _HWP_LINE_STYLE_NAMES.get(value & 0x3F, f"CODE_{value & 0x3F}"),
        "end_cap_code": (value >> 6) & 0x0F,
        "end_cap": _HWP_LINE_END_CAP_NAMES.get((value >> 6) & 0x0F, f"CODE_{(value >> 6) & 0x0F}"),
        "head_style_code": (value >> 10) & 0x3F,
        "head_style": _HWP_LINE_ARROW_NAMES.get((value >> 10) & 0x3F, f"CODE_{(value >> 10) & 0x3F}"),
        "tail_style_code": (value >> 16) & 0x3F,
        "tail_style": _HWP_LINE_ARROW_NAMES.get((value >> 16) & 0x3F, f"CODE_{(value >> 16) & 0x3F}"),
        "head_size_code": (value >> 22) & 0x0F,
        "head_size": _HWP_LINE_ARROW_SIZE_NAMES.get((value >> 22) & 0x0F, f"CODE_{(value >> 22) & 0x0F}"),
        "tail_size_code": (value >> 26) & 0x0F,
        "tail_size": _HWP_LINE_ARROW_SIZE_NAMES.get((value >> 26) & 0x0F, f"CODE_{(value >> 26) & 0x0F}"),
        "head_fill": bool(value & (1 << 30)),
        "tail_fill": bool(value & (1 << 31)),
    }

def _payload_u16_values(payload: bytes) -> list[int]:
    if len(payload) < 2:
        return []
    even_length = len(payload) - (len(payload) % 2)
    return list(struct.unpack("<" + "H" * (even_length // 2), payload[:even_length]))

def _payload_u32_values(payload: bytes) -> list[int]:
    if len(payload) < 4:
        return []
    aligned_length = len(payload) - (len(payload) % 4)
    return list(struct.unpack("<" + "I" * (aligned_length // 4), payload[:aligned_length]))

def _bit_flags(value: int, *, width: int = 32, start: int = 0, prefix: str = "bit") -> dict[str, bool]:
    return {f"{prefix}_{index}": bool(value & (1 << index)) for index in range(start, start + width)}

def _parse_hwp_colorref_or_none(payload: bytes) -> str:
    values = payload[:4].ljust(4, b"\x00")
    if values == b"\xff\xff\xff\xff":
        return "none"
    return _parse_colorref(values)

def _parse_shape_line_info_payload(payload: bytes) -> dict[str, object]:
    values = payload[:11].ljust(11, b"\x00")
    attributes = int.from_bytes(values[6:10], "little", signed=False)
    return {
        "color": _parse_colorref(values[0:4]),
        "width": int.from_bytes(values[4:6], "little", signed=False),
        "attributes": attributes,
        "attribute_flags": _parse_line_attribute_flags(attributes),
        "outline_style_code": values[10],
        "outline_style": _HWP_LINE_OUTLINE_STYLE_NAMES.get(values[10], f"CODE_{values[10]}"),
    }

def _parse_border_line_payload(payload: bytes) -> dict[str, object]:
    values = payload[:6].ljust(6, b"\x00")
    style_code = values[0]
    width_code = values[1]
    return {
        "style_code": style_code,
        "style": _HWP_LINE_STYLE_NAMES.get(style_code, f"CODE_{style_code}"),
        "width_code": width_code,
        "color": _parse_colorref(values[2:6]),
    }

def _parse_face_name_payload(payload: bytes) -> dict[str, object]:
    if not payload:
        return {"flags": 0, "name": ""}
    flags = payload[0]
    cursor = 1
    name, cursor = _parse_length_prefixed_utf16(payload, cursor)
    result: dict[str, object] = {
        "flags": flags,
        "name": name,
        "has_substitute_font": bool(flags & 0x80),
        "has_font_information": bool(flags & 0x40),
        "has_default_font": bool(flags & 0x20),
        "flag_bits": {
            "alternate_font": bool(flags & 0x80),
            "font_information": bool(flags & 0x40),
            "default_font": bool(flags & 0x20),
            "metric_font": bool(flags & 0x10),
        },
    }
    if result["has_substitute_font"] and cursor < len(payload):
        result["substitute_font_type"] = payload[cursor]
        cursor += 1
        substitute_name, cursor = _parse_length_prefixed_utf16(payload, cursor)
        result["substitute_font_name"] = substitute_name
    if result["has_font_information"] and cursor + 10 <= len(payload):
        panose = payload[cursor : cursor + 10]
        cursor += 10
        result["font_information"] = {
            "family_type": panose[0],
            "serif_style": panose[1],
            "weight": panose[2],
            "proportion": panose[3],
            "contrast": panose[4],
            "stroke_variation": panose[5],
            "arm_style": panose[6],
            "letterform": panose[7],
            "midline": panose[8],
            "x_height": panose[9],
        }
    if result["has_default_font"]:
        default_name, cursor = _parse_length_prefixed_utf16(payload, cursor)
        result["default_font_name"] = default_name
    if cursor < len(payload):
        result["trailing_payload"] = payload[cursor:]
    return result

def _parse_border_fill_payload(payload: bytes) -> dict[str, object]:
    values = payload
    border_flags = int.from_bytes(values[0:2], "little", signed=False) if len(values) >= 2 else 0
    result: dict[str, object] = {
        "border_flags": border_flags,
        "border_flag_bits": {
            "three_d": bool(border_flags & 0x0001),
            "shadow": bool(border_flags & 0x0002),
            "slash": (border_flags >> 2) & 0x07,
            "backslash": (border_flags >> 5) & 0x07,
            "crooked_slash": bool(border_flags & 0x0100),
            "crooked_backslash": bool(border_flags & 0x0200),
            "counter_slash": bool(border_flags & 0x0400),
        },
        "border_flag_words": {
            "low": border_flags & 0x00FF,
            "high": (border_flags >> 8) & 0x00FF,
        },
    }
    if len(values) >= 32:
        result["left"] = _parse_border_line_payload(values[2:8])
        result["right"] = _parse_border_line_payload(values[8:14])
        result["top"] = _parse_border_line_payload(values[14:20])
        result["bottom"] = _parse_border_line_payload(values[20:26])
        result["diagonal"] = _parse_border_line_payload(values[26:32])
    if len(values) >= 34:
        fill_flags = int.from_bytes(values[32:34], "little", signed=False)
        result["fill_flags"] = fill_flags
        result["fill_flag_bits"] = {
            "color_pattern": bool(fill_flags & 0x0001),
            "image_fill": bool(fill_flags & 0x0002),
            "gradation_fill": bool(fill_flags & 0x0004),
        }
        fill_payload = values[34:]
        result["fill"] = _parse_shape_fill_payload(fill_payload)
        if fill_payload:
            result["fill_payload"] = fill_payload
    result["u16_values"] = _payload_u16_values(values)
    result["u32_values"] = _payload_u32_values(values)
    return result

def _parse_character_shape_payload(payload: bytes) -> dict[str, object]:
    values = bytes(payload)
    attribute_raw = int.from_bytes(values[46:50].ljust(4, b"\x00"), "little", signed=False) if len(values) >= 50 else 0
    attribute_low = int.from_bytes(values[46:48].ljust(2, b"\x00"), "little", signed=False) if len(values) >= 48 else 0
    attribute_high = int.from_bytes(values[48:50].ljust(2, b"\x00"), "little", signed=False) if len(values) >= 50 else 0
    result: dict[str, object] = {
        "font_refs": {
            key: int.from_bytes(values[index * 2 : (index + 1) * 2].ljust(2, b"\x00"), "little", signed=False)
            for index, key in enumerate(_HWP_CHAR_FONT_ORDER)
        }
    }
    if len(values) >= 28:
        result["ratios"] = {
            key: int.from_bytes(values[14 + index : 15 + index], "little", signed=False)
            for index, key in enumerate(_HWP_CHAR_FONT_ORDER)
        }
        result["spacings"] = {
            key: int.from_bytes(values[21 + index : 22 + index], "little", signed=True)
            for index, key in enumerate(_HWP_CHAR_FONT_ORDER)
        }
    if len(values) >= 42:
        result["relative_sizes"] = {
            key: int.from_bytes(values[28 + index : 29 + index], "little", signed=False)
            for index, key in enumerate(_HWP_CHAR_FONT_ORDER)
        }
        result["position_offsets"] = {
            key: int.from_bytes(values[35 + index : 36 + index], "little", signed=True)
            for index, key in enumerate(_HWP_CHAR_FONT_ORDER)
        }
    if len(values) >= 50:
        result["base_size"] = int.from_bytes(values[42:46], "little", signed=False)
        result["attributes"] = attribute_raw
        result["attribute_bytes"] = list(values[46:50])
        result["attribute_words"] = {
            "low": attribute_low,
            "high": attribute_high,
        }
        result["attribute_sections"] = {
            "style_code": values[46],
            "outline_shadow_code": values[47],
            "effect_code": values[48],
            "decoration_code": values[49],
        }
        result["attribute_low_bits"] = {
            "bit_0": bool(attribute_low & 0x0001),
            "bit_1": bool(attribute_low & 0x0002),
            "bit_2": bool(attribute_low & 0x0004),
            "bit_3": bool(attribute_low & 0x0008),
            "bit_4": bool(attribute_low & 0x0010),
            "bit_5": bool(attribute_low & 0x0020),
            "bit_6": bool(attribute_low & 0x0040),
            "bit_7": bool(attribute_low & 0x0080),
            "bit_8": bool(attribute_low & 0x0100),
            "bit_9": bool(attribute_low & 0x0200),
            "bit_10": bool(attribute_low & 0x0400),
            "bit_11": bool(attribute_low & 0x0800),
            "bit_12": bool(attribute_low & 0x1000),
            "bit_13": bool(attribute_low & 0x2000),
            "bit_14": bool(attribute_low & 0x4000),
            "bit_15": bool(attribute_low & 0x8000),
        }
        result["attribute_high_bits"] = {
            "bit_16": bool(attribute_high & 0x0001),
            "bit_17": bool(attribute_high & 0x0002),
            "bit_18": bool(attribute_high & 0x0004),
            "bit_19": bool(attribute_high & 0x0008),
            "bit_20": bool(attribute_high & 0x0010),
            "bit_21": bool(attribute_high & 0x0020),
            "bit_22": bool(attribute_high & 0x0040),
            "bit_23": bool(attribute_high & 0x0080),
            "bit_24": bool(attribute_high & 0x0100),
            "bit_25": bool(attribute_high & 0x0200),
            "bit_26": bool(attribute_high & 0x0400),
            "bit_27": bool(attribute_high & 0x0800),
            "bit_28": bool(attribute_high & 0x1000),
            "bit_29": bool(attribute_high & 0x2000),
            "bit_30": bool(attribute_high & 0x4000),
            "bit_31": bool(attribute_high & 0x8000),
        }
    if len(values) >= 54:
        result["shadow_offset"] = {
            "x": int.from_bytes(values[50:52], "little", signed=True),
            "y": int.from_bytes(values[52:54], "little", signed=True),
        }
    if len(values) >= 58:
        result["text_color"] = _parse_hwp_colorref_or_none(values[54:58])
    if len(values) >= 60:
        result["tail_gap_58_60"] = values[58:60]
    if len(values) >= 64:
        result["shade_color"] = _parse_hwp_colorref_or_none(values[60:64])
    if len(values) >= 68:
        result["shadow_color"] = _parse_hwp_colorref_or_none(values[64:68])
    if len(values) >= 72:
        result["border_fill_id"] = int.from_bytes(values[68:72], "little", signed=False)
    if len(values) > 72:
        result["tail_reserved"] = values[72:]
    elif len(values) > 68:
        result["tail_reserved"] = values[68:]
    result["probable_effect_flags"] = {
        "attribute_bit_0": bool(attribute_raw & 0x00000001),
        "attribute_bit_1": bool(attribute_raw & 0x00000002),
        "attribute_bit_2": bool(attribute_raw & 0x00000004),
        "emboss": bool(attribute_raw & 0x00002000),
        "engrave": bool(attribute_raw & 0x00004000),
        "superscript": bool(attribute_raw & 0x00008000),
        "subscript": bool(attribute_raw & 0x00010000),
        "use_font_space": bool(attribute_raw & 0x20000000),
        "use_kerning": bool(attribute_raw & 0x40000000),
    }
    result["hwpx_compatible_flags"] = {
        "useFontSpace": bool(attribute_raw & 0x20000000),
        "useKerning": bool(attribute_raw & 0x40000000),
    }
    result["colors"] = {
        "text": result.get("text_color", "#000000"),
        "shade": result.get("shade_color", "none"),
        "shadow": result.get("shadow_color", "#C0C0C0"),
    }
    result["u16_values"] = _payload_u16_values(values)
    result["u32_values"] = _payload_u32_values(values)
    return result

def _parse_tab_def_payload(payload: bytes) -> dict[str, object]:
    values = bytes(payload)
    if len(values) < 8:
        return {"attributes": int.from_bytes(values[:4].ljust(4, b"\x00"), "little", signed=False), "tabs": []}
    attributes = int.from_bytes(values[0:4], "little", signed=False)
    tab_count = int.from_bytes(values[4:8], "little", signed=False)
    cursor = 8
    tabs: list[dict[str, object]] = []
    while len(tabs) < tab_count and cursor + 8 <= len(values):
        tabs.append(
            {
                "position": int.from_bytes(values[cursor : cursor + 4], "little", signed=True),
                "tab_type": values[cursor + 4],
                "tab_type_name": _HWP_TAB_TYPE_NAMES.get(values[cursor + 4], f"CODE_{values[cursor + 4]}"),
                "leader_type": values[cursor + 5],
                "leader_type_name": _HWP_TAB_LEADER_NAMES.get(values[cursor + 5], f"CODE_{values[cursor + 5]}"),
                "reserved": int.from_bytes(values[cursor + 6 : cursor + 8], "little", signed=False),
            }
        )
        cursor += 8
    result: dict[str, object] = {
        "attributes": attributes,
        "auto_tab_left": bool(attributes & 0x01),
        "auto_tab_right": bool(attributes & 0x02),
        "tabs": tabs,
    }
    if cursor < len(values):
        result["trailing_payload"] = values[cursor:]
    return result

def _parse_numbering_payload(payload: bytes) -> dict[str, object]:
    levels: list[dict[str, object]] = []
    values = bytes(payload)
    cursor = 0
    for _index in range(7):
        if cursor + 14 > len(values):
            break
        para_head = int.from_bytes(values[cursor : cursor + 4], "little", signed=False)
        cursor += 4
        width_adjust = int.from_bytes(values[cursor : cursor + 2], "little", signed=True)
        cursor += 2
        text_offset = int.from_bytes(values[cursor : cursor + 2], "little", signed=True)
        cursor += 2
        char_pr_id = int.from_bytes(values[cursor : cursor + 4], "little", signed=True)
        cursor += 4
        format_length = int.from_bytes(values[cursor : cursor + 2], "little", signed=False)
        cursor += 2
        end = min(cursor + format_length * 2, len(values))
        format_bytes = values[cursor:end]
        if len(format_bytes) % 2:
            format_bytes = format_bytes[:-1]
        cursor = end
        levels.append(
            {
                "level_index": len(levels),
                "para_head": para_head,
                "para_head_flags": _bit_flags(para_head),
                "width_adjust": width_adjust,
                "text_offset": text_offset,
                "char_pr_id": None if char_pr_id < 0 else char_pr_id,
                "format_length": format_length,
                "format": format_bytes.decode("utf-16-le", errors="ignore"),
                "format_is_empty": format_length == 0,
            }
        )
    unknown_short = int.from_bytes(values[cursor : cursor + 2].ljust(2, b"\x00"), "little", signed=False)
    cursor += 2 if cursor + 2 <= len(values) else 0
    start_numbers: list[int] = []
    while cursor + 4 <= len(values):
        start_numbers.append(int.from_bytes(values[cursor : cursor + 4], "little", signed=False))
        cursor += 4
    result: dict[str, object] = {
        "levels": levels,
        "unknown_short": unknown_short,
        "unknown_short_bits": _bit_flags(unknown_short, width=16),
        "start_numbers": start_numbers,
        "start_number_map": {
            f"level_{index}": value
            for index, value in enumerate(start_numbers)
        },
    }
    for index, level in enumerate(levels):
        if index < len(start_numbers):
            level["start_number"] = start_numbers[index]
    if cursor < len(values):
        result["trailing_payload"] = values[cursor:]
    result["u16_values"] = _payload_u16_values(values)
    result["u32_values"] = _payload_u32_values(values)
    return result

def _parse_bullet_payload(payload: bytes) -> dict[str, object]:
    values = bytes(payload)
    result: dict[str, object] = {}
    if len(values) >= 4:
        flags = int.from_bytes(values[0:4], "little", signed=False)
        result["flags"] = flags
        result["flag_bits"] = _bit_flags(flags)
    if len(values) >= 6:
        result["width_adjust"] = int.from_bytes(values[4:6], "little", signed=True)
    if len(values) >= 8:
        result["text_offset"] = int.from_bytes(values[6:8], "little", signed=True)
    if len(values) >= 12:
        char_pr_id = int.from_bytes(values[8:12], "little", signed=True)
        result["char_pr_id"] = None if char_pr_id < 0 else char_pr_id
    if len(values) >= 14:
        result["bullet_char"] = values[12:14].decode("utf-16-le", errors="ignore") or "\uf0a1"
        result["bullet_char_codepoint"] = ord(str(result["bullet_char"])[0])
    if len(values) > 14:
        result["unknown_tail"] = values[14:]
    result["flag_bytes"] = list(values[:4].ljust(4, b"\x00"))
    result["u16_values"] = _payload_u16_values(values)
    result["u32_values"] = _payload_u32_values(values)
    return result

def _parse_paragraph_style_payload(payload: bytes) -> dict[str, object]:
    values = bytes(payload)
    result: dict[str, object] = {}
    attributes = int.from_bytes(values[0:4].ljust(4, b"\x00"), "little", signed=False)
    result["attributes"] = attributes
    result["attribute_words"] = {
        "low": int.from_bytes(values[0:2].ljust(2, b"\x00"), "little", signed=False),
        "high": int.from_bytes(values[2:4].ljust(2, b"\x00"), "little", signed=False),
    }
    result["attribute_sections"] = {
        "byte_0": values[0] if len(values) >= 1 else 0,
        "byte_1": values[1] if len(values) >= 2 else 0,
        "byte_2": values[2] if len(values) >= 3 else 0,
        "byte_3": values[3] if len(values) >= 4 else 0,
    }
    result["alignment_code"] = attributes & 0x07
    result["alignment_horizontal"] = {
        0: "JUSTIFY",
        1: "LEFT",
        2: "RIGHT",
        3: "CENTER",
        4: "DISTRIBUTE",
        5: "DISTRIBUTE_SPACE",
    }.get(attributes & 0x07, f"CODE_{attributes & 0x07}")
    word_break_keep_word = bool(attributes & (1 << 8))
    result["break_setting"] = {
        "breakLatinWord": "KEEP_WORD" if word_break_keep_word else "BREAK_WORD",
        "breakNonLatinWord": "KEEP_WORD" if word_break_keep_word else "BREAK_WORD",
        "widowOrphan": "1" if attributes & (1 << 9) else "0",
        "keepWithNext": "1" if attributes & (1 << 10) else "0",
        "keepLines": "1" if attributes & (1 << 11) else "0",
        "pageBreakBefore": "1" if attributes & (1 << 12) else "0",
        "lineWrap": "SQUEEZE" if attributes & (1 << 15) else "BREAK",
    }
    result["auto_spacing"] = {
        "eAsianEng": "1" if attributes & (1 << 23) else "0",
        "eAsianNum": "1" if attributes & (1 << 24) else "0",
    }
    result["attribute_bits"] = {
        "snap_to_grid": bool(attributes & (1 << 7)),
        "word_break_keep_word": word_break_keep_word,
        "widow_orphan": bool(attributes & (1 << 9)),
        "keep_with_next": bool(attributes & (1 << 10)),
        "keep_lines": bool(attributes & (1 << 11)),
        "page_break_before": bool(attributes & (1 << 12)),
        "line_wrap_squeeze": bool(attributes & (1 << 15)),
        "auto_space_easian_eng": bool(attributes & (1 << 23)),
        "auto_space_easian_num": bool(attributes & (1 << 24)),
    }
    if len(values) >= 24:
        result["margins"] = {
            "intent": int.from_bytes(values[4:8], "little", signed=True),
            "left": int.from_bytes(values[8:12], "little", signed=True),
            "right": int.from_bytes(values[12:16], "little", signed=True),
            "prev": int.from_bytes(values[16:20], "little", signed=True),
            "next": int.from_bytes(values[20:24], "little", signed=True),
            "unit": "HWPUNIT",
        }
    if len(values) >= 28:
        result["line_spacing_primary"] = int.from_bytes(values[24:28], "little", signed=True)
    if len(values) >= 30:
        result["tab_def_id"] = int.from_bytes(values[28:30], "little", signed=False)
    if len(values) >= 32:
        result["numbering_bullet_id"] = int.from_bytes(values[30:32], "little", signed=False)
    if len(values) >= 34:
        result["border_fill_id"] = int.from_bytes(values[32:34], "little", signed=False)
    if len(values) >= 42:
        result["border_offsets"] = {
            "left": int.from_bytes(values[34:36], "little", signed=False),
            "right": int.from_bytes(values[36:38], "little", signed=False),
            "top": int.from_bytes(values[38:40], "little", signed=False),
            "bottom": int.from_bytes(values[40:42], "little", signed=False),
        }
    if len(values) >= 46:
        result["attribute2"] = int.from_bytes(values[42:46], "little", signed=False)
    if len(values) >= 50:
        result["attribute3"] = int.from_bytes(values[46:50], "little", signed=False)
    if len(values) >= 54:
        result["line_spacing"] = int.from_bytes(values[50:54], "little", signed=True)
        result["line_spacing_secondary"] = result["line_spacing"]
    if "line_spacing" not in result and "line_spacing_primary" in result:
        result["line_spacing"] = result["line_spacing_primary"]
    numbering_bullet_id = result.get("numbering_bullet_id", 0)
    normalized_numbering_bullet_id = numbering_bullet_id if isinstance(numbering_bullet_id, int) else 0
    result["heading"] = {
        "type": "NUMBER" if normalized_numbering_bullet_id > 0 else "NONE",
        "idRef": str(normalized_numbering_bullet_id),
        "level": "0",
        "source": "derived_from_numbering_bullet_id",
    }
    result["u32_values"] = _payload_u32_values(values)
    return result

def _parse_style_definition_payload(payload: bytes, *, style_id: int | None = None) -> dict[str, object]:
    cursor = 0
    name, cursor = _parse_length_prefixed_utf16(payload, cursor)
    english_name, cursor = _parse_length_prefixed_utf16(payload, cursor)
    style_word = int.from_bytes(payload[cursor : cursor + 2].ljust(2, b"\x00"), "little", signed=False)
    cursor += 2 if cursor + 2 <= len(payload) else 0
    lang_id = int.from_bytes(payload[cursor : cursor + 2].ljust(2, b"\x00"), "little", signed=False)
    cursor += 2 if cursor + 2 <= len(payload) else 0
    next_style_id = int.from_bytes(payload[cursor : cursor + 2].ljust(2, b"\x00"), "little", signed=False)
    cursor += 2 if cursor + 2 <= len(payload) else 0
    para_pr_id = int.from_bytes(payload[cursor : cursor + 2].ljust(2, b"\x00"), "little", signed=False)
    cursor += 2 if cursor + 2 <= len(payload) else 0
    char_pr_id = int.from_bytes(payload[cursor : cursor + 2].ljust(2, b"\x00"), "little", signed=False)
    style_type_code = style_word & 0x00FF
    resolved_style_id = style_word >> 8 if style_id is None else style_id
    result: dict[str, object] = {
        "name": name or f"Style {resolved_style_id}",
        "english_name": english_name,
        "style_id": resolved_style_id,
        "style_word": style_word,
        "style_word_bytes": {
            "low": style_word & 0x00FF,
            "high": (style_word >> 8) & 0x00FF,
        },
        "style_slot": resolved_style_id,
        "style_type_code": style_type_code,
        "style_type": {0: "PARA", 1: "CHAR"}.get(style_type_code, f"CODE_{style_type_code}"),
        "lang_id": lang_id,
        "next_style_id": next_style_id,
        "para_pr_id": para_pr_id,
        "char_pr_id": char_pr_id,
        "name_length": len(name),
        "english_name_length": len(english_name),
        "lang_id_hex": f"0x{lang_id:04X}",
    }
    if cursor < len(payload):
        result["trailing_payload"] = payload[cursor:]
    result["u16_values"] = _payload_u16_values(payload)
    return result

def _parse_shape_component_common_payload(payload: bytes) -> dict[str, object]:
    values = bytes(payload)
    result: dict[str, object] = {
        "reference_bounds": {
            "x": int.from_bytes(values[20:24].ljust(4, b"\x00"), "little", signed=False),
            "y": int.from_bytes(values[24:28].ljust(4, b"\x00"), "little", signed=False),
            "width": int.from_bytes(values[28:32].ljust(4, b"\x00"), "little", signed=False),
            "height": int.from_bytes(values[32:36].ljust(4, b"\x00"), "little", signed=False),
        }
    }
    if len(values) >= 202:
        result["line_color"] = _parse_colorref(values[196:200])
        result["line_width"] = int.from_bytes(values[200:202], "little", signed=False)
    if len(values) >= 217:
        result["fill_color"] = _parse_colorref(values[213:217])
    return result

def _shape_specific_payload_size(tag_id: int) -> int:
    if tag_id == TAG_SHAPE_COMPONENT_LINE:
        return 20
    if tag_id == TAG_SHAPE_COMPONENT_RECTANGLE:
        return 33
    if tag_id == TAG_SHAPE_COMPONENT_ELLIPSE:
        return 60
    if tag_id == TAG_SHAPE_COMPONENT_ARC:
        return 25
    return 0

def _shape_common_payload_size(tag_id: int) -> int:
    if tag_id == TAG_SHAPE_COMPONENT_LINE:
        return 11
    if tag_id in {
        TAG_SHAPE_COMPONENT_ELLIPSE,
        TAG_SHAPE_COMPONENT_ARC,
        TAG_SHAPE_COMPONENT_POLYGON,
        TAG_SHAPE_COMPONENT_TEXTART,
    }:
        return 0
    if tag_id in {
        TAG_SHAPE_COMPONENT_RECTANGLE,
        TAG_SHAPE_COMPONENT_CURVE,
        TAG_SHAPE_COMPONENT_CONTAINER,
    }:
        return 32
    return 0

def _parse_line_shape_payload(payload: bytes) -> dict[str, object]:
    values = bytes(payload)
    return {
        "start": {
            "x": int.from_bytes(values[0:4].ljust(4, b"\x00"), "little", signed=True),
            "y": int.from_bytes(values[4:8].ljust(4, b"\x00"), "little", signed=True),
        },
        "end": {
            "x": int.from_bytes(values[8:12].ljust(4, b"\x00"), "little", signed=True),
            "y": int.from_bytes(values[12:16].ljust(4, b"\x00"), "little", signed=True),
        },
        "attributes": int.from_bytes(values[16:20].ljust(4, b"\x00"), "little", signed=False),
        "u16_values": _payload_u16_values(values),
        "u32_values": _payload_u32_values(values),
    }

def _parse_connectline_shape_payload(payload: bytes) -> dict[str, object]:
    values = bytes(payload)
    return {
        "raw_payload": values,
        "u16_values": _payload_u16_values(values),
        "u32_values": _payload_u32_values(values),
        "variant": "hancom_connectline",
    }

def _parse_rectangle_shape_payload(payload: bytes) -> dict[str, object]:
    values = bytes(payload)
    return {
        "corner_radius": int.from_bytes(values[0:4].ljust(4, b"\x00"), "little", signed=False) if values else 0,
        "raw_payload": values,
        "u16_values": _payload_u16_values(values),
        "u32_values": _payload_u32_values(values),
    }

def _parse_ellipse_shape_payload(payload: bytes) -> dict[str, object]:
    values = bytes(payload)
    u32_values = list(struct.unpack("<15i", values[:60].ljust(60, b"\x00"))) if values else [0] * 15
    return {
        "arc_flags": u32_values[0],
        "center": {"x": u32_values[1], "y": u32_values[2]},
        "axis1": {"x": u32_values[3], "y": u32_values[4]},
        "axis2": {"x": u32_values[5], "y": u32_values[6]},
        "start": {"x": u32_values[7], "y": u32_values[8]},
        "end": {"x": u32_values[9], "y": u32_values[10]},
        "start_2": {"x": u32_values[11], "y": u32_values[12]},
        "end_2": {"x": u32_values[13], "y": u32_values[14]},
        "u16_values": _payload_u16_values(values),
        "u32_values": _payload_u32_values(values),
    }

def _parse_arc_shape_payload(payload: bytes) -> dict[str, object]:
    values = bytes(payload)
    u32_values = list(struct.unpack("<7i", values[:28].ljust(28, b"\x00"))) if values else [0] * 7
    return {
        "arc_type": u32_values[0],
        "center": {"x": u32_values[1], "y": u32_values[2]},
        "axis1": {"x": u32_values[3], "y": u32_values[4]},
        "axis2": {"x": u32_values[5], "y": u32_values[6]},
        "u16_values": _payload_u16_values(values),
        "u32_values": _payload_u32_values(values),
    }

def _parse_polygon_shape_payload(payload: bytes) -> dict[str, object]:
    values = bytes(payload)
    point_count = int.from_bytes(values[0:2].ljust(2, b"\x00"), "little", signed=False)
    cursor = 2
    x_values: list[int] = []
    y_values: list[int] = []
    for _ in range(point_count):
        if cursor + 4 > len(values):
            break
        x_values.append(int.from_bytes(values[cursor : cursor + 4], "little", signed=True))
        cursor += 4
    for _ in range(point_count):
        if cursor + 4 > len(values):
            break
        y_values.append(int.from_bytes(values[cursor : cursor + 4], "little", signed=True))
        cursor += 4
    return {
        "point_count": point_count,
        "points": [{"x": x_values[index], "y": y_values[index]} for index in range(min(len(x_values), len(y_values)))],
        "trailing_payload": values[cursor:] if cursor < len(values) else b"",
        "u16_values": _payload_u16_values(values),
        "u32_values": _payload_u32_values(values),
    }

def _parse_curve_shape_payload(payload: bytes) -> dict[str, object]:
    values = bytes(payload)
    return {
        "raw_payload": values,
        "u16_values": _payload_u16_values(values),
        "u32_values": _payload_u32_values(values),
    }

def _parse_container_shape_payload(payload: bytes) -> dict[str, object]:
    values = bytes(payload)
    return {
        "raw_payload": values,
        "u16_values": _payload_u16_values(values),
        "u32_values": _payload_u32_values(values),
    }
