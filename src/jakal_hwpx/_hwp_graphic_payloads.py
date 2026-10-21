from __future__ import annotations

import struct

from ._hwp_constants import (
    DEFAULT_HWP_SHAPE_TEXT,
    _CONTROL_NATIVE_METADATA_PREFIX,
    _DEFAULT_ARC_PAYLOAD,
    _DEFAULT_COMPLEX_SHAPE_COMPONENT_LINE_WIDTH,
    _DEFAULT_ELLIPSE_PAYLOAD,
    _DEFAULT_OLE_LINE_WIDTH,
    _DEFAULT_POLYGON_PAYLOAD,
    _DEFAULT_SHAPE_COMPONENT,
    _DEFAULT_SHAPE_CTRL_HEADER,
    _DEFAULT_SHAPE_LINE_PAYLOAD,
    _DEFAULT_SHAPE_LINE_WIDTH,
    _DEFAULT_SHAPE_RECTANGLE_PAYLOAD,
    _DEFAULT_TEXTART_FONT_NAME,
    _DEFAULT_TEXTART_FONT_STYLE,
    _HWP_COMPLEX_SHAPE_TEMPLATES,
    _HWP_DRAW_ASPECT_CODES,
    _HWP_LINE_ARROW_CODES,
    _HWP_LINE_ARROW_SIZE_CODES,
    _HWP_LINE_END_CAP_CODES,
    _HWP_LINE_OUTLINE_STYLE_CODES,
    _HWP_LINE_STYLE_CODES,
    _HWP_OLE_OBJECT_TYPE_CODES,
    _HWP_PICTURE_EFFECT_CODES,
    _HWP_PICTURE_EFFECT_NAMES,
    _MEMO_SHAPE_NATIVE_METADATA_PREFIX,
    _SHAPE_NATIVE_METADATA_PREFIX,
    _TEXTART_COMPONENT_REFERENCE_EXTENT,
)
from ._hwp_page_payloads import _build_colorref, _build_rgb_color


def _build_control_id_payload(control_id: str) -> bytes:
    normalized = control_id[:4].ljust(4)
    return normalized.encode("latin1", errors="replace")[::-1]


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
        uses_template_extents = kind in {"ellipse", "arc", "polygon"}
        if kind == "textart":
            if len(payload) >= 36:
                payload[28:32] = int(width).to_bytes(4, "little", signed=False)
                payload[32:36] = int(height).to_bytes(4, "little", signed=False)
            if len(payload) >= 40:
                payload[36:40] = (0x00080000).to_bytes(4, "little", signed=False)
            if len(payload) >= 48:
                payload[40:44] = (int(width) << 15).to_bytes(4, "little", signed=False)
                payload[44:48] = (int(height) << 15).to_bytes(4, "little", signed=False)
            if len(payload) >= 108 and _TEXTART_COMPONENT_REFERENCE_EXTENT:
                payload[100:108] = struct.pack("<d", float(width) / float(_TEXTART_COMPONENT_REFERENCE_EXTENT))
            if len(payload) >= 140 and _TEXTART_COMPONENT_REFERENCE_EXTENT:
                payload[132:140] = struct.pack("<d", float(height) / float(_TEXTART_COMPONENT_REFERENCE_EXTENT))
        elif not uses_template_extents and len(payload) >= 24:
            payload[20:24] = int(width).to_bytes(4, "little", signed=False)
        if kind != "textart" and not uses_template_extents and len(payload) >= 28:
            payload[24:28] = int(height).to_bytes(4, "little", signed=False)
        if kind != "textart" and not uses_template_extents and len(payload) >= 32:
            payload[28:32] = int(width).to_bytes(4, "little", signed=False)
        if kind != "textart" and not uses_template_extents and len(payload) >= 36:
            payload[32:36] = int(height).to_bytes(4, "little", signed=False)
        if len(payload) >= 200:
            payload[196:200] = _build_colorref(line_color)
        if len(payload) >= 202:
            resolved_line_width = (
                _DEFAULT_COMPLEX_SHAPE_COMPONENT_LINE_WIDTH if kind in {"ellipse", "arc", "polygon"} else line_width
            )
            payload[200:202] = max(0, min(int(resolved_line_width), 0xFFFF)).to_bytes(2, "little", signed=False)
        if kind == "textart" and len(payload) >= 217:
            payload[213:217] = _build_colorref(fill_color)
        return bytes(payload)
    payload = bytearray(_DEFAULT_SHAPE_COMPONENT)
    payload[20:24] = int(width).to_bytes(4, "little", signed=False)
    payload[24:28] = int(height).to_bytes(4, "little", signed=False)
    payload[28:32] = int(width).to_bytes(4, "little", signed=False)
    payload[32:36] = int(height).to_bytes(4, "little", signed=False)
    return bytes(payload)


def _build_prefixed_metadata_payload(prefix: str, **fields: object) -> bytes:
    serialized_fields: list[str] = []
    for name, value in fields.items():
        if value is None:
            continue
        if isinstance(value, bool):
            rendered = "1" if value else "0"
        else:
            rendered = str(value)
        serialized_fields.append(f"{name}={rendered}")
    serialized = ";".join(serialized_fields)
    text = f"{prefix};{serialized}" if serialized else prefix
    return text.encode("utf-16-le")


def _build_shape_native_metadata_payload(**fields: object) -> bytes:
    return _build_prefixed_metadata_payload(_SHAPE_NATIVE_METADATA_PREFIX, **fields)


def _build_control_native_metadata_payload(**fields: object) -> bytes:
    return _build_prefixed_metadata_payload(_CONTROL_NATIVE_METADATA_PREFIX, **fields)


def _build_memo_shape_native_metadata_payload(**fields: object) -> bytes:
    return _build_prefixed_metadata_payload(_MEMO_SHAPE_NATIVE_METADATA_PREFIX, **fields)


def _graphic_control_out_margin_values(payload: bytes) -> tuple[int, int, int, int]:
    if len(payload) < 36:
        return (0, 0, 0, 0)
    return (
        int.from_bytes(payload[28:30], "little", signed=False),
        int.from_bytes(payload[30:32], "little", signed=False),
        int.from_bytes(payload[32:34], "little", signed=False),
        int.from_bytes(payload[34:36], "little", signed=False),
    )


def _parse_graphic_control_out_margins(payload: bytes) -> dict[str, int]:
    left, right, top, bottom = _graphic_control_out_margin_values(payload)
    if not any((left, right, top, bottom)):
        return {}
    return {
        "left": left,
        "right": right,
        "top": top,
        "bottom": bottom,
    }


def _parse_graphic_control_layout(payload: bytes) -> dict[str, str]:
    if len(payload) < 16:
        return {}
    flow_bits = payload[7] & 0x03
    result = {
        "textWrap": "TOP_AND_BOTTOM" if payload[6] & 0x20 else "SQUARE",
        "textFlow": "LEFT_ONLY" if flow_bits == 0x01 else "RIGHT_ONLY" if flow_bits == 0x02 else "BOTH_SIDES",
        "treatAsChar": "1" if payload[4] & 0x01 else "0",
        "vertRelTo": "PARA" if payload[4] & 0x10 else "PAPER",
        "horzRelTo": "COLUMN" if payload[5] & 0x02 else "PAPER",
    }
    if payload[4] & 0x20:
        result["vertAlign"] = "CENTER"
    if payload[5] & 0x08:
        result["horzAlign"] = "RIGHT"
    vert_offset = int.from_bytes(payload[8:12], "little", signed=True)
    horz_offset = int.from_bytes(payload[12:16], "little", signed=True)
    if vert_offset:
        result["vertOffset"] = str(vert_offset)
    if horz_offset:
        result["horzOffset"] = str(horz_offset)
    return result


def _set_graphic_control_out_margins_payload(
    payload: bytes,
    *,
    left: int | str | None = None,
    right: int | str | None = None,
    top: int | str | None = None,
    bottom: int | str | None = None,
) -> bytes:
    updated = bytearray(payload)
    if len(updated) < 36:
        updated.extend(b"\x00" * (36 - len(updated)))
    current_left, current_right, current_top, current_bottom = _graphic_control_out_margin_values(bytes(updated))
    resolved = (
        current_left if left is None else int(left),
        current_right if right is None else int(right),
        current_top if top is None else int(top),
        current_bottom if bottom is None else int(bottom),
    )
    for offset, value in zip((28, 30, 32, 34), resolved, strict=True):
        updated[offset : offset + 2] = max(0, min(int(value), 0xFFFF)).to_bytes(2, "little", signed=False)
    return bytes(updated)


def _set_graphic_control_layout_payload(
    payload: bytes,
    *,
    text_wrap: str | None = None,
    text_flow: str | None = None,
    treat_as_char: bool | None = None,
    vert_rel_to: str | None = None,
    horz_rel_to: str | None = None,
    vert_align: str | None = None,
    horz_align: str | None = None,
    vert_offset: int | str | None = None,
    horz_offset: int | str | None = None,
) -> bytes:
    updated = bytearray(payload)
    if len(updated) < 16:
        updated.extend(b"\x00" * (16 - len(updated)))
    if text_wrap == "TOP_AND_BOTTOM":
        updated[6] |= 0x20
    elif text_wrap == "SQUARE":
        updated[6] &= ~0x20
    if text_flow == "BOTH_SIDES":
        updated[7] &= ~0x03
    elif text_flow == "LEFT_ONLY":
        updated[7] = (updated[7] & ~0x03) | 0x01
    elif text_flow == "RIGHT_ONLY":
        updated[7] = (updated[7] & ~0x03) | 0x02
    if treat_as_char is True:
        updated[4] |= 0x01
    elif treat_as_char is False:
        updated[4] &= ~0x01
    if vert_rel_to == "PARA":
        updated[4] |= 0x10
    elif vert_rel_to == "PAPER":
        updated[4] &= ~0x10
    if horz_rel_to == "COLUMN":
        updated[5] |= 0x02
    elif horz_rel_to == "PAPER":
        updated[5] &= ~0x02
    if vert_align == "CENTER":
        updated[4] |= 0x20
    elif vert_align == "TOP":
        updated[4] &= ~0x20
    if horz_align == "RIGHT":
        updated[5] |= 0x08
    elif horz_align == "LEFT":
        updated[5] &= ~0x08
    if vert_offset is not None:
        updated[8:12] = int(vert_offset).to_bytes(4, "little", signed=True)
    if horz_offset is not None:
        updated[12:16] = int(horz_offset).to_bytes(4, "little", signed=True)
    return bytes(updated)


def _shape_component_rotation_values(payload: bytes) -> tuple[int, int, int]:
    if len(payload) < 50:
        return (0, 0, 0)
    return (
        int.from_bytes(payload[40:42], "little", signed=False),
        int.from_bytes(payload[42:46], "little", signed=True),
        int.from_bytes(payload[46:50], "little", signed=True),
    )


def _shape_component_default_rotation_center(payload: bytes) -> tuple[int, int]:
    if len(payload) < 36:
        return (0, 0)
    return (
        int.from_bytes(payload[28:32], "little", signed=False) // 2,
        int.from_bytes(payload[32:36], "little", signed=False) // 2,
    )


def _parse_shape_component_rotation_payload(payload: bytes, *, include_center: bool = True) -> dict[str, str]:
    angle, center_x, center_y = _shape_component_rotation_values(payload)
    default_center_x, default_center_y = _shape_component_default_rotation_center(payload)
    has_nondefault_center = center_x != default_center_x or center_y != default_center_y
    if angle == 0 and not has_nondefault_center:
        return {}
    result: dict[str, str] = {"angle": str(angle)}
    if include_center and has_nondefault_center:
        result["centerX"] = str(center_x)
        result["centerY"] = str(center_y)
    return result


def _set_shape_component_rotation_payload(
    payload: bytes,
    *,
    angle: int | str | None = None,
    center_x: int | str | None = None,
    center_y: int | str | None = None,
) -> bytes:
    updated = bytearray(payload)
    if len(updated) < 50:
        updated.extend(b"\x00" * (50 - len(updated)))
    current_angle, current_center_x, current_center_y = _shape_component_rotation_values(bytes(updated))
    resolved_angle = current_angle if angle is None else int(angle)
    resolved_center_x = current_center_x if center_x is None else int(center_x)
    resolved_center_y = current_center_y if center_y is None else int(center_y)
    updated[40:42] = max(0, min(resolved_angle, 0xFFFF)).to_bytes(2, "little", signed=False)
    updated[42:46] = int(resolved_center_x).to_bytes(4, "little", signed=True)
    updated[46:50] = int(resolved_center_y).to_bytes(4, "little", signed=True)
    return bytes(updated)


def _picture_crop_values(payload: bytes) -> tuple[int, int, int, int]:
    if len(payload) < 60:
        return (0, 0, 0, 0)
    return (
        int.from_bytes(payload[44:48], "little", signed=False),
        int.from_bytes(payload[48:52], "little", signed=False),
        int.from_bytes(payload[52:56], "little", signed=False),
        int.from_bytes(payload[56:60], "little", signed=False),
    )


def _picture_default_clip_size(payload: bytes) -> tuple[int, int]:
    if len(payload) < 36:
        return (0, 0)
    return (
        int.from_bytes(payload[28:32], "little", signed=False),
        int.from_bytes(payload[32:36], "little", signed=False),
    )


def _parse_picture_crop_payload(payload: bytes) -> dict[str, int]:
    left, top, right, bottom = _picture_crop_values(payload)
    default_right, default_bottom = _picture_default_clip_size(payload)
    if left == 0 and top == 0:
        if right == 0 and bottom == 0:
            return {}
        if default_right and default_bottom and right == default_right and bottom == default_bottom:
            return {}
    return {
        "left": left,
        "right": right,
        "top": top,
        "bottom": bottom,
    }


def _set_picture_crop_payload(
    payload: bytes,
    *,
    left: int | str | None = None,
    right: int | str | None = None,
    top: int | str | None = None,
    bottom: int | str | None = None,
) -> bytes:
    updated = bytearray(payload)
    if len(updated) < 60:
        updated.extend(b"\x00" * (60 - len(updated)))
    current_left, current_top, current_right, current_bottom = _picture_crop_values(bytes(updated))
    resolved = (
        current_left if left is None else int(left),
        current_top if top is None else int(top),
        current_right if right is None else int(right),
        current_bottom if bottom is None else int(bottom),
    )
    for offset, value in zip((44, 48, 52, 56), resolved, strict=True):
        updated[offset : offset + 4] = max(0, int(value)).to_bytes(4, "little", signed=False)
    return bytes(updated)


def _picture_image_adjustment_values(payload: bytes) -> tuple[int, int, int, int]:
    contrast = payload[68] if len(payload) >= 69 else 0
    bright = payload[69] if len(payload) >= 70 else 0
    effect_code = payload[70] if len(payload) >= 71 else 0
    alpha = payload[90] if len(payload) >= 91 else 0
    return (contrast, bright, effect_code, alpha)


def _picture_effect_name(code: int) -> str:
    return _HWP_PICTURE_EFFECT_NAMES.get(code, f"CODE_{code}")


def _picture_effect_code(effect: str | None) -> int | None:
    if effect is None:
        return None
    return _HWP_PICTURE_EFFECT_CODES.get(str(effect).strip().upper())


def _parse_picture_image_adjustment_payload(payload: bytes) -> dict[str, str]:
    contrast, bright, effect_code, alpha = _picture_image_adjustment_values(payload)
    if contrast == 0 and bright == 0 and effect_code == 0 and alpha == 0:
        return {}
    return {
        "bright": str(bright),
        "contrast": str(contrast),
        "effect": _picture_effect_name(effect_code),
        "alpha": str(alpha),
    }


def _set_picture_image_adjustment_payload(
    payload: bytes,
    *,
    bright: int | str | None = None,
    contrast: int | str | None = None,
    effect_code: int | None = None,
    alpha: int | str | None = None,
) -> bytes:
    updated = bytearray(payload)
    if len(updated) < 91:
        updated.extend(b"\x00" * (91 - len(updated)))
    current_contrast, current_bright, current_effect_code, current_alpha = _picture_image_adjustment_values(bytes(updated))
    updated[68] = max(0, min(0xFF, current_contrast if contrast is None else int(contrast)))
    updated[69] = max(0, min(0xFF, current_bright if bright is None else int(bright)))
    updated[70] = max(0, min(0xFF, current_effect_code if effect_code is None else int(effect_code)))
    updated[90] = max(0, min(0xFF, current_alpha if alpha is None else int(alpha)))
    return bytes(updated)


def _set_picture_line_payload(
    payload: bytes,
    *,
    color: str | None = None,
    width: int | None = None,
) -> bytes:
    updated = bytearray(payload)
    if len(updated) < 8:
        updated.extend(b"\x00" * (8 - len(updated)))
    if color is not None:
        updated[0:4] = _build_rgb_color(color)
    if width is not None:
        updated[4:8] = max(0, int(width)).to_bytes(4, "little", signed=False)
    return bytes(updated)


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
    raw_text = text
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
    payload[14:18] = _build_rgb_color(line_color)
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
    if kind in {"line", "connectLine"}:
        common = _build_shape_line_info_payload(color=line_color, width=_DEFAULT_SHAPE_LINE_WIDTH)
        payload = bytearray(_DEFAULT_SHAPE_LINE_PAYLOAD)
        payload[8:12] = int(width).to_bytes(4, "little", signed=False)
        payload[12:16] = int(height).to_bytes(4, "little", signed=False)
        return common + bytes(payload)
    if kind == "textart":
        return _build_textart_specific_payload(text)
    return b""

for _compat_name in ["_build_graphic_control_header_payload","_build_minimal_shape_component_payload","_shape_control_header_template","_build_shape_component_payload","_build_prefixed_metadata_payload","_build_shape_native_metadata_payload","_build_control_native_metadata_payload","_build_memo_shape_native_metadata_payload","_graphic_control_out_margin_values","_parse_graphic_control_out_margins","_parse_graphic_control_layout","_set_graphic_control_out_margins_payload","_set_graphic_control_layout_payload","_shape_component_rotation_values","_shape_component_default_rotation_center","_parse_shape_component_rotation_payload","_set_shape_component_rotation_payload","_picture_crop_values","_picture_default_clip_size","_parse_picture_crop_payload","_set_picture_crop_payload","_picture_image_adjustment_values","_picture_effect_name","_picture_effect_code","_parse_picture_image_adjustment_payload","_set_picture_image_adjustment_payload","_set_picture_line_payload","_build_line_attribute_flags","_build_shape_line_info_payload","_build_shape_fill_payload","_build_filled_shape_common_payload","_build_ellipse_specific_payload","_build_arc_specific_payload","_build_polygon_specific_payload","_build_textart_specific_payload","_build_ole_attribute_flags","_build_ole_specific_payload","_build_shape_specific_payload"]:
    globals()[_compat_name].__module__ = "jakal_hwpx.hwp_binary"

