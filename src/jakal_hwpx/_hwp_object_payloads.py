from __future__ import annotations

import json
import struct
from collections.abc import Sequence
from typing import Any, cast

from .hwp_binary import (
    TAG_CTRL_DATA,
    TAG_FORM_OBJECT,
    TAG_MEMO_LIST,
    TAG_SHAPE_COMPONENT,
    TAG_SHAPE_COMPONENT_ARC,
    TAG_SHAPE_COMPONENT_ELLIPSE,
    TAG_SHAPE_COMPONENT_LINE,
    TAG_SHAPE_COMPONENT_RECTANGLE,
    RecordNode,
    _build_control_native_metadata_payload,
    _build_shape_native_metadata_payload,
    _parse_colorref,
    _payload_u16_values,
    _payload_u32_values,
)

_HWP_SHAPE_LINE_SPECIFIC_SIZE = 20
_HWP_SHAPE_RECTANGLE_SPECIFIC_SIZE = 33
_HWP_SHAPE_ELLIPSE_SPECIFIC_SIZE = 60
_HWP_SHAPE_ARC_SPECIFIC_SIZE = 28
_HWP_TEXTART_PREFIX_PAYLOAD = bytes.fromhex("00000000000000005d370000000000005d3700005d370000000000005d370000")
_HWP_TEXTART_TRAILING_PAYLOAD = bytes.fromhex(
    "01000000000000007800000064000000000000000000000000000000000000000000000000000000"
)
_HWP_OLE_DRAW_ASPECT_NAMES = {
    1: "CONTENT",
    2: "THUMBNAIL",
    4: "ICON",
    8: "DOCPRINT",
}
_HWP_OLE_OBJECT_TYPE_NAMES = {
    0: "UNKNOWN",
    1: "EMBEDDED",
    2: "LINK",
    3: "STATIC",
    4: "EQUATION",
}
_SHAPE_NATIVE_METADATA_PREFIX = "JAKAL_SHAPE_META"
_CONTROL_NATIVE_METADATA_PREFIX = "JAKAL_CTRL_META"
_HANCOM_CONNECTLINE_COMPONENT_SIGNATURE = b"loc$loc$"

def _parse_object_description(payload: bytes) -> str:
    if len(payload) < 46:
        return ""
    char_count = int.from_bytes(payload[44:46], "little", signed=False)
    encoded = payload[46 : 46 + char_count * 2]
    return encoded.decode("utf-16-le", errors="ignore")


def _replace_object_description(payload: bytes, description: str | None) -> bytes:
    encoded = str(description or "").encode("utf-16-le")
    base = bytes(payload[:44]).ljust(44, b"\x00")
    return base + (len(encoded) // 2).to_bytes(2, "little", signed=False) + encoded


def _set_payload_size(payload: bytes, *, width: int | None = None, height: int | None = None, width_offset: int, height_offset: int) -> bytes:
    updated = bytearray(payload)
    required = max(width_offset + 4 if width is not None else 0, height_offset + 4 if height is not None else 0)
    if required > len(updated):
        updated.extend(b"\x00" * (required - len(updated)))
    if width is not None:
        updated[width_offset : width_offset + 4] = int(width).to_bytes(4, "little", signed=False)
    if height is not None:
        updated[height_offset : height_offset + 4] = int(height).to_bytes(4, "little", signed=False)
    return bytes(updated)


def _parse_raw_payload_fields(payload: bytes) -> dict[str, object]:
    values = bytes(payload)
    result: dict[str, object] = {
        "raw_payload": values,
        "u16_values": _payload_u16_values(values),
        "u32_values": _payload_u32_values(values),
    }
    decoded = values.decode("utf-16-le", errors="ignore").replace("\x00", "")
    if decoded:
        result["utf16_text"] = decoded
    return result


def _parse_shape_fill_payload(payload: bytes) -> dict[str, str]:
    if len(payload) < 16:
        return {}
    fill_type = int.from_bytes(payload[0:4], "little", signed=False)
    result: dict[str, str] = {}
    if fill_type & 0x00000001:
        result["faceColor"] = _parse_colorref(payload[4:8])
        result["hatchColor"] = _parse_colorref(payload[8:12])
    return result


def _parse_shape_native_metadata(payload: bytes) -> dict[str, str]:
    if not payload:
        return {}
    text = payload.decode("utf-16-le", errors="ignore")
    if not text.startswith(_SHAPE_NATIVE_METADATA_PREFIX):
        return {}
    metadata: dict[str, str] = {}
    for field in text.split(";")[1:]:
        key, separator, value = field.partition("=")
        if separator and key:
            metadata[key] = value
    return metadata


def _collect_shape_native_metadata(control_node: RecordNode) -> dict[str, str]:
    metadata: dict[str, str] = {}
    for child in control_node.children:
        if child.tag_id != TAG_CTRL_DATA:
            continue
        metadata.update(_parse_shape_native_metadata(child.payload))
    return metadata


def _metadata_prefixed_str_map(metadata: dict[str, str], prefix: str) -> dict[str, str]:
    result: dict[str, str] = {}
    for key, value in metadata.items():
        if key.startswith(prefix):
            result[key[len(prefix) :]] = value
    return result


def _metadata_prefixed_int_map(metadata: dict[str, str], prefix: str) -> dict[str, int]:
    result: dict[str, int] = {}
    for key, value in metadata.items():
        if not key.startswith(prefix):
            continue
        try:
            result[key[len(prefix) :]] = int(value)
        except ValueError:
            continue
    return result


def _shape_metadata_node(control_node: RecordNode, *, create: bool = False) -> RecordNode | None:
    for child in control_node.children:
        if child.tag_id == TAG_CTRL_DATA and _parse_shape_native_metadata(child.payload):
            return child
    if not create:
        return None
    node = RecordNode(tag_id=TAG_CTRL_DATA, level=control_node.level + 1, payload=_build_shape_native_metadata_payload())
    insert_at = next((index for index, child in enumerate(control_node.children) if child.tag_id == TAG_SHAPE_COMPONENT), len(control_node.children))
    control_node.children.insert(insert_at, node)
    return node


def _update_shape_native_metadata(
    control_node: RecordNode,
    *,
    updates: dict[str, object | None],
    remove_prefixes: Sequence[str] = (),
) -> None:
    metadata = _collect_shape_native_metadata(control_node)
    for prefix in remove_prefixes:
        for key in list(metadata):
            if key.startswith(prefix):
                metadata.pop(key, None)
    for key, value in updates.items():
        if value is None:
            metadata.pop(key, None)
            continue
        if isinstance(value, bool):
            metadata[key] = "1" if value else "0"
        else:
            metadata[key] = str(value)
    node = _shape_metadata_node(control_node, create=bool(metadata))
    if node is None:
        return
    if metadata:
        node.payload = _build_shape_native_metadata_payload(**metadata)
        return
    control_node.children = [child for child in control_node.children if child is not node]


def _parse_control_native_metadata(payload: bytes) -> dict[str, str]:
    if not payload:
        return {}
    text = payload.decode("utf-16-le", errors="ignore")
    if not text.startswith(_CONTROL_NATIVE_METADATA_PREFIX):
        return {}
    metadata: dict[str, str] = {}
    for field in text.split(";")[1:]:
        key, separator, value = field.partition("=")
        if separator and key:
            metadata[key] = value
    return metadata


def _collect_control_native_metadata(control_node: RecordNode) -> dict[str, str]:
    metadata: dict[str, str] = {}
    for child in control_node.children:
        if child.tag_id != TAG_CTRL_DATA:
            continue
        metadata.update(_parse_control_native_metadata(child.payload))
    return metadata


def _control_metadata_node(control_node: RecordNode, *, create: bool = False) -> RecordNode | None:
    for child in control_node.children:
        if child.tag_id == TAG_CTRL_DATA and _parse_control_native_metadata(child.payload):
            return child
    if not create:
        return None
    node = RecordNode(tag_id=TAG_CTRL_DATA, level=control_node.level + 1, payload=_build_control_native_metadata_payload())
    insert_at = next(
        (index for index, child in enumerate(control_node.children) if child.tag_id in {TAG_FORM_OBJECT, TAG_MEMO_LIST}),
        len(control_node.children),
    )
    control_node.children.insert(insert_at, node)
    return node


def _update_control_native_metadata(
    control_node: RecordNode,
    *,
    updates: dict[str, object | None],
    remove_prefixes: Sequence[str] = (),
) -> None:
    metadata = _collect_control_native_metadata(control_node)
    for prefix in remove_prefixes:
        for key in list(metadata):
            if key.startswith(prefix):
                metadata.pop(key, None)
    for key, value in updates.items():
        if value is None:
            metadata.pop(key, None)
            continue
        if isinstance(value, bool):
            metadata[key] = "1" if value else "0"
        else:
            metadata[key] = str(value)
    node = _control_metadata_node(control_node, create=bool(metadata))
    if node is None:
        return
    if metadata:
        node.payload = _build_control_native_metadata_payload(**metadata)
        return
    control_node.children = [child for child in control_node.children if child is not node]


def _metadata_bool_value(metadata: dict[str, str], key: str, default: bool = False) -> bool:
    value = metadata.get(key)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "y", "on"}


def _metadata_bool_literal_is_valid(value: str | None) -> bool:
    if value is None:
        return True
    return value.strip().lower() in {"0", "1", "false", "true", "no", "yes", "n", "y", "off", "on"}


def _metadata_json_value(metadata: dict[str, str], key: str) -> tuple[object | None, str | None]:
    raw_value = metadata.get(key)
    if not raw_value:
        return None, None
    try:
        return json.loads(raw_value), None
    except json.JSONDecodeError as exc:
        return None, str(exc)


def _metadata_json_list(metadata: dict[str, str], key: str) -> list[object]:
    value, error = _metadata_json_value(metadata, key)
    if error is not None:
        return []
    return list(value) if isinstance(value, list) else []


def _metadata_json_dict_list(metadata: dict[str, str], key: str) -> list[dict[str, object]]:
    result: list[dict[str, object]] = []
    for item in _metadata_json_list(metadata, key):
        if isinstance(item, dict):
            result.append(dict(item))
    return result


def _graphic_layout_metadata_updates(
    *,
    text_wrap: str | None = None,
    text_flow: str | None = None,
    treat_as_char: bool | None = None,
    affect_line_spacing: bool | None = None,
    flow_with_text: bool | None = None,
    allow_overlap: bool | None = None,
    hold_anchor_and_so: bool | None = None,
    vert_rel_to: str | None = None,
    horz_rel_to: str | None = None,
    vert_align: str | None = None,
    horz_align: str | None = None,
    vert_offset: int | str | None = None,
    horz_offset: int | str | None = None,
) -> dict[str, object | None]:
    return {
        "layout.textWrap": text_wrap,
        "layout.textFlow": text_flow,
        "layout.treatAsChar": treat_as_char,
        "layout.affectLSpacing": affect_line_spacing,
        "layout.flowWithText": flow_with_text,
        "layout.allowOverlap": allow_overlap,
        "layout.holdAnchorAndSO": hold_anchor_and_so,
        "layout.vertRelTo": vert_rel_to,
        "layout.horzRelTo": horz_rel_to,
        "layout.vertAlign": vert_align,
        "layout.horzAlign": horz_align,
        "layout.vertOffset": vert_offset,
        "layout.horzOffset": horz_offset,
    }


def _graphic_rotation_metadata_updates(
    *,
    angle: int | str | None = None,
    center_x: int | str | None = None,
    center_y: int | str | None = None,
    rotate_image: bool | None = None,
) -> dict[str, object | None]:
    return {
        "rotation.angle": angle,
        "rotation.centerX": center_x,
        "rotation.centerY": center_y,
        "rotation.rotateimage": rotate_image,
    }


def _image_adjustment_metadata_updates(
    *,
    bright: int | str | None = None,
    contrast: int | str | None = None,
    effect: str | None = None,
    alpha: int | str | None = None,
) -> dict[str, object | None]:
    return {
        "image.bright": bright,
        "image.contrast": contrast,
        "image.effect": effect,
        "image.alpha": alpha,
    }


def _parse_connectline_payload_fields(payload: bytes) -> dict[str, object]:
    result = _parse_raw_payload_fields(payload)
    result["variant"] = "hancom_connectline"
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


def _encode_length_prefixed_utf16(value: str) -> bytes:
    encoded = str(value).encode("utf-16-le")
    return len(str(value)).to_bytes(2, "little", signed=False) + encoded


def _parse_textart_payload_fields(payload: bytes) -> dict[str, object]:
    prefix = payload[:32] if len(payload) >= 32 else _HWP_TEXTART_PREFIX_PAYLOAD
    cursor = 32
    text, cursor = _parse_length_prefixed_utf16(payload, cursor)
    font_name, cursor = _parse_length_prefixed_utf16(payload, cursor)
    font_style, cursor = _parse_length_prefixed_utf16(payload, cursor)
    trailing_payload = payload[cursor:] if cursor < len(payload) else b""
    return {
        "text": text.replace("\r", ""),
        "font_name": font_name,
        "font_style": font_style,
        "prefix_payload": prefix,
        "trailing_payload": trailing_payload,
    }


def _coerce_shape_point(
    value: dict[str, object] | Sequence[int] | None,
    *,
    default: tuple[int, int] = (0, 0),
) -> tuple[int, int]:
    if value is None:
        return default
    if isinstance(value, dict):
        return (
            int(cast(Any, value.get("x", default[0]))),
            int(cast(Any, value.get("y", default[1]))),
        )
    values = list(value)
    if len(values) < 2:
        return default
    return (int(values[0]), int(values[1]))


def _build_line_specific_payload(
    *,
    start: dict[str, object] | Sequence[int],
    end: dict[str, object] | Sequence[int],
    attributes: int = 0,
) -> bytes:
    start_x, start_y = _coerce_shape_point(start)
    end_x, end_y = _coerce_shape_point(end)
    payload = bytearray()
    payload.extend(int(start_x).to_bytes(4, "little", signed=True))
    payload.extend(int(start_y).to_bytes(4, "little", signed=True))
    payload.extend(int(end_x).to_bytes(4, "little", signed=True))
    payload.extend(int(end_y).to_bytes(4, "little", signed=True))
    payload.extend(int(attributes).to_bytes(4, "little", signed=False))
    return bytes(payload)


def _build_rectangle_specific_payload(*, corner_radius: int, trailing_payload: bytes = b"") -> bytes:
    return int(corner_radius).to_bytes(4, "little", signed=False) + bytes(trailing_payload)


def _build_ellipse_geometry_payload(
    *,
    arc_flags: int,
    center: dict[str, object] | Sequence[int],
    axis1: dict[str, object] | Sequence[int],
    axis2: dict[str, object] | Sequence[int],
    start: dict[str, object] | Sequence[int],
    end: dict[str, object] | Sequence[int],
    start_2: dict[str, object] | Sequence[int],
    end_2: dict[str, object] | Sequence[int],
) -> bytes:
    values = [int(arc_flags)]
    for point in (center, axis1, axis2, start, end, start_2, end_2):
        point_x, point_y = _coerce_shape_point(point)
        values.extend((point_x, point_y))
    return struct.pack("<15i", *values)


def _build_arc_geometry_payload(
    *,
    arc_type: int,
    center: dict[str, object] | Sequence[int],
    axis1: dict[str, object] | Sequence[int],
    axis2: dict[str, object] | Sequence[int],
) -> bytes:
    values = [int(arc_type)]
    for point in (center, axis1, axis2):
        point_x, point_y = _coerce_shape_point(point)
        values.extend((point_x, point_y))
    return struct.pack("<7i", *values)


def _build_polygon_points_payload(
    points: Sequence[dict[str, object] | Sequence[int]],
    *,
    trailing_payload: bytes = b"",
) -> bytes:
    payload = bytearray()
    payload.extend(len(points).to_bytes(2, "little", signed=False))
    for point in points:
        point_x, _point_y = _coerce_shape_point(point)
        payload.extend(int(point_x).to_bytes(4, "little", signed=True))
    for point in points:
        _point_x, point_y = _coerce_shape_point(point)
        payload.extend(int(point_y).to_bytes(4, "little", signed=True))
    payload.extend(trailing_payload)
    return bytes(payload)


def _build_textart_payload_fields(
    *,
    text: str,
    font_name: str,
    font_style: str,
    prefix_payload: bytes | None = None,
    trailing_payload: bytes | None = None,
) -> bytes:
    payload = bytearray(prefix_payload or _HWP_TEXTART_PREFIX_PAYLOAD)
    payload.extend(_encode_length_prefixed_utf16(text))
    payload.extend(_encode_length_prefixed_utf16(font_name))
    payload.extend(_encode_length_prefixed_utf16(font_style))
    payload.extend(trailing_payload if trailing_payload is not None else _HWP_TEXTART_TRAILING_PAYLOAD)
    return bytes(payload)


def _shape_specific_payload_size(tag_id: int) -> int:
    if tag_id == TAG_SHAPE_COMPONENT_LINE:
        return _HWP_SHAPE_LINE_SPECIFIC_SIZE
    if tag_id == TAG_SHAPE_COMPONENT_RECTANGLE:
        return _HWP_SHAPE_RECTANGLE_SPECIFIC_SIZE
    if tag_id == TAG_SHAPE_COMPONENT_ELLIPSE:
        return _HWP_SHAPE_ELLIPSE_SPECIFIC_SIZE
    if tag_id == TAG_SHAPE_COMPONENT_ARC:
        return _HWP_SHAPE_ARC_SPECIFIC_SIZE
    return 0


def _has_hancom_connectline_signature(control_node: RecordNode) -> bool:
    component = next((node for node in control_node.iter_descendants() if node.tag_id == TAG_SHAPE_COMPONENT), None)
    specific = next((node for node in control_node.iter_descendants() if node.tag_id == TAG_SHAPE_COMPONENT_LINE), None)
    if component is None or specific is None:
        return False
    return component.payload.startswith(_HANCOM_CONNECTLINE_COMPONENT_SIGNATURE)
