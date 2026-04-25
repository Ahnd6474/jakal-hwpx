from __future__ import annotations

import json
import struct
import tempfile
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from .document import HwpxDocument
from .hwp_binary import (
    _AUTO_NUMBER_TYPE_VALUES,
    _build_auto_number_payload,
    _build_colorref,
    _build_control_id_payload,
    _build_control_native_metadata_payload,
    _build_header_footer_control_payload,
    _build_ole_attribute_flags,
    _build_table_cell_list_header_payload,
    _build_table_record_payload,
    _build_page_num_payload,
    _build_rgb_color,
    _build_shape_native_metadata_payload,
    _parse_colorref,
    _parse_arc_shape_payload,
    _parse_chart_data_payload,
    _parse_container_shape_payload,
    _parse_curve_shape_payload,
    _parse_ellipse_shape_payload,
    _parse_form_object_payload,
    _parse_graphic_control_layout,
    _parse_graphic_control_out_margins,
    _parse_rgb_color,
    _build_hyperlink_ctrl_payload,
    _build_hyperlink_para_text_payload,
    _parse_line_shape_payload,
    _parse_memo_list_payload,
    _parse_picture_crop_payload,
    _parse_picture_image_adjustment_payload,
    _picture_effect_code,
    _parse_polygon_shape_payload,
    _parse_rectangle_shape_payload,
    _parse_shape_component_rotation_payload,
    _normalize_field_control_id,
    _payload_u16_values,
    _payload_u32_values,
    _set_graphic_control_out_margins_payload,
    _set_graphic_control_layout_payload,
    _set_picture_crop_payload,
    _set_picture_image_adjustment_payload,
    _set_picture_line_payload,
    _set_shape_component_rotation_payload,
    _shape_common_payload_size,
    _TableCellSpec,
    _TABLE_CELL_CHAR_SHAPE,
    _TABLE_CELL_LINE_SEG,
    _TABLE_CELL_PARA_HEADER_TAIL,
    DEFAULT_HWP_EQUATION_FONT,
    DEFAULT_HWP_TABLE_CELL_HEIGHT,
    DEFAULT_HWP_TABLE_CELL_WIDTH,
    DocInfoModel,
    HwpBinaryDocument,
    HwpBinaryFileHeader,
    HwpDocumentProperties,
    HwpParagraph,
    HwpStreamCapacity,
    ParagraphHeaderRecord,
    ParagraphTextRecord,
    RecordNode,
    SectionModel,
    SectionParagraphModel,
    TAG_CHART_DATA,
    TAG_CTRL_DATA,
    TAG_CTRL_HEADER,
    TAG_EQEDIT,
    TAG_FORM_OBJECT,
    TAG_LIST_HEADER,
    TAG_MEMO_LIST,
    TAG_PARA_CHAR_SHAPE,
    TAG_PARA_HEADER,
    TAG_PARA_LINE_SEG,
    TAG_SHAPE_COMPONENT,
    TAG_SHAPE_COMPONENT_ARC,
    TAG_SHAPE_COMPONENT_CONTAINER,
    TAG_SHAPE_COMPONENT_CURVE,
    TAG_SHAPE_COMPONENT_ELLIPSE,
    TAG_SHAPE_COMPONENT_LINE,
    TAG_SHAPE_COMPONENT_PICTURE,
    TAG_SHAPE_COMPONENT_OLE,
    TAG_SHAPE_COMPONENT_POLYGON,
    TAG_SHAPE_COMPONENT_RECTANGLE,
    TAG_SHAPE_COMPONENT_TEXTART,
    TAG_TABLE,
)
from .exceptions import HwpxValidationError, ValidationIssue
from .hwp_pure_profile import HwpPureProfile, append_feature_from_profile

HancomConverter = Callable[[str | Path, str | Path, str], Path]

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
_HWP_STRICT_LINT_HINTS = {
    "binary_closure": "Keep DocInfo BinData records and BinData/BINxxxx streams in sync. Re-add the picture/OLE data or remove the broken control.",
    "binary_media_type": "Use replace_binary(..., extension=...) or rebuild the BinData entry so the DocInfo extension matches the stream name.",
    "binary_reference": "Re-add the missing embedded binary or update the picture/OLE control to point at an existing BinData id.",
    "control_subtree": "Recreate the object with the public append/setter API, or restore the missing child record before saving.",
    "docinfo_mapping": "Repair DocInfo id mappings before saving. This usually means recreating the referenced memo/BinData entry through the HWP API.",
    "metadata_schema": "Use the semantic setter API instead of editing JAKAL metadata text directly; values must keep their documented type.",
    "payload_sanity": "The native HWP payload is truncated or malformed. Recreate that control or replace it from a known-good source document.",
}


def _issue(
    kind: str,
    message: str,
    *,
    section_index: int | None = None,
    paragraph_index: int | None = None,
    context: str | None = None,
    hint: str | None = None,
) -> ValidationIssue:
    return ValidationIssue(
        kind=kind,
        message=message,
        section_index=section_index,
        paragraph_index=paragraph_index,
        context=context,
        hint=hint or _HWP_STRICT_LINT_HINTS.get(kind),
    )


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
        return (int(value.get("x", default[0])), int(value.get("y", default[1])))
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


def _has_hancom_connectline_signature(control_node: RecordNode) -> bool:
    component = next((node for node in control_node.iter_descendants() if node.tag_id == TAG_SHAPE_COMPONENT), None)
    specific = next((node for node in control_node.iter_descendants() if node.tag_id == TAG_SHAPE_COMPONENT_LINE), None)
    if component is None or specific is None:
        return False
    return component.payload.startswith(_HANCOM_CONNECTLINE_COMPONENT_SIGNATURE)


class HwpDocument:
    def __init__(
        self,
        binary_document: HwpBinaryDocument,
        *,
        converter: HancomConverter | None = None,
        default_append_section_index: int = 0,
    ):
        self._binary_document = binary_document
        self._converter = converter
        self._default_append_section_index = default_append_section_index
        self._bridge_document: HwpxDocument | None = None
        self._hancom_document = None
        self._bridge_temp_dir: Path | None = None
        self._pure_profile: HwpPureProfile | None = None

    @classmethod
    def blank(cls, *, converter: HancomConverter | None = None) -> "HwpDocument":
        instance = cls.blank_from_profile(HwpPureProfile.load_bundled().root, converter=converter)
        return instance

    @classmethod
    def open(cls, path: str | Path, *, converter: HancomConverter | None = None) -> "HwpDocument":
        return cls(
            HwpBinaryDocument.open(path),
            converter=converter,
            default_append_section_index=HwpPureProfile.load_bundled().target_section_index,
        )

    @classmethod
    def blank_from_profile(
        cls,
        profile_root: str | Path,
        *,
        converter: HancomConverter | None = None,
    ) -> "HwpDocument":
        profile = HwpPureProfile.load(profile_root)
        instance = cls(
            HwpBinaryDocument.open(profile.base_path),
            converter=converter,
            default_append_section_index=profile.target_section_index,
        )
        instance._pure_profile = profile
        instance._binary_document.reset_body_sections_to_blank()
        return instance

    @property
    def source_path(self) -> Path:
        return self._binary_document.source_path

    def binary_document(self) -> HwpBinaryDocument:
        return self._binary_document

    def file_header(self) -> HwpBinaryFileHeader:
        return self._binary_document.file_header()

    def document_properties(self) -> HwpDocumentProperties:
        return self._binary_document.document_properties()

    def docinfo_model(self) -> DocInfoModel:
        return self._binary_document.docinfo_model()

    def list_stream_paths(self) -> list[str]:
        return self._binary_document.list_stream_paths()

    def bindata_stream_paths(self) -> list[str]:
        return self._binary_document.bindata_stream_paths()

    def stream_capacity(self, path: str) -> HwpStreamCapacity:
        data = self._binary_document.read_stream(path, decompress=self._binary_document._stream_is_compressed(path))
        return self._binary_document.stream_capacity(
            path,
            data,
            compress=self._binary_document._stream_is_compressed(path),
        )

    def preview_text(self) -> str:
        return self._binary_document.preview_text()

    def set_preview_text(self, value: str) -> None:
        self._invalidate_bridge()
        self._binary_document.set_preview_text(value)

    def sections(self) -> list["HwpSection"]:
        return [HwpSection(self, index) for index, _path in enumerate(self._binary_document.section_stream_paths())]

    def section(self, index: int) -> "HwpSection":
        return HwpSection(self, index)

    def section_model(self, section_index: int = 0) -> SectionModel:
        return self._binary_document.section_model(section_index)

    def ensure_section_count(self, section_count: int) -> None:
        self._invalidate_bridge()
        self._binary_document.ensure_section_count(section_count)

    def apply_section_settings(
        self,
        *,
        section_index: int = 0,
        page_width: int | None = None,
        page_height: int | None = None,
        landscape: str | None = None,
        margins: dict[str, int] | None = None,
        visibility: dict[str, str] | None = None,
        grid: dict[str, int] | None = None,
        start_numbers: dict[str, str] | None = None,
        numbering_shape_id: str | None = None,
        memo_shape_id: str | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.set_section_page_settings(
            section_index,
            page_width=page_width,
            page_height=page_height,
            landscape=landscape,
            margins=margins,
        )
        self._binary_document.set_section_definition_settings(
            section_index,
            visibility=visibility,
            grid=grid,
            start_numbers=start_numbers,
            numbering_shape_id=int(numbering_shape_id) if numbering_shape_id is not None else None,
            memo_shape_id=memo_shape_id,
        )

    def apply_section_page_border_fills(
        self,
        page_border_fills: list[dict[str, str | int]],
        *,
        section_index: int = 0,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.set_section_page_border_fills(section_index, page_border_fills)

    def apply_section_page_numbers(
        self,
        page_numbers: list[dict[str, str]],
        *,
        section_index: int = 0,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.set_section_page_numbers(section_index, page_numbers)

    def apply_section_note_settings(
        self,
        *,
        section_index: int = 0,
        footnote_pr: dict[str, object] | None = None,
        endnote_pr: dict[str, object] | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.set_section_note_settings(
            section_index,
            footnote_pr=footnote_pr,
            endnote_pr=endnote_pr,
        )

    def paragraphs(self, section_index: int = 0) -> list["HwpParagraphObject"]:
        return self.section(section_index).paragraphs()

    def controls(self, section_index: int | None = None) -> list["HwpControlObject"]:
        sections = [self.section(section_index)] if section_index is not None else self.sections()
        controls: list[HwpControlObject] = []
        for section in sections:
            controls.extend(section.controls())
        return controls

    def tables(self, section_index: int | None = None) -> list["HwpTableObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpTableObject)]

    def pictures(self, section_index: int | None = None) -> list["HwpPictureObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpPictureObject)]

    def hyperlinks(self, section_index: int | None = None) -> list["HwpHyperlinkObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpHyperlinkObject)]

    def bookmarks(self, section_index: int | None = None) -> list["HwpBookmarkObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpBookmarkObject)]

    def notes(self, section_index: int | None = None) -> list["HwpNoteObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpNoteObject)]

    def page_numbers(self, section_index: int | None = None) -> list["HwpPageNumObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpPageNumObject)]

    def fields(self, section_index: int | None = None) -> list["HwpFieldObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpFieldObject)]

    def mail_merge_fields(self, section_index: int | None = None) -> list["HwpFieldObject"]:
        return [field for field in self.fields(section_index) if field.is_mail_merge]

    def calculation_fields(self, section_index: int | None = None) -> list["HwpFieldObject"]:
        return [field for field in self.fields(section_index) if field.is_calculation]

    def cross_references(self, section_index: int | None = None) -> list["HwpFieldObject"]:
        return [field for field in self.fields(section_index) if field.is_cross_reference]

    def doc_property_fields(self, section_index: int | None = None) -> list["HwpFieldObject"]:
        return [field for field in self.fields(section_index) if field.is_doc_property]

    def date_fields(self, section_index: int | None = None) -> list["HwpFieldObject"]:
        return [field for field in self.fields(section_index) if field.is_date]

    def shapes(self, section_index: int | None = None) -> list["HwpShapeObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpShapeObject)]

    def lines(self, section_index: int | None = None) -> list["HwpLineShapeObject"]:
        return [
            shape
            for shape in self.shapes(section_index)
            if isinstance(shape, HwpLineShapeObject) and not isinstance(shape, HwpConnectLineShapeObject)
        ]

    def connect_lines(self, section_index: int | None = None) -> list["HwpConnectLineShapeObject"]:
        return [shape for shape in self.shapes(section_index) if isinstance(shape, HwpConnectLineShapeObject)]

    def rectangles(self, section_index: int | None = None) -> list["HwpRectangleShapeObject"]:
        return [shape for shape in self.shapes(section_index) if isinstance(shape, HwpRectangleShapeObject)]

    def ellipses(self, section_index: int | None = None) -> list["HwpEllipseShapeObject"]:
        return [shape for shape in self.shapes(section_index) if isinstance(shape, HwpEllipseShapeObject)]

    def arcs(self, section_index: int | None = None) -> list["HwpArcShapeObject"]:
        return [shape for shape in self.shapes(section_index) if isinstance(shape, HwpArcShapeObject)]

    def polygons(self, section_index: int | None = None) -> list["HwpPolygonShapeObject"]:
        return [shape for shape in self.shapes(section_index) if isinstance(shape, HwpPolygonShapeObject)]

    def curves(self, section_index: int | None = None) -> list["HwpCurveShapeObject"]:
        return [shape for shape in self.shapes(section_index) if isinstance(shape, HwpCurveShapeObject)]

    def containers(self, section_index: int | None = None) -> list["HwpContainerShapeObject"]:
        return [shape for shape in self.shapes(section_index) if isinstance(shape, HwpContainerShapeObject)]

    def textarts(self, section_index: int | None = None) -> list["HwpTextArtShapeObject"]:
        return [shape for shape in self.shapes(section_index) if isinstance(shape, HwpTextArtShapeObject)]

    def charts(self, section_index: int | None = None) -> list["HwpChartObject"]:
        return [shape for shape in self.shapes(section_index) if isinstance(shape, HwpChartObject)]

    def oles(self, section_index: int | None = None) -> list["HwpOleObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpOleObject)]

    def forms(self, section_index: int | None = None) -> list["HwpFormObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpFormObject)]

    def memos(self, section_index: int | None = None) -> list["HwpMemoObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpMemoObject)]

    def equations(self, section_index: int | None = None) -> list["HwpEquationObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpEquationObject)]

    def get_document_text(self) -> str:
        return self._binary_document.get_document_text()

    def replace_text_same_length(
        self,
        old: str,
        new: str,
        *,
        section_index: int | None = None,
        count: int = -1,
    ) -> int:
        return self._binary_document.replace_text_same_length(old, new, section_index=section_index, count=count)

    def append_paragraph(
        self,
        text: str,
        *,
        section_index: int = 0,
        para_shape_id: int = 0,
        style_id: int = 0,
        split_flags: int = 0,
        control_mask: int = 0,
    ) -> HwpParagraphObject:
        self._invalidate_bridge()
        paragraph = self._binary_document.append_paragraph(
            text,
            section_index=section_index,
            para_shape_id=para_shape_id,
            style_id=style_id,
            split_flags=split_flags,
            control_mask=control_mask,
        )
        return HwpParagraphObject(
            self,
            HwpParagraph(
                index=paragraph.index,
                section_index=paragraph.section_index,
                char_count=paragraph.header.char_count,
                control_mask=paragraph.header.control_mask,
                para_shape_id=paragraph.header.para_shape_id,
                style_id=paragraph.header.style_id,
                split_flags=paragraph.header.split_flags,
                raw_text=paragraph.raw_text,
                text=paragraph.text,
                text_record_offset=-1,
                text_record_size=len(paragraph.raw_text.encode("utf-16-le")),
            ),
        )

    def to_hwpx_document(self, *, force_refresh: bool = False) -> HwpxDocument:
        if self._bridge_document is not None and not force_refresh:
            return self._bridge_document

        self._bridge_document = self.to_hancom_document(force_refresh=force_refresh).to_hwpx_document()
        return self._bridge_document

    def to_hancom_document(self, *, force_refresh: bool = False):
        if self._hancom_document is not None and not force_refresh:
            return self._hancom_document

        from .hancom_document import HancomDocument

        self._hancom_document = HancomDocument.from_hwp_document(self, converter=self._converter)
        return self._hancom_document

    def save_as_hwpx(self, path: str | Path) -> Path:
        target_path = Path(path).expanduser().resolve()
        target_path.parent.mkdir(parents=True, exist_ok=True)
        self.to_hwpx_document(force_refresh=False).save(target_path)
        return target_path

    def save(self, path: str | Path | None = None) -> Path:
        if path is not None:
            target_path = Path(path).expanduser().resolve()
            if target_path.suffix.lower() == ".hwpx":
                return self.save_as_hwpx(target_path)
        else:
            target_path = self.source_path

        if self._bridge_document is None:
            return self._binary_document.save(path)

        self._sync_from_bridge_document()
        self._binary_document.save(target_path)
        self._binary_document.source_path = target_path
        return target_path

    def strict_lint_errors(self) -> list[ValidationIssue]:
        errors: list[ValidationIssue] = []
        docinfo = self.docinfo_model()
        docinfo_bindata: dict[int, object] = {}
        for record in docinfo.bin_data_records():
            storage_id = int(record.storage_id or 0)
            if storage_id <= 0:
                continue
            if storage_id in docinfo_bindata:
                errors.append(_issue("docinfo_mapping", f"Duplicate BinData docinfo storage id: {storage_id}", context="DocInfo"))
                continue
            docinfo_bindata[storage_id] = record

        stream_paths = self.bindata_stream_paths()
        bindata_streams: dict[int, str] = {}
        for stream_path in stream_paths:
            name = Path(stream_path).name
            if not name.startswith("BIN") or "." not in name:
                continue
            storage_token = name[3:].split(".", 1)[0]
            if not storage_token.isdigit():
                continue
            storage_id = int(storage_token)
            if storage_id in bindata_streams:
                errors.append(_issue("binary_closure", f"Duplicate BinData stream for storage id {storage_id}: {stream_path}", context="BinData"))
                continue
            bindata_streams[storage_id] = stream_path

        for storage_id, record in docinfo_bindata.items():
            stream_path = bindata_streams.get(storage_id)
            if stream_path is None:
                errors.append(_issue("binary_closure", f"DocInfo BinData id {storage_id} is missing a BinData stream.", context="DocInfo"))
                continue
            expected_extension = str(getattr(record, "extension", "") or "").lstrip(".").lower()
            actual_extension = Path(stream_path).suffix.lstrip(".").lower()
            if expected_extension and actual_extension and expected_extension != actual_extension:
                errors.append(
                    _issue(
                        "binary_media_type",
                        f"DocInfo BinData id {storage_id} expects .{expected_extension} but stream path is {stream_path}.",
                        context="BinData",
                    )
                )

        for storage_id, stream_path in bindata_streams.items():
            if storage_id not in docinfo_bindata:
                errors.append(
                    _issue(
                        "binary_closure",
                        f"BinData stream is not declared in DocInfo: {stream_path}",
                        context="BinData",
                    )
                )

        id_mappings = docinfo.id_mappings_record()
        max_bindata_id = max([0, *docinfo_bindata.keys(), *bindata_streams.keys()])
        if id_mappings.bin_data_count < max_bindata_id:
            errors.append(
                _issue(
                    "docinfo_mapping",
                    f"DocInfo id_mappings.bin_data_count={id_mappings.bin_data_count} is smaller than max BinData id {max_bindata_id}.",
                    context="DocInfo",
                )
            )

        memo_shape_records = docinfo.memo_shape_records()
        memo_shape_count = id_mappings.get_count(15)
        if memo_shape_count < len(memo_shape_records):
            errors.append(
                _issue(
                    "docinfo_mapping",
                    f"DocInfo id_mappings.memo_shapes={memo_shape_count} is smaller than actual memo shape record count {len(memo_shape_records)}.",
                    context="DocInfo",
                )
            )

        for section_index, section in enumerate(self.sections()):
            section_model = self.section_model(section_index)
            memo_shape_id = self.binary_document().section_definition_settings(section_index).get("memo_shape_id")
            if memo_shape_id not in (None, "", "0"):
                if not str(memo_shape_id).isdigit():
                    errors.append(
                        _issue(
                            "docinfo_mapping",
                            f"Section memo_shape_id must be numeric, got {memo_shape_id!r}.",
                            section_index=section_index,
                            context="SectionSettings",
                        )
                    )
                elif int(str(memo_shape_id)) >= len(memo_shape_records):
                    errors.append(
                        _issue(
                            "docinfo_mapping",
                            f"Section memo_shape_id {memo_shape_id} points to missing TAG_MEMO_SHAPE record.",
                            section_index=section_index,
                            context="SectionSettings",
                        )
                    )
            for paragraph in section_model.paragraphs():
                for control_node in paragraph.control_nodes():
                    control_id = _control_id(control_node)
                    context = f"control[{control_id or 'UNKNOWN'}]"
                    if control_id == "gso ":
                        has_shape_component = _control_has_descendant(control_node, TAG_SHAPE_COMPONENT)
                        has_picture = _control_has_descendant(control_node, TAG_SHAPE_COMPONENT_PICTURE)
                        has_ole = _control_has_descendant(control_node, TAG_SHAPE_COMPONENT_OLE)
                        has_chart = _control_has_descendant(control_node, TAG_CHART_DATA)
                        shape_metadata = _collect_shape_native_metadata(control_node)
                        expects_chart_data = any(key.startswith("chart.") for key in shape_metadata)
                        if not (has_shape_component or has_chart):
                            errors.append(
                                _issue(
                                    "control_subtree",
                                    "Graphic control does not contain a shape component or chart data record.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        if expects_chart_data and not has_chart:
                            errors.append(
                                _issue(
                                    "control_subtree",
                                    "Chart control metadata is present but TAG_CHART_DATA is missing.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                    )
                                )
                        if not _metadata_bool_literal_is_valid(shape_metadata.get("chart.legendVisible")):
                            errors.append(
                                _issue(
                                    "metadata_schema",
                                    "chart.legendVisible metadata must be a bool-like literal.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        chart_categories, chart_categories_error = _metadata_json_value(shape_metadata, "chart.categories")
                        if chart_categories_error is not None:
                            errors.append(
                                _issue(
                                    "metadata_schema",
                                    f"chart.categories metadata must be valid JSON: {chart_categories_error}",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        elif chart_categories is not None and not isinstance(chart_categories, list):
                            errors.append(
                                _issue(
                                    "metadata_schema",
                                    "chart.categories metadata must decode to a JSON list.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        chart_series, chart_series_error = _metadata_json_value(shape_metadata, "chart.series")
                        if chart_series_error is not None:
                            errors.append(
                                _issue(
                                    "metadata_schema",
                                    f"chart.series metadata must be valid JSON: {chart_series_error}",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        elif chart_series is not None:
                            if not isinstance(chart_series, list):
                                errors.append(
                                    _issue(
                                        "metadata_schema",
                                        "chart.series metadata must decode to a JSON list.",
                                        section_index=section_index,
                                        paragraph_index=paragraph.index,
                                        context=context,
                                    )
                                )
                            elif not all(isinstance(item, dict) for item in chart_series):
                                errors.append(
                                    _issue(
                                        "metadata_schema",
                                        "chart.series metadata must decode to a JSON list of objects.",
                                        section_index=section_index,
                                        paragraph_index=paragraph.index,
                                        context=context,
                                    )
                                )
                        if has_chart:
                            chart_nodes = [node for node in control_node.iter_descendants() if node.tag_id == TAG_CHART_DATA]
                            for chart_node in chart_nodes:
                                if len(chart_node.payload) % 2:
                                    errors.append(
                                        _issue(
                                            "payload_sanity",
                                            "Chart payload is not UTF-16 aligned.",
                                            section_index=section_index,
                                            paragraph_index=paragraph.index,
                                            context=context,
                                        )
                                    )
                        if has_picture:
                            picture_nodes = [node for node in control_node.iter_descendants() if node.tag_id == TAG_SHAPE_COMPONENT_PICTURE]
                            for picture_node in picture_nodes:
                                if len(picture_node.payload) < 73:
                                    errors.append(
                                        _issue(
                                            "payload_sanity",
                                            "Picture payload is shorter than the expected 73-byte minimum.",
                                            section_index=section_index,
                                            paragraph_index=paragraph.index,
                                            context=context,
                                        )
                                    )
                                else:
                                    storage_id = int.from_bytes(picture_node.payload[71:73], "little", signed=False)
                                    if storage_id and storage_id not in docinfo_bindata:
                                        errors.append(
                                            _issue(
                                                "binary_reference",
                                                f"Picture control references missing BinData id {storage_id}.",
                                                section_index=section_index,
                                                paragraph_index=paragraph.index,
                                                context=context,
                                            )
                                        )
                        if has_ole:
                            ole_nodes = [node for node in control_node.iter_descendants() if node.tag_id == TAG_SHAPE_COMPONENT_OLE]
                            for ole_node in ole_nodes:
                                if len(ole_node.payload) < 20:
                                    errors.append(
                                        _issue(
                                            "payload_sanity",
                                            "OLE payload is shorter than the expected 20-byte minimum.",
                                            section_index=section_index,
                                            paragraph_index=paragraph.index,
                                            context=context,
                                        )
                                    )
                                storage_id = int.from_bytes(ole_node.payload[12:14].ljust(2, b"\x00"), "little", signed=False)
                                if storage_id and storage_id not in docinfo_bindata:
                                    errors.append(
                                        _issue(
                                            "binary_reference",
                                            f"OLE control references missing BinData id {storage_id}.",
                                            section_index=section_index,
                                            paragraph_index=paragraph.index,
                                            context=context,
                                            )
                                        )
                    elif control_id == "mrof":
                        metadata = _collect_control_native_metadata(control_node)
                        form_nodes = [node for node in control_node.iter_descendants() if node.tag_id == TAG_FORM_OBJECT]
                        if not form_nodes:
                            errors.append(
                                _issue(
                                    "control_subtree",
                                    "Form control does not contain a TAG_FORM_OBJECT record.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        for form_node in form_nodes:
                            if len(form_node.payload) % 2:
                                errors.append(
                                    _issue(
                                        "payload_sanity",
                                        "Form payload is not UTF-16 aligned.",
                                        section_index=section_index,
                                        paragraph_index=paragraph.index,
                                        context=context,
                                    )
                                )
                        if not _metadata_bool_literal_is_valid(metadata.get("form.checked")):
                            errors.append(
                                _issue(
                                    "metadata_schema",
                                    "form.checked metadata must be a bool-like literal.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        if not _metadata_bool_literal_is_valid(metadata.get("form.editable")):
                            errors.append(
                                _issue(
                                    "metadata_schema",
                                    "form.editable metadata must be a bool-like literal.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        if not _metadata_bool_literal_is_valid(metadata.get("form.locked")):
                            errors.append(
                                _issue(
                                    "metadata_schema",
                                    "form.locked metadata must be a bool-like literal.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        form_items, form_items_error = _metadata_json_value(metadata, "form.items")
                        if form_items_error is not None:
                            errors.append(
                                _issue(
                                    "metadata_schema",
                                    f"form.items metadata must be valid JSON: {form_items_error}",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        elif form_items is not None and not isinstance(form_items, list):
                            errors.append(
                                _issue(
                                    "metadata_schema",
                                    "form.items metadata must decode to a JSON list.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                    elif control_id == "omem":
                        metadata = _collect_control_native_metadata(control_node)
                        memo_nodes = [node for node in control_node.iter_descendants() if node.tag_id == TAG_MEMO_LIST]
                        if not memo_nodes:
                            errors.append(
                                _issue(
                                    "control_subtree",
                                    "Memo control does not contain a TAG_MEMO_LIST record.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        for memo_node in memo_nodes:
                            if len(memo_node.payload) % 2:
                                errors.append(
                                    _issue(
                                        "payload_sanity",
                                        "Memo payload is not UTF-16 aligned.",
                                        section_index=section_index,
                                        paragraph_index=paragraph.index,
                                        context=context,
                                    )
                                )
                        memo_order = metadata.get("memo.order")
                        if memo_order not in (None, ""):
                            try:
                                int(memo_order)
                            except ValueError:
                                errors.append(
                                    _issue(
                                        "metadata_schema",
                                        "memo.order metadata must be an integer.",
                                        section_index=section_index,
                                        paragraph_index=paragraph.index,
                                        context=context,
                                    )
                                )
                        if not _metadata_bool_literal_is_valid(metadata.get("memo.visible")):
                            errors.append(
                                _issue(
                                    "metadata_schema",
                                    "memo.visible metadata must be a bool-like literal.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                    elif control_id == "tbl ":
                        table_nodes = [node for node in control_node.iter_descendants() if node.tag_id == TAG_TABLE]
                        if not table_nodes:
                            errors.append(
                                _issue(
                                    "control_subtree",
                                    "Table control does not contain a TAG_TABLE record.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        for table_node in table_nodes:
                            if len(table_node.payload) < 20:
                                errors.append(
                                    _issue(
                                        "payload_sanity",
                                        "Table payload is shorter than the expected minimum.",
                                        section_index=section_index,
                                        paragraph_index=paragraph.index,
                                        context=context,
                                    )
                                )
                    elif control_id in {"head", "foot", "fn  ", "en  "}:
                        if not any(isinstance(node, ParagraphHeaderRecord) for node in control_node.children):
                            errors.append(
                                _issue(
                                    "control_subtree",
                                    "Nested control is missing a writable paragraph header.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        if control_id in {"head", "foot"} and len(control_node.payload) < 8:
                            errors.append(
                                _issue(
                                    "payload_sanity",
                                    "Header/footer payload is shorter than the expected 8-byte minimum.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                    elif control_id == "pgnp" and len(control_node.payload) < 8:
                        errors.append(
                            _issue(
                                "payload_sanity",
                                "Page number payload is shorter than the expected 8-byte minimum.",
                                section_index=section_index,
                                paragraph_index=paragraph.index,
                                context=context,
                            )
                        )
                    elif control_id.startswith("%") and control_id != "%hlk":
                        if not _control_has_descendant(control_node, TAG_CTRL_DATA):
                            errors.append(
                                _issue(
                                    "control_subtree",
                                    "Field control is missing TAG_CTRL_DATA parameters.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )

        return errors

    def strict_validate(self) -> None:
        errors = self.strict_lint_errors()
        if errors:
            raise HwpxValidationError(errors)

    def strict_lint_report(self, *, include_none: bool = False) -> list[dict[str, object]]:
        return [issue.to_dict(include_none=include_none) for issue in self.strict_lint_errors()]

    def format_strict_lint_errors(self) -> str:
        issues = self.strict_lint_errors()
        if not issues:
            return ""
        lines: list[str] = []
        for index, issue in enumerate(issues, start=1):
            location = issue.location() or "document"
            lines.append(f"{index}. [{issue.code or issue.kind}] {location}: {issue.message}")
            if issue.hint:
                lines.append(f"   Hint: {issue.hint}")
        return "\n".join(lines)

    def load_pure_profile(self, profile_root: str | Path) -> HwpPureProfile:
        self._pure_profile = HwpPureProfile.load(profile_root)
        self._default_append_section_index = self._pure_profile.target_section_index
        return self._pure_profile

    def append_table(
        self,
        cell_text: str | None = None,
        *,
        rows: int = 1,
        cols: int = 1,
        cell_texts: Sequence[str] | Sequence[Sequence[str]] | None = None,
        row_heights: Sequence[int] | None = None,
        col_widths: Sequence[int] | None = None,
        cell_spans: dict[tuple[int, int], tuple[int, int]] | None = None,
        cell_border_fill_ids: dict[tuple[int, int], int] | None = None,
        table_border_fill_id: int = 1,
        section_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        target_section_index = self._default_append_section_index if section_index is None else section_index
        self._binary_document.append_table(
            cell_text,
            rows=rows,
            cols=cols,
            cell_texts=cell_texts,
            row_heights=row_heights,
            col_widths=col_widths,
            cell_spans=cell_spans,
            cell_border_fill_ids=cell_border_fill_ids,
            table_border_fill_id=table_border_fill_id,
            profile_root=None,
            section_index=target_section_index,
        )

    def append_picture(
        self,
        image_bytes: bytes | None = None,
        *,
        extension: str | None = None,
        width: int = 12000,
        height: int = 3200,
        section_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        target_section_index = self._default_append_section_index if section_index is None else section_index
        self._binary_document.append_picture(
            image_bytes,
            extension=extension,
            width=width,
            height=height,
            profile_root=None,
            section_index=target_section_index,
        )

    def append_hyperlink(
        self,
        url: str | None = None,
        *,
        text: str | None = None,
        metadata_fields: Sequence[str | int] | None = None,
        section_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        target_section_index = self._default_append_section_index if section_index is None else section_index
        self._binary_document.append_hyperlink(
            url,
            text=text,
            metadata_fields=metadata_fields,
            profile_root=None,
            section_index=target_section_index,
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
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_field(
            field_type=field_type,
            display_text=display_text,
            name=name,
            parameters=parameters,
            editable=editable,
            dirty=dirty,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_mail_merge_field(
        self,
        field_name: str,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self.append_field(
            field_type="MAILMERGE",
            display_text=display_text,
            name=field_name,
            parameters={"FieldName": field_name, "MergeField": field_name},
            editable=editable,
            dirty=dirty,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_calculation_field(
        self,
        expression: str,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self.append_field(
            field_type="FORMULA",
            display_text=display_text,
            parameters={"Expression": expression, "Command": expression},
            editable=editable,
            dirty=dirty,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_cross_reference(
        self,
        bookmark_name: str,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self.append_field(
            field_type="CROSSREF",
            display_text=display_text,
            name=bookmark_name,
            parameters={"BookmarkName": bookmark_name, "Path": bookmark_name},
            editable=editable,
            dirty=dirty,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_doc_property_field(
        self,
        property_name: str,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self.append_field(
            field_type="DOCPROPERTY",
            display_text=display_text,
            name=property_name,
            parameters={"FieldName": property_name},
            editable=editable,
            dirty=dirty,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_date_field(
        self,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self.append_field(
            field_type="DATE",
            display_text=display_text,
            editable=editable,
            dirty=dirty,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_bookmark(
        self,
        name: str,
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_bookmark(name, section_index=section_index, paragraph_index=paragraph_index)

    def append_note(
        self,
        text: str,
        *,
        kind: str = "footNote",
        number: int | str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_note(
            text,
            kind=kind,
            number=number,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_footnote(
        self,
        text: str,
        *,
        number: int | str | None = None,
        section_index: int = 0,
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
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self.append_note(
            text,
            kind="endNote",
            number=number,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_form(
        self,
        label: str = "",
        *,
        form_type: str = "INPUT",
        name: str | None = None,
        value: str | None = None,
        checked: bool = False,
        items: Sequence[str] | None = None,
        editable: bool = True,
        locked: bool = False,
        placeholder: str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_form(
            label,
            form_type=form_type,
            name=name,
            value=value,
            checked=checked,
            items=items,
            editable=editable,
            locked=locked,
            placeholder=placeholder,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_memo(
        self,
        text: str,
        *,
        author: str | None = None,
        memo_id: str | None = None,
        anchor_id: str | None = None,
        order: int | None = None,
        visible: bool = True,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_memo(
            text,
            author=author,
            memo_id=memo_id,
            anchor_id=anchor_id,
            order=order,
            visible=visible,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_chart(
        self,
        title: str = "",
        *,
        chart_type: str = "BAR",
        categories: Sequence[str] | None = None,
        series: Sequence[dict[str, object]] | None = None,
        data_ref: str | None = None,
        legend_visible: bool = True,
        width: int = 12000,
        height: int = 3200,
        shape_comment: str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_chart(
            title,
            chart_type=chart_type,
            categories=categories,
            series=series,
            data_ref=data_ref,
            legend_visible=legend_visible,
            width=width,
            height=height,
            shape_comment=shape_comment,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_auto_number(
        self,
        *,
        number: int | str = 1,
        number_type: str = "PAGE",
        kind: str = "newNum",
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_auto_number(
            number=number,
            number_type=number_type,
            kind=kind,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_header(
        self,
        text: str,
        *,
        apply_page_type: str = "BOTH",
        section_index: int = 0,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_header(text, apply_page_type=apply_page_type, section_index=section_index)

    def append_footer(
        self,
        text: str,
        *,
        apply_page_type: str = "BOTH",
        section_index: int = 0,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_footer(text, apply_page_type=apply_page_type, section_index=section_index)

    def append_equation(
        self,
        script: str,
        *,
        width: int = 4800,
        height: int = 2300,
        font: str = DEFAULT_HWP_EQUATION_FONT,
        shape_comment: str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_equation(
            script,
            width=width,
            height=height,
            font=font,
            shape_comment=shape_comment,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_shape(
        self,
        *,
        kind: str = "rect",
        text: str = "",
        width: int = 12000,
        height: int = 3200,
        fill_color: str = "#FFFFFF",
        line_color: str = "#000000",
        shape_comment: str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_shape(
            kind=kind,
            text=text,
            width=width,
            height=height,
            fill_color=fill_color,
            line_color=line_color,
            shape_comment=shape_comment,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_ole(
        self,
        name: str,
        data: bytes,
        *,
        width: int = 42001,
        height: int = 13501,
        shape_comment: str | None = None,
        object_type: str = "EMBEDDED",
        draw_aspect: str = "CONTENT",
        has_moniker: bool = False,
        eq_baseline: int = 0,
        line_color: str = "#000000",
        line_width: int = 0,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_ole(
            name,
            data,
            width=width,
            height=height,
            shape_comment=shape_comment,
            object_type=object_type,
            draw_aspect=draw_aspect,
            has_moniker=has_moniker,
            eq_baseline=eq_baseline,
            line_color=line_color,
            line_width=line_width,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def add_embedded_bindata(self, data: bytes, *, extension: str, storage_id: int | None = None) -> tuple[int, str]:
        self._invalidate_bridge()
        return self._binary_document.add_embedded_bindata(data, extension=extension, storage_id=storage_id)

    def remove_embedded_bindata(self, storage_id: int) -> bool:
        self._invalidate_bridge()
        return self._binary_document.remove_embedded_bindata(storage_id)

    def append_table_pure(self) -> None:
        self.append_table()

    def append_picture_pure(self) -> None:
        self.append_picture()

    def append_hyperlink_pure(self) -> None:
        self.append_hyperlink()

    def bridge(self):
        from .bridge import HwpHwpxBridge

        return HwpHwpxBridge.from_hwp(self, converter=self._converter)

    def __getattr__(self, name: str):
        if name.startswith("_"):
            raise AttributeError(name)
        bridge_document = self.to_hwpx_document()
        return getattr(bridge_document, name)

    def _ensure_bridge_workspace(self) -> Path:
        if self._bridge_temp_dir is None:
            self._bridge_temp_dir = Path(tempfile.mkdtemp(prefix="jakal_hwp_bridge_"))
        return self._bridge_temp_dir

    def _append_feature_pure(self, feature: str) -> None:
        self._invalidate_bridge()
        if self._pure_profile is None:
            self._pure_profile = HwpPureProfile.load_bundled()
        append_feature_from_profile(self._binary_document, self._pure_profile, feature)

    def _mutate_via_hwpx(self, mutate: Callable[[HwpxDocument], object]) -> object:
        bridge_document = self.to_hwpx_document(force_refresh=False)
        result = mutate(bridge_document)
        self._bridge_document = bridge_document
        self._sync_from_bridge_document()
        return result

    def _sync_from_bridge_document(self) -> None:
        if self._bridge_document is None:
            return
        from .hancom_document import HancomDocument

        rebuilt = HancomDocument.from_hwpx_document(self._bridge_document, converter=self._converter).to_hwp_document(
            converter=self._converter
        )
        self._binary_document = rebuilt.binary_document()
        self._hancom_document = None

    def _invalidate_bridge(self) -> None:
        self._bridge_document = None
        self._hancom_document = None


@dataclass(frozen=True)
class HwpSection:
    document: HwpDocument
    index: int

    def paragraphs(self) -> list["HwpParagraphObject"]:
        return [HwpParagraphObject(self.document, paragraph) for paragraph in self.document.binary_document().paragraphs(self.index)]

    def paragraph(self, index: int) -> "HwpParagraphObject":
        return self.paragraphs()[index]

    def records(self):
        return self.document.binary_document().section_records(self.index)

    def model(self) -> SectionModel:
        return self.document.binary_document().section_model(self.index)

    def controls(self) -> list["HwpControlObject"]:
        section_model = self.model()
        controls: list[HwpControlObject] = []
        for paragraph in section_model.paragraphs():
            control_ordinal = 0
            for control_node in paragraph.header.children:
                if control_node.tag_id != TAG_CTRL_HEADER:
                    continue
                wrapper = _build_control_wrapper(self.document, section_model, paragraph, control_node, control_ordinal)
                if wrapper is not None:
                    controls.append(wrapper)
                control_ordinal += 1
        return controls

    def tables(self) -> list["HwpTableObject"]:
        return [control for control in self.controls() if isinstance(control, HwpTableObject)]

    def pictures(self) -> list["HwpPictureObject"]:
        return [control for control in self.controls() if isinstance(control, HwpPictureObject)]

    def hyperlinks(self) -> list["HwpHyperlinkObject"]:
        return [control for control in self.controls() if isinstance(control, HwpHyperlinkObject)]

    def bookmarks(self) -> list["HwpBookmarkObject"]:
        return [control for control in self.controls() if isinstance(control, HwpBookmarkObject)]

    def notes(self) -> list["HwpNoteObject"]:
        return [control for control in self.controls() if isinstance(control, HwpNoteObject)]

    def page_numbers(self) -> list["HwpPageNumObject"]:
        return [control for control in self.controls() if isinstance(control, HwpPageNumObject)]

    def fields(self) -> list["HwpFieldObject"]:
        return [control for control in self.controls() if isinstance(control, HwpFieldObject)]

    def mail_merge_fields(self) -> list["HwpFieldObject"]:
        return [field for field in self.fields() if field.is_mail_merge]

    def calculation_fields(self) -> list["HwpFieldObject"]:
        return [field for field in self.fields() if field.is_calculation]

    def cross_references(self) -> list["HwpFieldObject"]:
        return [field for field in self.fields() if field.is_cross_reference]

    def doc_property_fields(self) -> list["HwpFieldObject"]:
        return [field for field in self.fields() if field.is_doc_property]

    def date_fields(self) -> list["HwpFieldObject"]:
        return [field for field in self.fields() if field.is_date]

    def shapes(self) -> list["HwpShapeObject"]:
        return [control for control in self.controls() if isinstance(control, HwpShapeObject)]

    def lines(self) -> list["HwpLineShapeObject"]:
        return [shape for shape in self.shapes() if isinstance(shape, HwpLineShapeObject) and not isinstance(shape, HwpConnectLineShapeObject)]

    def connect_lines(self) -> list["HwpConnectLineShapeObject"]:
        return [shape for shape in self.shapes() if isinstance(shape, HwpConnectLineShapeObject)]

    def rectangles(self) -> list["HwpRectangleShapeObject"]:
        return [shape for shape in self.shapes() if isinstance(shape, HwpRectangleShapeObject)]

    def ellipses(self) -> list["HwpEllipseShapeObject"]:
        return [shape for shape in self.shapes() if isinstance(shape, HwpEllipseShapeObject)]

    def arcs(self) -> list["HwpArcShapeObject"]:
        return [shape for shape in self.shapes() if isinstance(shape, HwpArcShapeObject)]

    def polygons(self) -> list["HwpPolygonShapeObject"]:
        return [shape for shape in self.shapes() if isinstance(shape, HwpPolygonShapeObject)]

    def curves(self) -> list["HwpCurveShapeObject"]:
        return [shape for shape in self.shapes() if isinstance(shape, HwpCurveShapeObject)]

    def containers(self) -> list["HwpContainerShapeObject"]:
        return [shape for shape in self.shapes() if isinstance(shape, HwpContainerShapeObject)]

    def textarts(self) -> list["HwpTextArtShapeObject"]:
        return [shape for shape in self.shapes() if isinstance(shape, HwpTextArtShapeObject)]

    def charts(self) -> list["HwpChartObject"]:
        return [shape for shape in self.shapes() if isinstance(shape, HwpChartObject)]

    def oles(self) -> list["HwpOleObject"]:
        return [control for control in self.controls() if isinstance(control, HwpOleObject)]

    def forms(self) -> list["HwpFormObject"]:
        return [control for control in self.controls() if isinstance(control, HwpFormObject)]

    def memos(self) -> list["HwpMemoObject"]:
        return [control for control in self.controls() if isinstance(control, HwpMemoObject)]

    def equations(self) -> list["HwpEquationObject"]:
        return [control for control in self.controls() if isinstance(control, HwpEquationObject)]

    def replace_text_same_length(self, old: str, new: str, *, count: int = -1) -> int:
        return self.document.replace_text_same_length(old, new, section_index=self.index, count=count)

    def append_paragraph(
        self,
        text: str,
        *,
        para_shape_id: int = 0,
        style_id: int = 0,
        split_flags: int = 0,
        control_mask: int = 0,
    ) -> "HwpParagraphObject":
        return self.document.append_paragraph(
            text,
            section_index=self.index,
            para_shape_id=para_shape_id,
            style_id=style_id,
            split_flags=split_flags,
            control_mask=control_mask,
        )

    def append_table(
        self,
        cell_text: str | None = None,
        *,
        rows: int = 1,
        cols: int = 1,
        cell_texts: Sequence[str] | Sequence[Sequence[str]] | None = None,
        row_heights: Sequence[int] | None = None,
        col_widths: Sequence[int] | None = None,
        cell_spans: dict[tuple[int, int], tuple[int, int]] | None = None,
        cell_border_fill_ids: dict[tuple[int, int], int] | None = None,
        table_border_fill_id: int = 1,
    ) -> None:
        self.document.append_table(
            cell_text,
            rows=rows,
            cols=cols,
            cell_texts=cell_texts,
            row_heights=row_heights,
            col_widths=col_widths,
            cell_spans=cell_spans,
            cell_border_fill_ids=cell_border_fill_ids,
            table_border_fill_id=table_border_fill_id,
            section_index=self.index,
        )

    def append_picture(
        self,
        image_bytes: bytes | None = None,
        *,
        extension: str | None = None,
        width: int = 12000,
        height: int = 3200,
    ) -> None:
        self.document.append_picture(
            image_bytes,
            extension=extension,
            width=width,
            height=height,
            section_index=self.index,
        )

    def append_hyperlink(
        self,
        url: str | None = None,
        *,
        text: str | None = None,
        metadata_fields: Sequence[str | int] | None = None,
    ) -> None:
        self.document.append_hyperlink(
            url,
            text=text,
            metadata_fields=metadata_fields,
            section_index=self.index,
        )

    def apply_settings(
        self,
        *,
        page_width: int | None = None,
        page_height: int | None = None,
        landscape: str | None = None,
        margins: dict[str, int] | None = None,
        visibility: dict[str, str] | None = None,
        grid: dict[str, int] | None = None,
        start_numbers: dict[str, str] | None = None,
        numbering_shape_id: str | None = None,
        memo_shape_id: str | None = None,
    ) -> None:
        self.document.apply_section_settings(
            section_index=self.index,
            page_width=page_width,
            page_height=page_height,
            landscape=landscape,
            margins=margins,
            visibility=visibility,
            grid=grid,
            start_numbers=start_numbers,
            numbering_shape_id=numbering_shape_id,
            memo_shape_id=memo_shape_id,
        )

    def apply_page_border_fills(self, page_border_fills: list[dict[str, str | int]]) -> None:
        self.document.apply_section_page_border_fills(page_border_fills, section_index=self.index)

    def apply_page_numbers(self, page_numbers: list[dict[str, str]]) -> None:
        self.document.apply_section_page_numbers(page_numbers, section_index=self.index)

    def apply_note_settings(
        self,
        *,
        footnote_pr: dict[str, object] | None = None,
        endnote_pr: dict[str, object] | None = None,
    ) -> None:
        self.document.apply_section_note_settings(
            section_index=self.index,
            footnote_pr=footnote_pr,
            endnote_pr=endnote_pr,
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
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_field(
            field_type=field_type,
            display_text=display_text,
            name=name,
            parameters=parameters,
            editable=editable,
            dirty=dirty,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_mail_merge_field(
        self,
        field_name: str,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_mail_merge_field(
            field_name,
            display_text=display_text,
            editable=editable,
            dirty=dirty,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_calculation_field(
        self,
        expression: str,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_calculation_field(
            expression,
            display_text=display_text,
            editable=editable,
            dirty=dirty,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_cross_reference(
        self,
        bookmark_name: str,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_cross_reference(
            bookmark_name,
            display_text=display_text,
            editable=editable,
            dirty=dirty,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_doc_property_field(
        self,
        property_name: str,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_doc_property_field(
            property_name,
            display_text=display_text,
            editable=editable,
            dirty=dirty,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_date_field(
        self,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_date_field(
            display_text=display_text,
            editable=editable,
            dirty=dirty,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_bookmark(self, name: str, *, paragraph_index: int | None = None) -> None:
        self.document.append_bookmark(name, section_index=self.index, paragraph_index=paragraph_index)

    def append_note(
        self,
        text: str,
        *,
        kind: str = "footNote",
        number: int | str | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_note(
            text,
            kind=kind,
            number=number,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_form(
        self,
        label: str = "",
        *,
        form_type: str = "INPUT",
        name: str | None = None,
        value: str | None = None,
        checked: bool = False,
        items: Sequence[str] | None = None,
        editable: bool = True,
        locked: bool = False,
        placeholder: str | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_form(
            label,
            form_type=form_type,
            name=name,
            value=value,
            checked=checked,
            items=items,
            editable=editable,
            locked=locked,
            placeholder=placeholder,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_memo(
        self,
        text: str,
        *,
        author: str | None = None,
        memo_id: str | None = None,
        anchor_id: str | None = None,
        order: int | None = None,
        visible: bool = True,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_memo(
            text,
            author=author,
            memo_id=memo_id,
            anchor_id=anchor_id,
            order=order,
            visible=visible,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_chart(
        self,
        title: str = "",
        *,
        chart_type: str = "BAR",
        categories: Sequence[str] | None = None,
        series: Sequence[dict[str, object]] | None = None,
        data_ref: str | None = None,
        legend_visible: bool = True,
        width: int = 12000,
        height: int = 3200,
        shape_comment: str | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_chart(
            title,
            chart_type=chart_type,
            categories=categories,
            series=series,
            data_ref=data_ref,
            legend_visible=legend_visible,
            width=width,
            height=height,
            shape_comment=shape_comment,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_footnote(
        self,
        text: str,
        *,
        number: int | str | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        self.append_note(text, kind="footNote", number=number, paragraph_index=paragraph_index)

    def append_endnote(
        self,
        text: str,
        *,
        number: int | str | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        self.append_note(text, kind="endNote", number=number, paragraph_index=paragraph_index)

    def append_auto_number(
        self,
        *,
        number: int | str = 1,
        number_type: str = "PAGE",
        kind: str = "newNum",
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_auto_number(
            number=number,
            number_type=number_type,
            kind=kind,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_header(self, text: str, *, apply_page_type: str = "BOTH") -> None:
        self.document.append_header(text, apply_page_type=apply_page_type, section_index=self.index)

    def append_footer(self, text: str, *, apply_page_type: str = "BOTH") -> None:
        self.document.append_footer(text, apply_page_type=apply_page_type, section_index=self.index)

    def append_equation(
        self,
        script: str,
        *,
        width: int = 4800,
        height: int = 2300,
        font: str = DEFAULT_HWP_EQUATION_FONT,
        shape_comment: str | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_equation(
            script,
            width=width,
            height=height,
            font=font,
            shape_comment=shape_comment,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_shape(
        self,
        *,
        kind: str = "rect",
        text: str = "",
        width: int = 12000,
        height: int = 3200,
        fill_color: str = "#FFFFFF",
        line_color: str = "#000000",
        shape_comment: str | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_shape(
            kind=kind,
            text=text,
            width=width,
            height=height,
            fill_color=fill_color,
            line_color=line_color,
            shape_comment=shape_comment,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_ole(
        self,
        name: str,
        data: bytes,
        *,
        width: int = 42001,
        height: int = 13501,
        shape_comment: str | None = None,
        object_type: str = "EMBEDDED",
        draw_aspect: str = "CONTENT",
        has_moniker: bool = False,
        eq_baseline: int = 0,
        line_color: str = "#000000",
        line_width: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_ole(
            name,
            data,
            width=width,
            height=height,
            shape_comment=shape_comment,
            object_type=object_type,
            draw_aspect=draw_aspect,
            has_moniker=has_moniker,
            eq_baseline=eq_baseline,
            line_color=line_color,
            line_width=line_width,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )


class HwpParagraphObject:
    def __init__(self, document: HwpDocument, paragraph: HwpParagraph):
        self._document = document
        self._section_index = paragraph.section_index
        self._index = paragraph.index

    @property
    def document(self) -> HwpDocument:
        return self._document

    @property
    def index(self) -> int:
        return self._index

    @property
    def section_index(self) -> int:
        return self._section_index

    def _snapshot(self) -> HwpParagraph:
        return self._document.binary_document().paragraphs(self._section_index)[self._index]

    @property
    def text(self) -> str:
        return self._snapshot().text

    @property
    def raw_text(self) -> str:
        return self._snapshot().raw_text

    @property
    def char_count(self) -> int:
        return self._snapshot().char_count

    @property
    def control_mask(self) -> int:
        return self._snapshot().control_mask

    @property
    def para_shape_id(self) -> int:
        return self._snapshot().para_shape_id

    @property
    def style_id(self) -> int:
        return self._snapshot().style_id

    @property
    def split_flags(self) -> int:
        return self._snapshot().split_flags

    @property
    def has_hidden_controls(self) -> bool:
        snapshot = self._snapshot()
        return snapshot.raw_text != snapshot.text

    def set_text_same_length(self, value: str) -> None:
        self._document.binary_document().set_paragraph_text_same_length(self._section_index, self._index, value)

    def replace_text_same_length(self, old: str, new: str, *, count: int = -1) -> int:
        return self._document.binary_document().replace_paragraph_text_same_length(
            self._section_index,
            self._index,
            old,
            new,
            count=count,
        )

    def set_text(self, value: str) -> None:
        model = self._document.section_model(self._section_index)
        paragraph = model.paragraphs()[self._index]
        text_record = paragraph.text_record()
        if text_record is None:
            raise ValueError("HWP paragraph does not contain a para text record.")
        text_record.set_raw_text(value)
        paragraph.header.char_count = len(value)
        self._document.binary_document().replace_section_model(self._section_index, model)
        self._document._invalidate_bridge()

    def replace_text(self, old: str, new: str, *, count: int = -1) -> int:
        model = self._document.section_model(self._section_index)
        paragraph = model.paragraphs()[self._index]
        text_record = paragraph.text_record()
        if text_record is None or not old:
            return 0
        raw_text = text_record.raw_text
        replaced = raw_text.count(old) if count < 0 else min(raw_text.count(old), count)
        if replaced == 0:
            return 0
        updated = raw_text.replace(old, new) if count < 0 else raw_text.replace(old, new, count)
        text_record.set_raw_text(updated)
        paragraph.header.char_count = len(updated)
        self._document.binary_document().replace_section_model(self._section_index, model)
        self._document._invalidate_bridge()
        return replaced


class HwpControlObject:
    def __init__(
        self,
        document: HwpDocument,
        paragraph: SectionParagraphModel,
        control_node: RecordNode,
        control_ordinal: int,
    ):
        self._document = document
        self._paragraph = paragraph
        self._control_ordinal = control_ordinal
        self._initial_control_node = control_node

    @property
    def document(self) -> HwpDocument:
        return self._document

    @property
    def section_index(self) -> int:
        return self._paragraph.section_index

    @property
    def paragraph_index(self) -> int:
        return self._paragraph.index

    @property
    def control_node(self) -> RecordNode:
        return self._live_context()[2]

    @property
    def control_id(self) -> str:
        payload = self.control_node.payload
        if len(payload) < 4:
            return ""
        return payload[:4][::-1].decode("latin1", errors="replace")

    def descendant_nodes(self) -> list[RecordNode]:
        return list(self.control_node.iter_descendants())

    def _live_context(self) -> tuple[SectionModel, SectionParagraphModel, RecordNode]:
        section_model = self._document.section_model(self.section_index)
        paragraph = section_model.paragraphs()[self.paragraph_index]
        controls = [child for child in paragraph.header.children if child.tag_id == TAG_CTRL_HEADER]
        control_node = controls[self._control_ordinal]
        return section_model, paragraph, control_node

    def _commit(self, section_model: SectionModel) -> None:
        self._document.binary_document().replace_section_model(self.section_index, section_model)
        self._document._invalidate_bridge()

    def _native_shape_metadata(self, control_node: RecordNode | None = None) -> dict[str, str]:
        target = self.control_node if control_node is None else control_node
        return _collect_shape_native_metadata(target)

    def _native_control_metadata(self, control_node: RecordNode | None = None) -> dict[str, str]:
        target = self.control_node if control_node is None else control_node
        return _collect_control_native_metadata(target)

    def _graphic_out_margins(self, control_node: RecordNode | None = None) -> dict[str, int]:
        target = self.control_node if control_node is None else control_node
        native = _parse_graphic_control_out_margins(target.payload)
        return native or _metadata_prefixed_int_map(_collect_shape_native_metadata(target), "outMargin.")

    def _graphic_layout(self, control_node: RecordNode | None = None) -> dict[str, str]:
        target = self.control_node if control_node is None else control_node
        native = _parse_graphic_control_layout(target.payload)
        metadata = _metadata_prefixed_str_map(_collect_shape_native_metadata(target), "layout.")
        if not native:
            return metadata
        merged = dict(metadata)
        merged.update(native)
        return merged

    def _set_graphic_out_margins(
        self,
        *,
        left: int | str | None = None,
        right: int | str | None = None,
        top: int | str | None = None,
        bottom: int | str | None = None,
    ) -> None:
        section_model, _paragraph, control_node = self._live_context()
        control_node.payload = _set_graphic_control_out_margins_payload(
            control_node.payload,
            left=left,
            right=right,
            top=top,
            bottom=bottom,
        )
        _update_shape_native_metadata(control_node, updates={}, remove_prefixes=("outMargin.",))
        self._commit(section_model)

    def _set_graphic_layout(
        self,
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
    ) -> None:
        section_model, _paragraph, control_node = self._live_context()
        current_layout = self._graphic_layout(control_node)
        native_text_wrap = text_wrap if text_wrap in {"TOP_AND_BOTTOM", "SQUARE"} else None
        effective_text_wrap = native_text_wrap or current_layout.get("textWrap")
        native_text_flow = text_flow if text_flow in {"BOTH_SIDES", "LEFT_ONLY", "RIGHT_ONLY"} and effective_text_wrap == "SQUARE" else None
        native_treat_as_char = treat_as_char
        native_vert_rel_to = vert_rel_to if vert_rel_to in {"PARA", "PAPER"} else None
        native_horz_rel_to = horz_rel_to if horz_rel_to in {"COLUMN", "PAPER"} else None
        native_vert_align = vert_align if vert_align in {"TOP", "CENTER"} else None
        native_horz_align = horz_align if horz_align in {"LEFT", "RIGHT"} else None
        control_node.payload = _set_graphic_control_layout_payload(
            control_node.payload,
            text_wrap=native_text_wrap,
            text_flow=native_text_flow,
            treat_as_char=native_treat_as_char,
            vert_rel_to=native_vert_rel_to,
            horz_rel_to=native_horz_rel_to,
            vert_align=native_vert_align,
            horz_align=native_horz_align,
            vert_offset=vert_offset,
            horz_offset=horz_offset,
        )
        metadata_updates = _graphic_layout_metadata_updates(
            text_wrap=text_wrap,
            text_flow=text_flow,
            treat_as_char=treat_as_char,
            affect_line_spacing=affect_line_spacing,
            flow_with_text=flow_with_text,
            allow_overlap=allow_overlap,
            hold_anchor_and_so=hold_anchor_and_so,
            vert_rel_to=vert_rel_to,
            horz_rel_to=horz_rel_to,
            vert_align=vert_align,
            horz_align=horz_align,
            vert_offset=vert_offset,
            horz_offset=horz_offset,
        )
        remove_prefixes: list[str] = []
        for key, native_value in (
            ("layout.textWrap", native_text_wrap),
            ("layout.textFlow", native_text_flow),
            ("layout.treatAsChar", native_treat_as_char),
            ("layout.vertRelTo", native_vert_rel_to),
            ("layout.horzRelTo", native_horz_rel_to),
            ("layout.vertAlign", vert_align if vert_align == "CENTER" else None),
            (
                "layout.horzAlign",
                horz_align if horz_align == "RIGHT" and effective_text_wrap != "SQUARE" else None,
            ),
            ("layout.vertOffset", vert_offset),
            ("layout.horzOffset", horz_offset),
        ):
            if native_value is not None:
                metadata_updates.pop(key, None)
                remove_prefixes.append(key)
        _update_shape_native_metadata(control_node, updates=metadata_updates, remove_prefixes=tuple(remove_prefixes))
        self._commit(section_model)

    def _write_native_shape_metadata(
        self,
        *,
        updates: dict[str, object | None],
        remove_prefixes: Sequence[str] = (),
    ) -> None:
        section_model, _paragraph, control_node = self._live_context()
        _update_shape_native_metadata(control_node, updates=updates, remove_prefixes=remove_prefixes)
        self._commit(section_model)

    def _write_native_control_metadata(
        self,
        *,
        updates: dict[str, object | None],
        remove_prefixes: Sequence[str] = (),
    ) -> None:
        section_model, _paragraph, control_node = self._live_context()
        _update_control_native_metadata(control_node, updates=updates, remove_prefixes=remove_prefixes)
        self._commit(section_model)


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


class HwpPictureObject(HwpControlObject):
    @property
    def shape_comment(self) -> str:
        return _parse_object_description(self.control_node.payload)

    @property
    def storage_id(self) -> int:
        payload = self._shape_picture_record().payload
        return int.from_bytes(payload[71:73], "little")

    def bindata_path(self) -> str:
        prefix = f"BinData/BIN{self.storage_id:04d}."
        for stream_path in self.document.bindata_stream_paths():
            if stream_path.startswith(prefix):
                return stream_path
        raise KeyError(f"Could not find BinData stream for storage id {self.storage_id}.")

    @property
    def extension(self) -> str | None:
        suffix = Path(self.bindata_path()).suffix
        return suffix.lstrip(".") or None

    def size(self) -> dict[str, int]:
        shape_component = self._shape_component_record()
        if shape_component is None or len(shape_component.payload) < 28:
            return {"width": 0, "height": 0}
        return {
            "width": int.from_bytes(shape_component.payload[20:24], "little", signed=False),
            "height": int.from_bytes(shape_component.payload[24:28], "little", signed=False),
        }

    @property
    def line_color(self) -> str:
        shape_component = self._shape_component_record()
        if shape_component is not None and len(shape_component.payload) >= 202:
            component_color = shape_component.payload[196:200]
            component_width = int.from_bytes(shape_component.payload[200:202], "little", signed=False)
            if any(component_color) or component_width:
                return _parse_colorref(component_color)
        picture_record = self._shape_picture_record()
        if len(picture_record.payload) >= 8 and any(picture_record.payload[0:8]):
            return _parse_rgb_color(picture_record.payload[0:4])
        if shape_component is None or len(shape_component.payload) < 200:
            return "#000000"
        return _parse_colorref(shape_component.payload[196:200])

    @property
    def line_width(self) -> int:
        shape_component = self._shape_component_record()
        if shape_component is not None and len(shape_component.payload) >= 202:
            component_color = shape_component.payload[196:200]
            component_width = int.from_bytes(shape_component.payload[200:202], "little", signed=False)
            if any(component_color) or component_width:
                return component_width
        picture_record = self._shape_picture_record()
        if len(picture_record.payload) >= 8 and any(picture_record.payload[0:8]):
            return int.from_bytes(picture_record.payload[4:8], "little", signed=False)
        if shape_component is None or len(shape_component.payload) < 202:
            return 0
        return int.from_bytes(shape_component.payload[200:202], "little", signed=False)

    def layout(self) -> dict[str, str]:
        return self._graphic_layout()

    def out_margins(self) -> dict[str, int]:
        return self._graphic_out_margins()

    def rotation(self) -> dict[str, str]:
        shape_component = self._shape_component_record()
        native = {} if shape_component is None else _parse_shape_component_rotation_payload(shape_component.payload, include_center=True)
        metadata = _metadata_prefixed_str_map(self._native_shape_metadata(), "rotation.")
        if not native:
            return metadata
        merged = dict(metadata)
        merged.update(native)
        return merged

    def image_adjustment(self) -> dict[str, str]:
        native = _parse_picture_image_adjustment_payload(self._shape_picture_record().payload)
        return native or _metadata_prefixed_str_map(self._native_shape_metadata(), "image.")

    def crop(self) -> dict[str, int]:
        native = _parse_picture_crop_payload(self._shape_picture_record().payload)
        return native or _metadata_prefixed_int_map(self._native_shape_metadata(), "crop.")

    def line_style(self) -> dict[str, str]:
        result = {
            "color": self.line_color,
            "width": str(self.line_width),
        }
        result.update(_metadata_prefixed_str_map(self._native_shape_metadata(), "lineStyle."))
        return result

    def binary_data(self) -> bytes:
        return self.document.binary_document().read_stream(self.bindata_path(), decompress=False)

    def replace_binary(self, data: bytes, *, extension: str | None = None) -> None:
        if extension is None or extension.lower().lstrip(".") == (self.extension or "").lower():
            self.document.binary_document().add_stream(self.bindata_path(), data)
        else:
            section_model, _paragraph, control_node = self._live_context()
            storage_id, _ = self.document.add_embedded_bindata(data, extension=extension)
            payload = bytearray(self._shape_picture_record(control_node).payload)
            payload[71:73] = int(storage_id).to_bytes(2, "little")
            self._shape_picture_record(control_node).payload = bytes(payload)
            self._commit(section_model)
            return
        self.document._invalidate_bridge()

    def set_shape_comment(self, value: str | None) -> None:
        section_model, _paragraph, control_node = self._live_context()
        control_node.payload = _replace_object_description(control_node.payload, value)
        self._commit(section_model)

    def set_layout(
        self,
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
    ) -> None:
        self._set_graphic_layout(
            text_wrap=text_wrap,
            text_flow=text_flow,
            treat_as_char=treat_as_char,
            affect_line_spacing=affect_line_spacing,
            flow_with_text=flow_with_text,
            allow_overlap=allow_overlap,
            hold_anchor_and_so=hold_anchor_and_so,
            vert_rel_to=vert_rel_to,
            horz_rel_to=horz_rel_to,
            vert_align=vert_align,
            horz_align=horz_align,
            vert_offset=vert_offset,
            horz_offset=horz_offset,
        )

    def set_out_margins(
        self,
        *,
        left: int | str | None = None,
        right: int | str | None = None,
        top: int | str | None = None,
        bottom: int | str | None = None,
    ) -> None:
        self._set_graphic_out_margins(left=left, right=right, top=top, bottom=bottom)

    def set_rotation(
        self,
        *,
        angle: int | str | None = None,
        center_x: int | str | None = None,
        center_y: int | str | None = None,
        rotate_image: bool | None = None,
    ) -> None:
        section_model, _paragraph, control_node = self._live_context()
        shape_component = self._shape_component_record(control_node)
        if shape_component is None:
            raise ValueError("HWP picture control does not contain a shape component record.")
        shape_component.payload = _set_shape_component_rotation_payload(
            shape_component.payload,
            angle=angle,
            center_x=center_x,
            center_y=center_y,
        )
        _update_shape_native_metadata(
            control_node,
            updates={"rotation.rotateimage": rotate_image},
            remove_prefixes=("rotation.angle", "rotation.centerX", "rotation.centerY"),
        )
        self._commit(section_model)

    def set_image_adjustment(
        self,
        *,
        bright: int | str | None = None,
        contrast: int | str | None = None,
        effect: str | None = None,
        alpha: int | str | None = None,
    ) -> None:
        section_model, _paragraph, control_node = self._live_context()
        effect_code = _picture_effect_code(effect)
        picture_record = self._shape_picture_record(control_node)
        picture_record.payload = _set_picture_image_adjustment_payload(
            picture_record.payload,
            bright=bright,
            contrast=contrast,
            effect_code=effect_code,
            alpha=alpha,
        )
        if effect is not None and effect_code is None:
            _update_shape_native_metadata(
                control_node,
                updates=_image_adjustment_metadata_updates(
                    bright=bright,
                    contrast=contrast,
                    effect=effect,
                    alpha=alpha,
                ),
            )
        else:
            _update_shape_native_metadata(control_node, updates={}, remove_prefixes=("image.",))
        self._commit(section_model)

    def set_crop(
        self,
        *,
        left: int | str | None = None,
        right: int | str | None = None,
        top: int | str | None = None,
        bottom: int | str | None = None,
    ) -> None:
        section_model, _paragraph, control_node = self._live_context()
        picture_record = self._shape_picture_record(control_node)
        picture_record.payload = _set_picture_crop_payload(
            picture_record.payload,
            left=left,
            right=right,
            top=top,
            bottom=bottom,
        )
        _update_shape_native_metadata(control_node, updates={}, remove_prefixes=("crop.",))
        self._commit(section_model)

    def set_size(self, *, width: int | None = None, height: int | None = None) -> None:
        section_model, _paragraph, control_node = self._live_context()
        control_node.payload = _set_payload_size(control_node.payload, width=width, height=height, width_offset=16, height_offset=20)
        record = self._shape_component_record(control_node)
        if record is None:
            raise ValueError("HWP picture control does not contain a shape component record.")
        record.payload = _set_payload_size(record.payload, width=width, height=height, width_offset=20, height_offset=24)
        self._commit(section_model)

    def set_line_style(
        self,
        *,
        color: str | None = None,
        width: int | None = None,
        style: str | None = None,
        end_cap: str | None = None,
        head_style: str | None = None,
        tail_style: str | None = None,
        head_fill: bool | None = None,
        tail_fill: bool | None = None,
        head_size: str | None = None,
        tail_size: str | None = None,
        outline_style: str | None = None,
        alpha: int | str | None = None,
    ) -> None:
        section_model, _paragraph, control_node = self._live_context()
        record = self._shape_component_record(control_node)
        if record is None:
            raise ValueError("HWP picture control does not contain a shape component record.")
        picture_record = self._shape_picture_record(control_node)
        payload = bytearray(record.payload)
        if len(payload) < 202:
            payload.extend(b"\x00" * (202 - len(payload)))
        if color is not None:
            payload[196:200] = _build_colorref(color)
        if width is not None:
            payload[200:202] = max(0, min(int(width), 0xFFFF)).to_bytes(2, "little", signed=False)
        record.payload = bytes(payload)
        picture_record.payload = _set_picture_line_payload(
            picture_record.payload,
            color=color,
            width=width,
        )
        _update_shape_native_metadata(
            control_node,
            updates={
                "lineStyle.style": style,
                "lineStyle.endCap": end_cap,
                "lineStyle.headStyle": head_style,
                "lineStyle.tailStyle": tail_style,
                "lineStyle.headfill": head_fill,
                "lineStyle.tailfill": tail_fill,
                "lineStyle.headSz": head_size,
                "lineStyle.tailSz": tail_size,
                "lineStyle.outlineStyle": outline_style,
                "lineStyle.alpha": alpha,
            },
        )
        self._commit(section_model)

    def _shape_picture_record(self, control_node: RecordNode | None = None) -> RecordNode:
        target = self.control_node if control_node is None else control_node
        for node in target.iter_descendants():
            if node.tag_id == TAG_SHAPE_COMPONENT_PICTURE:
                return node
        raise ValueError("HWP picture control does not contain a shape picture record.")

    def _shape_component_record(self, control_node: RecordNode | None = None) -> RecordNode | None:
        target = self.control_node if control_node is None else control_node
        for node in target.iter_descendants():
            if node.tag_id == TAG_SHAPE_COMPONENT:
                return node
        return None


class HwpBookmarkObject(HwpControlObject):
    @property
    def name(self) -> str:
        data_node = next((child for child in self.control_node.children if child.tag_id == TAG_CTRL_DATA), None)
        if data_node is None or len(data_node.payload) < 12:
            return ""
        name_size = int.from_bytes(data_node.payload[10:12], "little", signed=False)
        encoded_name = data_node.payload[12 : 12 + name_size * 2]
        return encoded_name.decode("utf-16-le", errors="ignore")

    def rename(self, value: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        data_node = next((child for child in control_node.children if child.tag_id == TAG_CTRL_DATA), None)
        if data_node is None:
            raise ValueError("HWP bookmark control does not contain ctrl data.")
        prefix = data_node.payload[:10]
        payload = bytearray(prefix)
        payload.extend(len(value).to_bytes(2, "little", signed=False))
        payload.extend(value.encode("utf-16-le"))
        data_node.payload = bytes(payload)
        self._commit(section_model)


class HwpFieldObject(HwpControlObject):
    @property
    def native_field_type(self) -> str:
        return self.control_id

    @property
    def field_type(self) -> str:
        return _semantic_field_type_from_native(self.native_field_type, self.parameters)

    @property
    def is_mail_merge(self) -> bool:
        return self.field_type == "MAILMERGE"

    @property
    def is_calculation(self) -> bool:
        return self.field_type == "FORMULA"

    @property
    def is_cross_reference(self) -> bool:
        return self.field_type == "CROSSREF"

    @property
    def is_doc_property(self) -> bool:
        return self.field_type == "DOCPROPERTY"

    @property
    def is_date(self) -> bool:
        return self.field_type == "DATE"

    @property
    def parameters(self) -> dict[str, str]:
        data_node = next((child for child in self.control_node.children if child.tag_id == TAG_CTRL_DATA), None)
        if data_node is None:
            return {}
        decoded = data_node.payload.decode("utf-16-le", errors="ignore")
        parameters: dict[str, str] = {}
        for entry in decoded.split(";"):
            key, separator, value = entry.partition("=")
            if separator and key:
                parameters[key] = value
        return parameters

    def get_parameter(self, name: str) -> str | None:
        return self.parameters.get(name)

    @property
    def name(self) -> str | None:
        for key in ("FieldName", "Name", "name"):
            value = self.parameters.get(key)
            if value:
                return value
        return None

    @property
    def merge_field_name(self) -> str | None:
        return self.get_parameter("MergeField") or self.get_parameter("FieldName") or self.name

    @property
    def expression(self) -> str | None:
        return self.get_parameter("Expression") or self.get_parameter("Command")

    @property
    def bookmark_name(self) -> str | None:
        return self.get_parameter("BookmarkName") or self.get_parameter("Path") or self.name

    @property
    def document_property_name(self) -> str | None:
        if not self.is_doc_property:
            return None
        return self.get_parameter("FieldName") or self.name

    @property
    def display_text(self) -> str:
        return self._paragraph.text.replace("\r", "")

    @property
    def editable(self) -> bool:
        return bool(self._paragraph.header.control_mask & 0x01)

    @property
    def dirty(self) -> bool:
        return bool(self._paragraph.header.control_mask & 0x01)

    def _write_parameters(self, section_model: SectionModel, control_node: RecordNode, parameters: dict[str, str]) -> None:
        data_node = next((child for child in control_node.children if child.tag_id == TAG_CTRL_DATA), None)
        if data_node is None:
            data_node = RecordNode(tag_id=TAG_CTRL_DATA, level=control_node.level + 1, payload=b"")
            control_node.add_child(data_node)
        parameter_text = ";".join(f"{key}={value}" for key, value in parameters.items())
        data_node.payload = parameter_text.encode("utf-16-le")
        self._commit(section_model)

    def set_parameter(self, name: str, value: str | int) -> None:
        section_model, _paragraph, control_node = self._live_context()
        parameters = dict(self.parameters)
        parameters[str(name)] = str(value)
        self._write_parameters(section_model, control_node, parameters)

    def set_name(self, value: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        parameters = dict(self.parameters)
        updated = False
        for key in ("FieldName", "Name", "name"):
            if key in parameters:
                parameters[key] = value
                updated = True
        if self.is_mail_merge:
            parameters["MergeField"] = value
        if self.is_cross_reference:
            if "BookmarkName" in parameters or not updated:
                parameters["BookmarkName"] = value
            if "Path" in parameters or not updated:
                parameters["Path"] = value
        if not updated and not self.is_cross_reference:
            parameters["FieldName"] = value
        self._write_parameters(section_model, control_node, parameters)

    def set_field_type(self, value: str) -> None:
        normalized = value.strip()
        if normalized.upper() in {"HYPERLINK", "%HLK"}:
            raise ValueError("Use HwpHyperlinkObject to edit hyperlink controls.")
        section_model, _paragraph, control_node = self._live_context()
        control_node.payload = _build_control_id_payload(_normalize_field_control_id(normalized))
        self._commit(section_model)

    def set_display_text(self, text: str) -> None:
        section_model, paragraph, _control_node = self._live_context()
        text_record = paragraph.text_record()
        if text_record is None:
            raise ValueError("HWP field paragraph does not contain a para text record.")
        raw_text = f"{text}\r" if text else ""
        text_record.set_raw_text(raw_text)
        paragraph.header.char_count = len(raw_text)
        self._commit(section_model)

    def configure_mail_merge(self, field_name: str, *, display_text: str | None = None) -> None:
        self.set_field_type("MAILMERGE")
        self.set_name(field_name)
        self.set_parameter("FieldName", field_name)
        self.set_parameter("MergeField", field_name)
        if display_text is not None:
            self.set_display_text(display_text)

    def configure_calculation(self, expression: str, *, display_text: str | None = None) -> None:
        self.set_field_type("FORMULA")
        self.set_parameter("Expression", expression)
        self.set_parameter("Command", expression)
        if display_text is not None:
            self.set_display_text(display_text)

    def configure_cross_reference(self, bookmark_name: str, *, display_text: str | None = None) -> None:
        self.set_field_type("CROSSREF")
        self.set_name(bookmark_name)
        self.set_parameter("BookmarkName", bookmark_name)
        self.set_parameter("Path", bookmark_name)
        if display_text is not None:
            self.set_display_text(display_text)

    def configure_doc_property(self, property_name: str, *, display_text: str | None = None) -> None:
        self.set_field_type("DOCPROPERTY")
        self.set_name(property_name)
        self.set_parameter("FieldName", property_name)
        if display_text is not None:
            self.set_display_text(display_text)

    def configure_date(self, *, display_text: str | None = None) -> None:
        self.set_field_type("DATE")
        if display_text is not None:
            self.set_display_text(display_text)


class HwpHyperlinkObject(HwpFieldObject):
    @property
    def command(self) -> str:
        payload = self.control_node.payload
        prefix_size = 9
        text_size = int.from_bytes(payload[prefix_size : prefix_size + 2], "little")
        start = prefix_size + 2
        end = start + text_size * 2
        return payload[start:end].decode("utf-16-le", errors="ignore")

    @property
    def url(self) -> str:
        command = self.command
        return command.split(";", 1)[0].replace("\\:", ":")

    @property
    def metadata_fields(self) -> list[str]:
        parts = self.command.split(";")
        if parts and parts[-1] == "":
            parts = parts[:-1]
        return parts[1:]

    @property
    def display_text(self) -> str:
        return self._paragraph.text.replace("\r", "")

    def set_url(self, url: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        control_node.payload = _build_hyperlink_ctrl_payload(url, metadata_fields=self.metadata_fields)
        self._commit(section_model)

    def set_metadata_fields(self, metadata_fields: Sequence[str | int]) -> None:
        section_model, _paragraph, control_node = self._live_context()
        control_node.payload = _build_hyperlink_ctrl_payload(self.url, metadata_fields=metadata_fields)
        self._commit(section_model)

    def set_display_text(self, text: str) -> None:
        section_model, paragraph, _control_node = self._live_context()
        text_record = paragraph.text_record()
        if text_record is None:
            raise ValueError("HWP hyperlink paragraph does not contain a para text record.")
        raw_text = _build_hyperlink_para_text_payload(text).decode("utf-16-le", errors="ignore")
        text_record.set_raw_text(raw_text)
        paragraph.header.char_count = len(raw_text)
        self._commit(section_model)


class HwpAutoNumberObject(HwpControlObject):
    @property
    def kind(self) -> str:
        return "newNum" if self.control_id == "nwno" else "autoNum"

    @property
    def number(self) -> str | None:
        payload = self.control_node.payload
        if self.control_id == "nwno" and len(payload) >= 10:
            return str(int.from_bytes(payload[8:10], "little"))
        if self.control_id == "atno":
            return "1"
        return None

    @property
    def number_type(self) -> str:
        if self.control_id == "atno" and len(self.control_node.payload) >= 12:
            code = int.from_bytes(self.control_node.payload[8:12], "little", signed=False)
            return _AUTO_NUMBER_TYPE_VALUES.get(code, "PAGE")
        return "PAGE"

    def set_number(self, value: int | str) -> None:
        if self.control_id != "nwno":
            raise ValueError("Only newNum controls support explicit numeric rewrites in the native HWP wrapper.")
        section_model, _paragraph, control_node = self._live_context()
        payload = bytearray(control_node.payload)
        payload[8:10] = int(value).to_bytes(2, "little", signed=False)
        control_node.payload = bytes(payload)
        self._commit(section_model)

    def set_number_type(self, value: str) -> None:
        normalized = str(value).upper()
        if self.control_id != "atno":
            if normalized != "PAGE":
                raise ValueError("Only autoNum controls support editable number_type in the native HWP wrapper.")
            return
        section_model, _paragraph, control_node = self._live_context()
        control_node.payload = _build_auto_number_payload("autoNum", number=1, number_type=normalized)
        self._commit(section_model)


class HwpHeaderFooterObject(HwpControlObject):
    @property
    def kind(self) -> str:
        return "header" if self.control_id == "head" else "footer"

    @property
    def text(self) -> str:
        return _collect_text_from_node(self.control_node).replace("\r", "").strip()

    @property
    def apply_page_type(self) -> str:
        payload = self.control_node.payload
        if len(payload) < 8:
            return "BOTH"
        code = int.from_bytes(payload[4:8], "little", signed=False)
        return {0: "BOTH", 1: "EVEN", 2: "ODD"}.get(code, "BOTH")

    def set_text(self, value: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        paragraph_headers = [node for node in control_node.children if isinstance(node, ParagraphHeaderRecord)]
        if not paragraph_headers:
            raise ValueError("HWP header/footer control does not contain a writable paragraph header.")
        target = paragraph_headers[0]
        text_record = next((child for child in target.children if isinstance(child, ParagraphTextRecord)), None)
        raw_text = f"{value}\r" if value else ""
        if text_record is None:
            text_record = ParagraphTextRecord(level=target.level + 1, raw_text=raw_text)
            target.children.insert(0, text_record)
        else:
            text_record.set_raw_text(raw_text)
        target.char_count = len(raw_text)
        target.sync_payload()
        self._commit(section_model)

    def set_apply_page_type(self, value: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        control_node.payload = _build_header_footer_control_payload(self.control_id, apply_page_type=str(value).upper())
        self._commit(section_model)


class HwpNoteObject(HwpControlObject):
    def _note_ordinal(self, numbering_type: str) -> int:
        count = 0
        target_control_id = self.control_id
        if numbering_type == "CONTINUOUS":
            section_indexes = range(self.section_index + 1)
        else:
            section_indexes = range(self.section_index, self.section_index + 1)
        for section_index in section_indexes:
            section_model = self.document.section_model(section_index)
            for paragraph in section_model.paragraphs():
                controls = [child for child in paragraph.header.children if child.tag_id == TAG_CTRL_HEADER]
                for ordinal, control in enumerate(controls):
                    payload = control.payload
                    if len(payload) < 4 or payload[:4][::-1].decode("latin1", errors="replace") != target_control_id:
                        continue
                    count += 1
                    if (
                        section_index == self.section_index
                        and paragraph.index == self.paragraph_index
                        and ordinal == self._control_ordinal
                    ):
                        return count
        return count

    @property
    def kind(self) -> str:
        return "footNote" if self.control_id == "fn  " else "endNote"

    @property
    def number(self) -> str | None:
        note_settings = self.document.binary_document().section_note_settings(self.section_index)
        settings_key = "footNotePr" if self.kind == "footNote" else "endNotePr"
        numbering = dict(note_settings.get(settings_key, {}).get("numbering", {}))
        numbering_type = str(numbering.get("type", "CONTINUOUS")).upper()
        if numbering_type == "ON_PAGE":
            return None
        try:
            start_number = int(str(numbering.get("newNum", "1")) or "1")
        except ValueError:
            start_number = 1
        ordinal = self._note_ordinal(numbering_type)
        if ordinal > 0:
            return str(start_number + ordinal - 1)
        return None

    @property
    def text(self) -> str:
        return _collect_text_from_node(self.control_node).replace("\r", "").strip()

    def set_text(self, value: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        paragraph_headers = [node for node in control_node.children if isinstance(node, ParagraphHeaderRecord)]
        if not paragraph_headers:
            raise ValueError("HWP note control does not contain a writable paragraph header.")
        target = paragraph_headers[0]
        text_record = next((child for child in target.children if isinstance(child, ParagraphTextRecord)), None)
        raw_text = f"{value}\r" if value else ""
        if text_record is None:
            text_record = ParagraphTextRecord(level=target.level + 1, raw_text=raw_text)
            target.children.insert(0, text_record)
        else:
            text_record.set_raw_text(raw_text)
        target.char_count = 0x80000000 | len(raw_text)
        target.sync_payload()
        self._commit(section_model)

    def set_number(self, value: int | str) -> None:
        note_settings = self.document.binary_document().section_note_settings(self.section_index)
        settings_key = "footNotePr" if self.kind == "footNote" else "endNotePr"
        note_pr = dict(note_settings.get(settings_key, {}))
        numbering = dict(note_pr.get("numbering", {}))
        numbering_type = str(numbering.get("type", "CONTINUOUS")).upper()
        if numbering_type == "ON_PAGE":
            raise ValueError("ON_PAGE note numbering cannot be rewritten by absolute note number.")
        ordinal = self._note_ordinal(numbering_type)
        if ordinal <= 0:
            raise ValueError("Failed to resolve the target native HWP note ordinal.")
        numbering["newNum"] = str(int(value) - ordinal + 1)
        note_pr["numbering"] = numbering
        if self.kind == "footNote":
            self.document.apply_section_note_settings(section_index=self.section_index, footnote_pr=note_pr)
        else:
            self.document.apply_section_note_settings(section_index=self.section_index, endnote_pr=note_pr)


class HwpPageNumObject(HwpControlObject):
    @property
    def page_number(self) -> dict[str, str]:
        payload = self.control_node.payload
        return {
            "pos": self.pos,
            "formatType": self.format_type,
            "sideChar": self.side_char if len(payload) >= 16 else "-",
        }

    @property
    def pos(self) -> str:
        payload = self.control_node.payload
        attributes = int.from_bytes(payload[4:8].ljust(4, b"\x00"), "little")
        positions = {
            0: "NONE",
            1: "TOP_LEFT",
            2: "TOP_CENTER",
            3: "TOP_RIGHT",
            4: "BOTTOM_LEFT",
            5: "BOTTOM_CENTER",
            6: "BOTTOM_RIGHT",
            7: "OUTSIDE_TOP",
            8: "OUTSIDE_BOTTOM",
            9: "INSIDE_TOP",
            10: "INSIDE_BOTTOM",
        }
        return positions.get((attributes >> 8) & 0x0F, "BOTTOM_CENTER")

    @property
    def format_type(self) -> str:
        payload = self.control_node.payload
        attributes = int.from_bytes(payload[4:8].ljust(4, b"\x00"), "little")
        formats = {
            0: "DIGIT",
            1: "CIRCLED_DIGIT",
            2: "ROMAN_CAPITAL",
            3: "ROMAN_SMALL",
            4: "LATIN_CAPITAL",
            5: "LATIN_SMALL",
        }
        return formats.get(attributes & 0xFF, "DIGIT")

    @property
    def side_char(self) -> str:
        payload = self.control_node.payload
        if len(payload) < 16:
            return "-"
        return payload[14:16].decode("utf-16-le", errors="ignore") or "-"

    def set_page_number(
        self,
        *,
        pos: str | None = None,
        format_type: str | None = None,
        side_char: str | None = None,
    ) -> None:
        section_model, _paragraph, control_node = self._live_context()
        page_number = dict(self.page_number)
        if pos is not None:
            page_number["pos"] = pos
        if format_type is not None:
            page_number["formatType"] = format_type
        if side_char is not None:
            page_number["sideChar"] = side_char
        control_node.payload = _build_page_num_payload(page_number)
        self._commit(section_model)

    def set_pos(self, value: str) -> None:
        self.set_page_number(pos=value)

    def set_format_type(self, value: str) -> None:
        self.set_page_number(format_type=value)

    def set_side_char(self, value: str) -> None:
        self.set_page_number(side_char=value)


class HwpEquationObject(HwpControlObject):
    @property
    def script(self) -> str:
        eqedit_record = self._eqedit_record()
        if eqedit_record is None:
            return ""
        decoded = eqedit_record.payload[4:].decode("utf-16-le", errors="ignore")
        script = decoded.split("Equation Version", 1)[0].replace("\x00", "").strip()
        if "ь" in script:
            script = script.split("ь", 1)[0]
        script = "".join(character for character in script if ord(character) >= 32 or character in "\t\n")
        if script.startswith("* "):
            script = script[2:]
        return script.rstrip("`").strip()

    @property
    def shape_comment(self) -> str:
        return _parse_object_description(self.control_node.payload)

    def layout(self) -> dict[str, str]:
        return self._graphic_layout()

    def out_margins(self) -> dict[str, int]:
        return self._graphic_out_margins()

    def rotation(self) -> dict[str, str]:
        return _metadata_prefixed_str_map(self._native_shape_metadata(), "rotation.")

    def size(self) -> dict[str, int]:
        payload = self.control_node.payload
        if len(payload) < 24:
            return {"width": 0, "height": 0}
        return {
            "width": int.from_bytes(payload[16:20], "little", signed=False),
            "height": int.from_bytes(payload[20:24], "little", signed=False),
        }

    @property
    def font(self) -> str:
        eqedit_record = self._eqedit_record()
        if eqedit_record is None:
            return DEFAULT_HWP_EQUATION_FONT
        decoded = eqedit_record.payload[4:].decode("utf-16-le", errors="ignore")
        marker_index = decoded.find("Equation Version")
        if marker_index < 0:
            return DEFAULT_HWP_EQUATION_FONT
        suffix = decoded[marker_index:]
        if "\x07" not in suffix:
            return DEFAULT_HWP_EQUATION_FONT
        font = suffix.split("\x07", 1)[1].replace("\x00", "").strip()
        return font or DEFAULT_HWP_EQUATION_FONT

    def _eqedit_record(self, control_node: RecordNode | None = None) -> RecordNode | None:
        target = self.control_node if control_node is None else control_node
        for node in target.iter_descendants():
            if node.tag_id == TAG_EQEDIT:
                return node
        return None

    def set_script(self, script: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        eqedit_record = self._eqedit_record(control_node)
        if eqedit_record is None:
            raise ValueError("HWP equation control does not contain an eqedit record.")
        decoded = eqedit_record.payload[4:].decode("utf-16-le", errors="ignore")
        prefix = "* " if decoded.startswith("* ") else ""
        marker_index = decoded.find("Equation Version")
        suffix = decoded[marker_index:] if marker_index >= 0 else ""
        eqedit_record.payload = eqedit_record.payload[:4] + (prefix + script + suffix).encode("utf-16-le")
        self._commit(section_model)

    def set_shape_comment(self, value: str | None) -> None:
        section_model, _paragraph, control_node = self._live_context()
        control_node.payload = _replace_object_description(control_node.payload, value)
        self._commit(section_model)

    def set_layout(
        self,
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
    ) -> None:
        self._set_graphic_layout(
            text_wrap=text_wrap,
            text_flow=text_flow,
            treat_as_char=treat_as_char,
            affect_line_spacing=affect_line_spacing,
            flow_with_text=flow_with_text,
            allow_overlap=allow_overlap,
            hold_anchor_and_so=hold_anchor_and_so,
            vert_rel_to=vert_rel_to,
            horz_rel_to=horz_rel_to,
            vert_align=vert_align,
            horz_align=horz_align,
            vert_offset=vert_offset,
            horz_offset=horz_offset,
        )

    def set_out_margins(
        self,
        *,
        left: int | str | None = None,
        right: int | str | None = None,
        top: int | str | None = None,
        bottom: int | str | None = None,
    ) -> None:
        self._set_graphic_out_margins(left=left, right=right, top=top, bottom=bottom)

    def set_rotation(
        self,
        *,
        angle: int | str | None = None,
        center_x: int | str | None = None,
        center_y: int | str | None = None,
        rotate_image: bool | None = None,
    ) -> None:
        self._write_native_shape_metadata(
            updates=_graphic_rotation_metadata_updates(
                angle=angle,
                center_x=center_x,
                center_y=center_y,
                rotate_image=rotate_image,
            )
        )

    def set_size(self, *, width: int | None = None, height: int | None = None) -> None:
        section_model, _paragraph, control_node = self._live_context()
        payload = bytearray(control_node.payload)
        if len(payload) < 24:
            raise ValueError("HWP equation control payload is shorter than expected.")
        if width is not None:
            payload[16:20] = int(width).to_bytes(4, "little", signed=False)
        if height is not None:
            payload[20:24] = int(height).to_bytes(4, "little", signed=False)
        control_node.payload = bytes(payload)
        self._commit(section_model)


class HwpShapeObject(HwpControlObject):
    @property
    def kind(self) -> str:
        return _shape_kind_from_control_node(self.control_node)

    @property
    def shape_comment(self) -> str:
        return _parse_object_description(self.control_node.payload)

    @property
    def text(self) -> str:
        metadata = self._shape_native_metadata()
        if "text" in metadata:
            return metadata["text"]
        if self.kind == "textart":
            specific_record = self._shape_specific_record()
            if specific_record is not None:
                native_text = _parse_textart_payload_text(specific_record.payload).strip()
                if native_text:
                    return native_text
        control_text = _collect_text_from_node(self.control_node).replace("\r", "").strip()
        if control_text:
            return control_text
        return self._paragraph.text.replace("\r", "")

    @property
    def line_color(self) -> str:
        metadata = self._shape_native_metadata()
        if "line_color" in metadata:
            return metadata["line_color"]
        shape_component = self._shape_component_record()
        if self.kind in {"ellipse", "arc", "polygon", "textart"} and shape_component is not None and len(shape_component.payload) >= 200:
            native_color = shape_component.payload[196:200]
            if any(native_color):
                return _parse_colorref(native_color)
        common_payload, _specific_payload = self._shape_specific_payload_parts()
        if len(common_payload) < 4:
            return "#000000"
        return _parse_colorref(common_payload[0:4])

    @property
    def fill_color(self) -> str:
        metadata = self._shape_native_metadata()
        if "fill_color" in metadata:
            return metadata["fill_color"]
        shape_component = self._shape_component_record()
        if self.kind == "textart" and shape_component is not None and len(shape_component.payload) >= 217:
            native_color = shape_component.payload[213:217]
            if any(native_color):
                return _parse_colorref(native_color)
        common_payload, _specific_payload = self._shape_specific_payload_parts()
        if len(common_payload) <= 11:
            return "#FFFFFF"
        fill_payload = common_payload[11:]
        return _parse_shape_fill_payload(fill_payload).get("faceColor", "#FFFFFF")

    def descendant_tag_ids(self) -> list[int]:
        return [node.tag_id for node in self.control_node.iter_descendants()]

    def layout(self) -> dict[str, str]:
        return self._graphic_layout()

    def out_margins(self) -> dict[str, int]:
        return self._graphic_out_margins()

    def rotation(self) -> dict[str, str]:
        shape_component = self._shape_component_record()
        native = {} if shape_component is None else _parse_shape_component_rotation_payload(shape_component.payload, include_center=True)
        metadata = _metadata_prefixed_str_map(self._shape_native_metadata(), "rotation.")
        if not native:
            return metadata
        merged = dict(metadata)
        merged.update(native)
        return merged

    def size(self) -> dict[str, int]:
        record = self._shape_component_record()
        width = 0
        height = 0
        if record is not None and len(record.payload) >= 28:
            width = int.from_bytes(record.payload[20:24], "little", signed=False)
            height = int.from_bytes(record.payload[24:28], "little", signed=False)
        if self.kind == "textart":
            native_width, native_height = self._control_size()
            if native_width or native_height:
                return {"width": native_width, "height": native_height}
        if width or height:
            return {"width": width, "height": height}
        native_width, native_height = self._control_size()
        return {"width": native_width, "height": native_height}

    def specific_fields(self) -> dict[str, object]:
        record = self._shape_specific_record()
        if record is None:
            return {}
        common_size = _shape_common_payload_size(record.tag_id)
        specific_payload = record.payload[common_size:] if common_size > 0 and len(record.payload) > common_size else record.payload
        if record.tag_id == TAG_SHAPE_COMPONENT_LINE:
            if self.kind == "connectLine" and _has_hancom_connectline_signature(self.control_node):
                return _parse_connectline_payload_fields(specific_payload or record.payload)
            return _parse_line_shape_payload(specific_payload)
        if record.tag_id == TAG_SHAPE_COMPONENT_RECTANGLE:
            return _parse_rectangle_shape_payload(specific_payload)
        if record.tag_id == TAG_SHAPE_COMPONENT_ELLIPSE:
            return _parse_ellipse_shape_payload(specific_payload)
        if record.tag_id == TAG_SHAPE_COMPONENT_ARC:
            return _parse_arc_shape_payload(specific_payload)
        if record.tag_id == TAG_SHAPE_COMPONENT_POLYGON:
            return _parse_polygon_shape_payload(specific_payload)
        if record.tag_id == TAG_SHAPE_COMPONENT_CURVE:
            return _parse_curve_shape_payload(specific_payload)
        if record.tag_id == TAG_SHAPE_COMPONENT_CONTAINER:
            return _parse_container_shape_payload(specific_payload)
        if record.tag_id == TAG_SHAPE_COMPONENT_TEXTART:
            return _parse_textart_payload_fields(specific_payload or record.payload)
        return {}

    def _control_size(self) -> tuple[int, int]:
        payload = self.control_node.payload
        if len(payload) < 24:
            return (0, 0)
        return (
            int.from_bytes(payload[16:20], "little", signed=False),
            int.from_bytes(payload[20:24], "little", signed=False),
        )

    def _shape_component_record(self, control_node: RecordNode | None = None) -> RecordNode | None:
        target = self.control_node if control_node is None else control_node
        for node in target.iter_descendants():
            if node.tag_id == TAG_SHAPE_COMPONENT:
                return node
        return None

    def _shape_native_metadata(self, control_node: RecordNode | None = None) -> dict[str, str]:
        target = self.control_node if control_node is None else control_node
        return _collect_shape_native_metadata(target)

    def _shape_specific_record(self, control_node: RecordNode | None = None) -> RecordNode | None:
        target = self.control_node if control_node is None else control_node
        shape_tags = {
            TAG_SHAPE_COMPONENT_LINE,
            TAG_SHAPE_COMPONENT_RECTANGLE,
            TAG_SHAPE_COMPONENT_ELLIPSE,
            TAG_SHAPE_COMPONENT_ARC,
            TAG_SHAPE_COMPONENT_POLYGON,
            TAG_SHAPE_COMPONENT_CURVE,
            TAG_SHAPE_COMPONENT_CONTAINER,
            TAG_SHAPE_COMPONENT_TEXTART,
            TAG_SHAPE_COMPONENT_OLE,
            TAG_SHAPE_COMPONENT_PICTURE,
        }
        for node in target.iter_descendants():
            if node.tag_id in shape_tags:
                return node
        return None

    def _shape_specific_payload_parts(self, control_node: RecordNode | None = None) -> tuple[bytes, bytes]:
        record = self._shape_specific_record(control_node)
        if record is None:
            return b"", b""
        common_size = _shape_common_payload_size(record.tag_id)
        if common_size <= 0:
            return b"", record.payload
        if len(record.payload) <= common_size:
            return record.payload, b""
        return record.payload[:common_size], record.payload[common_size:]

    def _shape_specific_payload(self, control_node: RecordNode | None = None) -> bytes:
        _common_payload, specific_payload = self._shape_specific_payload_parts(control_node)
        return specific_payload

    def _set_shape_specific_payload(self, specific_payload: bytes) -> None:
        section_model, _paragraph, control_node = self._live_context()
        record = self._shape_specific_record(control_node)
        if record is None:
            raise ValueError("HWP shape control does not contain a subtype-specific shape record.")
        common_size = _shape_common_payload_size(record.tag_id)
        common_payload = record.payload[:common_size] if common_size > 0 else b""
        record.payload = common_payload + bytes(specific_payload)
        self._commit(section_model)

    def set_shape_comment(self, value: str | None) -> None:
        section_model, _paragraph, control_node = self._live_context()
        control_node.payload = _replace_object_description(control_node.payload, value)
        self._commit(section_model)

    def set_layout(
        self,
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
    ) -> None:
        self._set_graphic_layout(
            text_wrap=text_wrap,
            text_flow=text_flow,
            treat_as_char=treat_as_char,
            affect_line_spacing=affect_line_spacing,
            flow_with_text=flow_with_text,
            allow_overlap=allow_overlap,
            hold_anchor_and_so=hold_anchor_and_so,
            vert_rel_to=vert_rel_to,
            horz_rel_to=horz_rel_to,
            vert_align=vert_align,
            horz_align=horz_align,
            vert_offset=vert_offset,
            horz_offset=horz_offset,
        )

    def set_out_margins(
        self,
        *,
        left: int | str | None = None,
        right: int | str | None = None,
        top: int | str | None = None,
        bottom: int | str | None = None,
    ) -> None:
        self._set_graphic_out_margins(left=left, right=right, top=top, bottom=bottom)

    def set_rotation(
        self,
        *,
        angle: int | str | None = None,
        center_x: int | str | None = None,
        center_y: int | str | None = None,
        rotate_image: bool | None = None,
    ) -> None:
        section_model, _paragraph, control_node = self._live_context()
        shape_component = self._shape_component_record(control_node)
        if shape_component is None:
            self._write_native_shape_metadata(
                updates=_graphic_rotation_metadata_updates(
                    angle=angle,
                    center_x=center_x,
                    center_y=center_y,
                    rotate_image=rotate_image,
                )
            )
            return
        shape_component.payload = _set_shape_component_rotation_payload(
            shape_component.payload,
            angle=angle,
            center_x=center_x,
            center_y=center_y,
        )
        _update_shape_native_metadata(
            control_node,
            updates={"rotation.rotateimage": rotate_image},
            remove_prefixes=("rotation.angle", "rotation.centerX", "rotation.centerY"),
        )
        self._commit(section_model)

    def set_size(self, *, width: int | None = None, height: int | None = None) -> None:
        section_model, _paragraph, control_node = self._live_context()
        control_node.payload = _set_payload_size(control_node.payload, width=width, height=height, width_offset=16, height_offset=20)
        record = self._shape_component_record(control_node)
        if record is None:
            raise ValueError("HWP shape control does not contain a shape component record.")
        record.payload = _set_payload_size(record.payload, width=width, height=height, width_offset=20, height_offset=24)
        self._commit(section_model)


class HwpLineShapeObject(HwpShapeObject):
    @property
    def kind(self) -> str:
        return "line"

    @property
    def start(self) -> dict[str, int]:
        return dict(self.specific_fields().get("start", {"x": 0, "y": 0}))

    @property
    def end(self) -> dict[str, int]:
        return dict(self.specific_fields().get("end", {"x": 0, "y": 0}))

    @property
    def attributes(self) -> int:
        return int(self.specific_fields().get("attributes", 0))

    def set_endpoints(
        self,
        *,
        start: dict[str, object] | Sequence[int] | None = None,
        end: dict[str, object] | Sequence[int] | None = None,
        attributes: int | None = None,
    ) -> None:
        payload = _build_line_specific_payload(
            start=start or self.start,
            end=end or self.end,
            attributes=self.attributes if attributes is None else int(attributes),
        )
        self._set_shape_specific_payload(payload)

    def set_start(self, value: dict[str, object] | Sequence[int]) -> None:
        self.set_endpoints(start=value)

    def set_end(self, value: dict[str, object] | Sequence[int]) -> None:
        self.set_endpoints(end=value)


class HwpConnectLineShapeObject(HwpLineShapeObject):
    @property
    def kind(self) -> str:
        return "connectLine"


class HwpRectangleShapeObject(HwpShapeObject):
    @property
    def kind(self) -> str:
        return "rect"

    @property
    def corner_radius(self) -> int:
        return int(self.specific_fields().get("corner_radius", 0))

    def set_corner_radius(self, value: int) -> None:
        current_specific = self._shape_specific_payload()
        trailing_payload = current_specific[4:] if len(current_specific) > 4 else b""
        self._set_shape_specific_payload(
            _build_rectangle_specific_payload(corner_radius=int(value), trailing_payload=trailing_payload)
        )


class HwpEllipseShapeObject(HwpShapeObject):
    @property
    def kind(self) -> str:
        return "ellipse"

    @property
    def arc_flags(self) -> int:
        return int(self.specific_fields().get("arc_flags", 0))

    @property
    def center(self) -> dict[str, int]:
        return dict(self.specific_fields().get("center", {"x": 0, "y": 0}))

    @property
    def axis1(self) -> dict[str, int]:
        return dict(self.specific_fields().get("axis1", {"x": 0, "y": 0}))

    @property
    def axis2(self) -> dict[str, int]:
        return dict(self.specific_fields().get("axis2", {"x": 0, "y": 0}))

    @property
    def start(self) -> dict[str, int]:
        return dict(self.specific_fields().get("start", {"x": 0, "y": 0}))

    @property
    def end(self) -> dict[str, int]:
        return dict(self.specific_fields().get("end", {"x": 0, "y": 0}))

    @property
    def start_2(self) -> dict[str, int]:
        return dict(self.specific_fields().get("start_2", {"x": 0, "y": 0}))

    @property
    def end_2(self) -> dict[str, int]:
        return dict(self.specific_fields().get("end_2", {"x": 0, "y": 0}))

    def set_geometry(
        self,
        *,
        arc_flags: int | None = None,
        center: dict[str, object] | Sequence[int] | None = None,
        axis1: dict[str, object] | Sequence[int] | None = None,
        axis2: dict[str, object] | Sequence[int] | None = None,
        start: dict[str, object] | Sequence[int] | None = None,
        end: dict[str, object] | Sequence[int] | None = None,
        start_2: dict[str, object] | Sequence[int] | None = None,
        end_2: dict[str, object] | Sequence[int] | None = None,
    ) -> None:
        self._set_shape_specific_payload(
            _build_ellipse_geometry_payload(
                arc_flags=self.arc_flags if arc_flags is None else int(arc_flags),
                center=center or self.center,
                axis1=axis1 or self.axis1,
                axis2=axis2 or self.axis2,
                start=start or self.start,
                end=end or self.end,
                start_2=start_2 or self.start_2,
                end_2=end_2 or self.end_2,
            )
        )

    def set_axes(
        self,
        *,
        axis1: dict[str, object] | Sequence[int] | None = None,
        axis2: dict[str, object] | Sequence[int] | None = None,
    ) -> None:
        self.set_geometry(axis1=axis1, axis2=axis2)


class HwpArcShapeObject(HwpShapeObject):
    @property
    def kind(self) -> str:
        return "arc"

    @property
    def arc_type(self) -> int:
        return int(self.specific_fields().get("arc_type", 0))

    @property
    def center(self) -> dict[str, int]:
        return dict(self.specific_fields().get("center", {"x": 0, "y": 0}))

    @property
    def axis1(self) -> dict[str, int]:
        return dict(self.specific_fields().get("axis1", {"x": 0, "y": 0}))

    @property
    def axis2(self) -> dict[str, int]:
        return dict(self.specific_fields().get("axis2", {"x": 0, "y": 0}))

    def set_geometry(
        self,
        *,
        arc_type: int | None = None,
        center: dict[str, object] | Sequence[int] | None = None,
        axis1: dict[str, object] | Sequence[int] | None = None,
        axis2: dict[str, object] | Sequence[int] | None = None,
    ) -> None:
        self._set_shape_specific_payload(
            _build_arc_geometry_payload(
                arc_type=self.arc_type if arc_type is None else int(arc_type),
                center=center or self.center,
                axis1=axis1 or self.axis1,
                axis2=axis2 or self.axis2,
            )
        )

    def set_axes(
        self,
        *,
        axis1: dict[str, object] | Sequence[int] | None = None,
        axis2: dict[str, object] | Sequence[int] | None = None,
    ) -> None:
        self.set_geometry(axis1=axis1, axis2=axis2)


class HwpPolygonShapeObject(HwpShapeObject):
    @property
    def kind(self) -> str:
        return "polygon"

    @property
    def point_count(self) -> int:
        return int(self.specific_fields().get("point_count", 0))

    @property
    def points(self) -> list[dict[str, int]]:
        points = self.specific_fields().get("points", [])
        return [dict(point) for point in points if isinstance(point, dict)]

    @property
    def trailing_payload(self) -> bytes:
        payload = self.specific_fields().get("trailing_payload", b"")
        return bytes(payload) if isinstance(payload, (bytes, bytearray)) else b""

    def set_points(
        self,
        points: Sequence[dict[str, object] | Sequence[int]],
        *,
        trailing_payload: bytes | None = None,
    ) -> None:
        self._set_shape_specific_payload(
            _build_polygon_points_payload(
                points,
                trailing_payload=self.trailing_payload if trailing_payload is None else trailing_payload,
            )
        )


class HwpCurveShapeObject(HwpShapeObject):
    @property
    def kind(self) -> str:
        return "curve"

    @property
    def raw_payload(self) -> bytes:
        payload = self.specific_fields().get("raw_payload", b"")
        return bytes(payload) if isinstance(payload, (bytes, bytearray)) else b""

    def set_raw_payload(self, payload: bytes) -> None:
        self._set_shape_specific_payload(bytes(payload))


class HwpContainerShapeObject(HwpShapeObject):
    @property
    def kind(self) -> str:
        return "container"

    @property
    def raw_payload(self) -> bytes:
        payload = self.specific_fields().get("raw_payload", b"")
        return bytes(payload) if isinstance(payload, (bytes, bytearray)) else b""

    def set_raw_payload(self, payload: bytes) -> None:
        self._set_shape_specific_payload(bytes(payload))


class HwpTextArtShapeObject(HwpShapeObject):
    @property
    def kind(self) -> str:
        return "textart"

    @property
    def native_text(self) -> str:
        return str(self.specific_fields().get("text", ""))

    @property
    def font_name(self) -> str:
        return str(self.specific_fields().get("font_name", ""))

    @property
    def font_style(self) -> str:
        return str(self.specific_fields().get("font_style", ""))

    def set_text(self, value: str) -> None:
        fields = self.specific_fields()
        self._set_shape_specific_payload(
            _build_textart_payload_fields(
                text=value,
                font_name=str(fields.get("font_name", "")),
                font_style=str(fields.get("font_style", "")),
                prefix_payload=fields.get("prefix_payload") if isinstance(fields.get("prefix_payload"), bytes) else None,
                trailing_payload=fields.get("trailing_payload") if isinstance(fields.get("trailing_payload"), bytes) else None,
            )
        )

    def set_style(self, *, font_name: str | None = None, font_style: str | None = None) -> None:
        fields = self.specific_fields()
        self._set_shape_specific_payload(
            _build_textart_payload_fields(
                text=str(fields.get("text", "")),
                font_name=str(font_name if font_name is not None else fields.get("font_name", "")),
                font_style=str(font_style if font_style is not None else fields.get("font_style", "")),
                prefix_payload=fields.get("prefix_payload") if isinstance(fields.get("prefix_payload"), bytes) else None,
                trailing_payload=fields.get("trailing_payload") if isinstance(fields.get("trailing_payload"), bytes) else None,
            )
        )


class HwpChartObject(HwpShapeObject):
    @property
    def kind(self) -> str:
        return "chart"

    def fields(self) -> dict[str, object]:
        result = _parse_chart_data_payload(self._chart_record().payload)
        metadata = self._shape_native_metadata()
        title = metadata.get("chart.title") or str(result.get("utf16_text", ""))
        if title:
            result["title"] = title
        if "chart.chartType" in metadata:
            result["chart_type"] = metadata["chart.chartType"]
        if "chart.dataRef" in metadata:
            result["data_ref"] = metadata["chart.dataRef"]
        if "chart.legendVisible" in metadata:
            result["legend_visible"] = _metadata_bool_value(metadata, "chart.legendVisible", True)
        categories = [str(item) for item in _metadata_json_list(metadata, "chart.categories")]
        if categories:
            result["categories"] = categories
        series = _metadata_json_dict_list(metadata, "chart.series")
        if series:
            result["series"] = series
        return result

    @property
    def raw_payload(self) -> bytes:
        payload = _parse_chart_data_payload(self._chart_record().payload).get("raw_payload", b"")
        return bytes(payload) if isinstance(payload, (bytes, bytearray)) else b""

    @property
    def utf16_text(self) -> str:
        return str(_parse_chart_data_payload(self._chart_record().payload).get("utf16_text", ""))

    @property
    def title(self) -> str:
        metadata = self._shape_native_metadata()
        return metadata.get("chart.title", self.utf16_text)

    @property
    def chart_type(self) -> str:
        metadata = self._shape_native_metadata()
        return metadata.get("chart.chartType", "BAR")

    @property
    def data_ref(self) -> str | None:
        metadata = self._shape_native_metadata()
        value = metadata.get("chart.dataRef")
        return value or None

    @property
    def legend_visible(self) -> bool:
        return _metadata_bool_value(self._shape_native_metadata(), "chart.legendVisible", True)

    @property
    def categories(self) -> list[str]:
        return [str(item) for item in _metadata_json_list(self._shape_native_metadata(), "chart.categories")]

    @property
    def series(self) -> list[dict[str, object]]:
        return _metadata_json_dict_list(self._shape_native_metadata(), "chart.series")

    def set_raw_payload(self, payload: bytes) -> None:
        section_model, _paragraph, control_node = self._live_context()
        self._chart_record(control_node).payload = bytes(payload)
        self._commit(section_model)

    def set_title(self, value: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        self._chart_record(control_node).payload = str(value).encode("utf-16-le")
        _update_shape_native_metadata(control_node, updates={"chart.title": value})
        self._commit(section_model)

    def set_chart_type(self, value: str) -> None:
        self._write_native_shape_metadata(updates={"chart.chartType": str(value).upper()})

    def set_data_ref(self, value: str | None) -> None:
        self._write_native_shape_metadata(updates={"chart.dataRef": value})

    def set_legend_visible(self, value: bool) -> None:
        self._write_native_shape_metadata(updates={"chart.legendVisible": bool(value)})

    def set_categories(self, values: Sequence[str]) -> None:
        serialized = json.dumps([str(value) for value in values], ensure_ascii=True, separators=(",", ":"))
        self._write_native_shape_metadata(updates={"chart.categories": serialized})

    def set_series(self, values: Sequence[dict[str, object]]) -> None:
        serialized = json.dumps([dict(value) for value in values], ensure_ascii=True, separators=(",", ":"))
        self._write_native_shape_metadata(updates={"chart.series": serialized})

    def _chart_record(self, control_node: RecordNode | None = None) -> RecordNode:
        target = self.control_node if control_node is None else control_node
        for node in target.iter_descendants():
            if node.tag_id == TAG_CHART_DATA:
                return node
        raise ValueError("HWP chart control does not contain chart data.")


class HwpOleObject(HwpShapeObject):
    @property
    def kind(self) -> str:
        return "ole"

    @property
    def storage_id(self) -> int:
        payload = self._ole_record().payload
        if len(payload) >= 14:
            return int.from_bytes(payload[12:14], "little", signed=False)
        return 0

    @property
    def object_type(self) -> str:
        flags = self._ole_flags()
        return _HWP_OLE_OBJECT_TYPE_NAMES.get((flags >> 16) & 0x3F, "UNKNOWN")

    @property
    def draw_aspect(self) -> str:
        flags = self._ole_flags()
        return _HWP_OLE_DRAW_ASPECT_NAMES.get(flags & 0xFF, "CONTENT")

    @property
    def has_moniker(self) -> bool:
        return bool(self._ole_flags() & (1 << 8))

    @property
    def eq_baseline(self) -> int:
        return (self._ole_flags() >> 9) & 0x7F

    @property
    def line_color(self) -> str:
        payload = self._ole_record().payload
        if len(payload) < 18:
            return "#000000"
        return _parse_rgb_color(payload[14:18])

    @property
    def line_width(self) -> int:
        payload = self._ole_record().payload
        if len(payload) < 20:
            return 0
        return int.from_bytes(payload[18:20], "little", signed=False)

    def line_style(self) -> dict[str, str]:
        metadata = _metadata_prefixed_str_map(self._shape_native_metadata(), "lineStyle.")
        result = {
            "color": self.line_color,
            "width": str(self.line_width),
        }
        result.update(metadata)
        return result

    def rotation(self) -> dict[str, str]:
        component = self._shape_component_record()
        native = {} if component is None else _parse_shape_component_rotation_payload(component.payload, include_center=False)
        metadata = _metadata_prefixed_str_map(self._shape_native_metadata(), "rotation.")
        if not native:
            return metadata
        merged = dict(metadata)
        merged.update(native)
        return merged

    def extent(self) -> dict[str, int]:
        payload = self._ole_record().payload
        if len(payload) >= 12:
            extent = {
                "x": int.from_bytes(payload[4:8], "little", signed=True),
                "y": int.from_bytes(payload[8:12], "little", signed=True),
            }
            if extent["x"] or extent["y"]:
                display_size = super().size()
                if extent != {"x": display_size.get("width", 0), "y": display_size.get("height", 0)}:
                    return extent
        return _metadata_prefixed_int_map(self._shape_native_metadata(), "extent.")

    def size(self) -> dict[str, int]:
        return super().size()

    def bindata_path(self) -> str:
        storage_id = self.storage_id
        if storage_id > 0:
            prefix = f"BinData/BIN{storage_id:04d}."
            for stream_path in self.document.bindata_stream_paths():
                if stream_path.startswith(prefix):
                    return stream_path
        bindata_path = next((stream_path for stream_path in self.document.bindata_stream_paths() if stream_path.endswith(".ole")), None)
        if bindata_path is None:
            raise ValueError("No OLE BinData stream is available.")
        return bindata_path

    def binary_data(self) -> bytes:
        return self.document.binary_document().read_stream(self.bindata_path(), decompress=False)

    def replace_binary(self, data: bytes, *, extension: str | None = None) -> None:
        self.document.binary_document().add_stream(self.bindata_path(), data)
        self.document._invalidate_bridge()

    def set_size(self, *, width: int | None = None, height: int | None = None) -> None:
        section_model, _paragraph, control_node = self._live_context()
        control_node.payload = _set_payload_size(control_node.payload, width=width, height=height, width_offset=16, height_offset=20)
        ole_record = self._ole_record(control_node)
        ole_record.payload = _set_payload_size(ole_record.payload, width=width, height=height, width_offset=4, height_offset=8)
        component = self._shape_component_record(control_node)
        if component is not None:
            component.payload = _set_payload_size(component.payload, width=width, height=height, width_offset=20, height_offset=24)
        self._commit(section_model)

    def set_object_metadata(
        self,
        *,
        object_type: str | None = None,
        draw_aspect: str | None = None,
        has_moniker: bool | None = None,
        eq_baseline: int | None = None,
    ) -> None:
        section_model, _paragraph, control_node = self._live_context()
        record = self._ole_record(control_node)
        payload = bytearray(record.payload)
        resolved_flags = _build_ole_attribute_flags(
            object_type=self.object_type if object_type is None else object_type,
            draw_aspect=self.draw_aspect if draw_aspect is None else draw_aspect,
            has_moniker=self.has_moniker if has_moniker is None else bool(has_moniker),
            eq_baseline=self.eq_baseline if eq_baseline is None else int(eq_baseline),
        )
        payload[0:4] = int(resolved_flags).to_bytes(4, "little", signed=False)
        record.payload = bytes(payload)
        self._commit(section_model)

    def set_line_style(
        self,
        *,
        color: str | None = None,
        width: int | None = None,
        style: str | None = None,
        end_cap: str | None = None,
        head_style: str | None = None,
        tail_style: str | None = None,
        head_fill: bool | None = None,
        tail_fill: bool | None = None,
        head_size: str | None = None,
        tail_size: str | None = None,
        outline_style: str | None = None,
        alpha: int | str | None = None,
    ) -> None:
        section_model, _paragraph, control_node = self._live_context()
        record = self._ole_record(control_node)
        payload = bytearray(record.payload)
        if color is not None:
            payload[14:18] = _build_rgb_color(color)
        if width is not None:
            payload[18:20] = max(0, min(int(width), 0xFFFF)).to_bytes(2, "little", signed=False)
        record.payload = bytes(payload)
        component = self._shape_component_record(control_node)
        if component is not None and len(component.payload) >= 202:
            component_payload = bytearray(component.payload)
            if color is not None:
                component_payload[196:200] = _build_colorref(color)
            if width is not None:
                component_payload[200:202] = max(0, min(int(width), 0xFFFF)).to_bytes(2, "little", signed=False)
            component.payload = bytes(component_payload)
        _update_shape_native_metadata(
            control_node,
            updates={
                "lineStyle.style": style,
                "lineStyle.endCap": end_cap,
                "lineStyle.headStyle": head_style,
                "lineStyle.tailStyle": tail_style,
                "lineStyle.headfill": head_fill,
                "lineStyle.tailfill": tail_fill,
                "lineStyle.headSz": head_size,
                "lineStyle.tailSz": tail_size,
                "lineStyle.outlineStyle": outline_style,
                "lineStyle.alpha": alpha,
            },
        )
        self._commit(section_model)

    def set_rotation(
        self,
        *,
        angle: int | str | None = None,
        center_x: int | str | None = None,
        center_y: int | str | None = None,
        rotate_image: bool | None = None,
    ) -> None:
        section_model, _paragraph, control_node = self._live_context()
        component = self._shape_component_record(control_node)
        if component is None:
            raise ValueError("HWP OLE control does not contain a shape component record.")
        component.payload = _set_shape_component_rotation_payload(component.payload, angle=angle)
        _update_shape_native_metadata(
            control_node,
            updates={
                "rotation.centerX": center_x,
                "rotation.centerY": center_y,
                "rotation.rotateimage": rotate_image,
            },
            remove_prefixes=("rotation.angle",),
        )
        self._commit(section_model)

    def set_extent(self, *, x: int | str | None = None, y: int | str | None = None) -> None:
        section_model, _paragraph, control_node = self._live_context()
        record = self._ole_record(control_node)
        payload = bytearray(record.payload)
        if len(payload) < 12:
            payload.extend(b"\x00" * (12 - len(payload)))
        current_x = int.from_bytes(payload[4:8], "little", signed=True)
        current_y = int.from_bytes(payload[8:12], "little", signed=True)
        payload[4:8] = int(current_x if x is None else int(x)).to_bytes(4, "little", signed=True)
        payload[8:12] = int(current_y if y is None else int(y)).to_bytes(4, "little", signed=True)
        record.payload = bytes(payload)
        _update_shape_native_metadata(control_node, updates={}, remove_prefixes=("extent.",))
        self._commit(section_model)

    def _ole_record(self, control_node: RecordNode | None = None) -> RecordNode:
        target = self.control_node if control_node is None else control_node
        for node in target.iter_descendants():
            if node.tag_id == TAG_SHAPE_COMPONENT_OLE:
                return node
        raise ValueError("HWP OLE control does not contain an OLE component record.")

    def _ole_flags(self) -> int:
        payload = self._ole_record().payload
        if len(payload) < 4:
            return 0
        return int.from_bytes(payload[0:4], "little", signed=False)


class HwpFormObject(HwpControlObject):
    @property
    def kind(self) -> str:
        return "form"

    def fields(self) -> dict[str, object]:
        result = _parse_form_object_payload(self._form_record().payload)
        metadata = self._native_control_metadata()
        label = metadata.get("form.label", str(result.get("utf16_text", "")))
        if label:
            result["label"] = label
        result["form_type"] = metadata.get("form.type", "INPUT")
        if metadata.get("form.name"):
            result["name"] = metadata["form.name"]
        result["value"] = metadata.get("form.value", label)
        result["checked"] = _metadata_bool_value(metadata, "form.checked", False)
        result["items"] = [str(item) for item in _metadata_json_list(metadata, "form.items")]
        result["editable"] = _metadata_bool_value(metadata, "form.editable", True)
        result["locked"] = _metadata_bool_value(metadata, "form.locked", False)
        if metadata.get("form.placeholder"):
            result["placeholder"] = metadata["form.placeholder"]
        return result

    @property
    def raw_payload(self) -> bytes:
        payload = _parse_form_object_payload(self._form_record().payload).get("raw_payload", b"")
        return bytes(payload) if isinstance(payload, (bytes, bytearray)) else b""

    @property
    def utf16_text(self) -> str:
        return str(_parse_form_object_payload(self._form_record().payload).get("utf16_text", ""))

    @property
    def label(self) -> str:
        metadata = self._native_control_metadata()
        return metadata.get("form.label", self.utf16_text)

    @property
    def text(self) -> str:
        return self.label

    @property
    def form_type(self) -> str:
        metadata = self._native_control_metadata()
        return metadata.get("form.type", "INPUT")

    @property
    def name(self) -> str | None:
        metadata = self._native_control_metadata()
        value = metadata.get("form.name")
        return value or None

    @property
    def value(self) -> str:
        metadata = self._native_control_metadata()
        return metadata.get("form.value", self.label)

    @property
    def checked(self) -> bool:
        return _metadata_bool_value(self._native_control_metadata(), "form.checked", False)

    @property
    def items(self) -> list[str]:
        return [str(item) for item in _metadata_json_list(self._native_control_metadata(), "form.items")]

    @property
    def editable(self) -> bool:
        return _metadata_bool_value(self._native_control_metadata(), "form.editable", True)

    @property
    def locked(self) -> bool:
        return _metadata_bool_value(self._native_control_metadata(), "form.locked", False)

    @property
    def placeholder(self) -> str | None:
        metadata = self._native_control_metadata()
        value = metadata.get("form.placeholder")
        return value or None

    def set_raw_payload(self, payload: bytes) -> None:
        section_model, _paragraph, control_node = self._live_context()
        self._form_record(control_node).payload = bytes(payload)
        self._commit(section_model)

    def set_label(self, value: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        self._form_record(control_node).payload = str(value).encode("utf-16-le")
        _update_control_native_metadata(control_node, updates={"form.label": value})
        self._commit(section_model)

    def set_text(self, value: str) -> None:
        self.set_label(value)

    def set_form_type(self, value: str) -> None:
        self._write_native_control_metadata(updates={"form.type": str(value).upper()})

    def set_name(self, value: str | None) -> None:
        self._write_native_control_metadata(updates={"form.name": value})

    def set_value(self, value: str | None) -> None:
        self._write_native_control_metadata(updates={"form.value": value})

    def set_checked(self, value: bool) -> None:
        self._write_native_control_metadata(updates={"form.checked": bool(value)})

    def set_items(self, values: Sequence[str]) -> None:
        serialized = json.dumps([str(value) for value in values], ensure_ascii=True, separators=(",", ":"))
        self._write_native_control_metadata(updates={"form.items": serialized})

    def set_editable(self, value: bool) -> None:
        self._write_native_control_metadata(updates={"form.editable": bool(value)})

    def set_locked(self, value: bool) -> None:
        self._write_native_control_metadata(updates={"form.locked": bool(value)})

    def set_placeholder(self, value: str | None) -> None:
        self._write_native_control_metadata(updates={"form.placeholder": value})

    def _form_record(self, control_node: RecordNode | None = None) -> RecordNode:
        target = self.control_node if control_node is None else control_node
        for node in target.iter_descendants():
            if node.tag_id == TAG_FORM_OBJECT:
                return node
        raise ValueError("HWP form control does not contain a form object record.")


class HwpMemoObject(HwpControlObject):
    @property
    def kind(self) -> str:
        return "memo"

    def fields(self) -> dict[str, object]:
        result = _parse_memo_list_payload(self._memo_record().payload)
        metadata = self._native_control_metadata()
        memo_text = str(result.get("utf16_text", "")).strip()
        result["text"] = memo_text or _collect_text_from_node(self.control_node).replace("\r", "")
        if metadata.get("memo.author"):
            result["author"] = metadata["memo.author"]
        if metadata.get("memo.memoId"):
            result["memo_id"] = metadata["memo.memoId"]
        if metadata.get("memo.anchorId"):
            result["anchor_id"] = metadata["memo.anchorId"]
        if metadata.get("memo.order"):
            try:
                result["order"] = int(metadata["memo.order"])
            except ValueError:
                pass
        result["visible"] = _metadata_bool_value(metadata, "memo.visible", True)
        return result

    @property
    def raw_payload(self) -> bytes:
        payload = _parse_memo_list_payload(self._memo_record().payload).get("raw_payload", b"")
        return bytes(payload) if isinstance(payload, (bytes, bytearray)) else b""

    @property
    def utf16_text(self) -> str:
        return str(_parse_memo_list_payload(self._memo_record().payload).get("utf16_text", ""))

    @property
    def text(self) -> str:
        memo_text = self.utf16_text.strip()
        return memo_text or _collect_text_from_node(self.control_node).replace("\r", "")

    @property
    def author(self) -> str | None:
        metadata = self._native_control_metadata()
        value = metadata.get("memo.author")
        return value or None

    @property
    def memo_id(self) -> str | None:
        metadata = self._native_control_metadata()
        value = metadata.get("memo.memoId")
        return value or None

    @property
    def anchor_id(self) -> str | None:
        metadata = self._native_control_metadata()
        value = metadata.get("memo.anchorId")
        return value or None

    @property
    def order(self) -> int | None:
        metadata = self._native_control_metadata()
        value = metadata.get("memo.order")
        if value is None:
            return None
        try:
            return int(value)
        except ValueError:
            return None

    @property
    def visible(self) -> bool:
        return _metadata_bool_value(self._native_control_metadata(), "memo.visible", True)

    def set_raw_payload(self, payload: bytes) -> None:
        section_model, _paragraph, control_node = self._live_context()
        self._memo_record(control_node).payload = bytes(payload)
        self._commit(section_model)

    def set_text(self, value: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        self._memo_record(control_node).payload = str(value).encode("utf-16-le")
        self._commit(section_model)

    def set_author(self, value: str | None) -> None:
        self._write_native_control_metadata(updates={"memo.author": value})

    def set_memo_id(self, value: str | None) -> None:
        self._write_native_control_metadata(updates={"memo.memoId": value})

    def set_anchor_id(self, value: str | None) -> None:
        self._write_native_control_metadata(updates={"memo.anchorId": value})

    def set_order(self, value: int | None) -> None:
        self._write_native_control_metadata(updates={"memo.order": value})

    def set_visible(self, value: bool) -> None:
        self._write_native_control_metadata(updates={"memo.visible": bool(value)})

    def _memo_record(self, control_node: RecordNode | None = None) -> RecordNode:
        target = self.control_node if control_node is None else control_node
        for node in target.iter_descendants():
            if node.tag_id == TAG_MEMO_LIST:
                return node
        raise ValueError("HWP memo control does not contain a memo list record.")


def _build_control_wrapper(
    document: HwpDocument,
    section_model: SectionModel,
    paragraph: SectionParagraphModel,
    control_node: RecordNode,
    control_ordinal: int,
) -> HwpControlObject | None:
    control_id = _control_id(control_node)
    if _control_has_descendant(control_node, TAG_FORM_OBJECT):
        return HwpFormObject(document, paragraph, control_node, control_ordinal)
    if _control_has_descendant(control_node, TAG_MEMO_LIST):
        return HwpMemoObject(document, paragraph, control_node, control_ordinal)
    if control_id == "tbl ":
        return HwpTableObject(document, paragraph, control_node, control_ordinal)
    if control_id == "bokm":
        return HwpBookmarkObject(document, paragraph, control_node, control_ordinal)
    if control_id in {"fn  ", "en  "}:
        return HwpNoteObject(document, paragraph, control_node, control_ordinal)
    if control_id == "%hlk":
        return HwpHyperlinkObject(document, paragraph, control_node, control_ordinal)
    if control_id.startswith("%"):
        return HwpFieldObject(document, paragraph, control_node, control_ordinal)
    if control_id in {"atno", "nwno"}:
        return HwpAutoNumberObject(document, paragraph, control_node, control_ordinal)
    if control_id in {"head", "foot"}:
        return HwpHeaderFooterObject(document, paragraph, control_node, control_ordinal)
    if control_id == "pgnp":
        return HwpPageNumObject(document, paragraph, control_node, control_ordinal)
    if control_id == "eqed":
        return HwpEquationObject(document, paragraph, control_node, control_ordinal)
    if control_id == "gso " and _control_has_descendant(control_node, TAG_SHAPE_COMPONENT_OLE):
        return HwpOleObject(document, paragraph, control_node, control_ordinal)
    if control_id == "gso " and any(node.tag_id == TAG_SHAPE_COMPONENT_PICTURE for node in control_node.iter_descendants()):
        return HwpPictureObject(document, paragraph, control_node, control_ordinal)
    if control_id == "gso " and _control_has_descendant(control_node, TAG_CHART_DATA):
        return HwpChartObject(document, paragraph, control_node, control_ordinal)
    if control_id == "gso ":
        wrapper_class = _shape_wrapper_class(control_node)
        return wrapper_class(document, paragraph, control_node, control_ordinal)
    return None


def _control_id(control_node: RecordNode) -> str:
    payload = control_node.payload
    if len(payload) < 4:
        return ""
    return payload[:4][::-1].decode("latin1", errors="replace")


def _collect_text_from_node(node: RecordNode | None) -> str:
    if node is None:
        return ""
    texts: list[str] = []
    for descendant in node.iter_descendants():
        if isinstance(descendant, ParagraphTextRecord):
            texts.append(descendant.text.replace("\r", ""))
    return "".join(texts)


def _control_has_descendant(control_node: RecordNode, tag_id: int) -> bool:
    return any(node.tag_id == tag_id for node in control_node.iter_descendants())


def _semantic_field_type_from_native(native_field_type: str, parameters: dict[str, str]) -> str:
    normalized = native_field_type.strip().upper()
    normalized_parameters = {str(key).upper(): str(value) for key, value in parameters.items()}
    if normalized in {"%DOC", "DOCPROPERTY"}:
        return "DOCPROPERTY"
    if normalized in {"%DAT", "DATE"}:
        return "DATE"
    if normalized in {"%MAI", "MAILMERGE", "MAIL_MERGE", "MERGEFIELD"}:
        return "MAILMERGE"
    if normalized in {"%FOR", "%CAL", "FORMULA", "CALCULATE", "CALC"}:
        return "FORMULA"
    if normalized in {"%REF", "%PAG", "%CRO", "REF", "PAGEREF", "BOOKMARKREF", "CROSSREF", "CROSS_REF"}:
        return "CROSSREF"
    if "MERGEFIELD" in normalized_parameters or ("FIELDNAME" in normalized_parameters and normalized == "%MAI"):
        return "MAILMERGE"
    if "EXPRESSION" in normalized_parameters or (normalized == "%FOR" and "COMMAND" in normalized_parameters):
        return "FORMULA"
    if "BOOKMARKNAME" in normalized_parameters:
        return "CROSSREF"
    return native_field_type


def _shape_kind_from_control_node(control_node: RecordNode) -> str:
    for data_node in control_node.children:
        if data_node.tag_id != TAG_CTRL_DATA:
            continue
        metadata = _parse_shape_native_metadata(data_node.payload)
        kind = metadata.get("kind")
        if kind:
            return kind
    if _has_hancom_connectline_signature(control_node):
        return "connectLine"
    tag_to_kind = {
        TAG_SHAPE_COMPONENT_LINE: "line",
        TAG_SHAPE_COMPONENT_RECTANGLE: "rect",
        TAG_SHAPE_COMPONENT_ELLIPSE: "ellipse",
        TAG_SHAPE_COMPONENT_ARC: "arc",
        TAG_SHAPE_COMPONENT_POLYGON: "polygon",
        TAG_SHAPE_COMPONENT_CURVE: "curve",
        TAG_SHAPE_COMPONENT_OLE: "ole",
        TAG_SHAPE_COMPONENT_PICTURE: "picture",
        TAG_SHAPE_COMPONENT_CONTAINER: "container",
        TAG_SHAPE_COMPONENT_TEXTART: "textart",
    }
    for node in control_node.iter_descendants():
        kind = tag_to_kind.get(node.tag_id)
        if kind is not None:
            return kind
    return "shape"


def _shape_wrapper_class(control_node: RecordNode) -> type[HwpShapeObject]:
    shape_kind = _shape_kind_from_control_node(control_node)
    return {
        "line": HwpLineShapeObject,
        "connectLine": HwpConnectLineShapeObject,
        "rect": HwpRectangleShapeObject,
        "ellipse": HwpEllipseShapeObject,
        "arc": HwpArcShapeObject,
        "polygon": HwpPolygonShapeObject,
        "curve": HwpCurveShapeObject,
        "container": HwpContainerShapeObject,
        "textart": HwpTextArtShapeObject,
    }.get(shape_kind, HwpShapeObject)
