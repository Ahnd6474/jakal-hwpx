"""Native HWP control identifiers, fields, markers, and semantic metadata codecs."""
from __future__ import annotations

import struct
from collections.abc import (
    Sequence,
)

from ._hwp_constants import (
    _AUTO_NUMBER_TYPE_CODES,
    _BOOKMARK_CTRL_DATA_PREFIX,
    _CONTROL_MARKER_CODES,
    _CONTROL_NATIVE_METADATA_PREFIX,
    _ENDNOTE_CTRL_HEADER,
    _EQUATION_VERSION_MARKER,
    _FOOTNOTE_CTRL_HEADER,
    _HANCOM_CONNECTLINE_COMPONENT_SIGNATURE,
    _HYPERLINK_CTRL_PREFIX,
    _HYPERLINK_CTRL_SUFFIX,
    _HYPERLINK_TEXT_END_MARKER,
    _HYPERLINK_TEXT_START_MARKER,
    _SHAPE_NATIVE_METADATA_PREFIX,
    DEFAULT_HWP_EQUATION_FONT,
    DEFAULT_HWP_HYPERLINK_METADATA_FIELDS,
    TAG_CTRL_DATA,
    TAG_SHAPE_COMPONENT,
    TAG_SHAPE_COMPONENT_ARC,
    TAG_SHAPE_COMPONENT_CONTAINER,
    TAG_SHAPE_COMPONENT_CURVE,
    TAG_SHAPE_COMPONENT_ELLIPSE,
    TAG_SHAPE_COMPONENT_LINE,
    TAG_SHAPE_COMPONENT_OLE,
    TAG_SHAPE_COMPONENT_PICTURE,
    TAG_SHAPE_COMPONENT_POLYGON,
    TAG_SHAPE_COMPONENT_RECTANGLE,
    TAG_SHAPE_COMPONENT_TEXTART,
)
from ._hwp_graphic_payloads import (
    _build_control_native_metadata_payload,
)
from ._hwp_page_payloads import (
    _HEADER_FOOTER_APPLY_PAGE_TYPE_CODES,
    _HEADER_FOOTER_APPLY_PAGE_TYPE_VALUES,
)
from ._hwp_record_base import (
    RecordNode,
)


def _build_control_id_payload(control_id: str) -> bytes:
    normalized = control_id[:4].ljust(4)
    return normalized.encode("latin1", errors="replace")[::-1]


def _parse_header_footer_apply_page_type(payload: bytes) -> str:
    if len(payload) < 8:
        return "BOTH"
    code = int.from_bytes(payload[4:8], "little", signed=False)
    return _HEADER_FOOTER_APPLY_PAGE_TYPE_VALUES.get(code, "BOTH")


def _build_header_footer_control_payload(
    control_id: str,
    *,
    apply_page_type: str = "BOTH",
    instance_id: int = 1,
) -> bytes:
    normalized_control_id = control_id[:4].ljust(4)
    apply_code = _HEADER_FOOTER_APPLY_PAGE_TYPE_CODES.get(str(apply_page_type).upper(), 0)
    if normalized_control_id not in {"head", "foot"}:
        raise ValueError(f"Unsupported header/footer control id: {control_id}")
    if instance_id < 1:
        raise ValueError("instance_id must be at least 1.")
    return (
        _build_control_id_payload(normalized_control_id)
        + apply_code.to_bytes(4, "little", signed=False)
        + instance_id.to_bytes(4, "little", signed=False)
    )


def _normalize_field_control_id(field_type: str) -> str:
    stripped = field_type.strip()
    if stripped.startswith("%"):
        return stripped[:4].ljust(4)
    alnum = "".join(character.lower() for character in stripped if character.isalnum())
    if not alnum:
        return "%fld"
    return ("%" + alnum[:3]).ljust(4)


def _semantic_field_type_from_native(native_field_type: str, parameters: dict[str, str] | None = None) -> str:
    normalized = native_field_type.strip().upper()
    normalized_parameters = {str(key).upper(): str(value) for key, value in (parameters or {}).items()}
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


def _build_equation_payload(script: str, *, font: str = DEFAULT_HWP_EQUATION_FONT) -> bytes:
    body = f"* {script}`{_EQUATION_VERSION_MARKER}\x07{font}"
    return b"\x00\x00\x00\x00" + body.encode("utf-16-le")


def _decode_control_id_payload(payload: bytes) -> str:
    if len(payload) < 4:
        return ""
    return payload[:4][::-1].decode("latin1", errors="replace")


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


def _parse_bookmark_ctrl_data_payload(payload: bytes) -> str:
    if len(payload) < 12:
        return ""
    name_size = int.from_bytes(payload[10:12], "little", signed=False)
    encoded_name = payload[12 : 12 + name_size * 2]
    return encoded_name.decode("utf-16-le", errors="ignore")


def _parse_hyperlink_command(payload: bytes) -> str:
    prefix_size = 9
    if len(payload) < prefix_size + 2:
        return ""
    text_size = int.from_bytes(payload[prefix_size : prefix_size + 2], "little", signed=False)
    start = prefix_size + 2
    end = start + text_size * 2
    return payload[start:end].decode("utf-16-le", errors="ignore")


def _control_has_descendant(control_node: "RecordNode", tag_id: int) -> bool:
    return any(node.tag_id == tag_id for node in control_node.iter_descendants())


def _shape_kind_from_control_node(control_node: "RecordNode") -> str:
    for child in control_node.children:
        if child.tag_id != TAG_CTRL_DATA:
            continue
        metadata = _parse_shape_native_metadata(child.payload)
        kind = metadata.get("kind")
        if kind:
            return kind
    component = next((node for node in control_node.iter_descendants() if node.tag_id == TAG_SHAPE_COMPONENT), None)
    specific = next((node for node in control_node.iter_descendants() if node.tag_id == TAG_SHAPE_COMPONENT_LINE), None)
    if component is not None and specific is not None and component.payload.startswith(_HANCOM_CONNECTLINE_COMPONENT_SIGNATURE):
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


def _parse_object_description(payload: bytes) -> str:
    if len(payload) < 46:
        return ""
    char_count = int.from_bytes(payload[44:46], "little", signed=False)
    encoded = payload[46 : 46 + char_count * 2]
    return encoded.decode("utf-16-le", errors="ignore")


def _parse_shape_native_metadata(payload: bytes) -> dict[str, str]:
    if not payload:
        return {}
    text = payload.decode("utf-16-le", errors="ignore")
    if not text.startswith(_SHAPE_NATIVE_METADATA_PREFIX):
        return {}
    metadata: dict[str, str] = {}
    for item in text.split(";")[1:]:
        key, separator, value = item.partition("=")
        if separator and key:
            metadata[key] = value
    return metadata


def _parse_control_native_metadata(payload: bytes) -> dict[str, str]:
    if not payload:
        return {}
    text = payload.decode("utf-16-le", errors="ignore")
    if not text.startswith(_CONTROL_NATIVE_METADATA_PREFIX):
        return {}
    metadata: dict[str, str] = {}
    for item in text.split(";")[1:]:
        key, separator, value = item.partition("=")
        if separator and key:
            metadata[key] = value
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


def _control_native_metadata(control_node: "RecordNode") -> dict[str, str]:
    metadata: dict[str, str] = {}
    for child in control_node.children:
        if child.tag_id != TAG_CTRL_DATA:
            continue
        metadata.update(_parse_control_native_metadata(child.payload))
    return metadata


def _update_control_native_metadata(
    control_node: "RecordNode",
    *,
    updates: dict[str, object | None],
    remove_prefixes: tuple[str, ...] = (),
) -> None:
    metadata = _control_native_metadata(control_node)
    for prefix in remove_prefixes:
        for key in [name for name in metadata if name.startswith(prefix)]:
            metadata.pop(key, None)
    for key, value in updates.items():
        if value is None:
            metadata.pop(key, None)
        elif isinstance(value, bool):
            metadata[key] = "1" if value else "0"
        else:
            metadata[key] = str(value)
    metadata_nodes = [
        child
        for child in control_node.children
        if child.tag_id == TAG_CTRL_DATA and _parse_control_native_metadata(child.payload)
    ]
    if not metadata:
        for node in metadata_nodes:
            control_node.children.remove(node)
        return
    payload = _build_control_native_metadata_payload(**metadata)
    if metadata_nodes:
        metadata_nodes[0].payload = payload
        for node in metadata_nodes[1:]:
            control_node.children.remove(node)
        return
    control_node.add_child(RecordNode(tag_id=TAG_CTRL_DATA, level=control_node.level + 1, payload=payload))


def _build_hyperlink_para_text_payload(display_text: str) -> bytes:
    return _HYPERLINK_TEXT_START_MARKER + display_text.encode("utf-16-le") + _HYPERLINK_TEXT_END_MARKER + "\r".encode("utf-16-le")


def _build_hyperlink_command(url: str, metadata_fields: Sequence[str | int] | None = None) -> str:
    fields = DEFAULT_HWP_HYPERLINK_METADATA_FIELDS if metadata_fields is None else tuple(str(value) for value in metadata_fields)
    return url.replace(":", "\\:") + ";" + ";".join(fields) + ";"


def _build_hyperlink_ctrl_payload(url: str, metadata_fields: Sequence[str | int] | None = None) -> bytes:
    normalized = _build_hyperlink_command(url, metadata_fields=metadata_fields)
    return _HYPERLINK_CTRL_PREFIX + len(normalized).to_bytes(2, "little") + normalized.encode("utf-16-le") + _HYPERLINK_CTRL_SUFFIX
