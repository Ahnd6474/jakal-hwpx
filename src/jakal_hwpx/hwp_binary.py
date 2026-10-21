from __future__ import annotations

import json
import struct
import zlib
from collections.abc import Sequence
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterator

import olefile

from ._cfb_writer import CfbLayout, capture_compound_file_layout, write_compound_file
from ._hwp_constants import (  # noqa: F401 - compatibility re-exports
    _AUTO_NUMBER_TYPE_CODES,
    _AUTO_NUMBER_TYPE_VALUES,
    _BLANK_DROP_CONTROL_IDS,
    _BOOKMARK_CTRL_DATA_PREFIX,
    _CONTROL_MARKER_CODES,
    _CONTROL_NATIVE_METADATA_PREFIX,
    _DEFAULT_ARC_CTRL_HEADER,
    _DEFAULT_ARC_PAYLOAD,
    _DEFAULT_ARC_SHAPE_COMPONENT,
    _DEFAULT_COMPLEX_SHAPE_COMPONENT_LINE_WIDTH,
    _DEFAULT_ELLIPSE_CTRL_HEADER,
    _DEFAULT_ELLIPSE_PAYLOAD,
    _DEFAULT_ELLIPSE_SHAPE_COMPONENT,
    _DEFAULT_EQUATION_CTRL_HEADER,
    _DEFAULT_OLE_LINE_WIDTH,
    _DEFAULT_POLYGON_CTRL_HEADER,
    _DEFAULT_POLYGON_PAYLOAD,
    _DEFAULT_POLYGON_SHAPE_COMPONENT,
    _DEFAULT_RECT_CTRL_HEADER,
    _DEFAULT_RECT_SHAPE_COMPONENT,
    _DEFAULT_SHAPE_COMPONENT,
    _DEFAULT_SHAPE_CTRL_HEADER,
    _DEFAULT_SHAPE_LINE_PAYLOAD,
    _DEFAULT_SHAPE_LINE_WIDTH,
    _DEFAULT_SHAPE_LIST_HEADER,
    _DEFAULT_SHAPE_RECTANGLE_PAYLOAD,
    _DEFAULT_SHAPE_TEXT_CHAR_SHAPE,
    _DEFAULT_SHAPE_TEXT_LINE_SEG,
    _DEFAULT_SHAPE_TEXT_PARA_CONTROL_MASK,
    _DEFAULT_SHAPE_TEXT_PARA_SHAPE_ID,
    _DEFAULT_SHAPE_TEXT_PARA_TRAILING,
    _DEFAULT_SHAPE_TEXT_SPLIT_FLAGS,
    _DEFAULT_SHAPE_TEXT_STYLE_ID,
    _DEFAULT_TEXTART_CTRL_HEADER,
    _DEFAULT_TEXTART_FONT_NAME,
    _DEFAULT_TEXTART_FONT_STYLE,
    _DEFAULT_TEXTART_SHAPE_COMPONENT,
    _ENDNOTE_CTRL_HEADER,
    _EQUATION_VERSION_MARKER,
    _FOOTER_CTRL_HEADER,
    _FOOTER_LIST_HEADER,
    _FOOTER_PARA_CHAR_SHAPE,
    _FOOTER_PARA_HEADER,
    _FOOTER_PARA_LINE_SEG,
    _FOOTNOTE_CTRL_HEADER,
    _HANCOM_CONNECTLINE_COMPONENT_SIGNATURE,
    _HEADER_CTRL_HEADER,
    _HEADER_LIST_HEADER,
    _HEADER_PARA_CHAR_SHAPE,
    _HEADER_PARA_HEADER,
    _HEADER_PARA_LINE_SEG,
    _HWP_CHAR_FONT_ORDER,
    _HWP_COMPLEX_SHAPE_TEMPLATES,
    _HWP_DRAW_ASPECT_CODES,
    _HWP_DRAW_ASPECT_NAMES,
    _HWP_FILLED_SHAPE_COMMON_PAYLOAD_SIZE,
    _HWP_LINE_ARROW_CODES,
    _HWP_LINE_ARROW_NAMES,
    _HWP_LINE_ARROW_SIZE_CODES,
    _HWP_LINE_ARROW_SIZE_NAMES,
    _HWP_LINE_END_CAP_CODES,
    _HWP_LINE_END_CAP_NAMES,
    _HWP_LINE_OUTLINE_STYLE_CODES,
    _HWP_LINE_OUTLINE_STYLE_NAMES,
    _HWP_LINE_STYLE_CODES,
    _HWP_LINE_STYLE_NAMES,
    _HWP_OLE_OBJECT_TYPE_CODES,
    _HWP_OLE_OBJECT_TYPE_NAMES,
    _HWP_PICTURE_EFFECT_CODES,
    _HWP_PICTURE_EFFECT_NAMES,
    _HWP_TAB_LEADER_NAMES,
    _HWP_TAB_TYPE_NAMES,
    _HYPERLINK_CHAR_SHAPE,
    _HYPERLINK_CTRL_PREFIX,
    _HYPERLINK_CTRL_SUFFIX,
    _HYPERLINK_LINE_SEG,
    _HYPERLINK_PARA_HEADER_TAIL,
    _HYPERLINK_TEXT_END_MARKER,
    _HYPERLINK_TEXT_START_MARKER,
    _MEMO_SHAPE_NATIVE_METADATA_PREFIX,
    _NOTE_CHAR_SHAPE,
    _NOTE_LINE_SEG,
    _NOTE_LIST_HEADER,
    _NOTE_PARA_HEADER,
    _PICTURE_CTRL_HEADER,
    _PICTURE_SHAPE_COMPONENT,
    _PICTURE_SHAPE_PICTURE,
    _PICTURE_TOP_CHAR_SHAPE,
    _PICTURE_TOP_LINE_SEG,
    _PICTURE_TOP_PARA_HEADER_TAIL,
    _PICTURE_TOP_PARA_TEXT,
    _SHAPE_NATIVE_METADATA_PREFIX,
    _TABLE_CELL_CHAR_SHAPE,
    _TABLE_CELL_LINE_SEG,
    _TABLE_CELL_PARA_HEADER_TAIL,
    _TABLE_CTRL_HEADER,
    _TABLE_TOP_CHAR_SHAPE,
    _TABLE_TOP_LINE_SEG,
    _TABLE_TOP_PARA_HEADER_TAIL,
    _TABLE_TOP_PARA_TEXT,
    _TEXTART_COMPONENT_REFERENCE_EXTENT,
    DEFAULT_HWP_EQUATION_FONT,
    DEFAULT_HWP_EQUATION_SCRIPT,
    DEFAULT_HWP_HYPERLINK_METADATA_FIELDS,
    DEFAULT_HWP_HYPERLINK_TEXT,
    DEFAULT_HWP_HYPERLINK_URL,
    DEFAULT_HWP_SHAPE_HEIGHT,
    DEFAULT_HWP_SHAPE_TEXT,
    DEFAULT_HWP_SHAPE_WIDTH,
    DEFAULT_HWP_TABLE_BORDER_FILL_ID,
    DEFAULT_HWP_TABLE_CELL_HEIGHT,
    DEFAULT_HWP_TABLE_CELL_MARGIN,
    DEFAULT_HWP_TABLE_CELL_WIDTH,
    DEFAULT_HWP_TABLE_COLS,
    DEFAULT_HWP_TABLE_ROWS,
    FILE_HEADER_SIGNATURE,
    FILE_HEADER_SIZE,
    FLAG_COMPRESSED,
    FLAG_DISTRIBUTABLE,
    FLAG_PASSWORD_PROTECTED,
    TAG_BIN_DATA,
    TAG_BORDER_FILL,
    TAG_BULLET,
    TAG_CHAR_SHAPE,
    TAG_CHART_DATA,
    TAG_COMPATIBLE_DOCUMENT,
    TAG_CTRL_DATA,
    TAG_CTRL_HEADER,
    TAG_DOC_DATA,
    TAG_DOCUMENT_PROPERTIES,
    TAG_EQEDIT,
    TAG_FACE_NAME,
    TAG_FOOTNOTE_SHAPE,
    TAG_FORM_OBJECT,
    TAG_ID_MAPPINGS,
    TAG_LAYOUT_COMPATIBILITY,
    TAG_LIST_HEADER,
    TAG_MEMO_LIST,
    TAG_MEMO_SHAPE,
    TAG_NAMES,
    TAG_NUMBERING,
    TAG_PAGE_BORDER_FILL,
    TAG_PAGE_DEF,
    TAG_PARA_CHAR_SHAPE,
    TAG_PARA_HEADER,
    TAG_PARA_LINE_SEG,
    TAG_PARA_RANGE_TAG,
    TAG_PARA_SHAPE,
    TAG_PARA_TEXT,
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
    TAG_STYLE,
    TAG_TAB_DEF,
    TAG_TABLE,
    TAG_TRACK_CHANGE,
    TAG_TRACK_CHANGE_AUTHOR,
)
from ._hwp_payload_codecs import (  # noqa: F401 - compatibility re-exports
    _bit_flags,
    _parse_arc_shape_payload,
    _parse_border_fill_payload,
    _parse_border_line_payload,
    _parse_bullet_payload,
    _parse_character_shape_payload,
    _parse_chart_data_payload,
    _parse_colorref,
    _parse_connectline_shape_payload,
    _parse_container_shape_payload,
    _parse_curve_shape_payload,
    _parse_ellipse_shape_payload,
    _parse_face_name_payload,
    _parse_form_object_payload,
    _parse_hwp_colorref_or_none,
    _parse_length_prefixed_utf16,
    _parse_line_attribute_flags,
    _parse_line_shape_payload,
    _parse_list_header_payload,
    _parse_memo_list_payload,
    _parse_memo_shape_native_metadata,
    _parse_memo_shape_payload,
    _parse_numbering_payload,
    _parse_paragraph_style_payload,
    _parse_polygon_shape_payload,
    _parse_rectangle_shape_payload,
    _parse_shape_component_common_payload,
    _parse_shape_fill_payload,
    _parse_shape_line_info_payload,
    _parse_style_definition_payload,
    _parse_tab_def_payload,
    _parse_textart_payload_text,
    _parse_track_change_author_payload,
    _parse_track_change_payload,
    _payload_u16_values,
    _payload_u32_values,
    _shape_common_payload_size,
    _shape_specific_payload_size,
)
from .exceptions import HwpBinaryEditError, InvalidHwpFileError


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


from ._hwp_graphic_payloads import (  # noqa: E402, F401 - compatibility re-exports
    _build_graphic_control_header_payload,
    _build_minimal_shape_component_payload,
    _shape_control_header_template,
    _build_shape_component_payload,
    _build_prefixed_metadata_payload,
    _build_shape_native_metadata_payload,
    _build_control_native_metadata_payload,
    _build_memo_shape_native_metadata_payload,
    _graphic_control_out_margin_values,
    _parse_graphic_control_out_margins,
    _parse_graphic_control_layout,
    _set_graphic_control_out_margins_payload,
    _set_graphic_control_layout_payload,
    _shape_component_rotation_values,
    _shape_component_default_rotation_center,
    _parse_shape_component_rotation_payload,
    _set_shape_component_rotation_payload,
    _picture_crop_values,
    _picture_default_clip_size,
    _parse_picture_crop_payload,
    _set_picture_crop_payload,
    _picture_image_adjustment_values,
    _picture_effect_name,
    _picture_effect_code,
    _parse_picture_image_adjustment_payload,
    _set_picture_image_adjustment_payload,
    _set_picture_line_payload,
    _build_line_attribute_flags,
    _build_shape_line_info_payload,
    _build_shape_fill_payload,
    _build_filled_shape_common_payload,
    _build_ellipse_specific_payload,
    _build_arc_specific_payload,
    _build_polygon_specific_payload,
    _build_textart_specific_payload,
    _build_ole_attribute_flags,
    _build_ole_specific_payload,
    _build_shape_specific_payload,
)


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
        char_count=(source_header.char_count & 0x80000000) | len(blank_raw_text),
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
        char_count=(source_header.char_count & 0x80000000) | len(blank_raw_text),
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

    paragraph = SectionParagraphModel(section_index=model.section_index, index=target_index, header=header)
    if target_index == len(current_paragraphs):
        model.roots.append(header)
        current_paragraphs.append(paragraph)
    else:
        root_index = model._root_insert_index_for_paragraph(target_index)
        model.roots.insert(root_index, header)
        model._paragraphs_cache = None

    return paragraph


def _find_section_definition_control_node(model: "SectionModel") -> ControlHeaderRecord | None:
    paragraphs = model.paragraphs()
    if not paragraphs:
        return None
    for child in paragraphs[0].header.children:
        if child.tag_id == TAG_CTRL_HEADER and _decode_control_id_payload(child.payload) == "secd":
            return child
    return None


def _normalize_section_definition_paragraph_for_multi_section(model: "SectionModel") -> None:
    paragraphs = model.paragraphs()
    if not paragraphs or _find_section_definition_control_node(model) is None:
        return
    header = paragraphs[0].header
    header.char_count |= 0x80000000
    header.children = [child for child in header.children if child.tag_id != TAG_PARA_LINE_SEG]
    trailing = bytearray(header.trailing_payload)
    if len(trailing) > 4:
        trailing[4] = 0
        header.trailing_payload = bytes(trailing)
    header.sync_payload()


def _find_child_record(node: RecordNode, tag_id: int) -> RecordNode | None:
    return next((child for child in node.children if child.tag_id == tag_id), None)


def _find_paragraph_control_node(model: "SectionModel", control_id: str) -> ControlHeaderRecord | None:
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


from ._hwp_page_payloads import (  # noqa: E402, F401 - compatibility re-exports
    _SECTION_DEF_HIDE_HEADER_FLAG,
    _SECTION_DEF_HIDE_FOOTER_FLAG,
    _SECTION_DEF_HIDE_MASTER_PAGE_FLAG,
    _SECTION_DEF_HIDE_PAGE_NUM_FLAG,
    _SECTION_DEF_HIDE_EMPTY_LINE_FLAG,
    _SECTION_DEF_WONGGOJI_FORMAT_FLAG,
    _SECTION_PAGE_STARTS_ON_MASK,
    _SECTION_PAGE_STARTS_ON_CODES,
    _SECTION_PAGE_STARTS_ON_VALUES,
    _SECTION_DEFINITION_STARTS_ON_SHIFT,
    _SECTION_DEFINITION_STARTS_ON_MASK,
    _SECTION_DEFINITION_STARTS_ON_CODES,
    _SECTION_DEFINITION_STARTS_ON_VALUES,
    _HEADER_FOOTER_APPLY_PAGE_TYPE_CODES,
    _HEADER_FOOTER_APPLY_PAGE_TYPE_VALUES,
    _PAGE_HIDING_FLAGS,
    _SECTION_DEF_BORDER_FIRST_FLAG,
    _SECTION_DEF_FILL_FIRST_FLAG,
    _PAGE_NUM_POSITION_CODES,
    _PAGE_NUM_POSITION_VALUES,
    _NUMBER_TYPE_CODES,
    _NUMBER_TYPE_VALUES,
    _NOTE_PLACEMENT_CODES,
    _NOTE_NUMBERING_CODES,
    _NOTE_NUMBERING_VALUES,
    _NOTE_LINE_TYPE_CODES,
    _NOTE_LINE_TYPE_VALUES,
    _NOTE_LINE_WIDTH_CODES,
    _NOTE_LINE_WIDTH_VALUES,
    _parse_section_definition_payload,
    _build_section_definition_payload,
    _parse_page_starts_on_payload,
    _build_page_starts_on_payload,
    _parse_page_hiding_payload,
    _build_page_hiding_payload,
    _parse_page_num_payload,
    _build_page_num_payload,
    _parse_note_shape_payload,
    _build_note_shape_payload,
    _build_colorref,
    _parse_rgb_color,
    _build_rgb_color,
    _parse_page_def_payload,
    _build_page_def_payload,
    _SECTION_PAGE_BORDER_FILL_TYPES,
    _PAGE_BORDER_FILL_TEXT_BORDER_FLAG,
    _PAGE_BORDER_FILL_HEADER_INSIDE_FLAG,
    _PAGE_BORDER_FILL_FOOTER_INSIDE_FLAG,
    _PAGE_BORDER_FILL_BORDER_AREA_FLAG,
    _parse_page_border_fill_payload,
    _build_page_border_fill_payload,
)


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


def _collect_text_from_node(node: "RecordNode" | None) -> str:
    if node is None:
        return ""
    texts: list[str] = []
    for descendant in node.iter_descendants():
        if isinstance(descendant, ParagraphTextRecord):
            texts.append(descendant.text.replace("\r", ""))
    return "".join(texts)


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


class OpaquePayloadRecord(TypedRecord):
    RECORD_TAG_ID = -1

    def __init__(
        self,
        *,
        level: int,
        payload: bytes,
        header_size: int = 4,
        offset: int = -1,
        children: list[RecordNode] | None = None,
    ) -> None:
        super().__init__(
            tag_id=self.RECORD_TAG_ID,
            level=level,
            payload=payload,
            header_size=header_size,
            offset=offset,
            children=list(children or []),
        )

    @classmethod
    def from_record(cls, record: HwpRecord) -> "OpaquePayloadRecord":
        return cls(
            level=record.level,
            payload=record.payload,
            header_size=record.header_size,
            offset=record.offset,
        )


class FaceNameRecord(OpaquePayloadRecord):
    RECORD_TAG_ID = TAG_FACE_NAME

    def fields(self) -> dict[str, object]:
        return _parse_face_name_payload(self.payload)


class BorderFillRecord(OpaquePayloadRecord):
    RECORD_TAG_ID = TAG_BORDER_FILL

    def fields(self) -> dict[str, object]:
        return _parse_border_fill_payload(self.payload)


class CharShapeRecord(OpaquePayloadRecord):
    RECORD_TAG_ID = TAG_CHAR_SHAPE

    def fields(self) -> dict[str, object]:
        return _parse_character_shape_payload(self.payload)


class TabDefRecord(OpaquePayloadRecord):
    RECORD_TAG_ID = TAG_TAB_DEF

    def fields(self) -> dict[str, object]:
        return _parse_tab_def_payload(self.payload)


class NumberingRecord(OpaquePayloadRecord):
    RECORD_TAG_ID = TAG_NUMBERING

    def fields(self) -> dict[str, object]:
        return _parse_numbering_payload(self.payload)


class BulletRecord(OpaquePayloadRecord):
    RECORD_TAG_ID = TAG_BULLET

    def fields(self) -> dict[str, object]:
        return _parse_bullet_payload(self.payload)


class ParaShapeRecord(OpaquePayloadRecord):
    RECORD_TAG_ID = TAG_PARA_SHAPE

    def fields(self) -> dict[str, object]:
        return _parse_paragraph_style_payload(self.payload)


class StyleRecord(OpaquePayloadRecord):
    RECORD_TAG_ID = TAG_STYLE

    def fields(self, *, style_id: int | None = None) -> dict[str, object]:
        return _parse_style_definition_payload(self.payload, style_id=style_id)


class ListHeaderRecord(OpaquePayloadRecord):
    RECORD_TAG_ID = TAG_LIST_HEADER

    def fields(self) -> dict[str, object]:
        return _parse_list_header_payload(self.payload)


class DocDataRecord(OpaquePayloadRecord):
    RECORD_TAG_ID = TAG_DOC_DATA


class CompatibleDocumentRecord(OpaquePayloadRecord):
    RECORD_TAG_ID = TAG_COMPATIBLE_DOCUMENT


class LayoutCompatibilityRecord(OpaquePayloadRecord):
    RECORD_TAG_ID = TAG_LAYOUT_COMPATIBILITY


class MemoShapeRecord(OpaquePayloadRecord):
    RECORD_TAG_ID = TAG_MEMO_SHAPE

    def fields(self) -> dict[str, object]:
        return _parse_memo_shape_payload(self.payload)


class FormObjectRecord(OpaquePayloadRecord):
    RECORD_TAG_ID = TAG_FORM_OBJECT

    def fields(self) -> dict[str, object]:
        return _parse_form_object_payload(self.payload)


class MemoListRecord(OpaquePayloadRecord):
    RECORD_TAG_ID = TAG_MEMO_LIST

    def fields(self) -> dict[str, object]:
        return _parse_memo_list_payload(self.payload)


class ChartDataRecord(OpaquePayloadRecord):
    RECORD_TAG_ID = TAG_CHART_DATA

    def fields(self) -> dict[str, object]:
        return _parse_chart_data_payload(self.payload)


class TrackChangeRecord(OpaquePayloadRecord):
    RECORD_TAG_ID = TAG_TRACK_CHANGE

    def fields(self) -> dict[str, object]:
        return _parse_track_change_payload(self.payload)


class TrackChangeAuthorRecord(OpaquePayloadRecord):
    RECORD_TAG_ID = TAG_TRACK_CHANGE_AUTHOR

    def fields(self) -> dict[str, object]:
        return _parse_track_change_author_payload(self.payload)


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


class ControlDataRecord(TypedRecord):
    def __init__(
        self,
        *,
        level: int,
        payload: bytes,
        header_size: int = 4,
        offset: int = -1,
        children: list[RecordNode] | None = None,
    ) -> None:
        super().__init__(
            tag_id=TAG_CTRL_DATA,
            level=level,
            payload=payload,
            header_size=header_size,
            offset=offset,
            children=list(children or []),
        )

    @classmethod
    def from_record(cls, record: HwpRecord) -> "ControlDataRecord":
        return cls(
            level=record.level,
            payload=record.payload,
            header_size=record.header_size,
            offset=record.offset,
        )

    def decoded_text(self) -> str:
        return self.payload.decode("utf-16-le", errors="ignore")

    def bookmark_name(self) -> str:
        return _parse_bookmark_ctrl_data_payload(self.payload)

    def native_metadata(self) -> dict[str, str]:
        return _parse_shape_native_metadata(self.payload)


class ControlHeaderRecord(TypedRecord):
    def __init__(
        self,
        *,
        level: int,
        payload: bytes,
        header_size: int = 4,
        offset: int = -1,
        children: list[RecordNode] | None = None,
    ) -> None:
        super().__init__(
            tag_id=TAG_CTRL_HEADER,
            level=level,
            payload=payload,
            header_size=header_size,
            offset=offset,
            children=list(children or []),
        )

    @classmethod
    def from_record(cls, record: HwpRecord) -> "ControlHeaderRecord":
        return cls(
            level=record.level,
            payload=record.payload,
            header_size=record.header_size,
            offset=record.offset,
        )

    @property
    def control_id(self) -> str:
        return _decode_control_id_payload(self.payload)

    def control_data_records(self) -> list[ControlDataRecord]:
        return [child for child in self.children if isinstance(child, ControlDataRecord)]

    def first_control_data(self) -> ControlDataRecord | None:
        return next(iter(self.control_data_records()), None)

    def list_header_records(self) -> list["ListHeaderRecord"]:
        return [child for child in self.children if isinstance(child, ListHeaderRecord)]

    def first_list_header(self) -> "ListHeaderRecord | None":
        return next(iter(self.list_header_records()), None)

    def nested_paragraphs(self) -> list["ParagraphHeaderRecord"]:
        return [node for node in self.iter_descendants() if isinstance(node, ParagraphHeaderRecord)]

    def text(self) -> str:
        return _collect_text_from_node(self).replace("\r", "").strip()


class SectionDefinitionControlRecord(ControlHeaderRecord):
    def settings(self) -> dict[str, object]:
        values = _parse_section_definition_payload(self.payload)
        attributes = int(values.get("attributes", 0))
        return {
            "attributes": attributes,
            "grid": {
                "space_columns": int(values.get("space_columns", 0)),
                "line_grid": int(values.get("line_grid", 0)),
                "char_grid": int(values.get("char_grid", 0)),
                "tab_stop": int(values.get("tab_stop", 0)),
                "wonggoji_format": bool(attributes & _SECTION_DEF_WONGGOJI_FORMAT_FLAG),
            },
            "start_numbers": {
                "page": int(values.get("page", 0)),
                "pic": int(values.get("pic", 0)),
                "tbl": int(values.get("tbl", 0)),
                "equation": int(values.get("equation", 0)),
            },
            "numbering_shape_id": int(values.get("numbering_shape_id", 0)),
            "language": int(values.get("language", 0)),
        }

    def page_def_record(self) -> "PageDefRecord | None":
        return next((child for child in self.children if isinstance(child, PageDefRecord)), None)

    def page_border_fill_records(self) -> list["PageBorderFillRecord"]:
        return [child for child in self.children if isinstance(child, PageBorderFillRecord)]

    def note_shape_records(self) -> list["FootnoteShapeRecord"]:
        return [child for child in self.children if isinstance(child, FootnoteShapeRecord)]

    def page_border_fills(self) -> list[dict[str, str | int]]:
        fills: list[dict[str, str | int]] = []
        for index, record in enumerate(self.page_border_fill_records()):
            border_type = _SECTION_PAGE_BORDER_FILL_TYPES[index] if index < len(_SECTION_PAGE_BORDER_FILL_TYPES) else f"EXTRA_{index}"
            fills.append(record.page_border_fill(border_type=border_type))
        return fills

    def note_settings(self) -> dict[str, dict[str, object]]:
        records = self.note_shape_records()
        settings: dict[str, dict[str, object]] = {}
        if len(records) >= 1:
            settings["footNotePr"] = records[0].note_shape(kind="footNotePr")
        if len(records) >= 2:
            settings["endNotePr"] = records[1].note_shape(kind="endNotePr")
        return settings


class PageHidingControlRecord(ControlHeaderRecord):
    def visibility(self) -> dict[str, str]:
        return _parse_page_hiding_payload(self.payload)


class PageStartsOnControlRecord(ControlHeaderRecord):
    def page_starts_on(self) -> str:
        return _parse_page_starts_on_payload(self.payload)


class PageNumberControlRecord(ControlHeaderRecord):
    def pos(self) -> str:
        return self.page_number().get("pos", "BOTTOM_CENTER")

    def format_type(self) -> str:
        return self.page_number().get("formatType", "DIGIT")

    def side_char(self) -> str:
        return self.page_number().get("sideChar", "-")

    def page_number(self) -> dict[str, str]:
        return _parse_page_num_payload(self.payload)

    def fields(self) -> dict[str, object]:
        values = self.page_number()
        values.update(
            {
                "pos": self.pos(),
                "format_type": self.format_type(),
                "side_char": self.side_char(),
            }
        )
        return values


class BookmarkControlRecord(ControlHeaderRecord):
    def name(self) -> str:
        data = self.first_control_data()
        return "" if data is None else data.bookmark_name()


class FieldControlRecord(ControlHeaderRecord):
    def native_field_type(self) -> str:
        return self.control_id

    def name(self) -> str | None:
        for key in ("FieldName", "Name", "name"):
            value = self.get_parameter(key)
            if value:
                return value
        return None

    def parameters(self) -> dict[str, str]:
        parameters: dict[str, str] = {}
        for entry in self.control_data_text().split(";"):
            key, separator, value = entry.partition("=")
            if separator and key:
                parameters[key] = value
        return parameters

    def get_parameter(self, name: str) -> str | None:
        return self.parameters().get(name)

    def field_type(self) -> str:
        return _semantic_field_type_from_native(self.native_field_type(), self.parameters())

    def is_mail_merge(self) -> bool:
        return self.field_type() == "MAILMERGE"

    def is_calculation(self) -> bool:
        return self.field_type() == "FORMULA"

    def is_cross_reference(self) -> bool:
        return self.field_type() == "CROSSREF"

    def is_doc_property(self) -> bool:
        return self.field_type() == "DOCPROPERTY"

    def is_date(self) -> bool:
        return self.field_type() == "DATE"

    def merge_field_name(self) -> str | None:
        return self.get_parameter("MergeField") or self.get_parameter("FieldName") or self.name()

    def expression(self) -> str | None:
        return self.get_parameter("Expression") or self.get_parameter("Command")

    def bookmark_name(self) -> str | None:
        return self.get_parameter("BookmarkName") or self.get_parameter("Path") or self.name()

    def document_property_name(self) -> str | None:
        if not self.is_doc_property():
            return None
        return self.get_parameter("FieldName") or self.name()

    def command(self) -> str:
        return _parse_hyperlink_command(self.payload) if len(self.payload) >= 9 else ""

    def control_data_text(self) -> str:
        data = self.first_control_data()
        return "" if data is None else data.decoded_text()

    def fields(self) -> dict[str, object]:
        return {
            "field_type": self.field_type(),
            "native_field_type": self.native_field_type(),
            "name": self.name(),
            "parameters": self.parameters(),
            "is_mail_merge": self.is_mail_merge(),
            "is_calculation": self.is_calculation(),
            "is_cross_reference": self.is_cross_reference(),
            "is_doc_property": self.is_doc_property(),
            "is_date": self.is_date(),
            "merge_field_name": self.merge_field_name(),
            "expression": self.expression(),
            "bookmark_name": self.bookmark_name(),
            "document_property_name": self.document_property_name(),
            "command": self.command(),
            "display_text": self.text(),
            "control_data_text": self.control_data_text(),
        }


class HyperlinkControlRecord(FieldControlRecord):
    def command(self) -> str:
        return _parse_hyperlink_command(self.payload)

    def url(self) -> str:
        return self.command().split(";", 1)[0].replace("\\:", ":")

    def metadata_fields(self) -> list[str]:
        parts = self.command().split(";")
        if parts and parts[-1] == "":
            parts = parts[:-1]
        return parts[1:]

    def display_text(self) -> str:
        return self.text()

    def fields(self) -> dict[str, object]:
        values = super().fields()
        values.update(
            {
                "url": self.url(),
                "metadata_fields": self.metadata_fields(),
            }
        )
        return values


class AutoNumberControlRecord(ControlHeaderRecord):
    def kind(self) -> str:
        return "newNum" if self.control_id == "nwno" else "autoNum"

    def number(self) -> str | None:
        if self.control_id == "nwno" and len(self.payload) >= 10:
            return str(int.from_bytes(self.payload[8:10], "little", signed=False))
        if self.control_id == "atno":
            return "1"
        return None

    def number_type(self) -> str:
        if self.control_id == "atno" and len(self.payload) >= 12:
            code = int.from_bytes(self.payload[8:12], "little", signed=False)
            return _AUTO_NUMBER_TYPE_VALUES.get(code, "PAGE")
        return "PAGE"

    def fields(self) -> dict[str, object]:
        return {
            "kind": self.kind(),
            "number": self.number(),
            "number_type": self.number_type(),
        }


class HeaderFooterControlRecord(ControlHeaderRecord):
    def kind(self) -> str:
        return "header" if self.control_id == "head" else "footer"

    def apply_page_type(self) -> str:
        return _parse_header_footer_apply_page_type(self.payload)

    def paragraph_texts(self) -> list[str]:
        return [_collect_text_from_node(paragraph).replace("\r", "").strip() for paragraph in self.nested_paragraphs()]

    def list_header(self) -> dict[str, object]:
        header = self.first_list_header()
        return {} if header is None else header.fields()

    def fields(self) -> dict[str, object]:
        return {
            "kind": self.kind(),
            "apply_page_type": self.apply_page_type(),
            "text": self.text(),
            "paragraph_texts": self.paragraph_texts(),
            "list_header": self.list_header(),
        }


class NoteControlRecord(ControlHeaderRecord):
    def kind(self) -> str:
        return "footNote" if self.control_id == "fn  " else "endNote"

    def paragraph_texts(self) -> list[str]:
        return [_collect_text_from_node(paragraph).replace("\r", "").strip() for paragraph in self.nested_paragraphs()]

    def list_header(self) -> dict[str, object]:
        header = self.first_list_header()
        return {} if header is None else header.fields()

    def fields(self) -> dict[str, object]:
        return {
            "kind": self.kind(),
            "text": self.text(),
            "paragraph_texts": self.paragraph_texts(),
            "list_header": self.list_header(),
        }


class GraphicControlRecord(ControlHeaderRecord):
    def graphic_size(self) -> dict[str, int]:
        if len(self.payload) < 24:
            return {"width": 0, "height": 0}
        return {
            "width": int.from_bytes(self.payload[16:20], "little", signed=False),
            "height": int.from_bytes(self.payload[20:24], "little", signed=False),
        }

    def description(self) -> str:
        return _parse_object_description(self.payload)

    def graphic_kind(self) -> str:
        return _shape_kind_from_control_node(self)

    def native_metadata(self) -> dict[str, str]:
        metadata: dict[str, str] = {}
        for data in self.control_data_records():
            metadata.update(data.native_metadata())
        return metadata

    def layout(self) -> dict[str, str]:
        native = _parse_graphic_control_layout(self.payload)
        metadata = _metadata_prefixed_str_map(self.native_metadata(), "layout.")
        if not native:
            return metadata
        merged = dict(metadata)
        merged.update(native)
        return merged

    def out_margins(self) -> dict[str, int]:
        native = _parse_graphic_control_out_margins(self.payload)
        return native or _metadata_prefixed_int_map(self.native_metadata(), "outMargin.")

    def rotation(self) -> dict[str, str]:
        return _metadata_prefixed_str_map(self.native_metadata(), "rotation.")

    def shape_component_record(self) -> "ShapeComponentRecord | None":
        return next((node for node in self.iter_descendants() if isinstance(node, ShapeComponentRecord) and node.tag_id == TAG_SHAPE_COMPONENT), None)

    def specific_shape_record(self) -> "ShapeComponentRecord | None":
        return next(
            (
                node
                for node in self.iter_descendants()
                if isinstance(
                    node,
                    (
                        PictureShapeComponentRecord,
                        OleShapeComponentRecord,
                        TextArtShapeComponentRecord,
                        ShapeComponentRecord,
                    ),
                )
                and node.tag_id
                in {
                    TAG_SHAPE_COMPONENT_LINE,
                    TAG_SHAPE_COMPONENT_RECTANGLE,
                    TAG_SHAPE_COMPONENT_ELLIPSE,
                    TAG_SHAPE_COMPONENT_ARC,
                    TAG_SHAPE_COMPONENT_POLYGON,
                    TAG_SHAPE_COMPONENT_CURVE,
                    TAG_SHAPE_COMPONENT_CONTAINER,
                    TAG_SHAPE_COMPONENT_OLE,
                    TAG_SHAPE_COMPONENT_PICTURE,
                    TAG_SHAPE_COMPONENT_TEXTART,
                }
            ),
            None,
        )


class EquationControlRecord(GraphicControlRecord):
    def eqedit_record(self) -> "EqEditRecord | None":
        return next((node for node in self.iter_descendants() if isinstance(node, EqEditRecord)), None)

    def script(self) -> str:
        eqedit_record = self.eqedit_record()
        if eqedit_record is None:
            return ""
        return eqedit_record.script()

    def font(self) -> str:
        eqedit_record = self.eqedit_record()
        if eqedit_record is None:
            return DEFAULT_HWP_EQUATION_FONT
        return eqedit_record.font()


class TableControlRecord(ControlHeaderRecord):
    pass


class PictureControlRecord(GraphicControlRecord):
    def picture_record(self) -> "PictureShapeComponentRecord | None":
        return next((node for node in self.iter_descendants() if isinstance(node, PictureShapeComponentRecord)), None)

    def storage_id(self) -> int:
        record = self.picture_record()
        return 0 if record is None else record.storage_id()

    def picture_size(self) -> dict[str, int]:
        component = self.shape_component_record()
        if component is None:
            return self.graphic_size()
        size = component.component_size()
        return size if any(size.values()) else self.graphic_size()

    def rotation(self) -> dict[str, str]:
        component = self.shape_component_record()
        native = {} if component is None else _parse_shape_component_rotation_payload(component.payload, include_center=True)
        if not native:
            return super().rotation()
        metadata = super().rotation()
        merged = dict(metadata)
        merged.update(native)
        return merged

    def image_adjustment(self) -> dict[str, str]:
        record = self.picture_record()
        native = {} if record is None else record.image_adjustment()
        return native or _metadata_prefixed_str_map(self.native_metadata(), "image.")

    def crop(self) -> dict[str, int]:
        record = self.picture_record()
        native = {} if record is None else record.crop()
        return native or _metadata_prefixed_int_map(self.native_metadata(), "crop.")


class ShapeControlRecord(GraphicControlRecord):
    def native_metadata(self) -> dict[str, str]:
        return super().native_metadata()

    def rotation(self) -> dict[str, str]:
        component = self.shape_component_record()
        native = {} if component is None else _parse_shape_component_rotation_payload(component.payload, include_center=True)
        if not native:
            return super().rotation()
        metadata = super().rotation()
        merged = dict(metadata)
        merged.update(native)
        return merged

    def text(self) -> str:
        metadata = self.native_metadata()
        if "text" in metadata:
            return metadata["text"]
        specific = self.specific_shape_record()
        if isinstance(specific, TextArtShapeComponentRecord):
            value = specific.text().strip()
            if value:
                return value
        return self.text_content()

    def text_content(self) -> str:
        return _collect_text_from_node(self).replace("\r", "").strip()

    def line_color(self) -> str:
        metadata = self.native_metadata()
        if "line_color" in metadata:
            return metadata["line_color"]
        component = self.shape_component_record()
        if component is not None and self.graphic_kind() in {"ellipse", "arc", "polygon", "textart"}:
            common = component.common_fields()
            line_color = common.get("line_color")
            if isinstance(line_color, str) and line_color:
                return line_color
        specific = self.specific_shape_record()
        if isinstance(specific, ShapeComponentRecord):
            common_payload, _specific_payload = specific.payload_parts()
            if len(common_payload) >= 4:
                return _parse_colorref(common_payload[0:4])
        return "#000000"

    def fill_color(self) -> str:
        metadata = self.native_metadata()
        if "fill_color" in metadata:
            return metadata["fill_color"]
        component = self.shape_component_record()
        if component is not None and self.graphic_kind() == "textart":
            common = component.common_fields()
            fill_color = common.get("fill_color")
            if isinstance(fill_color, str) and fill_color:
                return fill_color
        specific = self.specific_shape_record()
        if isinstance(specific, ShapeComponentRecord):
            common_payload, _specific_payload = specific.payload_parts()
            if len(common_payload) > 11:
                return _parse_shape_fill_payload(common_payload[11:]).get("faceColor", "#FFFFFF")
        return "#FFFFFF"

    def shape_size(self) -> dict[str, int]:
        component = self.shape_component_record()
        if component is None:
            return self.graphic_size()
        size = component.component_size()
        if self.graphic_kind() == "textart":
            graphic = self.graphic_size()
            return graphic if any(graphic.values()) else size
        return size if any(size.values()) else self.graphic_size()

    def specific_fields(self) -> dict[str, object]:
        specific = self.specific_shape_record()
        if specific is None or not hasattr(specific, "fields"):
            return {}
        if self.graphic_kind() == "connectLine" and isinstance(specific, LineShapeComponentRecord):
            _common_payload, specific_payload = specific.payload_parts()
            return _parse_connectline_shape_payload(specific_payload or specific.payload)
        return specific.fields()  # type: ignore[no-any-return]


class OleControlRecord(ShapeControlRecord):
    def ole_record(self) -> "OleShapeComponentRecord | None":
        return next((node for node in self.iter_descendants() if isinstance(node, OleShapeComponentRecord)), None)

    def storage_id(self) -> int:
        record = self.ole_record()
        return 0 if record is None else record.storage_id()

    def object_type(self) -> str:
        record = self.ole_record()
        return "UNKNOWN" if record is None else record.object_type()

    def draw_aspect(self) -> str:
        record = self.ole_record()
        return "CONTENT" if record is None else record.draw_aspect()

    def has_moniker(self) -> bool:
        record = self.ole_record()
        return False if record is None else record.has_moniker()

    def eq_baseline(self) -> int:
        record = self.ole_record()
        return 0 if record is None else record.eq_baseline()

    def line_width(self) -> int:
        record = self.ole_record()
        return 0 if record is None else record.line_width()

    def line_color(self) -> str:
        record = self.ole_record()
        return "#000000" if record is None else record.line_color()

    def rotation(self) -> dict[str, str]:
        component = self.shape_component_record()
        native = {} if component is None else _parse_shape_component_rotation_payload(component.payload, include_center=False)
        if not native:
            return super().rotation()
        metadata = super().rotation()
        merged = dict(metadata)
        merged.update(native)
        return merged

    def shape_size(self) -> dict[str, int]:
        return super().shape_size()

    def extent(self) -> dict[str, int]:
        record = self.ole_record()
        if record is not None:
            extent = record.ole_size()
            graphic_size = self.graphic_size()
            if any(extent.values()) and extent != graphic_size:
                return {
                    "x": extent["width"],
                    "y": extent["height"],
                }
        return _metadata_prefixed_int_map(self.native_metadata(), "extent.")


class PageDefRecord(TypedRecord):
    def __init__(self, *, level: int, payload: bytes, header_size: int = 4, offset: int = -1, children: list[RecordNode] | None = None) -> None:
        super().__init__(tag_id=TAG_PAGE_DEF, level=level, payload=payload, header_size=header_size, offset=offset, children=list(children or []))

    @classmethod
    def from_record(cls, record: HwpRecord) -> "PageDefRecord":
        return cls(level=record.level, payload=record.payload, header_size=record.header_size, offset=record.offset)

    def page_settings(self) -> dict[str, int]:
        return _parse_page_def_payload(self.payload)


class FootnoteShapeRecord(TypedRecord):
    def __init__(self, *, level: int, payload: bytes, header_size: int = 4, offset: int = -1, children: list[RecordNode] | None = None) -> None:
        super().__init__(
            tag_id=TAG_FOOTNOTE_SHAPE,
            level=level,
            payload=payload,
            header_size=header_size,
            offset=offset,
            children=list(children or []),
        )

    @classmethod
    def from_record(cls, record: HwpRecord) -> "FootnoteShapeRecord":
        return cls(level=record.level, payload=record.payload, header_size=record.header_size, offset=record.offset)

    def note_shape(self, *, kind: str) -> dict[str, object]:
        return _parse_note_shape_payload(self.payload, kind=kind)


class PageBorderFillRecord(TypedRecord):
    def __init__(self, *, level: int, payload: bytes, header_size: int = 4, offset: int = -1, children: list[RecordNode] | None = None) -> None:
        super().__init__(
            tag_id=TAG_PAGE_BORDER_FILL,
            level=level,
            payload=payload,
            header_size=header_size,
            offset=offset,
            children=list(children or []),
        )

    @classmethod
    def from_record(cls, record: HwpRecord) -> "PageBorderFillRecord":
        return cls(level=record.level, payload=record.payload, header_size=record.header_size, offset=record.offset)

    def page_border_fill(self, *, border_type: str = "UNKNOWN") -> dict[str, str | int]:
        return _parse_page_border_fill_payload(self.payload, border_type=border_type)


class EqEditRecord(TypedRecord):
    def __init__(self, *, level: int, payload: bytes, header_size: int = 4, offset: int = -1, children: list[RecordNode] | None = None) -> None:
        super().__init__(tag_id=TAG_EQEDIT, level=level, payload=payload, header_size=header_size, offset=offset, children=list(children or []))

    @classmethod
    def from_record(cls, record: HwpRecord) -> "EqEditRecord":
        return cls(level=record.level, payload=record.payload, header_size=record.header_size, offset=record.offset)

    def script(self) -> str:
        decoded = self.payload[4:].decode("utf-16-le", errors="ignore")
        script = decoded.split("Equation Version", 1)[0].replace("\x00", "").strip()
        if "ь" in script:
            script = script.split("ь", 1)[0]
        script = "".join(character for character in script if ord(character) >= 32 or character in "\t\n")
        if script.startswith("* "):
            script = script[2:]
        return script.rstrip("`").strip()

    def font(self) -> str:
        decoded = self.payload[4:].decode("utf-16-le", errors="ignore")
        marker_index = decoded.find("Equation Version")
        if marker_index < 0:
            return DEFAULT_HWP_EQUATION_FONT
        suffix = decoded[marker_index:]
        if "\x07" not in suffix:
            return DEFAULT_HWP_EQUATION_FONT
        font = suffix.split("\x07", 1)[1].replace("\x00", "").strip()
        return font or DEFAULT_HWP_EQUATION_FONT


class ShapeComponentRecord(TypedRecord):
    def __init__(self, *, tag_id: int = TAG_SHAPE_COMPONENT, level: int, payload: bytes, header_size: int = 4, offset: int = -1, children: list[RecordNode] | None = None) -> None:
        super().__init__(tag_id=tag_id, level=level, payload=payload, header_size=header_size, offset=offset, children=list(children or []))

    @classmethod
    def from_record(cls, record: HwpRecord) -> "ShapeComponentRecord":
        return cls(tag_id=record.tag_id, level=record.level, payload=record.payload, header_size=record.header_size, offset=record.offset)

    def component_size(self) -> dict[str, int]:
        if len(self.payload) < 28:
            return {"width": 0, "height": 0}
        return {
            "width": int.from_bytes(self.payload[20:24], "little", signed=False),
            "height": int.from_bytes(self.payload[24:28], "little", signed=False),
        }

    def common_fields(self) -> dict[str, object]:
        if self.tag_id != TAG_SHAPE_COMPONENT:
            return {}
        return _parse_shape_component_common_payload(self.payload)

    def payload_parts(self) -> tuple[bytes, bytes]:
        common_size = _shape_common_payload_size(self.tag_id)
        if common_size <= 0:
            return (b"", self.payload)
        if len(self.payload) <= common_size:
            return (self.payload, b"")
        return (self.payload[:common_size], self.payload[common_size:])


class LineShapeComponentRecord(ShapeComponentRecord):
    def fields(self) -> dict[str, object]:
        _common_payload, specific_payload = self.payload_parts()
        return _parse_line_shape_payload(specific_payload or self.payload)


class RectangleShapeComponentRecord(ShapeComponentRecord):
    def fields(self) -> dict[str, object]:
        _common_payload, specific_payload = self.payload_parts()
        return _parse_rectangle_shape_payload(specific_payload or self.payload)


class EllipseShapeComponentRecord(ShapeComponentRecord):
    def fields(self) -> dict[str, object]:
        return _parse_ellipse_shape_payload(self.payload)


class ArcShapeComponentRecord(ShapeComponentRecord):
    def fields(self) -> dict[str, object]:
        return _parse_arc_shape_payload(self.payload)


class PolygonShapeComponentRecord(ShapeComponentRecord):
    def fields(self) -> dict[str, object]:
        return _parse_polygon_shape_payload(self.payload)


class CurveShapeComponentRecord(ShapeComponentRecord):
    def fields(self) -> dict[str, object]:
        return _parse_curve_shape_payload(self.payload)


class ContainerShapeComponentRecord(ShapeComponentRecord):
    def fields(self) -> dict[str, object]:
        return _parse_container_shape_payload(self.payload)


class OleShapeComponentRecord(ShapeComponentRecord):
    def flags(self) -> int:
        return int.from_bytes(self.payload[0:4].ljust(4, b"\x00"), "little", signed=False)

    def storage_id(self) -> int:
        if len(self.payload) >= 14:
            return int.from_bytes(self.payload[12:14], "little", signed=False)
        return 0

    def object_type(self) -> str:
        return _HWP_OLE_OBJECT_TYPE_NAMES.get((self.flags() >> 16) & 0x3F, "UNKNOWN")

    def draw_aspect(self) -> str:
        return _HWP_DRAW_ASPECT_NAMES.get(self.flags() & 0xFF, "CONTENT")

    def has_moniker(self) -> bool:
        return bool(self.flags() & (1 << 8))

    def eq_baseline(self) -> int:
        return (self.flags() >> 9) & 0x7F

    def ole_size(self) -> dict[str, int]:
        return {
            "width": int.from_bytes(self.payload[4:8].ljust(4, b"\x00"), "little", signed=True),
            "height": int.from_bytes(self.payload[8:12].ljust(4, b"\x00"), "little", signed=True),
        }

    def line_color(self) -> str:
        return _parse_rgb_color(self.payload[14:18]) if len(self.payload) >= 18 else "#000000"

    def line_width(self) -> int:
        return int.from_bytes(self.payload[18:20].ljust(2, b"\x00"), "little", signed=False)

    def line_attribute_flags(self) -> dict[str, object]:
        raw = int.from_bytes(self.payload[20:24].ljust(4, b"\x00"), "little", signed=False)
        return _parse_line_attribute_flags(raw)


class PictureShapeComponentRecord(ShapeComponentRecord):
    def line_color(self) -> str:
        return _parse_rgb_color(self.payload[0:4]) if len(self.payload) >= 4 else "#000000"

    def line_width(self) -> int:
        return int.from_bytes(self.payload[4:8].ljust(4, b"\x00"), "little", signed=False)

    def crop(self) -> dict[str, int]:
        return _parse_picture_crop_payload(self.payload)

    def image_adjustment(self) -> dict[str, str]:
        return _parse_picture_image_adjustment_payload(self.payload)

    def storage_id(self) -> int:
        return int.from_bytes(self.payload[71:73].ljust(2, b"\x00"), "little", signed=False)


class TextArtShapeComponentRecord(ShapeComponentRecord):
    def text(self) -> str:
        return _parse_textart_payload_text(self.payload)

    def font_name(self) -> str:
        if len(self.payload) < 34:
            return ""
        text_length = int.from_bytes(self.payload[32:34], "little", signed=False)
        cursor = 34 + text_length * 2
        font_name, _ = _parse_length_prefixed_utf16(self.payload, cursor)
        return font_name

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


from ._hwp_record_tree import (  # noqa: E402, F401 - compatibility re-exports
    DocInfoModel,
    SectionModel,
    SectionParagraphModel,
    _iter_records,
    _specialize_record_node,
    _specialize_record_tree,
    _strip_para_text_controls,
    build_record_tree,
    flatten_record_tree,
    record_node_from_record,
)


class HwpBinaryDocument:
    def __init__(
        self,
        source_path: Path,
        streams: dict[str, bytes],
        file_header: HwpBinaryFileHeader,
        *,
        compound_layout: CfbLayout | None = None,
        original_file_bytes: bytes | None = None,
    ) -> None:
        self.source_path = source_path
        self._streams = dict(streams)
        self._original_sizes = {path: len(data) for path, data in streams.items()}
        self._file_header = file_header
        self._compound_layout = compound_layout
        self._original_file_bytes = original_file_bytes
        self._dirty = False
        self._section_model_batch_depth = 0
        self._section_model_cache: dict[int, SectionModel] = {}
        self._section_model_dirty: set[int] = set()
        self._docinfo_model_batch_depth = 0
        self._docinfo_model_cache: DocInfoModel | None = None
        self._docinfo_model_cache_dirty = False

    @classmethod
    def open(cls, path: str | Path) -> "HwpBinaryDocument":
        input_path = Path(path).expanduser().resolve()
        if not input_path.exists():
            raise FileNotFoundError(input_path)
        if not olefile.isOleFile(str(input_path)):
            raise InvalidHwpFileError(f"{input_path} is not a valid OLE Compound File HWP document.")

        original_file_bytes = input_path.read_bytes()

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

        return cls(
            source_path=input_path,
            streams=streams,
            file_header=file_header,
            compound_layout=capture_compound_file_layout(ole, original_file_bytes=original_file_bytes),
            original_file_bytes=original_file_bytes,
        )

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
        previous = self._streams.get(normalized)
        self._streams[normalized] = data
        self._original_sizes.setdefault(normalized, len(data))
        if previous != data:
            self._dirty = True

    def remove_stream(self, path: str | tuple[str, ...] | list[str]) -> None:
        normalized = _normalize_stream_path(path)
        removed = self._streams.pop(normalized, None)
        self._original_sizes.pop(normalized, None)
        if removed is not None:
            self._dirty = True

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

        if compress is None:
            compress = self._stream_is_compressed(normalized)

        if compress:
            compressor = zlib.compressobj(level=9, wbits=-15)
            payload = compressor.compress(data) + compressor.flush()
        else:
            payload = data

        if self._streams[normalized] != payload:
            self._streams[normalized] = payload
            self._dirty = True

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
        if self._docinfo_model_batch_depth > 0:
            if self._docinfo_model_cache is None:
                self._docinfo_model_cache = DocInfoModel.from_records(self.docinfo_records())
            return self._docinfo_model_cache
        return DocInfoModel.from_records(self.docinfo_records())

    def replace_docinfo_model(self, model: DocInfoModel) -> None:
        if self._docinfo_model_batch_depth > 0:
            self._docinfo_model_cache = model
            self._docinfo_model_cache_dirty = True
            self._dirty = True
            return
        self.write_stream("DocInfo", model.to_bytes(), compress=True)

    def begin_docinfo_model_batch(self) -> None:
        self._docinfo_model_batch_depth += 1

    def end_docinfo_model_batch(self) -> None:
        if self._docinfo_model_batch_depth <= 0:
            return
        self._docinfo_model_batch_depth -= 1
        if self._docinfo_model_batch_depth > 0:
            return
        model = self._docinfo_model_cache
        dirty = self._docinfo_model_cache_dirty
        self._docinfo_model_cache = None
        self._docinfo_model_cache_dirty = False
        if dirty and model is not None:
            self.write_stream("DocInfo", model.to_bytes(), compress=True)

    @contextmanager
    def docinfo_model_batch(self) -> Iterator[None]:
        self.begin_docinfo_model_batch()
        try:
            yield
        finally:
            self.end_docinfo_model_batch()

    def section_records(self, section_index: int) -> list[HwpRecord]:
        path = self.section_stream_paths()[section_index]
        return list(_iter_records(self.read_stream(path)))

    def replace_section_records(self, section_index: int, records: list[HwpRecord]) -> None:
        path = self.section_stream_paths()[section_index]
        raw = b"".join(record.to_bytes() for record in records)
        self.write_stream(path, raw, compress=True)

    def section_model(self, section_index: int = 0) -> SectionModel:
        if self._section_model_batch_depth > 0:
            cached = self._section_model_cache.get(section_index)
            if cached is not None:
                return cached
            model = SectionModel.from_records(section_index, self.section_records(section_index))
            self._section_model_cache[section_index] = model
            return model
        return SectionModel.from_records(section_index, self.section_records(section_index))

    def replace_section_model(self, section_index: int, model: SectionModel) -> None:
        if self._section_model_batch_depth > 0:
            self._section_model_cache[section_index] = model
            self._section_model_dirty.add(section_index)
            self._dirty = True
            return
        self.replace_section_records(section_index, model.to_records())

    def begin_section_model_batch(self) -> None:
        self._section_model_batch_depth += 1

    def end_section_model_batch(self) -> None:
        if self._section_model_batch_depth <= 0:
            return
        self._section_model_batch_depth -= 1
        if self._section_model_batch_depth > 0:
            return
        cached_models = self._section_model_cache
        dirty_indexes = self._section_model_dirty
        self._section_model_cache = {}
        self._section_model_dirty = set()
        for section_index in dirty_indexes:
            self.replace_section_records(section_index, cached_models[section_index].to_records())

    @contextmanager
    def section_model_batch(self) -> Iterator[None]:
        self.begin_section_model_batch()
        try:
            yield
        finally:
            self.end_section_model_batch()

    def ensure_section_count(self, section_count: int) -> None:
        if section_count < 1:
            raise ValueError("section_count must be at least 1.")

        current_count = len(self.section_stream_paths())
        if current_count == 0:
            raise InvalidHwpFileError("HWP document does not contain any BodyText/Section streams.")
        if current_count >= section_count:
            return

        if section_count > 1:
            for existing_section_index in range(current_count):
                existing_model = self.section_model(existing_section_index)
                _normalize_section_definition_paragraph_for_multi_section(existing_model)
                self.replace_section_model(existing_section_index, existing_model)

        while current_count < section_count:
            template_index = 1 if current_count > 1 else current_count - 1
            template_model = _build_blank_section_model(self.section_model(template_index), section_index=current_count)
            _normalize_section_definition_paragraph_for_multi_section(template_model)
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
        metadata = _control_native_metadata(section_definition)
        page_hiding = _parse_page_hiding_payload(_find_paragraph_control_node(model, "pghd").payload) if _find_paragraph_control_node(model, "pghd") is not None else {}
        attributes = int(values.get("attributes", 0))
        page_starts_on_code = (
            attributes & _SECTION_DEFINITION_STARTS_ON_MASK
        ) >> _SECTION_DEFINITION_STARTS_ON_SHIFT
        page_starts_on = _SECTION_DEFINITION_STARTS_ON_VALUES.get(page_starts_on_code, "BOTH")
        page_starts_on_control = _find_paragraph_control_node(model, "pgct")
        if page_starts_on_code == 0 and page_starts_on_control is not None:
            page_starts_on = _parse_page_starts_on_payload(page_starts_on_control.payload)
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
            "memo_shape_id": metadata.get("section.memoShapeId") or None,
        }

    def set_section_definition_settings(
        self,
        section_index: int,
        *,
        visibility: dict[str, str] | None = None,
        grid: dict[str, int] | None = None,
        start_numbers: dict[str, str] | None = None,
        numbering_shape_id: int | None = None,
        memo_shape_id: int | str | None = None,
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
        if memo_shape_id is not None:
            _update_control_native_metadata(
                section_definition,
                updates={"section.memoShapeId": memo_shape_id},
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
            page_starts_on_control = _find_paragraph_control_node(model, "pgct")
            if page_starts_on_control is not None:
                model.paragraphs()[0].header.children.remove(page_starts_on_control)
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
        native_parameters = dict(parameters or {})
        if name and not any(key in native_parameters for key in ("FieldName", "Name", "name")):
            native_parameters["FieldName"] = name
        if native_parameters:
            parameter_text = ";".join(f"{key}={value}" for key, value in native_parameters.items())
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
        apply_page_type: str = "BOTH",
        section_index: int | None = None,
    ) -> None:
        self._upsert_header_footer_control("head", text=text, apply_page_type=apply_page_type, section_index=section_index)

    def append_footer(
        self,
        text: str,
        *,
        apply_page_type: str = "BOTH",
        section_index: int | None = None,
    ) -> None:
        self._upsert_header_footer_control("foot", text=text, apply_page_type=apply_page_type, section_index=section_index)

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
        model._paragraphs_cache = None
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
        control_node = RecordNode(tag_id=TAG_CTRL_HEADER, level=1, payload=_build_control_id_payload("mrof"))
        semantic_metadata = {
            "form.type": str(form_type).upper(),
            "form.name": name,
            "form.value": value,
            "form.checked": checked,
            "form.editable": editable,
            "form.locked": locked,
            "form.placeholder": placeholder,
            "form.label": label,
        }
        if items:
            semantic_metadata["form.items"] = json.dumps(list(items), ensure_ascii=True, separators=(",", ":"))
        if any(field is not None for field in semantic_metadata.values()):
            control_node.add_child(
                RecordNode(
                    tag_id=TAG_CTRL_DATA,
                    level=2,
                    payload=_build_control_native_metadata_payload(**semantic_metadata),
                )
            )
        control_node.add_child(RecordNode(tag_id=TAG_FORM_OBJECT, level=2, payload=str(label).encode("utf-16-le")))
        paragraph.header.add_child(control_node)
        self.replace_section_model(target_section_index, model)

    def append_memo(
        self,
        text: str,
        *,
        author: str | None = None,
        memo_id: str | None = None,
        anchor_id: str | None = None,
        order: int | None = None,
        visible: bool = True,
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
        control_node = RecordNode(tag_id=TAG_CTRL_HEADER, level=1, payload=_build_control_id_payload("omem"))
        semantic_metadata = {
            "memo.author": author,
            "memo.memoId": memo_id,
            "memo.anchorId": anchor_id,
            "memo.order": order,
            "memo.visible": visible,
        }
        if any(field is not None for field in semantic_metadata.values()):
            control_node.add_child(
                RecordNode(
                    tag_id=TAG_CTRL_DATA,
                    level=2,
                    payload=_build_control_native_metadata_payload(**semantic_metadata),
                )
            )
        control_node.add_child(RecordNode(tag_id=TAG_MEMO_LIST, level=2, payload=str(text).encode("utf-16-le")))
        paragraph.header.add_child(control_node)
        self.replace_section_model(target_section_index, model)

    def append_chart(
        self,
        title: str = "",
        *,
        chart_type: str = "BAR",
        categories: Sequence[str] | None = None,
        series: Sequence[dict[str, object]] | None = None,
        data_ref: str | None = None,
        legend_visible: bool = True,
        width: int = DEFAULT_HWP_SHAPE_WIDTH,
        height: int = DEFAULT_HWP_SHAPE_HEIGHT,
        shape_comment: str | None = None,
        section_index: int | None = None,
        paragraph_index: int | None = None,
    ) -> None:
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
                _shape_control_header_template("rect"),
                "gso ",
                width=width,
                height=height,
                description=shape_comment,
            ),
        )
        semantic_metadata = {
            "chart.title": title,
            "chart.chartType": str(chart_type).upper(),
            "chart.dataRef": data_ref,
            "chart.legendVisible": legend_visible,
        }
        if categories:
            semantic_metadata["chart.categories"] = json.dumps(list(categories), ensure_ascii=True, separators=(",", ":"))
        if series:
            semantic_metadata["chart.series"] = json.dumps(list(series), ensure_ascii=True, separators=(",", ":"))
        control_node.add_child(
            RecordNode(
                tag_id=TAG_CTRL_DATA,
                level=2,
                payload=_build_shape_native_metadata_payload(**semantic_metadata),
            )
        )
        shape_component = RecordNode(
            tag_id=TAG_SHAPE_COMPONENT,
            level=2,
            payload=_build_shape_component_payload("rect", width, height),
        )
        shape_component.add_child(RecordNode(tag_id=TAG_CHART_DATA, level=3, payload=str(title).encode("utf-16-le")))
        control_node.add_child(shape_component)
        paragraph.header.add_child(control_node)
        self.replace_section_model(target_section_index, model)

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
            "connectLine": TAG_SHAPE_COMPONENT_LINE,
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
        if kind == "connectLine":
            metadata_fields["kind"] = "connectLine"
        normalized_fill_color = str(fill_color).strip().upper()
        if kind in {"ellipse", "arc", "polygon"} and normalized_fill_color and normalized_fill_color != "#FFFFFF":
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
            model._paragraphs_cache = None
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

    def save(self, path: str | Path | None = None, *, preserve_original_bytes: bool = True) -> Path:
        target_path = self.source_path if path is None else Path(path).expanduser().resolve()
        self._write_to_path(target_path, preserve_original_bytes=preserve_original_bytes)
        if path is not None:
            self.source_path = target_path
        return target_path

    def save_copy(self, path: str | Path, *, preserve_original_bytes: bool = True) -> Path:
        target_path = Path(path).expanduser().resolve()
        self._write_to_path(target_path, preserve_original_bytes=preserve_original_bytes)
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

    def _write_to_path(self, target_path: Path, *, preserve_original_bytes: bool) -> None:
        target_path.parent.mkdir(parents=True, exist_ok=True)
        if preserve_original_bytes and not self._dirty and self._original_file_bytes is not None:
            target_path.write_bytes(self._original_file_bytes)
            return
        write_compound_file(target_path, self._streams, layout=self._compound_layout)
        self._original_file_bytes = target_path.read_bytes()
        self._dirty = False

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
        apply_page_type: str,
        section_index: int | None,
    ) -> None:
        target_section_index = self._resolve_append_section_index(None, section_index)
        model = self.section_model(target_section_index)
        paragraphs = model.paragraphs()
        if not paragraphs:
            model.append_paragraph("\r")
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
            instance_id = 1 + sum(
                1
                for child in target_paragraph.header.children
                if child.tag_id == TAG_CTRL_HEADER
                and _decode_control_id_payload(child.payload) in {"head", "foot"}
            )
            control_node = _build_header_footer_control_node(
                control_id,
                apply_page_type=apply_page_type,
                instance_id=instance_id,
            )
            target_paragraph.header.add_child(control_node)
            _append_control_marker_to_paragraph(target_paragraph, control_id)
        else:
            instance_id = int.from_bytes(control_node.payload[8:12].ljust(4, b"\x00"), "little") or 1
            control_node.payload = _build_header_footer_control_payload(
                control_id,
                apply_page_type=apply_page_type,
                instance_id=instance_id,
            )
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
                size=24,
                header_size=4,
                offset=-1,
                payload=(0x80000000 | top_char_count).to_bytes(4, "little") + _PICTURE_TOP_PARA_HEADER_TAIL,
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


def _build_header_footer_control_node(
    control_id: str,
    *,
    apply_page_type: str = "BOTH",
    instance_id: int = 1,
) -> RecordNode:
    if control_id == "head":
        control_payload = _build_header_footer_control_payload(
            control_id,
            apply_page_type=apply_page_type,
            instance_id=instance_id,
        )
        list_payload = _HEADER_LIST_HEADER
        paragraph_payload = _HEADER_PARA_HEADER
        char_shape_payload = _HEADER_PARA_CHAR_SHAPE
        line_seg_payload = _HEADER_PARA_LINE_SEG
    elif control_id == "foot":
        control_payload = _build_header_footer_control_payload(
            control_id,
            apply_page_type=apply_page_type,
            instance_id=instance_id,
        )
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
    if line_seg_payload:
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
    paragraph.header.char_count = (paragraph.header.char_count & 0x80000000) | len(updated)
    if control_id in {"head", "foot"}:
        paragraph.header.control_mask |= 0x00010000
    if control_id == "head":
        paragraph.header.children = [
            child for child in paragraph.header.children if child.tag_id != TAG_PARA_LINE_SEG
        ]
        trailing = bytearray(paragraph.header.trailing_payload)
        if len(trailing) > 4:
            trailing[4] = 0
            paragraph.header.trailing_payload = bytes(trailing)
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
    paragraph_header.char_count = (paragraph_header.char_count & 0x80000000) | len(raw_text)
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
