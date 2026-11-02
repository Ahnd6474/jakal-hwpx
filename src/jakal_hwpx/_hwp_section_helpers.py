"""Native HWP section cloning, blank paragraphs, and control-tree operations."""
from __future__ import annotations

import struct
from typing import (
    TYPE_CHECKING,
    Iterator,
)

from ._hwp_constants import (
    _BLANK_DROP_CONTROL_IDS,
    _DEFAULT_SHAPE_LIST_HEADER,
    _DEFAULT_SHAPE_TEXT_CHAR_SHAPE,
    _DEFAULT_SHAPE_TEXT_LINE_SEG,
    _DEFAULT_SHAPE_TEXT_PARA_CONTROL_MASK,
    _DEFAULT_SHAPE_TEXT_PARA_SHAPE_ID,
    _DEFAULT_SHAPE_TEXT_PARA_TRAILING,
    _DEFAULT_SHAPE_TEXT_SPLIT_FLAGS,
    _DEFAULT_SHAPE_TEXT_STYLE_ID,
    TAG_CTRL_HEADER,
    TAG_LIST_HEADER,
    TAG_PARA_CHAR_SHAPE,
    TAG_PARA_LINE_SEG,
)
from ._hwp_control_payloads import (
    _build_control_id_payload,
    _decode_control_id_payload,
)
from ._hwp_paragraph_records import (
    ParagraphHeaderRecord,
    ParagraphTextRecord,
)
from ._hwp_record_base import (
    RecordNode,
)
from .exceptions import (
    HwpBinaryEditError,
)

if TYPE_CHECKING:
    from ._hwp_record_tree import (
        SectionModel,
        SectionParagraphModel,
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
    from ._hwp_record_tree import (
        SectionModel,
    )

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
    from ._hwp_record_tree import (
        SectionParagraphModel,
    )

    current_paragraphs = model.paragraphs()
    target_index = len(current_paragraphs) if paragraph_index is None or paragraph_index < 0 else paragraph_index
    if target_index < 0 or target_index > len(current_paragraphs):
        raise IndexError(f"paragraph_index {target_index} is out of range for section {model.section_index}.")

    if current_paragraphs:
        header = _build_blank_paragraph_header(current_paragraphs[0])
    else:
        header = ParagraphHeaderRecord(
            level=0, char_count=1, control_mask=0, para_shape_id=0, style_id=0, split_flags=0,
        )
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


def _find_section_definition_control_node(model: "SectionModel") -> RecordNode | None:

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


def _collect_text_from_node(node: "RecordNode" | None) -> str:
    if node is None:
        return ""
    texts: list[str] = []
    for descendant in node.iter_descendants():
        if isinstance(descendant, ParagraphTextRecord):
            texts.append(descendant.text.replace("\r", ""))
    return "".join(texts)
