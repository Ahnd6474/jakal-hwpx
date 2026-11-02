"""Native HWP nested-story controls and synthetic paragraph line layout."""
from __future__ import annotations

from ._hwp_constants import (
    _FOOTER_LIST_HEADER,
    _FOOTER_PARA_CHAR_SHAPE,
    _FOOTER_PARA_HEADER,
    _FOOTER_PARA_LINE_SEG,
    _HEADER_LIST_HEADER,
    _HEADER_PARA_CHAR_SHAPE,
    _HEADER_PARA_HEADER,
    _HEADER_PARA_LINE_SEG,
    _NOTE_CHAR_SHAPE,
    _NOTE_LINE_SEG,
    _NOTE_LIST_HEADER,
    _NOTE_PARA_HEADER,
    TAG_CTRL_HEADER,
    TAG_LIST_HEADER,
    TAG_PARA_CHAR_SHAPE,
    TAG_PARA_HEADER,
    TAG_PARA_LINE_SEG,
)
from ._hwp_control_payloads import (
    _build_control_marker_text,
    _build_header_footer_control_payload,
    _build_note_ctrl_payload,
)
from ._hwp_paragraph_records import (
    ParagraphHeaderRecord,
    ParagraphTextRecord,
)
from ._hwp_record_base import (
    HwpRecord,
    RecordNode,
)
from ._hwp_record_tree import (
    SectionModel,
    SectionParagraphModel,
)
from .exceptions import HwpBinaryEditError


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
        paragraph.header.add_child(
            RecordNode(tag_id=TAG_PARA_LINE_SEG, level=paragraph.header.level + 1, payload=payload)
        )
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
