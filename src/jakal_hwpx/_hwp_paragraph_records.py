"""Native HWP paragraph headers and editable text records."""
from __future__ import annotations

from ._hwp_constants import (
    TAG_PARA_HEADER,
    TAG_PARA_TEXT,
)
from ._hwp_record_base import (
    HwpRecord,
    RecordNode,
    TypedRecord,
)
from .exceptions import (
    HwpBinaryEditError,
)

def _replace_with_limit(source: str, old: str, new: str, count: int) -> tuple[str, int]:
    if count < 0:
        return source.replace(old, new), source.count(old)

    replaced = 0
    result = source
    while replaced < count and old in result:
        result = result.replace(old, new, 1)
        replaced += 1
    return result, replaced


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
        from ._hwp_record_tree import (
            _strip_para_text_controls,
        )

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


# Preserve historical import and pickle identities.
for _public_type in (ParagraphHeaderRecord, ParagraphTextRecord):
    _public_type.__module__ = "jakal_hwpx.hwp_binary"
del _public_type
