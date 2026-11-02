"""Raw HWP records, common record nodes, and immutable document metadata."""
from __future__ import annotations

from dataclasses import (
    dataclass,
    field,
)
from typing import (
    Iterator,
    cast,
)

from ._hwp_constants import (
    FLAG_COMPRESSED,
    FLAG_DISTRIBUTABLE,
    FLAG_PASSWORD_PROTECTED,
    TAG_NAMES,
)
from .exceptions import (
    InvalidHwpFileError,
)

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
            version_bytes=cast(tuple[int, int, int, int], tuple(data[32:36])),
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
        from ._hwp_record_tree import (
            record_node_from_record,
        )

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


# Preserve historical import and pickle identities.
for _public_type in (HwpBinaryFileHeader, HwpRecord, HwpDocumentProperties, HwpParagraph, HwpStreamCapacity, RecordNode, UnknownRecord, TypedRecord, OpaquePayloadRecord):
    _public_type.__module__ = "jakal_hwpx.hwp_binary"
del _public_type
