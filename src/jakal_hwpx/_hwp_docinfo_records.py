"""Native HWP DocInfo records, style slots, mappings, and binary references."""
from __future__ import annotations

from ._hwp_constants import (
    TAG_BIN_DATA,
    TAG_BORDER_FILL,
    TAG_BULLET,
    TAG_CHAR_SHAPE,
    TAG_CHART_DATA,
    TAG_COMPATIBLE_DOCUMENT,
    TAG_DOC_DATA,
    TAG_DOCUMENT_PROPERTIES,
    TAG_FACE_NAME,
    TAG_FORM_OBJECT,
    TAG_ID_MAPPINGS,
    TAG_LAYOUT_COMPATIBILITY,
    TAG_LIST_HEADER,
    TAG_MEMO_LIST,
    TAG_MEMO_SHAPE,
    TAG_NUMBERING,
    TAG_PARA_SHAPE,
    TAG_STYLE,
    TAG_TAB_DEF,
    TAG_TRACK_CHANGE,
    TAG_TRACK_CHANGE_AUTHOR,
)
from ._hwp_payload_codecs import (
    _parse_border_fill_payload,
    _parse_bullet_payload,
    _parse_character_shape_payload,
    _parse_chart_data_payload,
    _parse_face_name_payload,
    _parse_form_object_payload,
    _parse_list_header_payload,
    _parse_memo_list_payload,
    _parse_memo_shape_payload,
    _parse_numbering_payload,
    _parse_paragraph_style_payload,
    _parse_style_definition_payload,
    _parse_tab_def_payload,
    _parse_track_change_author_payload,
    _parse_track_change_payload,
)
from ._hwp_record_base import (
    HwpDocumentProperties,
    HwpRecord,
    OpaquePayloadRecord,
    RecordNode,
    TypedRecord,
)
from .exceptions import (
    InvalidHwpFileError,
)

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


# Preserve historical import and pickle identities.
for _public_type in (FaceNameRecord, BorderFillRecord, CharShapeRecord, TabDefRecord, NumberingRecord, BulletRecord, ParaShapeRecord, StyleRecord, ListHeaderRecord, DocDataRecord, CompatibleDocumentRecord, LayoutCompatibilityRecord, MemoShapeRecord, FormObjectRecord, MemoListRecord, ChartDataRecord, TrackChangeRecord, TrackChangeAuthorRecord, DocumentPropertiesRecord, IdMappingsRecord, BinDataRecord):
    _public_type.__module__ = "jakal_hwpx.hwp_binary"
del _public_type
