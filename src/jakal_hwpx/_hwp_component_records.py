"""Native HWP page, equation, and shape component record types."""
from __future__ import annotations

from ._hwp_constants import (
    _HWP_DRAW_ASPECT_NAMES,
    _HWP_OLE_OBJECT_TYPE_NAMES,
    DEFAULT_HWP_EQUATION_FONT,
    TAG_EQEDIT,
    TAG_FOOTNOTE_SHAPE,
    TAG_PAGE_BORDER_FILL,
    TAG_PAGE_DEF,
    TAG_SHAPE_COMPONENT,
)
from ._hwp_graphic_payloads import (
    _parse_picture_crop_payload,
    _parse_picture_image_adjustment_payload,
)
from ._hwp_page_payloads import (
    _parse_note_shape_payload,
    _parse_page_border_fill_payload,
    _parse_page_def_payload,
    _parse_rgb_color,
)
from ._hwp_payload_codecs import (
    _parse_arc_shape_payload,
    _parse_container_shape_payload,
    _parse_curve_shape_payload,
    _parse_ellipse_shape_payload,
    _parse_length_prefixed_utf16,
    _parse_line_attribute_flags,
    _parse_line_shape_payload,
    _parse_polygon_shape_payload,
    _parse_rectangle_shape_payload,
    _parse_shape_component_common_payload,
    _parse_textart_payload_text,
    _shape_common_payload_size,
)
from ._hwp_record_base import (
    HwpRecord,
    RecordNode,
    TypedRecord,
)


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


# Preserve historical import and pickle identities.
for _public_type in (PageDefRecord, FootnoteShapeRecord, PageBorderFillRecord, EqEditRecord, ShapeComponentRecord, LineShapeComponentRecord, RectangleShapeComponentRecord, EllipseShapeComponentRecord, ArcShapeComponentRecord, PolygonShapeComponentRecord, CurveShapeComponentRecord, ContainerShapeComponentRecord, OleShapeComponentRecord, PictureShapeComponentRecord, TextArtShapeComponentRecord):
    _public_type.__module__ = "jakal_hwpx.hwp_binary"
del _public_type
