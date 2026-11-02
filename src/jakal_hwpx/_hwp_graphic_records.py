"""Native HWP graphic, equation, picture, shape, and OLE control records."""
from __future__ import annotations

from typing import cast

from ._hwp_component_records import (
    EqEditRecord,
    LineShapeComponentRecord,
    OleShapeComponentRecord,
    PictureShapeComponentRecord,
    ShapeComponentRecord,
    TextArtShapeComponentRecord,
)
from ._hwp_constants import (
    DEFAULT_HWP_EQUATION_FONT,
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
from ._hwp_control_payloads import (
    _metadata_prefixed_int_map,
    _metadata_prefixed_str_map,
    _parse_object_description,
    _shape_kind_from_control_node,
)
from ._hwp_control_records import (
    ControlHeaderRecord,
)
from ._hwp_graphic_payloads import (
    _parse_graphic_control_layout,
    _parse_graphic_control_out_margins,
    _parse_shape_component_rotation_payload,
)
from ._hwp_payload_codecs import (
    _parse_colorref,
    _parse_connectline_shape_payload,
    _parse_shape_fill_payload,
)
from ._hwp_section_helpers import (
    _collect_text_from_node,
)


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
                return cast(str, _parse_shape_fill_payload(common_payload[11:]).get("faceColor", "#FFFFFF"))
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


# Preserve historical import and pickle identities.
for _public_type in (GraphicControlRecord, EquationControlRecord, TableControlRecord, PictureControlRecord, ShapeControlRecord, OleControlRecord):
    _public_type.__module__ = "jakal_hwpx.hwp_binary"
del _public_type
