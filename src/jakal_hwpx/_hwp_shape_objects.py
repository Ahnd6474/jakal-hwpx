from __future__ import annotations

from collections.abc import Sequence
from typing import TYPE_CHECKING

from ._hwp_object_payloads import (  # noqa: F401 - compatibility re-exports
    _CONTROL_NATIVE_METADATA_PREFIX,
    _HANCOM_CONNECTLINE_COMPONENT_SIGNATURE,
    _HWP_OLE_DRAW_ASPECT_NAMES,
    _HWP_OLE_OBJECT_TYPE_NAMES,
    _HWP_SHAPE_ARC_SPECIFIC_SIZE,
    _HWP_SHAPE_ELLIPSE_SPECIFIC_SIZE,
    _HWP_SHAPE_LINE_SPECIFIC_SIZE,
    _HWP_SHAPE_RECTANGLE_SPECIFIC_SIZE,
    _HWP_TEXTART_PREFIX_PAYLOAD,
    _HWP_TEXTART_TRAILING_PAYLOAD,
    _SHAPE_NATIVE_METADATA_PREFIX,
    _build_arc_geometry_payload,
    _build_ellipse_geometry_payload,
    _build_line_specific_payload,
    _build_polygon_points_payload,
    _build_rectangle_specific_payload,
    _build_textart_payload_fields,
    _coerce_shape_point,
    _collect_control_native_metadata,
    _collect_shape_native_metadata,
    _control_metadata_node,
    _encode_length_prefixed_utf16,
    _graphic_layout_metadata_updates,
    _graphic_rotation_metadata_updates,
    _has_hancom_connectline_signature,
    _image_adjustment_metadata_updates,
    _metadata_bool_literal_is_valid,
    _metadata_bool_value,
    _metadata_json_dict_list,
    _metadata_json_list,
    _metadata_json_value,
    _metadata_prefixed_int_map,
    _metadata_prefixed_str_map,
    _parse_connectline_payload_fields,
    _parse_control_native_metadata,
    _parse_length_prefixed_utf16,
    _parse_object_description,
    _parse_raw_payload_fields,
    _parse_shape_fill_payload,
    _parse_shape_native_metadata,
    _parse_textart_payload_fields,
    _parse_textart_payload_text,
    _replace_object_description,
    _set_payload_size,
    _shape_metadata_node,
    _shape_specific_payload_size,
    _update_control_native_metadata,
    _update_shape_native_metadata,
)
from .hwp_binary import (
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
    RecordNode,
    _collect_text_from_node,
    _build_colorref,
    _parse_arc_shape_payload,
    _parse_colorref,
    _parse_container_shape_payload,
    _parse_curve_shape_payload,
    _parse_ellipse_shape_payload,
    _parse_line_shape_payload,
    _parse_polygon_shape_payload,
    _parse_rectangle_shape_payload,
    _parse_shape_component_rotation_payload,
    _set_shape_component_rotation_payload,
    _shape_kind_from_control_node,
    _shape_common_payload_size,
)

if TYPE_CHECKING:
    pass

from ._hwp_objects import HwpControlObject


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

    @property
    def line_width(self) -> int:
        metadata = self._shape_native_metadata()
        if "line_width" in metadata:
            try:
                return int(metadata["line_width"])
            except ValueError:
                pass
        shape_component = self._shape_component_record()
        if shape_component is not None and len(shape_component.payload) >= 202:
            return int.from_bytes(shape_component.payload[200:202], "little", signed=False)
        return 33

    def descendant_tag_ids(self) -> list[int]:
        return [node.tag_id for node in self.control_node.iter_descendants()]

    def layout(self) -> dict[str, str]:
        return self._graphic_layout()

    def out_margins(self) -> dict[str, int]:
        return self._graphic_out_margins()

    def text_margins(self) -> dict[str, int]:
        return _metadata_prefixed_int_map(self._shape_native_metadata(), "textMargin.")

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

    def line_style(self) -> dict[str, str]:
        metadata = _metadata_prefixed_str_map(self._shape_native_metadata(), "lineStyle.")
        result = {
            "color": self.line_color,
            "width": str(self.line_width),
        }
        result.update(metadata)
        return result

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
        shape_component = self._shape_component_record(control_node)
        if shape_component is not None and len(shape_component.payload) >= 202:
            payload = bytearray(shape_component.payload)
            if color is not None:
                payload[196:200] = _build_colorref(color)
            if width is not None:
                payload[200:202] = max(0, min(int(width), 0xFFFF)).to_bytes(2, "little", signed=False)
            shape_component.payload = bytes(payload)
        _update_shape_native_metadata(
            control_node,
            updates={
                "line_width": width,
                "line_color": color,
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

    def set_text_margins(
        self,
        *,
        left: int | str | None = None,
        right: int | str | None = None,
        top: int | str | None = None,
        bottom: int | str | None = None,
    ) -> None:
        section_model, _paragraph, control_node = self._live_context()
        _update_shape_native_metadata(
            control_node,
            updates={
                "textMargin.left": left,
                "textMargin.right": right,
                "textMargin.top": top,
                "textMargin.bottom": bottom,
            },
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

