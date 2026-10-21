from __future__ import annotations

import json
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
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
    _AUTO_NUMBER_TYPE_VALUES,
    DEFAULT_HWP_EQUATION_FONT,
    TAG_CHART_DATA,
    TAG_CTRL_DATA,
    TAG_CTRL_HEADER,
    TAG_EQEDIT,
    TAG_FORM_OBJECT,
    TAG_LIST_HEADER,
    TAG_MEMO_LIST,
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
    HwpParagraph,
    ParagraphHeaderRecord,
    ParagraphTextRecord,
    RecordNode,
    SectionModel,
    SectionParagraphModel,
    _build_auto_number_payload,
    _build_bookmark_ctrl_data_payload,
    _build_colorref,
    _build_control_id_payload,
    _build_control_marker_text,
    _build_header_footer_control_payload,
    _build_hyperlink_ctrl_payload,
    _build_hyperlink_para_text_payload,
    _build_ole_attribute_flags,
    _build_page_num_payload,
    _build_rgb_color,
    _normalize_field_control_id,
    _parse_chart_data_payload,
    _parse_colorref,
    _parse_form_object_payload,
    _parse_graphic_control_layout,
    _parse_graphic_control_out_margins,
    _parse_memo_list_payload,
    _parse_picture_crop_payload,
    _parse_picture_image_adjustment_payload,
    _parse_rgb_color,
    _parse_shape_component_rotation_payload,
    _picture_effect_code,
    _set_graphic_control_layout_payload,
    _set_graphic_control_out_margins_payload,
    _set_picture_crop_payload,
    _set_picture_image_adjustment_payload,
    _set_picture_line_payload,
    _set_shape_component_rotation_payload,
)

if TYPE_CHECKING:
    from .hwp_document import HwpDocument


@dataclass(frozen=True)
class HwpSection:
    document: HwpDocument
    index: int

    def paragraphs(self) -> list["HwpParagraphObject"]:
        return [HwpParagraphObject(self.document, paragraph) for paragraph in self.document.binary_document().paragraphs(self.index)]

    def paragraph(self, index: int) -> "HwpParagraphObject":
        return self.paragraphs()[index]

    def records(self):
        return self.document.binary_document().section_records(self.index)

    def model(self) -> SectionModel:
        return self.document.binary_document().section_model(self.index)

    def controls(self) -> list["HwpControlObject"]:
        section_model = self.model()
        controls: list[HwpControlObject] = []
        for paragraph in section_model.paragraphs():
            control_ordinal = 0
            for control_node in paragraph.header.children:
                if control_node.tag_id != TAG_CTRL_HEADER:
                    continue
                wrapper = _build_control_wrapper(self.document, section_model, paragraph, control_node, control_ordinal)
                if wrapper is not None:
                    controls.append(wrapper)
                control_ordinal += 1
        return controls

    def tables(self) -> list["HwpTableObject"]:
        return [control for control in self.controls() if isinstance(control, HwpTableObject)]

    def pictures(self) -> list["HwpPictureObject"]:
        return [control for control in self.controls() if isinstance(control, HwpPictureObject)]

    def hyperlinks(self) -> list["HwpHyperlinkObject"]:
        return [control for control in self.controls() if isinstance(control, HwpHyperlinkObject)]

    def bookmarks(self) -> list["HwpBookmarkObject"]:
        return [control for control in self.controls() if isinstance(control, HwpBookmarkObject)]

    def notes(self) -> list["HwpNoteObject"]:
        return [control for control in self.controls() if isinstance(control, HwpNoteObject)]

    def page_numbers(self) -> list["HwpPageNumObject"]:
        return [control for control in self.controls() if isinstance(control, HwpPageNumObject)]

    def fields(self) -> list["HwpFieldObject"]:
        return [control for control in self.controls() if isinstance(control, HwpFieldObject)]

    def mail_merge_fields(self) -> list["HwpFieldObject"]:
        return [field for field in self.fields() if field.is_mail_merge]

    def calculation_fields(self) -> list["HwpFieldObject"]:
        return [field for field in self.fields() if field.is_calculation]

    def cross_references(self) -> list["HwpFieldObject"]:
        return [field for field in self.fields() if field.is_cross_reference]

    def doc_property_fields(self) -> list["HwpFieldObject"]:
        return [field for field in self.fields() if field.is_doc_property]

    def date_fields(self) -> list["HwpFieldObject"]:
        return [field for field in self.fields() if field.is_date]

    def shapes(self) -> list["HwpShapeObject"]:
        return [control for control in self.controls() if isinstance(control, HwpShapeObject)]

    def lines(self) -> list["HwpLineShapeObject"]:
        return [shape for shape in self.shapes() if isinstance(shape, HwpLineShapeObject) and not isinstance(shape, HwpConnectLineShapeObject)]

    def connect_lines(self) -> list["HwpConnectLineShapeObject"]:
        return [shape for shape in self.shapes() if isinstance(shape, HwpConnectLineShapeObject)]

    def rectangles(self) -> list["HwpRectangleShapeObject"]:
        return [shape for shape in self.shapes() if isinstance(shape, HwpRectangleShapeObject)]

    def ellipses(self) -> list["HwpEllipseShapeObject"]:
        return [shape for shape in self.shapes() if isinstance(shape, HwpEllipseShapeObject)]

    def arcs(self) -> list["HwpArcShapeObject"]:
        return [shape for shape in self.shapes() if isinstance(shape, HwpArcShapeObject)]

    def polygons(self) -> list["HwpPolygonShapeObject"]:
        return [shape for shape in self.shapes() if isinstance(shape, HwpPolygonShapeObject)]

    def curves(self) -> list["HwpCurveShapeObject"]:
        return [shape for shape in self.shapes() if isinstance(shape, HwpCurveShapeObject)]

    def containers(self) -> list["HwpContainerShapeObject"]:
        return [shape for shape in self.shapes() if isinstance(shape, HwpContainerShapeObject)]

    def textarts(self) -> list["HwpTextArtShapeObject"]:
        return [shape for shape in self.shapes() if isinstance(shape, HwpTextArtShapeObject)]

    def charts(self) -> list["HwpChartObject"]:
        return [shape for shape in self.shapes() if isinstance(shape, HwpChartObject)]

    def oles(self) -> list["HwpOleObject"]:
        return [control for control in self.controls() if isinstance(control, HwpOleObject)]

    def forms(self) -> list["HwpFormObject"]:
        return [control for control in self.controls() if isinstance(control, HwpFormObject)]

    def memos(self) -> list["HwpMemoObject"]:
        return [control for control in self.controls() if isinstance(control, HwpMemoObject)]

    def equations(self) -> list["HwpEquationObject"]:
        return [control for control in self.controls() if isinstance(control, HwpEquationObject)]

    def replace_text_same_length(self, old: str, new: str, *, count: int = -1) -> int:
        return self.document.replace_text_same_length(old, new, section_index=self.index, count=count)

    def append_paragraph(
        self,
        text: str,
        *,
        para_shape_id: int = 0,
        style_id: int = 0,
        split_flags: int = 0,
        control_mask: int = 0,
    ) -> "HwpParagraphObject":
        return self.document.append_paragraph(
            text,
            section_index=self.index,
            para_shape_id=para_shape_id,
            style_id=style_id,
            split_flags=split_flags,
            control_mask=control_mask,
        )

    def append_table(
        self,
        cell_text: str | None = None,
        *,
        rows: int = 1,
        cols: int = 1,
        cell_texts: Sequence[str] | Sequence[Sequence[str]] | None = None,
        row_heights: Sequence[int] | None = None,
        col_widths: Sequence[int] | None = None,
        cell_spans: dict[tuple[int, int], tuple[int, int]] | None = None,
        cell_border_fill_ids: dict[tuple[int, int], int] | None = None,
        table_border_fill_id: int = 1,
    ) -> None:
        self.document.append_table(
            cell_text,
            rows=rows,
            cols=cols,
            cell_texts=cell_texts,
            row_heights=row_heights,
            col_widths=col_widths,
            cell_spans=cell_spans,
            cell_border_fill_ids=cell_border_fill_ids,
            table_border_fill_id=table_border_fill_id,
            section_index=self.index,
        )

    def append_picture(
        self,
        image_bytes: bytes | None = None,
        *,
        extension: str | None = None,
        width: int = 12000,
        height: int = 3200,
    ) -> None:
        self.document.append_picture(
            image_bytes,
            extension=extension,
            width=width,
            height=height,
            section_index=self.index,
        )

    def append_hyperlink(
        self,
        url: str | None = None,
        *,
        text: str | None = None,
        metadata_fields: Sequence[str | int] | None = None,
    ) -> None:
        self.document.append_hyperlink(
            url,
            text=text,
            metadata_fields=metadata_fields,
            section_index=self.index,
        )

    def apply_settings(
        self,
        *,
        page_width: int | None = None,
        page_height: int | None = None,
        landscape: str | None = None,
        margins: dict[str, int] | None = None,
        visibility: dict[str, str] | None = None,
        grid: dict[str, int] | None = None,
        start_numbers: dict[str, str] | None = None,
        numbering_shape_id: str | None = None,
        memo_shape_id: str | None = None,
    ) -> None:
        self.document.apply_section_settings(
            section_index=self.index,
            page_width=page_width,
            page_height=page_height,
            landscape=landscape,
            margins=margins,
            visibility=visibility,
            grid=grid,
            start_numbers=start_numbers,
            numbering_shape_id=numbering_shape_id,
            memo_shape_id=memo_shape_id,
        )

    def apply_page_border_fills(self, page_border_fills: list[dict[str, str | int]]) -> None:
        self.document.apply_section_page_border_fills(page_border_fills, section_index=self.index)

    def apply_page_numbers(self, page_numbers: list[dict[str, str]]) -> None:
        self.document.apply_section_page_numbers(page_numbers, section_index=self.index)

    def apply_note_settings(
        self,
        *,
        footnote_pr: dict[str, object] | None = None,
        endnote_pr: dict[str, object] | None = None,
    ) -> None:
        self.document.apply_section_note_settings(
            section_index=self.index,
            footnote_pr=footnote_pr,
            endnote_pr=endnote_pr,
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
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_field(
            field_type=field_type,
            display_text=display_text,
            name=name,
            parameters=parameters,
            editable=editable,
            dirty=dirty,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_mail_merge_field(
        self,
        field_name: str,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_mail_merge_field(
            field_name,
            display_text=display_text,
            editable=editable,
            dirty=dirty,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_calculation_field(
        self,
        expression: str,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_calculation_field(
            expression,
            display_text=display_text,
            editable=editable,
            dirty=dirty,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_cross_reference(
        self,
        bookmark_name: str,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_cross_reference(
            bookmark_name,
            display_text=display_text,
            editable=editable,
            dirty=dirty,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_doc_property_field(
        self,
        property_name: str,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_doc_property_field(
            property_name,
            display_text=display_text,
            editable=editable,
            dirty=dirty,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_date_field(
        self,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_date_field(
            display_text=display_text,
            editable=editable,
            dirty=dirty,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_bookmark(self, name: str, *, paragraph_index: int | None = None) -> None:
        self.document.append_bookmark(name, section_index=self.index, paragraph_index=paragraph_index)

    def append_note(
        self,
        text: str,
        *,
        kind: str = "footNote",
        number: int | str | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_note(
            text,
            kind=kind,
            number=number,
            section_index=self.index,
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
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_form(
            label,
            form_type=form_type,
            name=name,
            value=value,
            checked=checked,
            items=items,
            editable=editable,
            locked=locked,
            placeholder=placeholder,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_memo(
        self,
        text: str,
        *,
        author: str | None = None,
        memo_id: str | None = None,
        anchor_id: str | None = None,
        order: int | None = None,
        visible: bool = True,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_memo(
            text,
            author=author,
            memo_id=memo_id,
            anchor_id=anchor_id,
            order=order,
            visible=visible,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_chart(
        self,
        title: str = "",
        *,
        chart_type: str = "BAR",
        categories: Sequence[str] | None = None,
        series: Sequence[dict[str, object]] | None = None,
        data_ref: str | None = None,
        legend_visible: bool = True,
        width: int = 12000,
        height: int = 3200,
        shape_comment: str | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_chart(
            title,
            chart_type=chart_type,
            categories=categories,
            series=series,
            data_ref=data_ref,
            legend_visible=legend_visible,
            width=width,
            height=height,
            shape_comment=shape_comment,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_footnote(
        self,
        text: str,
        *,
        number: int | str | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        self.append_note(text, kind="footNote", number=number, paragraph_index=paragraph_index)

    def append_endnote(
        self,
        text: str,
        *,
        number: int | str | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        self.append_note(text, kind="endNote", number=number, paragraph_index=paragraph_index)

    def append_auto_number(
        self,
        *,
        number: int | str = 1,
        number_type: str = "PAGE",
        kind: str = "newNum",
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_auto_number(
            number=number,
            number_type=number_type,
            kind=kind,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_header(self, text: str, *, apply_page_type: str = "BOTH") -> "HwpHeaderFooterObject | None":
        return self.document.append_header(text, apply_page_type=apply_page_type, section_index=self.index)

    def append_footer(self, text: str, *, apply_page_type: str = "BOTH") -> "HwpHeaderFooterObject | None":
        return self.document.append_footer(text, apply_page_type=apply_page_type, section_index=self.index)

    def append_equation(
        self,
        script: str,
        *,
        width: int = 4800,
        height: int = 2300,
        font: str = DEFAULT_HWP_EQUATION_FONT,
        shape_comment: str | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_equation(
            script,
            width=width,
            height=height,
            font=font,
            shape_comment=shape_comment,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_shape(
        self,
        *,
        kind: str = "rect",
        text: str = "",
        width: int = 12000,
        height: int = 3200,
        fill_color: str = "#FFFFFF",
        line_color: str = "#000000",
        shape_comment: str | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_shape(
            kind=kind,
            text=text,
            width=width,
            height=height,
            fill_color=fill_color,
            line_color=line_color,
            shape_comment=shape_comment,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_ole(
        self,
        name: str,
        data: bytes,
        *,
        width: int = 42001,
        height: int = 13501,
        shape_comment: str | None = None,
        object_type: str = "EMBEDDED",
        draw_aspect: str = "CONTENT",
        has_moniker: bool = False,
        eq_baseline: int = 0,
        line_color: str = "#000000",
        line_width: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_ole(
            name,
            data,
            width=width,
            height=height,
            shape_comment=shape_comment,
            object_type=object_type,
            draw_aspect=draw_aspect,
            has_moniker=has_moniker,
            eq_baseline=eq_baseline,
            line_color=line_color,
            line_width=line_width,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )


class HwpParagraphObject:
    def __init__(self, document: HwpDocument, paragraph: HwpParagraph):
        self._document = document
        self._section_index = paragraph.section_index
        self._index = paragraph.index

    @property
    def document(self) -> HwpDocument:
        return self._document

    @property
    def index(self) -> int:
        return self._index

    @property
    def section_index(self) -> int:
        return self._section_index

    def _snapshot(self) -> HwpParagraph:
        return self._document.binary_document().paragraphs(self._section_index)[self._index]

    @property
    def text(self) -> str:
        return self._snapshot().text

    @property
    def raw_text(self) -> str:
        return self._snapshot().raw_text

    @property
    def char_count(self) -> int:
        return self._snapshot().char_count

    @property
    def control_mask(self) -> int:
        return self._snapshot().control_mask

    @property
    def para_shape_id(self) -> int:
        return self._snapshot().para_shape_id

    @property
    def style_id(self) -> int:
        return self._snapshot().style_id

    @property
    def split_flags(self) -> int:
        return self._snapshot().split_flags

    @property
    def has_hidden_controls(self) -> bool:
        snapshot = self._snapshot()
        return snapshot.raw_text != snapshot.text

    def set_text_same_length(self, value: str) -> None:
        self._document.binary_document().set_paragraph_text_same_length(self._section_index, self._index, value)

    def replace_text_same_length(self, old: str, new: str, *, count: int = -1) -> int:
        return self._document.binary_document().replace_paragraph_text_same_length(
            self._section_index,
            self._index,
            old,
            new,
            count=count,
        )

    def set_text(self, value: str) -> None:
        model = self._document.section_model(self._section_index)
        paragraph = model.paragraphs()[self._index]
        text_record = paragraph.text_record()
        if text_record is None:
            raise ValueError("HWP paragraph does not contain a para text record.")
        text_record.set_raw_text(value)
        paragraph.header.char_count = len(value)
        self._document.binary_document().replace_section_model(self._section_index, model)
        self._document._invalidate_bridge()

    def replace_text(self, old: str, new: str, *, count: int = -1) -> int:
        model = self._document.section_model(self._section_index)
        paragraph = model.paragraphs()[self._index]
        text_record = paragraph.text_record()
        if text_record is None or not old:
            return 0
        raw_text = text_record.raw_text
        replaced = raw_text.count(old) if count < 0 else min(raw_text.count(old), count)
        if replaced == 0:
            return 0
        updated = raw_text.replace(old, new) if count < 0 else raw_text.replace(old, new, count)
        text_record.set_raw_text(updated)
        paragraph.header.char_count = len(updated)
        self._document.binary_document().replace_section_model(self._section_index, model)
        self._document._invalidate_bridge()
        return replaced


class HwpControlObject:
    def __init__(
        self,
        document: HwpDocument,
        paragraph: SectionParagraphModel,
        control_node: RecordNode,
        control_ordinal: int,
    ):
        self._document = document
        self._paragraph = paragraph
        self._control_ordinal = control_ordinal
        self._initial_control_node = control_node
        self._document_revision = getattr(document, "_mutation_revision", 0)

    @property
    def document(self) -> HwpDocument:
        return self._document

    @property
    def section_index(self) -> int:
        return self._paragraph.section_index

    @property
    def paragraph_index(self) -> int:
        return self._paragraph.index

    @property
    def control_node(self) -> RecordNode:
        if self._document_revision == getattr(self._document, "_mutation_revision", 0):
            return self._initial_control_node
        return self._live_context()[2]

    @property
    def control_id(self) -> str:
        payload = self.control_node.payload
        if len(payload) < 4:
            return ""
        return payload[:4][::-1].decode("latin1", errors="replace")

    def descendant_nodes(self) -> list[RecordNode]:
        return list(self.control_node.iter_descendants())

    def _live_context(self) -> tuple[SectionModel, SectionParagraphModel, RecordNode]:
        section_model = self._document.section_model(self.section_index)
        paragraph = section_model.paragraphs()[self.paragraph_index]
        controls = [child for child in paragraph.header.children if child.tag_id == TAG_CTRL_HEADER]
        control_node = controls[self._control_ordinal]
        return section_model, paragraph, control_node

    def _commit(self, section_model: SectionModel) -> None:
        self._document.binary_document().replace_section_model(self.section_index, section_model)
        self._document._invalidate_bridge()

    def _native_shape_metadata(self, control_node: RecordNode | None = None) -> dict[str, str]:
        target = self.control_node if control_node is None else control_node
        return _collect_shape_native_metadata(target)

    def _native_control_metadata(self, control_node: RecordNode | None = None) -> dict[str, str]:
        target = self.control_node if control_node is None else control_node
        return _collect_control_native_metadata(target)

    def _graphic_out_margins(self, control_node: RecordNode | None = None) -> dict[str, int]:
        target = self.control_node if control_node is None else control_node
        native = _parse_graphic_control_out_margins(target.payload)
        return native or _metadata_prefixed_int_map(_collect_shape_native_metadata(target), "outMargin.")

    def _graphic_layout(self, control_node: RecordNode | None = None) -> dict[str, str]:
        target = self.control_node if control_node is None else control_node
        native = _parse_graphic_control_layout(target.payload)
        metadata = _metadata_prefixed_str_map(_collect_shape_native_metadata(target), "layout.")
        if not native:
            return metadata
        merged = dict(metadata)
        merged.update(native)
        return merged

    def _set_graphic_out_margins(
        self,
        *,
        left: int | str | None = None,
        right: int | str | None = None,
        top: int | str | None = None,
        bottom: int | str | None = None,
    ) -> None:
        section_model, _paragraph, control_node = self._live_context()
        control_node.payload = _set_graphic_control_out_margins_payload(
            control_node.payload,
            left=left,
            right=right,
            top=top,
            bottom=bottom,
        )
        _update_shape_native_metadata(control_node, updates={}, remove_prefixes=("outMargin.",))
        self._commit(section_model)

    def _set_graphic_layout(
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
        section_model, _paragraph, control_node = self._live_context()
        current_layout = self._graphic_layout(control_node)
        native_text_wrap = text_wrap if text_wrap in {"TOP_AND_BOTTOM", "SQUARE"} else None
        effective_text_wrap = native_text_wrap or current_layout.get("textWrap")
        native_text_flow = text_flow if text_flow in {"BOTH_SIDES", "LEFT_ONLY", "RIGHT_ONLY"} and effective_text_wrap == "SQUARE" else None
        native_treat_as_char = treat_as_char
        native_vert_rel_to = vert_rel_to if vert_rel_to in {"PARA", "PAPER"} else None
        native_horz_rel_to = horz_rel_to if horz_rel_to in {"COLUMN", "PAPER"} else None
        native_vert_align = vert_align if vert_align in {"TOP", "CENTER"} else None
        native_horz_align = horz_align if horz_align in {"LEFT", "RIGHT"} else None
        control_node.payload = _set_graphic_control_layout_payload(
            control_node.payload,
            text_wrap=native_text_wrap,
            text_flow=native_text_flow,
            treat_as_char=native_treat_as_char,
            vert_rel_to=native_vert_rel_to,
            horz_rel_to=native_horz_rel_to,
            vert_align=native_vert_align,
            horz_align=native_horz_align,
            vert_offset=vert_offset,
            horz_offset=horz_offset,
        )
        metadata_updates = _graphic_layout_metadata_updates(
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
        remove_prefixes: list[str] = []
        for key, native_value in (
            ("layout.textWrap", native_text_wrap),
            ("layout.textFlow", native_text_flow),
            ("layout.treatAsChar", native_treat_as_char),
            ("layout.vertRelTo", native_vert_rel_to),
            ("layout.horzRelTo", native_horz_rel_to),
            ("layout.vertAlign", vert_align if vert_align == "CENTER" else None),
            (
                "layout.horzAlign",
                horz_align if horz_align == "RIGHT" and effective_text_wrap != "SQUARE" else None,
            ),
            ("layout.vertOffset", vert_offset),
            ("layout.horzOffset", horz_offset),
        ):
            if native_value is not None:
                metadata_updates.pop(key, None)
                remove_prefixes.append(key)
        _update_shape_native_metadata(control_node, updates=metadata_updates, remove_prefixes=tuple(remove_prefixes))
        self._commit(section_model)

    def _write_native_shape_metadata(
        self,
        *,
        updates: dict[str, object | None],
        remove_prefixes: Sequence[str] = (),
    ) -> None:
        section_model, _paragraph, control_node = self._live_context()
        _update_shape_native_metadata(control_node, updates=updates, remove_prefixes=remove_prefixes)
        self._commit(section_model)

    def _write_native_control_metadata(
        self,
        *,
        updates: dict[str, object | None],
        remove_prefixes: Sequence[str] = (),
    ) -> None:
        section_model, _paragraph, control_node = self._live_context()
        _update_control_native_metadata(control_node, updates=updates, remove_prefixes=remove_prefixes)
        self._commit(section_model)


from ._hwp_table_objects import HwpTableCellObject, HwpTableObject  # noqa: E402, F401


class HwpPictureObject(HwpControlObject):
    @property
    def shape_comment(self) -> str:
        return _parse_object_description(self.control_node.payload)

    @property
    def storage_id(self) -> int:
        payload = self._shape_picture_record().payload
        return int.from_bytes(payload[71:73], "little")

    def bindata_path(self) -> str:
        prefix = f"BinData/BIN{self.storage_id:04d}."
        for stream_path in self.document.bindata_stream_paths():
            if stream_path.startswith(prefix):
                return stream_path
        raise KeyError(f"Could not find BinData stream for storage id {self.storage_id}.")

    @property
    def extension(self) -> str | None:
        suffix = Path(self.bindata_path()).suffix
        return suffix.lstrip(".") or None

    def size(self) -> dict[str, int]:
        shape_component = self._shape_component_record()
        if shape_component is None or len(shape_component.payload) < 28:
            return {"width": 0, "height": 0}
        return {
            "width": int.from_bytes(shape_component.payload[20:24], "little", signed=False),
            "height": int.from_bytes(shape_component.payload[24:28], "little", signed=False),
        }

    @property
    def line_color(self) -> str:
        shape_component = self._shape_component_record()
        if shape_component is not None and len(shape_component.payload) >= 202:
            component_color = shape_component.payload[196:200]
            component_width = int.from_bytes(shape_component.payload[200:202], "little", signed=False)
            if any(component_color) or component_width:
                return _parse_colorref(component_color)
        picture_record = self._shape_picture_record()
        if len(picture_record.payload) >= 8 and any(picture_record.payload[0:8]):
            return _parse_rgb_color(picture_record.payload[0:4])
        if shape_component is None or len(shape_component.payload) < 200:
            return "#000000"
        return _parse_colorref(shape_component.payload[196:200])

    @property
    def line_width(self) -> int:
        shape_component = self._shape_component_record()
        if shape_component is not None and len(shape_component.payload) >= 202:
            component_color = shape_component.payload[196:200]
            component_width = int.from_bytes(shape_component.payload[200:202], "little", signed=False)
            if any(component_color) or component_width:
                return component_width
        picture_record = self._shape_picture_record()
        if len(picture_record.payload) >= 8 and any(picture_record.payload[0:8]):
            return int.from_bytes(picture_record.payload[4:8], "little", signed=False)
        if shape_component is None or len(shape_component.payload) < 202:
            return 0
        return int.from_bytes(shape_component.payload[200:202], "little", signed=False)

    def layout(self) -> dict[str, str]:
        return self._graphic_layout()

    def out_margins(self) -> dict[str, int]:
        return self._graphic_out_margins()

    def rotation(self) -> dict[str, str]:
        shape_component = self._shape_component_record()
        native = {} if shape_component is None else _parse_shape_component_rotation_payload(shape_component.payload, include_center=True)
        metadata = _metadata_prefixed_str_map(self._native_shape_metadata(), "rotation.")
        if not native:
            return metadata
        merged = dict(metadata)
        merged.update(native)
        return merged

    def image_adjustment(self) -> dict[str, str]:
        native = _parse_picture_image_adjustment_payload(self._shape_picture_record().payload)
        return native or _metadata_prefixed_str_map(self._native_shape_metadata(), "image.")

    def crop(self) -> dict[str, int]:
        native = _parse_picture_crop_payload(self._shape_picture_record().payload)
        return native or _metadata_prefixed_int_map(self._native_shape_metadata(), "crop.")

    def line_style(self) -> dict[str, str]:
        result = {
            "color": self.line_color,
            "width": str(self.line_width),
        }
        result.update(_metadata_prefixed_str_map(self._native_shape_metadata(), "lineStyle."))
        return result

    def binary_data(self) -> bytes:
        return self.document.binary_document().read_stream(self.bindata_path(), decompress=False)

    def replace_binary(self, data: bytes, *, extension: str | None = None) -> None:
        if extension is None or extension.lower().lstrip(".") == (self.extension or "").lower():
            self.document.binary_document().add_stream(self.bindata_path(), data)
        else:
            section_model, _paragraph, control_node = self._live_context()
            storage_id, _ = self.document.add_embedded_bindata(data, extension=extension)
            payload = bytearray(self._shape_picture_record(control_node).payload)
            payload[71:73] = int(storage_id).to_bytes(2, "little")
            self._shape_picture_record(control_node).payload = bytes(payload)
            self._commit(section_model)
            return
        self.document._invalidate_bridge()

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
            raise ValueError("HWP picture control does not contain a shape component record.")
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

    def set_image_adjustment(
        self,
        *,
        bright: int | str | None = None,
        contrast: int | str | None = None,
        effect: str | None = None,
        alpha: int | str | None = None,
    ) -> None:
        section_model, _paragraph, control_node = self._live_context()
        effect_code = _picture_effect_code(effect)
        picture_record = self._shape_picture_record(control_node)
        picture_record.payload = _set_picture_image_adjustment_payload(
            picture_record.payload,
            bright=bright,
            contrast=contrast,
            effect_code=effect_code,
            alpha=alpha,
        )
        if effect is not None and effect_code is None:
            _update_shape_native_metadata(
                control_node,
                updates=_image_adjustment_metadata_updates(
                    bright=bright,
                    contrast=contrast,
                    effect=effect,
                    alpha=alpha,
                ),
            )
        else:
            _update_shape_native_metadata(control_node, updates={}, remove_prefixes=("image.",))
        self._commit(section_model)

    def set_crop(
        self,
        *,
        left: int | str | None = None,
        right: int | str | None = None,
        top: int | str | None = None,
        bottom: int | str | None = None,
    ) -> None:
        section_model, _paragraph, control_node = self._live_context()
        picture_record = self._shape_picture_record(control_node)
        picture_record.payload = _set_picture_crop_payload(
            picture_record.payload,
            left=left,
            right=right,
            top=top,
            bottom=bottom,
        )
        _update_shape_native_metadata(control_node, updates={}, remove_prefixes=("crop.",))
        self._commit(section_model)

    def set_size(self, *, width: int | None = None, height: int | None = None) -> None:
        section_model, _paragraph, control_node = self._live_context()
        control_node.payload = _set_payload_size(control_node.payload, width=width, height=height, width_offset=16, height_offset=20)
        record = self._shape_component_record(control_node)
        if record is None:
            raise ValueError("HWP picture control does not contain a shape component record.")
        record.payload = _set_payload_size(record.payload, width=width, height=height, width_offset=20, height_offset=24)
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
        record = self._shape_component_record(control_node)
        if record is None:
            raise ValueError("HWP picture control does not contain a shape component record.")
        picture_record = self._shape_picture_record(control_node)
        payload = bytearray(record.payload)
        if len(payload) < 202:
            payload.extend(b"\x00" * (202 - len(payload)))
        if color is not None:
            payload[196:200] = _build_colorref(color)
        if width is not None:
            payload[200:202] = max(0, min(int(width), 0xFFFF)).to_bytes(2, "little", signed=False)
        record.payload = bytes(payload)
        picture_record.payload = _set_picture_line_payload(
            picture_record.payload,
            color=color,
            width=width,
        )
        _update_shape_native_metadata(
            control_node,
            updates={
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

    def _shape_picture_record(self, control_node: RecordNode | None = None) -> RecordNode:
        target = self.control_node if control_node is None else control_node
        for node in target.iter_descendants():
            if node.tag_id == TAG_SHAPE_COMPONENT_PICTURE:
                return node
        raise ValueError("HWP picture control does not contain a shape picture record.")

    def _shape_component_record(self, control_node: RecordNode | None = None) -> RecordNode | None:
        target = self.control_node if control_node is None else control_node
        for node in target.iter_descendants():
            if node.tag_id == TAG_SHAPE_COMPONENT:
                return node
        return None


class HwpBookmarkObject(HwpControlObject):
    @property
    def name(self) -> str:
        data_node = next((child for child in self.control_node.children if child.tag_id == TAG_CTRL_DATA), None)
        if data_node is None or len(data_node.payload) < 12:
            return ""
        name_size = int.from_bytes(data_node.payload[10:12], "little", signed=False)
        encoded_name = data_node.payload[12 : 12 + name_size * 2]
        return encoded_name.decode("utf-16-le", errors="ignore")

    def rename(self, value: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        data_node = next((child for child in control_node.children if child.tag_id == TAG_CTRL_DATA), None)
        if data_node is None:
            raise ValueError("HWP bookmark control does not contain ctrl data.")
        prefix = data_node.payload[:10]
        payload = bytearray(prefix)
        payload.extend(len(value).to_bytes(2, "little", signed=False))
        payload.extend(value.encode("utf-16-le"))
        data_node.payload = bytes(payload)
        self._commit(section_model)


class HwpFieldObject(HwpControlObject):
    @property
    def native_field_type(self) -> str:
        return self.control_id

    @property
    def field_type(self) -> str:
        return _semantic_field_type_from_native(self.native_field_type, self.parameters)

    @property
    def is_mail_merge(self) -> bool:
        return self.field_type == "MAILMERGE"

    @property
    def is_calculation(self) -> bool:
        return self.field_type == "FORMULA"

    @property
    def is_cross_reference(self) -> bool:
        return self.field_type == "CROSSREF"

    @property
    def is_doc_property(self) -> bool:
        return self.field_type == "DOCPROPERTY"

    @property
    def is_date(self) -> bool:
        return self.field_type == "DATE"

    @property
    def parameters(self) -> dict[str, str]:
        data_node = next((child for child in self.control_node.children if child.tag_id == TAG_CTRL_DATA), None)
        if data_node is None:
            return {}
        decoded = data_node.payload.decode("utf-16-le", errors="ignore")
        parameters: dict[str, str] = {}
        for entry in decoded.split(";"):
            key, separator, value = entry.partition("=")
            if separator and key:
                parameters[key] = value
        return parameters

    def get_parameter(self, name: str) -> str | None:
        return self.parameters.get(name)

    @property
    def name(self) -> str | None:
        for key in ("FieldName", "Name", "name"):
            value = self.parameters.get(key)
            if value:
                return value
        return None

    @property
    def merge_field_name(self) -> str | None:
        return self.get_parameter("MergeField") or self.get_parameter("FieldName") or self.name

    @property
    def expression(self) -> str | None:
        return self.get_parameter("Expression") or self.get_parameter("Command")

    @property
    def bookmark_name(self) -> str | None:
        return self.get_parameter("BookmarkName") or self.get_parameter("Path") or self.name

    @property
    def document_property_name(self) -> str | None:
        if not self.is_doc_property:
            return None
        return self.get_parameter("FieldName") or self.name

    @property
    def display_text(self) -> str:
        return self._paragraph.text.replace("\r", "")

    @property
    def editable(self) -> bool:
        return bool(self._paragraph.header.control_mask & 0x01)

    @property
    def dirty(self) -> bool:
        return bool(self._paragraph.header.control_mask & 0x01)

    def _write_parameters(self, section_model: SectionModel, control_node: RecordNode, parameters: dict[str, str]) -> None:
        data_node = next((child for child in control_node.children if child.tag_id == TAG_CTRL_DATA), None)
        if data_node is None:
            data_node = RecordNode(tag_id=TAG_CTRL_DATA, level=control_node.level + 1, payload=b"")
            control_node.add_child(data_node)
        parameter_text = ";".join(f"{key}={value}" for key, value in parameters.items())
        data_node.payload = parameter_text.encode("utf-16-le")
        self._commit(section_model)

    def set_parameter(self, name: str, value: str | int) -> None:
        section_model, _paragraph, control_node = self._live_context()
        parameters = dict(self.parameters)
        parameters[str(name)] = str(value)
        self._write_parameters(section_model, control_node, parameters)

    def set_name(self, value: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        parameters = dict(self.parameters)
        updated = False
        for key in ("FieldName", "Name", "name"):
            if key in parameters:
                parameters[key] = value
                updated = True
        if self.is_mail_merge:
            parameters["MergeField"] = value
        if self.is_cross_reference:
            if "BookmarkName" in parameters or not updated:
                parameters["BookmarkName"] = value
            if "Path" in parameters or not updated:
                parameters["Path"] = value
        if not updated and not self.is_cross_reference:
            parameters["FieldName"] = value
        self._write_parameters(section_model, control_node, parameters)

    def set_field_type(self, value: str) -> None:
        normalized = value.strip()
        if normalized.upper() in {"HYPERLINK", "%HLK"}:
            raise ValueError("Use HwpHyperlinkObject to edit hyperlink controls.")
        section_model, _paragraph, control_node = self._live_context()
        control_node.payload = _build_control_id_payload(_normalize_field_control_id(normalized))
        self._commit(section_model)

    def set_display_text(self, text: str) -> None:
        section_model, paragraph, _control_node = self._live_context()
        text_record = paragraph.text_record()
        if text_record is None:
            raise ValueError("HWP field paragraph does not contain a para text record.")
        raw_text = f"{text}\r" if text else ""
        text_record.set_raw_text(raw_text)
        paragraph.header.char_count = len(raw_text)
        self._commit(section_model)

    def configure_mail_merge(self, field_name: str, *, display_text: str | None = None) -> None:
        self.set_field_type("MAILMERGE")
        self.set_name(field_name)
        self.set_parameter("FieldName", field_name)
        self.set_parameter("MergeField", field_name)
        if display_text is not None:
            self.set_display_text(display_text)

    def configure_calculation(self, expression: str, *, display_text: str | None = None) -> None:
        self.set_field_type("FORMULA")
        self.set_parameter("Expression", expression)
        self.set_parameter("Command", expression)
        if display_text is not None:
            self.set_display_text(display_text)

    def configure_cross_reference(self, bookmark_name: str, *, display_text: str | None = None) -> None:
        self.set_field_type("CROSSREF")
        self.set_name(bookmark_name)
        self.set_parameter("BookmarkName", bookmark_name)
        self.set_parameter("Path", bookmark_name)
        if display_text is not None:
            self.set_display_text(display_text)

    def configure_doc_property(self, property_name: str, *, display_text: str | None = None) -> None:
        self.set_field_type("DOCPROPERTY")
        self.set_name(property_name)
        self.set_parameter("FieldName", property_name)
        if display_text is not None:
            self.set_display_text(display_text)

    def configure_date(self, *, display_text: str | None = None) -> None:
        self.set_field_type("DATE")
        if display_text is not None:
            self.set_display_text(display_text)


class HwpHyperlinkObject(HwpFieldObject):
    @property
    def command(self) -> str:
        payload = self.control_node.payload
        prefix_size = 9
        text_size = int.from_bytes(payload[prefix_size : prefix_size + 2], "little")
        start = prefix_size + 2
        end = start + text_size * 2
        return payload[start:end].decode("utf-16-le", errors="ignore")

    @property
    def url(self) -> str:
        command = self.command
        return command.split(";", 1)[0].replace("\\:", ":")

    @property
    def metadata_fields(self) -> list[str]:
        parts = self.command.split(";")
        if parts and parts[-1] == "":
            parts = parts[:-1]
        return parts[1:]

    @property
    def display_text(self) -> str:
        return self._paragraph.text.replace("\r", "")

    def set_url(self, url: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        control_node.payload = _build_hyperlink_ctrl_payload(url, metadata_fields=self.metadata_fields)
        self._commit(section_model)

    def set_metadata_fields(self, metadata_fields: Sequence[str | int]) -> None:
        section_model, _paragraph, control_node = self._live_context()
        control_node.payload = _build_hyperlink_ctrl_payload(self.url, metadata_fields=metadata_fields)
        self._commit(section_model)

    def set_display_text(self, text: str) -> None:
        section_model, paragraph, _control_node = self._live_context()
        text_record = paragraph.text_record()
        if text_record is None:
            raise ValueError("HWP hyperlink paragraph does not contain a para text record.")
        raw_text = _build_hyperlink_para_text_payload(text).decode("utf-16-le", errors="ignore")
        text_record.set_raw_text(raw_text)
        paragraph.header.char_count = len(raw_text)
        self._commit(section_model)


class HwpAutoNumberObject(HwpControlObject):
    @property
    def kind(self) -> str:
        return "newNum" if self.control_id == "nwno" else "autoNum"

    @property
    def number(self) -> str | None:
        payload = self.control_node.payload
        if self.control_id == "nwno" and len(payload) >= 10:
            return str(int.from_bytes(payload[8:10], "little"))
        if self.control_id == "atno":
            return "1"
        return None

    @property
    def number_type(self) -> str:
        if self.control_id == "atno" and len(self.control_node.payload) >= 12:
            code = int.from_bytes(self.control_node.payload[8:12], "little", signed=False)
            return _AUTO_NUMBER_TYPE_VALUES.get(code, "PAGE")
        return "PAGE"

    def set_number(self, value: int | str) -> None:
        if self.control_id != "nwno":
            raise ValueError("Only newNum controls support explicit numeric rewrites in the native HWP wrapper.")
        section_model, _paragraph, control_node = self._live_context()
        payload = bytearray(control_node.payload)
        payload[8:10] = int(value).to_bytes(2, "little", signed=False)
        control_node.payload = bytes(payload)
        self._commit(section_model)

    def set_number_type(self, value: str) -> None:
        normalized = str(value).upper()
        if self.control_id != "atno":
            if normalized != "PAGE":
                raise ValueError("Only autoNum controls support editable number_type in the native HWP wrapper.")
            return
        section_model, _paragraph, control_node = self._live_context()
        control_node.payload = _build_auto_number_payload("autoNum", number=1, number_type=normalized)
        self._commit(section_model)


class HwpHeaderFooterObject(HwpControlObject):
    @property
    def kind(self) -> str:
        return "header" if self.control_id == "head" else "footer"

    @property
    def text(self) -> str:
        return _collect_text_from_node(self.control_node).replace("\r", "").strip()

    @property
    def apply_page_type(self) -> str:
        payload = self.control_node.payload
        if len(payload) < 8:
            return "BOTH"
        code = int.from_bytes(payload[4:8], "little", signed=False)
        return {0: "BOTH", 1: "EVEN", 2: "ODD"}.get(code, "BOTH")

    def set_text(self, value: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        paragraph_headers = [node for node in control_node.children if isinstance(node, ParagraphHeaderRecord)]
        if not paragraph_headers:
            raise ValueError("HWP header/footer control does not contain a writable paragraph header.")
        target = paragraph_headers[0]
        text_record = next((child for child in target.children if isinstance(child, ParagraphTextRecord)), None)
        raw_text = f"{value}\r" if value else ""
        if text_record is None:
            text_record = ParagraphTextRecord(level=target.level + 1, raw_text=raw_text)
            target.children.insert(0, text_record)
        else:
            text_record.set_raw_text(raw_text)
        target.char_count = (target.char_count & 0x80000000) | len(raw_text)
        target.sync_payload()
        self._commit(section_model)

    def set_apply_page_type(self, value: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        instance_id = int.from_bytes(control_node.payload[8:12].ljust(4, b"\x00"), "little") or 1
        control_node.payload = _build_header_footer_control_payload(
            self.control_id,
            apply_page_type=str(value).upper(),
            instance_id=instance_id,
        )
        self._commit(section_model)

    def set_nested_blocks(self, blocks: Sequence[object]) -> None:
        from .hancom_document import AutoNumber as HancomAutoNumber
        from .hancom_document import Bookmark as HancomBookmark
        from .hancom_document import Field as HancomField
        from .hancom_document import Hyperlink as HancomHyperlink
        from .hancom_document import Paragraph as HancomParagraph

        section_model, _paragraph, control_node = self._live_context()
        list_header = next((child.clone() for child in control_node.children if child.tag_id == TAG_LIST_HEADER), None)
        template = next((child for child in control_node.children if isinstance(child, ParagraphHeaderRecord)), None)
        if template is None:
            raise ValueError("HWP header/footer control does not contain a writable paragraph header.")

        grouped: dict[int, list[object]] = {}
        for block in blocks:
            grouped.setdefault(int(getattr(block, "nested_paragraph_index", 0)), []).append(block)
        if not grouped:
            grouped[0] = [HancomParagraph(text="")]

        new_children: list[RecordNode] = []
        if list_header is not None:
            new_children.append(list_header)

        for paragraph_index in sorted(grouped):
            paragraph_blocks = grouped[paragraph_index]
            paragraph_header = ParagraphHeaderRecord(
                level=template.level,
                char_count=template.char_count & 0x80000000,
                control_mask=template.control_mask,
                para_shape_id=template.para_shape_id,
                style_id=template.style_id,
                split_flags=template.split_flags,
                header_size=template.header_size,
                offset=template.offset,
                trailing_payload=template.trailing_payload,
            )
            raw_text_parts: list[str] = []
            control_children: list[RecordNode] = []
            hyperlink_only_payload: str | None = None

            for block in paragraph_blocks:
                if isinstance(block, HancomParagraph):
                    if block.text:
                        raw_text_parts.append(block.text)
                    continue
                if isinstance(block, HancomBookmark):
                    raw_text_parts.append(_build_control_marker_text("bokm"))
                    control_node_child = RecordNode(tag_id=TAG_CTRL_HEADER, level=template.level + 1, payload=_build_control_id_payload("bokm"))
                    control_node_child.add_child(
                        RecordNode(
                            tag_id=TAG_CTRL_DATA,
                            level=template.level + 2,
                            payload=_build_bookmark_ctrl_data_payload(block.name),
                        )
                    )
                    control_children.append(control_node_child)
                    continue
                if isinstance(block, HancomAutoNumber):
                    control_id = "nwno" if str(block.kind).strip() == "newNum" else "atno"
                    raw_text_parts.append(_build_control_marker_text(control_id))
                    control_children.append(
                        RecordNode(
                            tag_id=TAG_CTRL_HEADER,
                            level=template.level + 1,
                            payload=_build_auto_number_payload(
                                str(block.kind).strip(),
                                number=block.number,
                                number_type=block.number_type,
                            ),
                        )
                    )
                    continue
                if isinstance(block, HancomHyperlink):
                    display_text = block.display_text or block.target
                    if len(paragraph_blocks) == 1:
                        hyperlink_only_payload = _build_hyperlink_para_text_payload(display_text).decode("utf-16-le", errors="ignore")
                    else:
                        raw_text_parts.append(display_text)
                    control_children.append(
                        RecordNode(
                            tag_id=TAG_CTRL_HEADER,
                            level=template.level + 1,
                            payload=_build_hyperlink_ctrl_payload(block.target, metadata_fields=block.metadata_fields),
                        )
                    )
                    continue
                if isinstance(block, HancomField):
                    display_text = block.display_text or block.name or block.effective_native_field_type
                    raw_text_parts.append(display_text)
                    field_node = RecordNode(
                        tag_id=TAG_CTRL_HEADER,
                        level=template.level + 1,
                        payload=_build_control_id_payload(_normalize_field_control_id(block.effective_native_field_type)),
                    )
                    parameters = dict(block.parameters)
                    if block.name and not any(key in parameters for key in ("FieldName", "Name", "name")):
                        parameters["FieldName"] = block.name
                    if parameters:
                        parameter_text = ";".join(f"{key}={value}" for key, value in parameters.items())
                        field_node.add_child(
                            RecordNode(
                                tag_id=TAG_CTRL_DATA,
                                level=template.level + 2,
                                payload=parameter_text.encode("utf-16-le"),
                            )
                        )
                    control_children.append(field_node)
                    continue
                if hasattr(block, "text") and getattr(block, "text", None):
                    raw_text_parts.append(str(getattr(block, "text")))

            raw_text = hyperlink_only_payload or (("".join(raw_text_parts)) + "\r" if raw_text_parts else "\r")
            paragraph_header.char_count = (paragraph_header.char_count & 0x80000000) | len(raw_text)
            paragraph_header.add_child(ParagraphTextRecord(level=template.level + 1, raw_text=raw_text))
            for child in template.children:
                if isinstance(child, ParagraphTextRecord) or child.tag_id == TAG_CTRL_HEADER:
                    continue
                paragraph_header.add_child(child.clone())
            for child in control_children:
                paragraph_header.add_child(child)
            paragraph_header.sync_payload()
            new_children.append(paragraph_header)

        control_node.children = new_children
        self._commit(section_model)


class HwpNoteObject(HwpControlObject):
    def _note_ordinal(self, numbering_type: str) -> int:
        count = 0
        target_control_id = self.control_id
        if numbering_type == "CONTINUOUS":
            section_indexes = range(self.section_index + 1)
        else:
            section_indexes = range(self.section_index, self.section_index + 1)
        for section_index in section_indexes:
            section_model = self.document.section_model(section_index)
            for paragraph in section_model.paragraphs():
                controls = [child for child in paragraph.header.children if child.tag_id == TAG_CTRL_HEADER]
                for ordinal, control in enumerate(controls):
                    payload = control.payload
                    if len(payload) < 4 or payload[:4][::-1].decode("latin1", errors="replace") != target_control_id:
                        continue
                    count += 1
                    if (
                        section_index == self.section_index
                        and paragraph.index == self.paragraph_index
                        and ordinal == self._control_ordinal
                    ):
                        return count
        return count

    @property
    def kind(self) -> str:
        return "footNote" if self.control_id == "fn  " else "endNote"

    @property
    def number(self) -> str | None:
        note_settings = self.document.binary_document().section_note_settings(self.section_index)
        settings_key = "footNotePr" if self.kind == "footNote" else "endNotePr"
        numbering = dict(note_settings.get(settings_key, {}).get("numbering", {}))
        numbering_type = str(numbering.get("type", "CONTINUOUS")).upper()
        if numbering_type == "ON_PAGE":
            return None
        try:
            start_number = int(str(numbering.get("newNum", "1")) or "1")
        except ValueError:
            start_number = 1
        ordinal = self._note_ordinal(numbering_type)
        if ordinal > 0:
            return str(start_number + ordinal - 1)
        return None

    @property
    def text(self) -> str:
        return _collect_text_from_node(self.control_node).replace("\r", "").strip()

    def set_text(self, value: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        paragraph_headers = [node for node in control_node.children if isinstance(node, ParagraphHeaderRecord)]
        if not paragraph_headers:
            raise ValueError("HWP note control does not contain a writable paragraph header.")
        target = paragraph_headers[0]
        text_record = next((child for child in target.children if isinstance(child, ParagraphTextRecord)), None)
        raw_text = f"{value}\r" if value else ""
        if text_record is None:
            text_record = ParagraphTextRecord(level=target.level + 1, raw_text=raw_text)
            target.children.insert(0, text_record)
        else:
            text_record.set_raw_text(raw_text)
        target.char_count = 0x80000000 | len(raw_text)
        target.sync_payload()
        self._commit(section_model)

    def set_number(self, value: int | str) -> None:
        note_settings = self.document.binary_document().section_note_settings(self.section_index)
        settings_key = "footNotePr" if self.kind == "footNote" else "endNotePr"
        note_pr = dict(note_settings.get(settings_key, {}))
        numbering = dict(note_pr.get("numbering", {}))
        numbering_type = str(numbering.get("type", "CONTINUOUS")).upper()
        if numbering_type == "ON_PAGE":
            raise ValueError("ON_PAGE note numbering cannot be rewritten by absolute note number.")
        ordinal = self._note_ordinal(numbering_type)
        if ordinal <= 0:
            raise ValueError("Failed to resolve the target native HWP note ordinal.")
        numbering["newNum"] = str(int(value) - ordinal + 1)
        note_pr["numbering"] = numbering
        if self.kind == "footNote":
            self.document.apply_section_note_settings(section_index=self.section_index, footnote_pr=note_pr)
        else:
            self.document.apply_section_note_settings(section_index=self.section_index, endnote_pr=note_pr)


class HwpPageNumObject(HwpControlObject):
    @property
    def page_number(self) -> dict[str, str]:
        payload = self.control_node.payload
        return {
            "pos": self.pos,
            "formatType": self.format_type,
            "sideChar": self.side_char if len(payload) >= 16 else "-",
        }

    @property
    def pos(self) -> str:
        payload = self.control_node.payload
        attributes = int.from_bytes(payload[4:8].ljust(4, b"\x00"), "little")
        positions = {
            0: "NONE",
            1: "TOP_LEFT",
            2: "TOP_CENTER",
            3: "TOP_RIGHT",
            4: "BOTTOM_LEFT",
            5: "BOTTOM_CENTER",
            6: "BOTTOM_RIGHT",
            7: "OUTSIDE_TOP",
            8: "OUTSIDE_BOTTOM",
            9: "INSIDE_TOP",
            10: "INSIDE_BOTTOM",
        }
        return positions.get((attributes >> 8) & 0x0F, "BOTTOM_CENTER")

    @property
    def format_type(self) -> str:
        payload = self.control_node.payload
        attributes = int.from_bytes(payload[4:8].ljust(4, b"\x00"), "little")
        formats = {
            0: "DIGIT",
            1: "CIRCLED_DIGIT",
            2: "ROMAN_CAPITAL",
            3: "ROMAN_SMALL",
            4: "LATIN_CAPITAL",
            5: "LATIN_SMALL",
        }
        return formats.get(attributes & 0xFF, "DIGIT")

    @property
    def side_char(self) -> str:
        payload = self.control_node.payload
        if len(payload) < 16:
            return "-"
        return payload[14:16].decode("utf-16-le", errors="ignore") or "-"

    def set_page_number(
        self,
        *,
        pos: str | None = None,
        format_type: str | None = None,
        side_char: str | None = None,
    ) -> None:
        section_model, _paragraph, control_node = self._live_context()
        page_number = dict(self.page_number)
        if pos is not None:
            page_number["pos"] = pos
        if format_type is not None:
            page_number["formatType"] = format_type
        if side_char is not None:
            page_number["sideChar"] = side_char
        control_node.payload = _build_page_num_payload(page_number)
        self._commit(section_model)

    def set_pos(self, value: str) -> None:
        self.set_page_number(pos=value)

    def set_format_type(self, value: str) -> None:
        self.set_page_number(format_type=value)

    def set_side_char(self, value: str) -> None:
        self.set_page_number(side_char=value)


class HwpEquationObject(HwpControlObject):
    @property
    def script(self) -> str:
        eqedit_record = self._eqedit_record()
        if eqedit_record is None:
            return ""
        decoded = eqedit_record.payload[4:].decode("utf-16-le", errors="ignore")
        script = decoded.split("Equation Version", 1)[0].replace("\x00", "").strip()
        if "ь" in script:
            script = script.split("ь", 1)[0]
        script = "".join(character for character in script if ord(character) >= 32 or character in "\t\n")
        if script.startswith("* "):
            script = script[2:]
        return script.rstrip("`").strip()

    @property
    def shape_comment(self) -> str:
        return _parse_object_description(self.control_node.payload)

    def layout(self) -> dict[str, str]:
        return self._graphic_layout()

    def out_margins(self) -> dict[str, int]:
        return self._graphic_out_margins()

    def rotation(self) -> dict[str, str]:
        return _metadata_prefixed_str_map(self._native_shape_metadata(), "rotation.")

    def size(self) -> dict[str, int]:
        payload = self.control_node.payload
        if len(payload) < 24:
            return {"width": 0, "height": 0}
        return {
            "width": int.from_bytes(payload[16:20], "little", signed=False),
            "height": int.from_bytes(payload[20:24], "little", signed=False),
        }

    @property
    def font(self) -> str:
        eqedit_record = self._eqedit_record()
        if eqedit_record is None:
            return DEFAULT_HWP_EQUATION_FONT
        decoded = eqedit_record.payload[4:].decode("utf-16-le", errors="ignore")
        marker_index = decoded.find("Equation Version")
        if marker_index < 0:
            return DEFAULT_HWP_EQUATION_FONT
        suffix = decoded[marker_index:]
        if "\x07" not in suffix:
            return DEFAULT_HWP_EQUATION_FONT
        font = suffix.split("\x07", 1)[1].replace("\x00", "").strip()
        return font or DEFAULT_HWP_EQUATION_FONT

    def _eqedit_record(self, control_node: RecordNode | None = None) -> RecordNode | None:
        target = self.control_node if control_node is None else control_node
        for node in target.iter_descendants():
            if node.tag_id == TAG_EQEDIT:
                return node
        return None

    def set_script(self, script: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        eqedit_record = self._eqedit_record(control_node)
        if eqedit_record is None:
            raise ValueError("HWP equation control does not contain an eqedit record.")
        decoded = eqedit_record.payload[4:].decode("utf-16-le", errors="ignore")
        prefix = "* " if decoded.startswith("* ") else ""
        marker_index = decoded.find("Equation Version")
        suffix = decoded[marker_index:] if marker_index >= 0 else ""
        eqedit_record.payload = eqedit_record.payload[:4] + (prefix + script + suffix).encode("utf-16-le")
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
        self._write_native_shape_metadata(
            updates=_graphic_rotation_metadata_updates(
                angle=angle,
                center_x=center_x,
                center_y=center_y,
                rotate_image=rotate_image,
            )
        )

    def set_size(self, *, width: int | None = None, height: int | None = None) -> None:
        section_model, _paragraph, control_node = self._live_context()
        payload = bytearray(control_node.payload)
        if len(payload) < 24:
            raise ValueError("HWP equation control payload is shorter than expected.")
        if width is not None:
            payload[16:20] = int(width).to_bytes(4, "little", signed=False)
        if height is not None:
            payload[20:24] = int(height).to_bytes(4, "little", signed=False)
        control_node.payload = bytes(payload)
        self._commit(section_model)


from ._hwp_shape_objects import (  # noqa: E402, F401
    HwpArcShapeObject,
    HwpConnectLineShapeObject,
    HwpContainerShapeObject,
    HwpCurveShapeObject,
    HwpEllipseShapeObject,
    HwpLineShapeObject,
    HwpPolygonShapeObject,
    HwpRectangleShapeObject,
    HwpShapeObject,
    HwpTextArtShapeObject,
)


class HwpChartObject(HwpShapeObject):
    @property
    def kind(self) -> str:
        return "chart"

    def fields(self) -> dict[str, object]:
        result = _parse_chart_data_payload(self._chart_record().payload)
        metadata = self._shape_native_metadata()
        title = metadata.get("chart.title") or str(result.get("utf16_text", ""))
        if title:
            result["title"] = title
        if "chart.chartType" in metadata:
            result["chart_type"] = metadata["chart.chartType"]
        if "chart.dataRef" in metadata:
            result["data_ref"] = metadata["chart.dataRef"]
        if "chart.legendVisible" in metadata:
            result["legend_visible"] = _metadata_bool_value(metadata, "chart.legendVisible", True)
        categories = [str(item) for item in _metadata_json_list(metadata, "chart.categories")]
        if categories:
            result["categories"] = categories
        series = _metadata_json_dict_list(metadata, "chart.series")
        if series:
            result["series"] = series
        return result

    @property
    def raw_payload(self) -> bytes:
        payload = _parse_chart_data_payload(self._chart_record().payload).get("raw_payload", b"")
        return bytes(payload) if isinstance(payload, (bytes, bytearray)) else b""

    @property
    def utf16_text(self) -> str:
        return str(_parse_chart_data_payload(self._chart_record().payload).get("utf16_text", ""))

    @property
    def title(self) -> str:
        metadata = self._shape_native_metadata()
        return metadata.get("chart.title", self.utf16_text)

    @property
    def chart_type(self) -> str:
        metadata = self._shape_native_metadata()
        return metadata.get("chart.chartType", "BAR")

    @property
    def data_ref(self) -> str | None:
        metadata = self._shape_native_metadata()
        value = metadata.get("chart.dataRef")
        return value or None

    @property
    def legend_visible(self) -> bool:
        return _metadata_bool_value(self._shape_native_metadata(), "chart.legendVisible", True)

    @property
    def categories(self) -> list[str]:
        return [str(item) for item in _metadata_json_list(self._shape_native_metadata(), "chart.categories")]

    @property
    def series(self) -> list[dict[str, object]]:
        return _metadata_json_dict_list(self._shape_native_metadata(), "chart.series")

    def set_raw_payload(self, payload: bytes) -> None:
        section_model, _paragraph, control_node = self._live_context()
        self._chart_record(control_node).payload = bytes(payload)
        self._commit(section_model)

    def set_title(self, value: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        self._chart_record(control_node).payload = str(value).encode("utf-16-le")
        _update_shape_native_metadata(control_node, updates={"chart.title": value})
        self._commit(section_model)

    def set_chart_type(self, value: str) -> None:
        self._write_native_shape_metadata(updates={"chart.chartType": str(value).upper()})

    def set_data_ref(self, value: str | None) -> None:
        self._write_native_shape_metadata(updates={"chart.dataRef": value})

    def set_legend_visible(self, value: bool) -> None:
        self._write_native_shape_metadata(updates={"chart.legendVisible": bool(value)})

    def set_categories(self, values: Sequence[str]) -> None:
        serialized = json.dumps([str(value) for value in values], ensure_ascii=True, separators=(",", ":"))
        self._write_native_shape_metadata(updates={"chart.categories": serialized})

    def set_series(self, values: Sequence[dict[str, object]]) -> None:
        serialized = json.dumps([dict(value) for value in values], ensure_ascii=True, separators=(",", ":"))
        self._write_native_shape_metadata(updates={"chart.series": serialized})

    def _chart_record(self, control_node: RecordNode | None = None) -> RecordNode:
        target = self.control_node if control_node is None else control_node
        for node in target.iter_descendants():
            if node.tag_id == TAG_CHART_DATA:
                return node
        raise ValueError("HWP chart control does not contain chart data.")


class HwpOleObject(HwpShapeObject):
    @property
    def kind(self) -> str:
        return "ole"

    @property
    def storage_id(self) -> int:
        payload = self._ole_record().payload
        if len(payload) >= 14:
            return int.from_bytes(payload[12:14], "little", signed=False)
        return 0

    @property
    def object_type(self) -> str:
        flags = self._ole_flags()
        return _HWP_OLE_OBJECT_TYPE_NAMES.get((flags >> 16) & 0x3F, "UNKNOWN")

    @property
    def draw_aspect(self) -> str:
        flags = self._ole_flags()
        return _HWP_OLE_DRAW_ASPECT_NAMES.get(flags & 0xFF, "CONTENT")

    @property
    def has_moniker(self) -> bool:
        return bool(self._ole_flags() & (1 << 8))

    @property
    def eq_baseline(self) -> int:
        return (self._ole_flags() >> 9) & 0x7F

    @property
    def line_color(self) -> str:
        payload = self._ole_record().payload
        if len(payload) < 18:
            return "#000000"
        return _parse_rgb_color(payload[14:18])

    @property
    def line_width(self) -> int:
        payload = self._ole_record().payload
        if len(payload) < 20:
            return 0
        return int.from_bytes(payload[18:20], "little", signed=False)

    def line_style(self) -> dict[str, str]:
        metadata = _metadata_prefixed_str_map(self._shape_native_metadata(), "lineStyle.")
        result = {
            "color": self.line_color,
            "width": str(self.line_width),
        }
        result.update(metadata)
        return result

    def rotation(self) -> dict[str, str]:
        component = self._shape_component_record()
        native = {} if component is None else _parse_shape_component_rotation_payload(component.payload, include_center=False)
        metadata = _metadata_prefixed_str_map(self._shape_native_metadata(), "rotation.")
        if not native:
            return metadata
        merged = dict(metadata)
        merged.update(native)
        return merged

    def extent(self) -> dict[str, int]:
        payload = self._ole_record().payload
        if len(payload) >= 12:
            extent = {
                "x": int.from_bytes(payload[4:8], "little", signed=True),
                "y": int.from_bytes(payload[8:12], "little", signed=True),
            }
            if extent["x"] or extent["y"]:
                display_size = super().size()
                if extent != {"x": display_size.get("width", 0), "y": display_size.get("height", 0)}:
                    return extent
        return _metadata_prefixed_int_map(self._shape_native_metadata(), "extent.")

    def size(self) -> dict[str, int]:
        return super().size()

    def bindata_path(self) -> str:
        storage_id = self.storage_id
        if storage_id > 0:
            prefix = f"BinData/BIN{storage_id:04d}."
            for stream_path in self.document.bindata_stream_paths():
                if stream_path.startswith(prefix):
                    return stream_path
        bindata_path = next((stream_path for stream_path in self.document.bindata_stream_paths() if stream_path.endswith(".ole")), None)
        if bindata_path is None:
            raise ValueError("No OLE BinData stream is available.")
        return bindata_path

    def binary_data(self) -> bytes:
        return self.document.binary_document().read_stream(self.bindata_path(), decompress=False)

    def replace_binary(self, data: bytes, *, extension: str | None = None) -> None:
        self.document.binary_document().add_stream(self.bindata_path(), data)
        self.document._invalidate_bridge()

    def set_size(self, *, width: int | None = None, height: int | None = None) -> None:
        section_model, _paragraph, control_node = self._live_context()
        control_node.payload = _set_payload_size(control_node.payload, width=width, height=height, width_offset=16, height_offset=20)
        ole_record = self._ole_record(control_node)
        ole_record.payload = _set_payload_size(ole_record.payload, width=width, height=height, width_offset=4, height_offset=8)
        component = self._shape_component_record(control_node)
        if component is not None:
            component.payload = _set_payload_size(component.payload, width=width, height=height, width_offset=20, height_offset=24)
        self._commit(section_model)

    def set_object_metadata(
        self,
        *,
        object_type: str | None = None,
        draw_aspect: str | None = None,
        has_moniker: bool | None = None,
        eq_baseline: int | None = None,
    ) -> None:
        section_model, _paragraph, control_node = self._live_context()
        record = self._ole_record(control_node)
        payload = bytearray(record.payload)
        resolved_flags = _build_ole_attribute_flags(
            object_type=self.object_type if object_type is None else object_type,
            draw_aspect=self.draw_aspect if draw_aspect is None else draw_aspect,
            has_moniker=self.has_moniker if has_moniker is None else bool(has_moniker),
            eq_baseline=self.eq_baseline if eq_baseline is None else int(eq_baseline),
        )
        payload[0:4] = int(resolved_flags).to_bytes(4, "little", signed=False)
        record.payload = bytes(payload)
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
        record = self._ole_record(control_node)
        payload = bytearray(record.payload)
        if color is not None:
            payload[14:18] = _build_rgb_color(color)
        if width is not None:
            payload[18:20] = max(0, min(int(width), 0xFFFF)).to_bytes(2, "little", signed=False)
        record.payload = bytes(payload)
        component = self._shape_component_record(control_node)
        if component is not None and len(component.payload) >= 202:
            component_payload = bytearray(component.payload)
            if color is not None:
                component_payload[196:200] = _build_colorref(color)
            if width is not None:
                component_payload[200:202] = max(0, min(int(width), 0xFFFF)).to_bytes(2, "little", signed=False)
            component.payload = bytes(component_payload)
        _update_shape_native_metadata(
            control_node,
            updates={
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

    def set_rotation(
        self,
        *,
        angle: int | str | None = None,
        center_x: int | str | None = None,
        center_y: int | str | None = None,
        rotate_image: bool | None = None,
    ) -> None:
        section_model, _paragraph, control_node = self._live_context()
        component = self._shape_component_record(control_node)
        if component is None:
            raise ValueError("HWP OLE control does not contain a shape component record.")
        component.payload = _set_shape_component_rotation_payload(component.payload, angle=angle)
        _update_shape_native_metadata(
            control_node,
            updates={
                "rotation.centerX": center_x,
                "rotation.centerY": center_y,
                "rotation.rotateimage": rotate_image,
            },
            remove_prefixes=("rotation.angle",),
        )
        self._commit(section_model)

    def set_extent(self, *, x: int | str | None = None, y: int | str | None = None) -> None:
        section_model, _paragraph, control_node = self._live_context()
        record = self._ole_record(control_node)
        payload = bytearray(record.payload)
        if len(payload) < 12:
            payload.extend(b"\x00" * (12 - len(payload)))
        current_x = int.from_bytes(payload[4:8], "little", signed=True)
        current_y = int.from_bytes(payload[8:12], "little", signed=True)
        payload[4:8] = int(current_x if x is None else int(x)).to_bytes(4, "little", signed=True)
        payload[8:12] = int(current_y if y is None else int(y)).to_bytes(4, "little", signed=True)
        record.payload = bytes(payload)
        _update_shape_native_metadata(control_node, updates={}, remove_prefixes=("extent.",))
        self._commit(section_model)

    def _ole_record(self, control_node: RecordNode | None = None) -> RecordNode:
        target = self.control_node if control_node is None else control_node
        for node in target.iter_descendants():
            if node.tag_id == TAG_SHAPE_COMPONENT_OLE:
                return node
        raise ValueError("HWP OLE control does not contain an OLE component record.")

    def _ole_flags(self) -> int:
        payload = self._ole_record().payload
        if len(payload) < 4:
            return 0
        return int.from_bytes(payload[0:4], "little", signed=False)


class HwpFormObject(HwpControlObject):
    @property
    def kind(self) -> str:
        return "form"

    def fields(self) -> dict[str, object]:
        result = _parse_form_object_payload(self._form_record().payload)
        metadata = self._native_control_metadata()
        label = metadata.get("form.label", str(result.get("utf16_text", "")))
        if label:
            result["label"] = label
        result["form_type"] = metadata.get("form.type", "INPUT")
        if metadata.get("form.name"):
            result["name"] = metadata["form.name"]
        result["value"] = metadata.get("form.value", label)
        result["checked"] = _metadata_bool_value(metadata, "form.checked", False)
        result["items"] = [str(item) for item in _metadata_json_list(metadata, "form.items")]
        result["editable"] = _metadata_bool_value(metadata, "form.editable", True)
        result["locked"] = _metadata_bool_value(metadata, "form.locked", False)
        if metadata.get("form.placeholder"):
            result["placeholder"] = metadata["form.placeholder"]
        return result

    @property
    def raw_payload(self) -> bytes:
        payload = _parse_form_object_payload(self._form_record().payload).get("raw_payload", b"")
        return bytes(payload) if isinstance(payload, (bytes, bytearray)) else b""

    @property
    def utf16_text(self) -> str:
        return str(_parse_form_object_payload(self._form_record().payload).get("utf16_text", ""))

    @property
    def label(self) -> str:
        metadata = self._native_control_metadata()
        return metadata.get("form.label", self.utf16_text)

    @property
    def text(self) -> str:
        return self.label

    @property
    def form_type(self) -> str:
        metadata = self._native_control_metadata()
        return metadata.get("form.type", "INPUT")

    @property
    def name(self) -> str | None:
        metadata = self._native_control_metadata()
        value = metadata.get("form.name")
        return value or None

    @property
    def value(self) -> str:
        metadata = self._native_control_metadata()
        return metadata.get("form.value", self.label)

    @property
    def checked(self) -> bool:
        return _metadata_bool_value(self._native_control_metadata(), "form.checked", False)

    @property
    def items(self) -> list[str]:
        return [str(item) for item in _metadata_json_list(self._native_control_metadata(), "form.items")]

    @property
    def editable(self) -> bool:
        return _metadata_bool_value(self._native_control_metadata(), "form.editable", True)

    @property
    def locked(self) -> bool:
        return _metadata_bool_value(self._native_control_metadata(), "form.locked", False)

    @property
    def placeholder(self) -> str | None:
        metadata = self._native_control_metadata()
        value = metadata.get("form.placeholder")
        return value or None

    def set_raw_payload(self, payload: bytes) -> None:
        section_model, _paragraph, control_node = self._live_context()
        self._form_record(control_node).payload = bytes(payload)
        self._commit(section_model)

    def set_label(self, value: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        self._form_record(control_node).payload = str(value).encode("utf-16-le")
        _update_control_native_metadata(control_node, updates={"form.label": value})
        self._commit(section_model)

    def set_text(self, value: str) -> None:
        self.set_label(value)

    def set_form_type(self, value: str) -> None:
        self._write_native_control_metadata(updates={"form.type": str(value).upper()})

    def set_name(self, value: str | None) -> None:
        self._write_native_control_metadata(updates={"form.name": value})

    def set_value(self, value: str | None) -> None:
        self._write_native_control_metadata(updates={"form.value": value})

    def set_checked(self, value: bool) -> None:
        self._write_native_control_metadata(updates={"form.checked": bool(value)})

    def set_items(self, values: Sequence[str]) -> None:
        serialized = json.dumps([str(value) for value in values], ensure_ascii=True, separators=(",", ":"))
        self._write_native_control_metadata(updates={"form.items": serialized})

    def set_editable(self, value: bool) -> None:
        self._write_native_control_metadata(updates={"form.editable": bool(value)})

    def set_locked(self, value: bool) -> None:
        self._write_native_control_metadata(updates={"form.locked": bool(value)})

    def set_placeholder(self, value: str | None) -> None:
        self._write_native_control_metadata(updates={"form.placeholder": value})

    def _form_record(self, control_node: RecordNode | None = None) -> RecordNode:
        target = self.control_node if control_node is None else control_node
        for node in target.iter_descendants():
            if node.tag_id == TAG_FORM_OBJECT:
                return node
        raise ValueError("HWP form control does not contain a form object record.")


class HwpMemoObject(HwpControlObject):
    @property
    def kind(self) -> str:
        return "memo"

    def fields(self) -> dict[str, object]:
        result = _parse_memo_list_payload(self._memo_record().payload)
        metadata = self._native_control_metadata()
        memo_text = str(result.get("utf16_text", "")).strip()
        result["text"] = memo_text or _collect_text_from_node(self.control_node).replace("\r", "")
        if metadata.get("memo.author"):
            result["author"] = metadata["memo.author"]
        if metadata.get("memo.memoId"):
            result["memo_id"] = metadata["memo.memoId"]
        if metadata.get("memo.anchorId"):
            result["anchor_id"] = metadata["memo.anchorId"]
        if metadata.get("memo.order"):
            try:
                result["order"] = int(metadata["memo.order"])
            except ValueError:
                pass
        result["visible"] = _metadata_bool_value(metadata, "memo.visible", True)
        return result

    @property
    def raw_payload(self) -> bytes:
        payload = _parse_memo_list_payload(self._memo_record().payload).get("raw_payload", b"")
        return bytes(payload) if isinstance(payload, (bytes, bytearray)) else b""

    @property
    def utf16_text(self) -> str:
        return str(_parse_memo_list_payload(self._memo_record().payload).get("utf16_text", ""))

    @property
    def text(self) -> str:
        memo_text = self.utf16_text.strip()
        return memo_text or _collect_text_from_node(self.control_node).replace("\r", "")

    @property
    def author(self) -> str | None:
        metadata = self._native_control_metadata()
        value = metadata.get("memo.author")
        return value or None

    @property
    def memo_id(self) -> str | None:
        metadata = self._native_control_metadata()
        value = metadata.get("memo.memoId")
        return value or None

    @property
    def anchor_id(self) -> str | None:
        metadata = self._native_control_metadata()
        value = metadata.get("memo.anchorId")
        return value or None

    @property
    def order(self) -> int | None:
        metadata = self._native_control_metadata()
        value = metadata.get("memo.order")
        if value is None:
            return None
        try:
            return int(value)
        except ValueError:
            return None

    @property
    def visible(self) -> bool:
        return _metadata_bool_value(self._native_control_metadata(), "memo.visible", True)

    def set_raw_payload(self, payload: bytes) -> None:
        section_model, _paragraph, control_node = self._live_context()
        self._memo_record(control_node).payload = bytes(payload)
        self._commit(section_model)

    def set_text(self, value: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        self._memo_record(control_node).payload = str(value).encode("utf-16-le")
        self._commit(section_model)

    def set_author(self, value: str | None) -> None:
        self._write_native_control_metadata(updates={"memo.author": value})

    def set_memo_id(self, value: str | None) -> None:
        self._write_native_control_metadata(updates={"memo.memoId": value})

    def set_anchor_id(self, value: str | None) -> None:
        self._write_native_control_metadata(updates={"memo.anchorId": value})

    def set_order(self, value: int | None) -> None:
        self._write_native_control_metadata(updates={"memo.order": value})

    def set_visible(self, value: bool) -> None:
        self._write_native_control_metadata(updates={"memo.visible": bool(value)})

    def _memo_record(self, control_node: RecordNode | None = None) -> RecordNode:
        target = self.control_node if control_node is None else control_node
        for node in target.iter_descendants():
            if node.tag_id == TAG_MEMO_LIST:
                return node
        raise ValueError("HWP memo control does not contain a memo list record.")


def _build_control_wrapper(
    document: HwpDocument,
    section_model: SectionModel,
    paragraph: SectionParagraphModel,
    control_node: RecordNode,
    control_ordinal: int,
) -> HwpControlObject | None:
    control_id = _control_id(control_node)
    if _control_has_descendant(control_node, TAG_FORM_OBJECT):
        return HwpFormObject(document, paragraph, control_node, control_ordinal)
    if _control_has_descendant(control_node, TAG_MEMO_LIST):
        return HwpMemoObject(document, paragraph, control_node, control_ordinal)
    if control_id == "tbl ":
        return HwpTableObject(document, paragraph, control_node, control_ordinal)
    if control_id == "bokm":
        return HwpBookmarkObject(document, paragraph, control_node, control_ordinal)
    if control_id in {"fn  ", "en  "}:
        return HwpNoteObject(document, paragraph, control_node, control_ordinal)
    if control_id == "%hlk":
        return HwpHyperlinkObject(document, paragraph, control_node, control_ordinal)
    if control_id.startswith("%"):
        return HwpFieldObject(document, paragraph, control_node, control_ordinal)
    if control_id in {"atno", "nwno"}:
        return HwpAutoNumberObject(document, paragraph, control_node, control_ordinal)
    if control_id in {"head", "foot"}:
        return HwpHeaderFooterObject(document, paragraph, control_node, control_ordinal)
    if control_id == "pgnp":
        return HwpPageNumObject(document, paragraph, control_node, control_ordinal)
    if control_id == "eqed":
        return HwpEquationObject(document, paragraph, control_node, control_ordinal)
    if control_id == "gso " and _control_has_descendant(control_node, TAG_SHAPE_COMPONENT_OLE):
        return HwpOleObject(document, paragraph, control_node, control_ordinal)
    if control_id == "gso " and any(node.tag_id == TAG_SHAPE_COMPONENT_PICTURE for node in control_node.iter_descendants()):
        return HwpPictureObject(document, paragraph, control_node, control_ordinal)
    if control_id == "gso " and _control_has_descendant(control_node, TAG_CHART_DATA):
        return HwpChartObject(document, paragraph, control_node, control_ordinal)
    if control_id == "gso ":
        wrapper_class = _shape_wrapper_class(control_node)
        return wrapper_class(document, paragraph, control_node, control_ordinal)
    return None


def _control_id(control_node: RecordNode) -> str:
    payload = control_node.payload
    if len(payload) < 4:
        return ""
    return payload[:4][::-1].decode("latin1", errors="replace")


def _collect_text_from_node(node: RecordNode | None) -> str:
    if node is None:
        return ""
    texts: list[str] = []
    for descendant in node.iter_descendants():
        if isinstance(descendant, ParagraphTextRecord):
            texts.append(descendant.text.replace("\r", ""))
    return "".join(texts)


def _control_has_descendant(control_node: RecordNode, tag_id: int) -> bool:
    return any(node.tag_id == tag_id for node in control_node.iter_descendants())


def _semantic_field_type_from_native(native_field_type: str, parameters: dict[str, str]) -> str:
    normalized = native_field_type.strip().upper()
    normalized_parameters = {str(key).upper(): str(value) for key, value in parameters.items()}
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


def _shape_kind_from_control_node(control_node: RecordNode) -> str:
    for data_node in control_node.children:
        if data_node.tag_id != TAG_CTRL_DATA:
            continue
        metadata = _parse_shape_native_metadata(data_node.payload)
        kind = metadata.get("kind")
        if kind:
            return kind
    if _has_hancom_connectline_signature(control_node):
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


def _shape_wrapper_class(control_node: RecordNode) -> type[HwpShapeObject]:
    shape_kind = _shape_kind_from_control_node(control_node)
    return {
        "line": HwpLineShapeObject,
        "connectLine": HwpConnectLineShapeObject,
        "rect": HwpRectangleShapeObject,
        "ellipse": HwpEllipseShapeObject,
        "arc": HwpArcShapeObject,
        "polygon": HwpPolygonShapeObject,
        "curve": HwpCurveShapeObject,
        "container": HwpContainerShapeObject,
        "textart": HwpTextArtShapeObject,
    }.get(shape_kind, HwpShapeObject)


# Keep established public type paths stable for repr, pickle, and introspection.
for _public_type in (
    HwpSection,
    HwpParagraphObject,
    HwpControlObject,
    HwpTableCellObject,
    HwpTableObject,
    HwpPictureObject,
    HwpBookmarkObject,
    HwpFieldObject,
    HwpHyperlinkObject,
    HwpAutoNumberObject,
    HwpHeaderFooterObject,
    HwpNoteObject,
    HwpPageNumObject,
    HwpEquationObject,
    HwpShapeObject,
    HwpLineShapeObject,
    HwpConnectLineShapeObject,
    HwpRectangleShapeObject,
    HwpEllipseShapeObject,
    HwpArcShapeObject,
    HwpPolygonShapeObject,
    HwpCurveShapeObject,
    HwpContainerShapeObject,
    HwpTextArtShapeObject,
    HwpChartObject,
    HwpOleObject,
    HwpFormObject,
    HwpMemoObject,
):
    _public_type.__module__ = "jakal_hwpx.hwp_document"
del _public_type

