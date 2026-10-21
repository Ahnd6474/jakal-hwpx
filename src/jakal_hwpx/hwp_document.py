from __future__ import annotations

import tempfile
from collections.abc import Sequence
from pathlib import Path
from typing import Callable

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
from .document import HwpxDocument
from .exceptions import HwpxValidationError, ValidationIssue
from .hwp_binary import (  # noqa: F401 - compatibility re-exports
    _AUTO_NUMBER_TYPE_VALUES,
    _TABLE_CELL_CHAR_SHAPE,
    _TABLE_CELL_LINE_SEG,
    _TABLE_CELL_PARA_HEADER_TAIL,
    DEFAULT_HWP_EQUATION_FONT,
    DEFAULT_HWP_TABLE_CELL_HEIGHT,
    DEFAULT_HWP_TABLE_CELL_WIDTH,
    TAG_CHART_DATA,
    TAG_CTRL_DATA,
    TAG_CTRL_HEADER,
    TAG_EQEDIT,
    TAG_FORM_OBJECT,
    TAG_LIST_HEADER,
    TAG_MEMO_LIST,
    TAG_PARA_CHAR_SHAPE,
    TAG_PARA_HEADER,
    TAG_PARA_LINE_SEG,
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
    TAG_TABLE,
    DocInfoModel,
    HwpBinaryDocument,
    HwpBinaryFileHeader,
    HwpDocumentProperties,
    HwpParagraph,
    HwpStreamCapacity,
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
    _build_table_cell_list_header_payload,
    _build_table_record_payload,
    _normalize_field_control_id,
    _parse_arc_shape_payload,
    _parse_chart_data_payload,
    _parse_colorref,
    _parse_container_shape_payload,
    _parse_curve_shape_payload,
    _parse_ellipse_shape_payload,
    _parse_form_object_payload,
    _parse_graphic_control_layout,
    _parse_graphic_control_out_margins,
    _parse_line_shape_payload,
    _parse_memo_list_payload,
    _parse_picture_crop_payload,
    _parse_picture_image_adjustment_payload,
    _parse_polygon_shape_payload,
    _parse_rectangle_shape_payload,
    _parse_rgb_color,
    _parse_shape_component_rotation_payload,
    _picture_effect_code,
    _set_graphic_control_layout_payload,
    _set_graphic_control_out_margins_payload,
    _set_picture_crop_payload,
    _set_picture_image_adjustment_payload,
    _set_picture_line_payload,
    _set_shape_component_rotation_payload,
    _shape_common_payload_size,
    _TableCellSpec,
)
from .hwp_pure_profile import HwpPureProfile, append_feature_from_profile

HancomConverter = Callable[[str | Path, str | Path, str], Path]

_HWP_STRICT_LINT_HINTS = {
    "binary_closure": "Keep DocInfo BinData records and BinData/BINxxxx streams in sync. Re-add the picture/OLE data or remove the broken control.",
    "binary_media_type": "Use replace_binary(..., extension=...) or rebuild the BinData entry so the DocInfo extension matches the stream name.",
    "binary_reference": "Re-add the missing embedded binary or update the picture/OLE control to point at an existing BinData id.",
    "control_subtree": "Recreate the object with the public append/setter API, or restore the missing child record before saving.",
    "docinfo_mapping": "Repair DocInfo id mappings before saving. This usually means recreating the referenced memo/BinData entry through the HWP API.",
    "docinfo_reference": "Repair the paragraph header or restore the missing DocInfo record before saving.",
    "metadata_schema": "Use the semantic setter API instead of editing JAKAL metadata text directly; values must keep their documented type.",
    "payload_sanity": "The native HWP payload is truncated or malformed. Recreate that control or replace it from a known-good source document.",
}


def _issue(
    kind: str,
    message: str,
    *,
    section_index: int | None = None,
    paragraph_index: int | None = None,
    context: str | None = None,
    hint: str | None = None,
) -> ValidationIssue:
    return ValidationIssue(
        kind=kind,
        message=message,
        section_index=section_index,
        paragraph_index=paragraph_index,
        context=context,
        hint=hint or _HWP_STRICT_LINT_HINTS.get(kind),
    )


class HwpDocument:
    def __init__(
        self,
        binary_document: HwpBinaryDocument,
        *,
        converter: HancomConverter | None = None,
        default_append_section_index: int = 0,
    ):
        self._binary_document = binary_document
        self._converter = converter
        self._default_append_section_index = default_append_section_index
        self._bridge_document: HwpxDocument | None = None
        self._hancom_document = None
        self._bridge_temp_dir: Path | None = None
        self._pure_profile: HwpPureProfile | None = None
        self._mutation_revision = 0

    @classmethod
    def blank(cls, *, converter: HancomConverter | None = None) -> "HwpDocument":
        instance = cls.blank_from_profile(HwpPureProfile.load_bundled().root, converter=converter)
        return instance

    @classmethod
    def open(cls, path: str | Path, *, converter: HancomConverter | None = None) -> "HwpDocument":
        return cls(
            HwpBinaryDocument.open(path),
            converter=converter,
            default_append_section_index=HwpPureProfile.load_bundled().target_section_index,
        )

    @classmethod
    def blank_from_profile(
        cls,
        profile_root: str | Path,
        *,
        converter: HancomConverter | None = None,
    ) -> "HwpDocument":
        profile = HwpPureProfile.load(profile_root)
        instance = cls(
            HwpBinaryDocument.open(profile.base_path),
            converter=converter,
            default_append_section_index=profile.target_section_index,
        )
        instance._pure_profile = profile
        instance._binary_document.set_preview_text("")
        return instance

    @property
    def source_path(self) -> Path:
        return self._binary_document.source_path

    def binary_document(self) -> HwpBinaryDocument:
        return self._binary_document

    def file_header(self) -> HwpBinaryFileHeader:
        return self._binary_document.file_header()

    def document_properties(self) -> HwpDocumentProperties:
        return self._binary_document.document_properties()

    def docinfo_model(self) -> DocInfoModel:
        return self._binary_document.docinfo_model()

    def list_stream_paths(self) -> list[str]:
        return self._binary_document.list_stream_paths()

    def bindata_stream_paths(self) -> list[str]:
        return self._binary_document.bindata_stream_paths()

    def stream_capacity(self, path: str) -> HwpStreamCapacity:
        data = self._binary_document.read_stream(path, decompress=self._binary_document._stream_is_compressed(path))
        return self._binary_document.stream_capacity(
            path,
            data,
            compress=self._binary_document._stream_is_compressed(path),
        )

    def preview_text(self) -> str:
        return self._binary_document.preview_text()

    def set_preview_text(self, value: str) -> None:
        self._invalidate_bridge()
        self._binary_document.set_preview_text(value)

    def sections(self) -> list["HwpSection"]:
        return [HwpSection(self, index) for index, _path in enumerate(self._binary_document.section_stream_paths())]

    def section(self, index: int) -> "HwpSection":
        return HwpSection(self, index)

    def section_model(self, section_index: int = 0) -> SectionModel:
        return self._binary_document.section_model(section_index)

    def ensure_section_count(self, section_count: int) -> None:
        self._invalidate_bridge()
        self._binary_document.ensure_section_count(section_count)

    def apply_section_settings(
        self,
        *,
        section_index: int = 0,
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
        self._invalidate_bridge()
        self._binary_document.set_section_page_settings(
            section_index,
            page_width=page_width,
            page_height=page_height,
            landscape=landscape,
            margins=margins,
        )
        self._binary_document.set_section_definition_settings(
            section_index,
            visibility=visibility,
            grid=grid,
            start_numbers=start_numbers,
            numbering_shape_id=int(numbering_shape_id) if numbering_shape_id is not None else None,
            memo_shape_id=memo_shape_id,
        )

    def apply_section_page_border_fills(
        self,
        page_border_fills: list[dict[str, str | int]],
        *,
        section_index: int = 0,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.set_section_page_border_fills(section_index, page_border_fills)

    def apply_section_page_numbers(
        self,
        page_numbers: list[dict[str, str]],
        *,
        section_index: int = 0,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.set_section_page_numbers(section_index, page_numbers)

    def apply_section_note_settings(
        self,
        *,
        section_index: int = 0,
        footnote_pr: dict[str, object] | None = None,
        endnote_pr: dict[str, object] | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.set_section_note_settings(
            section_index,
            footnote_pr=footnote_pr,
            endnote_pr=endnote_pr,
        )

    def paragraphs(self, section_index: int = 0) -> list["HwpParagraphObject"]:
        return self.section(section_index).paragraphs()

    def controls(self, section_index: int | None = None) -> list["HwpControlObject"]:
        sections = [self.section(section_index)] if section_index is not None else self.sections()
        controls: list[HwpControlObject] = []
        for section in sections:
            controls.extend(section.controls())
        return controls

    def _last_appended_control(
        self,
        section_index: int,
        paragraph_index: int | None,
        expected_type: type["HwpControlObject"],
    ) -> "HwpControlObject | None":
        model = self.section_model(section_index)
        if paragraph_index is None or paragraph_index < 0:
            header = next((root for root in reversed(model.roots) if isinstance(root, ParagraphHeaderRecord)), None)
            if header is None:
                return None
            resolved_index = next((item.index for item in model.paragraphs() if item.header is header), 0)
            paragraph = SectionParagraphModel(section_index=section_index, index=resolved_index, header=header)
        else:
            paragraphs = model.paragraphs()
            if paragraph_index >= len(paragraphs):
                return None
            paragraph = paragraphs[paragraph_index]
        controls = [child for child in paragraph.header.children if child.tag_id == TAG_CTRL_HEADER]
        if not controls:
            return None
        ordinal = len(controls) - 1
        wrapper = _build_control_wrapper(self, model, paragraph, controls[ordinal], ordinal)
        return wrapper if isinstance(wrapper, expected_type) else None

    def tables(self, section_index: int | None = None) -> list["HwpTableObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpTableObject)]

    def pictures(self, section_index: int | None = None) -> list["HwpPictureObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpPictureObject)]

    def hyperlinks(self, section_index: int | None = None) -> list["HwpHyperlinkObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpHyperlinkObject)]

    def bookmarks(self, section_index: int | None = None) -> list["HwpBookmarkObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpBookmarkObject)]

    def notes(self, section_index: int | None = None) -> list["HwpNoteObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpNoteObject)]

    def page_numbers(self, section_index: int | None = None) -> list["HwpPageNumObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpPageNumObject)]

    def fields(self, section_index: int | None = None) -> list["HwpFieldObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpFieldObject)]

    def mail_merge_fields(self, section_index: int | None = None) -> list["HwpFieldObject"]:
        return [field for field in self.fields(section_index) if field.is_mail_merge]

    def calculation_fields(self, section_index: int | None = None) -> list["HwpFieldObject"]:
        return [field for field in self.fields(section_index) if field.is_calculation]

    def cross_references(self, section_index: int | None = None) -> list["HwpFieldObject"]:
        return [field for field in self.fields(section_index) if field.is_cross_reference]

    def doc_property_fields(self, section_index: int | None = None) -> list["HwpFieldObject"]:
        return [field for field in self.fields(section_index) if field.is_doc_property]

    def date_fields(self, section_index: int | None = None) -> list["HwpFieldObject"]:
        return [field for field in self.fields(section_index) if field.is_date]

    def shapes(self, section_index: int | None = None) -> list["HwpShapeObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpShapeObject)]

    def lines(self, section_index: int | None = None) -> list["HwpLineShapeObject"]:
        return [
            shape
            for shape in self.shapes(section_index)
            if isinstance(shape, HwpLineShapeObject) and not isinstance(shape, HwpConnectLineShapeObject)
        ]

    def connect_lines(self, section_index: int | None = None) -> list["HwpConnectLineShapeObject"]:
        return [shape for shape in self.shapes(section_index) if isinstance(shape, HwpConnectLineShapeObject)]

    def rectangles(self, section_index: int | None = None) -> list["HwpRectangleShapeObject"]:
        return [shape for shape in self.shapes(section_index) if isinstance(shape, HwpRectangleShapeObject)]

    def ellipses(self, section_index: int | None = None) -> list["HwpEllipseShapeObject"]:
        return [shape for shape in self.shapes(section_index) if isinstance(shape, HwpEllipseShapeObject)]

    def arcs(self, section_index: int | None = None) -> list["HwpArcShapeObject"]:
        return [shape for shape in self.shapes(section_index) if isinstance(shape, HwpArcShapeObject)]

    def polygons(self, section_index: int | None = None) -> list["HwpPolygonShapeObject"]:
        return [shape for shape in self.shapes(section_index) if isinstance(shape, HwpPolygonShapeObject)]

    def curves(self, section_index: int | None = None) -> list["HwpCurveShapeObject"]:
        return [shape for shape in self.shapes(section_index) if isinstance(shape, HwpCurveShapeObject)]

    def containers(self, section_index: int | None = None) -> list["HwpContainerShapeObject"]:
        return [shape for shape in self.shapes(section_index) if isinstance(shape, HwpContainerShapeObject)]

    def textarts(self, section_index: int | None = None) -> list["HwpTextArtShapeObject"]:
        return [shape for shape in self.shapes(section_index) if isinstance(shape, HwpTextArtShapeObject)]

    def charts(self, section_index: int | None = None) -> list["HwpChartObject"]:
        return [shape for shape in self.shapes(section_index) if isinstance(shape, HwpChartObject)]

    def oles(self, section_index: int | None = None) -> list["HwpOleObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpOleObject)]

    def forms(self, section_index: int | None = None) -> list["HwpFormObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpFormObject)]

    def memos(self, section_index: int | None = None) -> list["HwpMemoObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpMemoObject)]

    def equations(self, section_index: int | None = None) -> list["HwpEquationObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpEquationObject)]

    def get_document_text(self) -> str:
        return self._binary_document.get_document_text()

    def replace_text_same_length(
        self,
        old: str,
        new: str,
        *,
        section_index: int | None = None,
        count: int = -1,
    ) -> int:
        return self._binary_document.replace_text_same_length(old, new, section_index=section_index, count=count)

    def append_paragraph(
        self,
        text: str,
        *,
        section_index: int = 0,
        para_shape_id: int = 0,
        style_id: int = 0,
        split_flags: int = 0,
        control_mask: int = 0,
    ) -> HwpParagraphObject:
        self._invalidate_bridge()
        paragraph = self._binary_document.append_paragraph(
            text,
            section_index=section_index,
            para_shape_id=para_shape_id,
            style_id=style_id,
            split_flags=split_flags,
            control_mask=control_mask,
        )
        return HwpParagraphObject(
            self,
            HwpParagraph(
                index=paragraph.index,
                section_index=paragraph.section_index,
                char_count=paragraph.header.char_count,
                control_mask=paragraph.header.control_mask,
                para_shape_id=paragraph.header.para_shape_id,
                style_id=paragraph.header.style_id,
                split_flags=paragraph.header.split_flags,
                raw_text=paragraph.raw_text,
                text=paragraph.text,
                text_record_offset=-1,
                text_record_size=len(paragraph.raw_text.encode("utf-16-le")),
            ),
        )

    def to_hwpx_document(self, *, force_refresh: bool = False) -> HwpxDocument:
        if self._bridge_document is not None and not force_refresh:
            return self._bridge_document

        self._bridge_document = self.to_hancom_document(force_refresh=force_refresh).to_hwpx_document()
        return self._bridge_document

    def to_hancom_document(self, *, force_refresh: bool = False):
        if self._hancom_document is not None and not force_refresh:
            return self._hancom_document

        from .hancom_document import HancomDocument

        self._hancom_document = HancomDocument.from_hwp_document(self, converter=self._converter)
        return self._hancom_document

    def save_as_hwpx(self, path: str | Path) -> Path:
        target_path = Path(path).expanduser().resolve()
        target_path.parent.mkdir(parents=True, exist_ok=True)
        self.to_hwpx_document(force_refresh=False).save(target_path)
        return target_path

    def save(self, path: str | Path | None = None) -> Path:
        if path is not None:
            target_path = Path(path).expanduser().resolve()
            if target_path.suffix.lower() == ".hwpx":
                return self.save_as_hwpx(target_path)
        else:
            target_path = self.source_path

        if self._bridge_document is None:
            return self._binary_document.save(path)

        self._sync_from_bridge_document()
        self._binary_document.save(target_path)
        self._binary_document.source_path = target_path
        return target_path

    def strict_lint_errors(self) -> list[ValidationIssue]:
        errors: list[ValidationIssue] = []
        docinfo = self.docinfo_model()
        docinfo_bindata: dict[int, object] = {}
        for record in docinfo.bin_data_records():
            storage_id = int(record.storage_id or 0)
            if storage_id <= 0:
                continue
            if storage_id in docinfo_bindata:
                errors.append(_issue("docinfo_mapping", f"Duplicate BinData docinfo storage id: {storage_id}", context="DocInfo"))
                continue
            docinfo_bindata[storage_id] = record

        stream_paths = self.bindata_stream_paths()
        bindata_streams: dict[int, str] = {}
        for stream_path in stream_paths:
            name = Path(stream_path).name
            if not name.startswith("BIN") or "." not in name:
                continue
            storage_token = name[3:].split(".", 1)[0]
            if not storage_token.isdigit():
                continue
            storage_id = int(storage_token)
            if storage_id in bindata_streams:
                errors.append(_issue("binary_closure", f"Duplicate BinData stream for storage id {storage_id}: {stream_path}", context="BinData"))
                continue
            bindata_streams[storage_id] = stream_path

        for storage_id, record in docinfo_bindata.items():
            stream_path = bindata_streams.get(storage_id)
            if stream_path is None:
                errors.append(_issue("binary_closure", f"DocInfo BinData id {storage_id} is missing a BinData stream.", context="DocInfo"))
                continue
            expected_extension = str(getattr(record, "extension", "") or "").lstrip(".").lower()
            actual_extension = Path(stream_path).suffix.lstrip(".").lower()
            if expected_extension and actual_extension and expected_extension != actual_extension:
                errors.append(
                    _issue(
                        "binary_media_type",
                        f"DocInfo BinData id {storage_id} expects .{expected_extension} but stream path is {stream_path}.",
                        context="BinData",
                    )
                )

        for storage_id, stream_path in bindata_streams.items():
            if storage_id not in docinfo_bindata:
                errors.append(
                    _issue(
                        "binary_closure",
                        f"BinData stream is not declared in DocInfo: {stream_path}",
                        context="BinData",
                    )
                )

        id_mappings = docinfo.id_mappings_record()
        count_specs = [
            (8, "border_fills", len(docinfo.border_fill_records())),
            (9, "char_shapes", len(docinfo.char_shape_records())),
            (10, "tab_defs", len(docinfo.tab_def_records())),
            (11, "numberings", len(docinfo.numbering_records())),
            (12, "bullets", len(docinfo.bullet_records())),
            (13, "para_shapes", len(docinfo.para_shape_records())),
            (14, "styles", len(docinfo.style_records())),
            (15, "memo_shapes", len(docinfo.memo_shape_records())),
            (16, "track_changes", len(docinfo.track_change_records())),
            (17, "track_change_authors", len(docinfo.track_change_author_records())),
        ]
        for count_index, count_name, actual_count in count_specs:
            declared_count = id_mappings.get_count(count_index)
            if declared_count < actual_count:
                errors.append(
                    _issue(
                        "docinfo_mapping",
                        f"DocInfo id_mappings.{count_name}={declared_count} is smaller than actual {count_name} record count {actual_count}.",
                        context="DocInfo",
                    )
                )

        max_bindata_id = max([0, *docinfo_bindata.keys(), *bindata_streams.keys()])
        if id_mappings.bin_data_count < max_bindata_id:
            errors.append(
                _issue(
                    "docinfo_mapping",
                    f"DocInfo id_mappings.bin_data_count={id_mappings.bin_data_count} is smaller than max BinData id {max_bindata_id}.",
                    context="DocInfo",
                )
            )

        memo_shape_records = docinfo.memo_shape_records()
        memo_shape_count = id_mappings.get_count(15)
        if memo_shape_count < len(memo_shape_records):
            errors.append(
                _issue(
                    "docinfo_mapping",
                    f"DocInfo id_mappings.memo_shapes={memo_shape_count} is smaller than actual memo shape record count {len(memo_shape_records)}.",
                    context="DocInfo",
                )
            )

        para_shape_count = len(docinfo.para_shape_records())
        style_count = len(docinfo.style_records())
        for section_index, section in enumerate(self.sections()):
            section_model = self.section_model(section_index)
            memo_shape_id = self.binary_document().section_definition_settings(section_index).get("memo_shape_id")
            if memo_shape_id not in (None, "", "0"):
                if not str(memo_shape_id).isdigit():
                    errors.append(
                        _issue(
                            "docinfo_mapping",
                            f"Section memo_shape_id must be numeric, got {memo_shape_id!r}.",
                            section_index=section_index,
                            context="SectionSettings",
                        )
                    )
                elif int(str(memo_shape_id)) >= len(memo_shape_records):
                    errors.append(
                        _issue(
                            "docinfo_mapping",
                            f"Section memo_shape_id {memo_shape_id} points to missing TAG_MEMO_SHAPE record.",
                            section_index=section_index,
                            context="SectionSettings",
                        )
                    )
            for paragraph in section_model.paragraphs():
                if paragraph.header.para_shape_id >= para_shape_count:
                    errors.append(
                        _issue(
                            "docinfo_reference",
                            f"Paragraph references missing para shape id {paragraph.header.para_shape_id}.",
                            section_index=section_index,
                            paragraph_index=paragraph.index,
                            context="ParagraphHeader",
                        )
                    )
                if paragraph.header.style_id >= style_count:
                    errors.append(
                        _issue(
                            "docinfo_reference",
                            f"Paragraph references missing style id {paragraph.header.style_id}.",
                            section_index=section_index,
                            paragraph_index=paragraph.index,
                            context="ParagraphHeader",
                        )
                    )
                for control_node in paragraph.control_nodes():
                    control_id = _control_id(control_node)
                    context = f"control[{control_id or 'UNKNOWN'}]"
                    if control_id == "gso ":
                        has_shape_component = _control_has_descendant(control_node, TAG_SHAPE_COMPONENT)
                        has_picture = _control_has_descendant(control_node, TAG_SHAPE_COMPONENT_PICTURE)
                        has_ole = _control_has_descendant(control_node, TAG_SHAPE_COMPONENT_OLE)
                        has_chart = _control_has_descendant(control_node, TAG_CHART_DATA)
                        shape_metadata = _collect_shape_native_metadata(control_node)
                        expects_chart_data = any(key.startswith("chart.") for key in shape_metadata)
                        if not (has_shape_component or has_chart):
                            errors.append(
                                _issue(
                                    "control_subtree",
                                    "Graphic control does not contain a shape component or chart data record.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        if expects_chart_data and not has_chart:
                            errors.append(
                                _issue(
                                    "control_subtree",
                                    "Chart control metadata is present but TAG_CHART_DATA is missing.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                    )
                                )
                        if not _metadata_bool_literal_is_valid(shape_metadata.get("chart.legendVisible")):
                            errors.append(
                                _issue(
                                    "metadata_schema",
                                    "chart.legendVisible metadata must be a bool-like literal.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        chart_categories, chart_categories_error = _metadata_json_value(shape_metadata, "chart.categories")
                        if chart_categories_error is not None:
                            errors.append(
                                _issue(
                                    "metadata_schema",
                                    f"chart.categories metadata must be valid JSON: {chart_categories_error}",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        elif chart_categories is not None and not isinstance(chart_categories, list):
                            errors.append(
                                _issue(
                                    "metadata_schema",
                                    "chart.categories metadata must decode to a JSON list.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        chart_series, chart_series_error = _metadata_json_value(shape_metadata, "chart.series")
                        if chart_series_error is not None:
                            errors.append(
                                _issue(
                                    "metadata_schema",
                                    f"chart.series metadata must be valid JSON: {chart_series_error}",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        elif chart_series is not None:
                            if not isinstance(chart_series, list):
                                errors.append(
                                    _issue(
                                        "metadata_schema",
                                        "chart.series metadata must decode to a JSON list.",
                                        section_index=section_index,
                                        paragraph_index=paragraph.index,
                                        context=context,
                                    )
                                )
                            elif not all(isinstance(item, dict) for item in chart_series):
                                errors.append(
                                    _issue(
                                        "metadata_schema",
                                        "chart.series metadata must decode to a JSON list of objects.",
                                        section_index=section_index,
                                        paragraph_index=paragraph.index,
                                        context=context,
                                    )
                                )
                        if has_chart:
                            chart_nodes = [node for node in control_node.iter_descendants() if node.tag_id == TAG_CHART_DATA]
                            for chart_node in chart_nodes:
                                if len(chart_node.payload) % 2:
                                    errors.append(
                                        _issue(
                                            "payload_sanity",
                                            "Chart payload is not UTF-16 aligned.",
                                            section_index=section_index,
                                            paragraph_index=paragraph.index,
                                            context=context,
                                        )
                                    )
                        if has_picture:
                            picture_nodes = [node for node in control_node.iter_descendants() if node.tag_id == TAG_SHAPE_COMPONENT_PICTURE]
                            for picture_node in picture_nodes:
                                if len(picture_node.payload) < 73:
                                    errors.append(
                                        _issue(
                                            "payload_sanity",
                                            "Picture payload is shorter than the expected 73-byte minimum.",
                                            section_index=section_index,
                                            paragraph_index=paragraph.index,
                                            context=context,
                                        )
                                    )
                                else:
                                    storage_id = int.from_bytes(picture_node.payload[71:73], "little", signed=False)
                                    if storage_id and storage_id not in docinfo_bindata:
                                        errors.append(
                                            _issue(
                                                "binary_reference",
                                                f"Picture control references missing BinData id {storage_id}.",
                                                section_index=section_index,
                                                paragraph_index=paragraph.index,
                                                context=context,
                                            )
                                        )
                        if has_ole:
                            ole_nodes = [node for node in control_node.iter_descendants() if node.tag_id == TAG_SHAPE_COMPONENT_OLE]
                            for ole_node in ole_nodes:
                                if len(ole_node.payload) < 20:
                                    errors.append(
                                        _issue(
                                            "payload_sanity",
                                            "OLE payload is shorter than the expected 20-byte minimum.",
                                            section_index=section_index,
                                            paragraph_index=paragraph.index,
                                            context=context,
                                        )
                                    )
                                storage_id = int.from_bytes(ole_node.payload[12:14].ljust(2, b"\x00"), "little", signed=False)
                                if storage_id and storage_id not in docinfo_bindata:
                                    errors.append(
                                        _issue(
                                            "binary_reference",
                                            f"OLE control references missing BinData id {storage_id}.",
                                            section_index=section_index,
                                            paragraph_index=paragraph.index,
                                            context=context,
                                            )
                                        )
                    elif control_id == "mrof":
                        metadata = _collect_control_native_metadata(control_node)
                        form_nodes = [node for node in control_node.iter_descendants() if node.tag_id == TAG_FORM_OBJECT]
                        if not form_nodes:
                            errors.append(
                                _issue(
                                    "control_subtree",
                                    "Form control does not contain a TAG_FORM_OBJECT record.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        for form_node in form_nodes:
                            if len(form_node.payload) % 2:
                                errors.append(
                                    _issue(
                                        "payload_sanity",
                                        "Form payload is not UTF-16 aligned.",
                                        section_index=section_index,
                                        paragraph_index=paragraph.index,
                                        context=context,
                                    )
                                )
                        if not _metadata_bool_literal_is_valid(metadata.get("form.checked")):
                            errors.append(
                                _issue(
                                    "metadata_schema",
                                    "form.checked metadata must be a bool-like literal.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        if not _metadata_bool_literal_is_valid(metadata.get("form.editable")):
                            errors.append(
                                _issue(
                                    "metadata_schema",
                                    "form.editable metadata must be a bool-like literal.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        if not _metadata_bool_literal_is_valid(metadata.get("form.locked")):
                            errors.append(
                                _issue(
                                    "metadata_schema",
                                    "form.locked metadata must be a bool-like literal.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        form_items, form_items_error = _metadata_json_value(metadata, "form.items")
                        if form_items_error is not None:
                            errors.append(
                                _issue(
                                    "metadata_schema",
                                    f"form.items metadata must be valid JSON: {form_items_error}",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        elif form_items is not None and not isinstance(form_items, list):
                            errors.append(
                                _issue(
                                    "metadata_schema",
                                    "form.items metadata must decode to a JSON list.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                    elif control_id == "omem":
                        metadata = _collect_control_native_metadata(control_node)
                        memo_nodes = [node for node in control_node.iter_descendants() if node.tag_id == TAG_MEMO_LIST]
                        if not memo_nodes:
                            errors.append(
                                _issue(
                                    "control_subtree",
                                    "Memo control does not contain a TAG_MEMO_LIST record.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        for memo_node in memo_nodes:
                            if len(memo_node.payload) % 2:
                                errors.append(
                                    _issue(
                                        "payload_sanity",
                                        "Memo payload is not UTF-16 aligned.",
                                        section_index=section_index,
                                        paragraph_index=paragraph.index,
                                        context=context,
                                    )
                                )
                        memo_order = metadata.get("memo.order")
                        if memo_order not in (None, ""):
                            try:
                                int(memo_order)
                            except ValueError:
                                errors.append(
                                    _issue(
                                        "metadata_schema",
                                        "memo.order metadata must be an integer.",
                                        section_index=section_index,
                                        paragraph_index=paragraph.index,
                                        context=context,
                                    )
                                )
                        if not _metadata_bool_literal_is_valid(metadata.get("memo.visible")):
                            errors.append(
                                _issue(
                                    "metadata_schema",
                                    "memo.visible metadata must be a bool-like literal.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                    elif control_id == "tbl ":
                        table_nodes = [node for node in control_node.iter_descendants() if node.tag_id == TAG_TABLE]
                        if not table_nodes:
                            errors.append(
                                _issue(
                                    "control_subtree",
                                    "Table control does not contain a TAG_TABLE record.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        for table_node in table_nodes:
                            if len(table_node.payload) < 20:
                                errors.append(
                                    _issue(
                                        "payload_sanity",
                                        "Table payload is shorter than the expected minimum.",
                                        section_index=section_index,
                                        paragraph_index=paragraph.index,
                                        context=context,
                                    )
                                )
                    elif control_id in {"head", "foot", "fn  ", "en  "}:
                        if not any(isinstance(node, ParagraphHeaderRecord) for node in control_node.children):
                            errors.append(
                                _issue(
                                    "control_subtree",
                                    "Nested control is missing a writable paragraph header.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        if control_id in {"head", "foot"} and len(control_node.payload) < 8:
                            errors.append(
                                _issue(
                                    "payload_sanity",
                                    "Header/footer payload is shorter than the expected 8-byte minimum.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                    elif control_id == "pgnp" and len(control_node.payload) < 8:
                        errors.append(
                            _issue(
                                "payload_sanity",
                                "Page number payload is shorter than the expected 8-byte minimum.",
                                section_index=section_index,
                                paragraph_index=paragraph.index,
                                context=context,
                            )
                        )
                    elif control_id.startswith("%") and control_id != "%hlk":
                        if not _control_has_descendant(control_node, TAG_CTRL_DATA):
                            errors.append(
                                _issue(
                                    "control_subtree",
                                    "Field control is missing TAG_CTRL_DATA parameters.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )

        return errors

    def strict_validate(self) -> None:
        errors = self.strict_lint_errors()
        if errors:
            raise HwpxValidationError(errors)

    def strict_lint_report(self, *, include_none: bool = False) -> list[dict[str, object]]:
        return [issue.to_dict(include_none=include_none) for issue in self.strict_lint_errors()]

    def format_strict_lint_errors(self) -> str:
        issues = self.strict_lint_errors()
        if not issues:
            return ""
        lines: list[str] = []
        for index, issue in enumerate(issues, start=1):
            location = issue.location() or "document"
            lines.append(f"{index}. [{issue.code or issue.kind}] {location}: {issue.message}")
            if issue.hint:
                lines.append(f"   Hint: {issue.hint}")
        return "\n".join(lines)

    def load_pure_profile(self, profile_root: str | Path) -> HwpPureProfile:
        self._pure_profile = HwpPureProfile.load(profile_root)
        self._default_append_section_index = self._pure_profile.target_section_index
        return self._pure_profile

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
        section_index: int | None = None,
    ) -> "HwpTableObject | None":
        self._invalidate_bridge()
        target_section_index = self._default_append_section_index if section_index is None else section_index
        self._binary_document.append_table(
            cell_text,
            rows=rows,
            cols=cols,
            cell_texts=cell_texts,
            row_heights=row_heights,
            col_widths=col_widths,
            cell_spans=cell_spans,
            cell_border_fill_ids=cell_border_fill_ids,
            table_border_fill_id=table_border_fill_id,
            profile_root=None,
            section_index=target_section_index,
        )
        return self._last_appended_control(target_section_index, None, HwpTableObject)  # type: ignore[return-value]

    def append_picture(
        self,
        image_bytes: bytes | None = None,
        *,
        extension: str | None = None,
        width: int = 12000,
        height: int = 3200,
        section_index: int | None = None,
    ) -> "HwpPictureObject | None":
        self._invalidate_bridge()
        target_section_index = self._default_append_section_index if section_index is None else section_index
        self._binary_document.append_picture(
            image_bytes,
            extension=extension,
            width=width,
            height=height,
            profile_root=None,
            section_index=target_section_index,
        )
        return self._last_appended_control(target_section_index, None, HwpPictureObject)  # type: ignore[return-value]

    def append_hyperlink(
        self,
        url: str | None = None,
        *,
        text: str | None = None,
        metadata_fields: Sequence[str | int] | None = None,
        section_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        target_section_index = self._default_append_section_index if section_index is None else section_index
        self._binary_document.append_hyperlink(
            url,
            text=text,
            metadata_fields=metadata_fields,
            profile_root=None,
            section_index=target_section_index,
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
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_field(
            field_type=field_type,
            display_text=display_text,
            name=name,
            parameters=parameters,
            editable=editable,
            dirty=dirty,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_mail_merge_field(
        self,
        field_name: str,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self.append_field(
            field_type="MAILMERGE",
            display_text=display_text,
            name=field_name,
            parameters={"FieldName": field_name, "MergeField": field_name},
            editable=editable,
            dirty=dirty,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_calculation_field(
        self,
        expression: str,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self.append_field(
            field_type="FORMULA",
            display_text=display_text,
            parameters={"Expression": expression, "Command": expression},
            editable=editable,
            dirty=dirty,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_cross_reference(
        self,
        bookmark_name: str,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self.append_field(
            field_type="CROSSREF",
            display_text=display_text,
            name=bookmark_name,
            parameters={"BookmarkName": bookmark_name, "Path": bookmark_name},
            editable=editable,
            dirty=dirty,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_doc_property_field(
        self,
        property_name: str,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self.append_field(
            field_type="DOCPROPERTY",
            display_text=display_text,
            name=property_name,
            parameters={"FieldName": property_name},
            editable=editable,
            dirty=dirty,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_date_field(
        self,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self.append_field(
            field_type="DATE",
            display_text=display_text,
            editable=editable,
            dirty=dirty,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_bookmark(
        self,
        name: str,
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_bookmark(name, section_index=section_index, paragraph_index=paragraph_index)

    def append_note(
        self,
        text: str,
        *,
        kind: str = "footNote",
        number: int | str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_note(
            text,
            kind=kind,
            number=number,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_footnote(
        self,
        text: str,
        *,
        number: int | str | None = None,
        section_index: int = 0,
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
        section_index: int = 0,
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
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_form(
            label,
            form_type=form_type,
            name=name,
            value=value,
            checked=checked,
            items=items,
            editable=editable,
            locked=locked,
            placeholder=placeholder,
            section_index=section_index,
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
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_memo(
            text,
            author=author,
            memo_id=memo_id,
            anchor_id=anchor_id,
            order=order,
            visible=visible,
            section_index=section_index,
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
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> "HwpChartObject | None":
        self._invalidate_bridge()
        self._binary_document.append_chart(
            title,
            chart_type=chart_type,
            categories=categories,
            series=series,
            data_ref=data_ref,
            legend_visible=legend_visible,
            width=width,
            height=height,
            shape_comment=shape_comment,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )
        return self._last_appended_control(section_index, paragraph_index, HwpChartObject)  # type: ignore[return-value]

    def append_auto_number(
        self,
        *,
        number: int | str = 1,
        number_type: str = "PAGE",
        kind: str = "newNum",
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_auto_number(
            number=number,
            number_type=number_type,
            kind=kind,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_header(
        self,
        text: str,
        *,
        apply_page_type: str = "BOTH",
        section_index: int = 0,
    ) -> "HwpHeaderFooterObject | None":
        self._invalidate_bridge()
        self._binary_document.append_header(text, apply_page_type=apply_page_type, section_index=section_index)
        controls = [
            control
            for control in self.section(section_index).controls()
            if isinstance(control, HwpHeaderFooterObject) and control.kind == "header"
        ]
        return controls[-1] if controls else None

    def append_footer(
        self,
        text: str,
        *,
        apply_page_type: str = "BOTH",
        section_index: int = 0,
    ) -> "HwpHeaderFooterObject | None":
        self._invalidate_bridge()
        self._binary_document.append_footer(text, apply_page_type=apply_page_type, section_index=section_index)
        controls = [
            control
            for control in self.section(section_index).controls()
            if isinstance(control, HwpHeaderFooterObject) and control.kind == "footer"
        ]
        return controls[-1] if controls else None

    def append_equation(
        self,
        script: str,
        *,
        width: int = 4800,
        height: int = 2300,
        font: str = DEFAULT_HWP_EQUATION_FONT,
        shape_comment: str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> "HwpEquationObject | None":
        self._invalidate_bridge()
        self._binary_document.append_equation(
            script,
            width=width,
            height=height,
            font=font,
            shape_comment=shape_comment,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )
        return self._last_appended_control(section_index, paragraph_index, HwpEquationObject)  # type: ignore[return-value]

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
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> "HwpShapeObject | None":
        self._invalidate_bridge()
        self._binary_document.append_shape(
            kind=kind,
            text=text,
            width=width,
            height=height,
            fill_color=fill_color,
            line_color=line_color,
            shape_comment=shape_comment,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )
        return self._last_appended_control(section_index, paragraph_index, HwpShapeObject)  # type: ignore[return-value]

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
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> "HwpOleObject | None":
        self._invalidate_bridge()
        self._binary_document.append_ole(
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
            section_index=section_index,
            paragraph_index=paragraph_index,
        )
        return self._last_appended_control(section_index, paragraph_index, HwpOleObject)  # type: ignore[return-value]

    def add_embedded_bindata(self, data: bytes, *, extension: str, storage_id: int | None = None) -> tuple[int, str]:
        self._invalidate_bridge()
        return self._binary_document.add_embedded_bindata(data, extension=extension, storage_id=storage_id)

    def remove_embedded_bindata(self, storage_id: int) -> bool:
        self._invalidate_bridge()
        return self._binary_document.remove_embedded_bindata(storage_id)

    def append_table_pure(self) -> None:
        self.append_table()

    def append_picture_pure(self) -> None:
        self.append_picture()

    def append_hyperlink_pure(self) -> None:
        self.append_hyperlink()

    def bridge(self):
        from .bridge import HwpHwpxBridge

        return HwpHwpxBridge.from_hwp(self, converter=self._converter)

    def __getattr__(self, name: str):
        if name.startswith("_"):
            raise AttributeError(name)
        bridge_document = self.to_hwpx_document()
        return getattr(bridge_document, name)

    def _ensure_bridge_workspace(self) -> Path:
        if self._bridge_temp_dir is None:
            self._bridge_temp_dir = Path(tempfile.mkdtemp(prefix="jakal_hwp_bridge_"))
        return self._bridge_temp_dir

    def _append_feature_pure(self, feature: str) -> None:
        self._invalidate_bridge()
        if self._pure_profile is None:
            self._pure_profile = HwpPureProfile.load_bundled()
        append_feature_from_profile(self._binary_document, self._pure_profile, feature)

    def _mutate_via_hwpx(self, mutate: Callable[[HwpxDocument], object]) -> object:
        bridge_document = self.to_hwpx_document(force_refresh=False)
        result = mutate(bridge_document)
        self._bridge_document = bridge_document
        self._sync_from_bridge_document()
        return result

    def _sync_from_bridge_document(self) -> None:
        if self._bridge_document is None:
            return
        from .hancom_document import HancomDocument

        rebuilt = HancomDocument.from_hwpx_document(self._bridge_document, converter=self._converter).to_hwp_document(
            converter=self._converter
        )
        self._binary_document = rebuilt.binary_document()
        self._hancom_document = None

    def _invalidate_bridge(self) -> None:
        self._mutation_revision += 1
        self._bridge_document = None
        self._hancom_document = None


from ._hwp_objects import (  # noqa: E402, F401 - late import avoids runtime cycle; compatibility re-exports
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
    _build_control_wrapper,
    _collect_text_from_node,
    _control_has_descendant,
    _control_id,
    _semantic_field_type_from_native,
    _shape_kind_from_control_node,
    _shape_wrapper_class,
)
