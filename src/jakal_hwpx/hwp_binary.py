from __future__ import annotations

import zlib
from contextlib import contextmanager
from pathlib import Path
from typing import Callable, Iterator

import olefile

from ._cfb_writer import CfbLayout, capture_compound_file_layout, write_compound_file
from ._hwp_component_records import (  # noqa: F401 - compatibility re-exports
    ArcShapeComponentRecord,
    ContainerShapeComponentRecord,
    CurveShapeComponentRecord,
    EllipseShapeComponentRecord,
    EqEditRecord,
    FootnoteShapeRecord,
    LineShapeComponentRecord,
    OleShapeComponentRecord,
    PageBorderFillRecord,
    PageDefRecord,
    PictureShapeComponentRecord,
    PolygonShapeComponentRecord,
    RectangleShapeComponentRecord,
    ShapeComponentRecord,
    TextArtShapeComponentRecord,
)
from ._hwp_constants import (  # noqa: F401 - compatibility re-exports
    _AUTO_NUMBER_TYPE_CODES,
    _AUTO_NUMBER_TYPE_VALUES,
    _BLANK_DROP_CONTROL_IDS,
    _BOOKMARK_CTRL_DATA_PREFIX,
    _CONTROL_MARKER_CODES,
    _CONTROL_NATIVE_METADATA_PREFIX,
    _DEFAULT_ARC_CTRL_HEADER,
    _DEFAULT_ARC_PAYLOAD,
    _DEFAULT_ARC_SHAPE_COMPONENT,
    _DEFAULT_COMPLEX_SHAPE_COMPONENT_LINE_WIDTH,
    _DEFAULT_ELLIPSE_CTRL_HEADER,
    _DEFAULT_ELLIPSE_PAYLOAD,
    _DEFAULT_ELLIPSE_SHAPE_COMPONENT,
    _DEFAULT_EQUATION_CTRL_HEADER,
    _DEFAULT_OLE_LINE_WIDTH,
    _DEFAULT_POLYGON_CTRL_HEADER,
    _DEFAULT_POLYGON_PAYLOAD,
    _DEFAULT_POLYGON_SHAPE_COMPONENT,
    _DEFAULT_RECT_CTRL_HEADER,
    _DEFAULT_RECT_SHAPE_COMPONENT,
    _DEFAULT_SHAPE_COMPONENT,
    _DEFAULT_SHAPE_CTRL_HEADER,
    _DEFAULT_SHAPE_LINE_PAYLOAD,
    _DEFAULT_SHAPE_LINE_WIDTH,
    _DEFAULT_SHAPE_LIST_HEADER,
    _DEFAULT_SHAPE_RECTANGLE_PAYLOAD,
    _DEFAULT_SHAPE_TEXT_CHAR_SHAPE,
    _DEFAULT_SHAPE_TEXT_LINE_SEG,
    _DEFAULT_SHAPE_TEXT_PARA_CONTROL_MASK,
    _DEFAULT_SHAPE_TEXT_PARA_SHAPE_ID,
    _DEFAULT_SHAPE_TEXT_PARA_TRAILING,
    _DEFAULT_SHAPE_TEXT_SPLIT_FLAGS,
    _DEFAULT_SHAPE_TEXT_STYLE_ID,
    _DEFAULT_TEXTART_CTRL_HEADER,
    _DEFAULT_TEXTART_FONT_NAME,
    _DEFAULT_TEXTART_FONT_STYLE,
    _DEFAULT_TEXTART_SHAPE_COMPONENT,
    _ENDNOTE_CTRL_HEADER,
    _EQUATION_VERSION_MARKER,
    _FOOTER_CTRL_HEADER,
    _FOOTER_LIST_HEADER,
    _FOOTER_PARA_CHAR_SHAPE,
    _FOOTER_PARA_HEADER,
    _FOOTER_PARA_LINE_SEG,
    _FOOTNOTE_CTRL_HEADER,
    _HANCOM_CONNECTLINE_COMPONENT_SIGNATURE,
    _HEADER_CTRL_HEADER,
    _HEADER_LIST_HEADER,
    _HEADER_PARA_CHAR_SHAPE,
    _HEADER_PARA_HEADER,
    _HEADER_PARA_LINE_SEG,
    _HWP_CHAR_FONT_ORDER,
    _HWP_COMPLEX_SHAPE_TEMPLATES,
    _HWP_DRAW_ASPECT_CODES,
    _HWP_DRAW_ASPECT_NAMES,
    _HWP_FILLED_SHAPE_COMMON_PAYLOAD_SIZE,
    _HWP_LINE_ARROW_CODES,
    _HWP_LINE_ARROW_NAMES,
    _HWP_LINE_ARROW_SIZE_CODES,
    _HWP_LINE_ARROW_SIZE_NAMES,
    _HWP_LINE_END_CAP_CODES,
    _HWP_LINE_END_CAP_NAMES,
    _HWP_LINE_OUTLINE_STYLE_CODES,
    _HWP_LINE_OUTLINE_STYLE_NAMES,
    _HWP_LINE_STYLE_CODES,
    _HWP_LINE_STYLE_NAMES,
    _HWP_OLE_OBJECT_TYPE_CODES,
    _HWP_OLE_OBJECT_TYPE_NAMES,
    _HWP_PICTURE_EFFECT_CODES,
    _HWP_PICTURE_EFFECT_NAMES,
    _HWP_TAB_LEADER_NAMES,
    _HWP_TAB_TYPE_NAMES,
    _HYPERLINK_CHAR_SHAPE,
    _HYPERLINK_CTRL_PREFIX,
    _HYPERLINK_CTRL_SUFFIX,
    _HYPERLINK_LINE_SEG,
    _HYPERLINK_PARA_HEADER_TAIL,
    _HYPERLINK_TEXT_END_MARKER,
    _HYPERLINK_TEXT_START_MARKER,
    _MEMO_SHAPE_NATIVE_METADATA_PREFIX,
    _NOTE_CHAR_SHAPE,
    _NOTE_LINE_SEG,
    _NOTE_LIST_HEADER,
    _NOTE_PARA_HEADER,
    _PICTURE_CTRL_HEADER,
    _PICTURE_SHAPE_COMPONENT,
    _PICTURE_SHAPE_PICTURE,
    _PICTURE_TOP_CHAR_SHAPE,
    _PICTURE_TOP_LINE_SEG,
    _PICTURE_TOP_PARA_HEADER_TAIL,
    _PICTURE_TOP_PARA_TEXT,
    _SHAPE_NATIVE_METADATA_PREFIX,
    _TABLE_CELL_CHAR_SHAPE,
    _TABLE_CELL_LINE_SEG,
    _TABLE_CELL_PARA_HEADER_TAIL,
    _TABLE_CTRL_HEADER,
    _TABLE_TOP_CHAR_SHAPE,
    _TABLE_TOP_LINE_SEG,
    _TABLE_TOP_PARA_HEADER_TAIL,
    _TABLE_TOP_PARA_TEXT,
    _TEXTART_COMPONENT_REFERENCE_EXTENT,
    DEFAULT_HWP_EQUATION_FONT,
    DEFAULT_HWP_EQUATION_SCRIPT,
    DEFAULT_HWP_HYPERLINK_METADATA_FIELDS,
    DEFAULT_HWP_HYPERLINK_TEXT,
    DEFAULT_HWP_HYPERLINK_URL,
    DEFAULT_HWP_SHAPE_HEIGHT,
    DEFAULT_HWP_SHAPE_TEXT,
    DEFAULT_HWP_SHAPE_WIDTH,
    DEFAULT_HWP_TABLE_BORDER_FILL_ID,
    DEFAULT_HWP_TABLE_CELL_HEIGHT,
    DEFAULT_HWP_TABLE_CELL_MARGIN,
    DEFAULT_HWP_TABLE_CELL_WIDTH,
    DEFAULT_HWP_TABLE_COLS,
    DEFAULT_HWP_TABLE_ROWS,
    FILE_HEADER_SIGNATURE,
    FILE_HEADER_SIZE,
    FLAG_COMPRESSED,
    FLAG_DISTRIBUTABLE,
    FLAG_PASSWORD_PROTECTED,
    TAG_BIN_DATA,
    TAG_BORDER_FILL,
    TAG_BULLET,
    TAG_CHAR_SHAPE,
    TAG_CHART_DATA,
    TAG_COMPATIBLE_DOCUMENT,
    TAG_CTRL_DATA,
    TAG_CTRL_HEADER,
    TAG_DOC_DATA,
    TAG_DOCUMENT_PROPERTIES,
    TAG_EQEDIT,
    TAG_FACE_NAME,
    TAG_FOOTNOTE_SHAPE,
    TAG_FORM_OBJECT,
    TAG_ID_MAPPINGS,
    TAG_LAYOUT_COMPATIBILITY,
    TAG_LIST_HEADER,
    TAG_MEMO_LIST,
    TAG_MEMO_SHAPE,
    TAG_NAMES,
    TAG_NUMBERING,
    TAG_PAGE_BORDER_FILL,
    TAG_PAGE_DEF,
    TAG_PARA_CHAR_SHAPE,
    TAG_PARA_HEADER,
    TAG_PARA_LINE_SEG,
    TAG_PARA_RANGE_TAG,
    TAG_PARA_SHAPE,
    TAG_PARA_TEXT,
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
    TAG_STYLE,
    TAG_TAB_DEF,
    TAG_TABLE,
    TAG_TRACK_CHANGE,
    TAG_TRACK_CHANGE_AUTHOR,
)
from ._hwp_control_payloads import (  # noqa: F401 - compatibility re-exports
    _build_auto_number_payload,
    _build_bookmark_ctrl_data_payload,
    _build_control_id_payload,
    _build_control_marker_text,
    _build_equation_payload,
    _build_header_footer_control_payload,
    _build_hyperlink_command,
    _build_hyperlink_ctrl_payload,
    _build_hyperlink_para_text_payload,
    _build_note_ctrl_payload,
    _control_has_descendant,
    _control_native_metadata,
    _decode_control_id_payload,
    _metadata_prefixed_int_map,
    _metadata_prefixed_str_map,
    _normalize_field_control_id,
    _parse_bookmark_ctrl_data_payload,
    _parse_control_native_metadata,
    _parse_header_footer_apply_page_type,
    _parse_hyperlink_command,
    _parse_object_description,
    _parse_shape_native_metadata,
    _semantic_field_type_from_native,
    _shape_kind_from_control_node,
    _update_control_native_metadata,
)
from ._hwp_control_records import (  # noqa: F401 - compatibility re-exports
    AutoNumberControlRecord,
    BookmarkControlRecord,
    ControlDataRecord,
    ControlHeaderRecord,
    FieldControlRecord,
    HeaderFooterControlRecord,
    HyperlinkControlRecord,
    NoteControlRecord,
    PageHidingControlRecord,
    PageNumberControlRecord,
    PageStartsOnControlRecord,
    SectionDefinitionControlRecord,
)
from ._hwp_docinfo_records import (  # noqa: F401 - compatibility re-exports
    ID_MAPPINGS_INDEX_NAMES,
    BinDataRecord,
    BorderFillRecord,
    BulletRecord,
    CharShapeRecord,
    ChartDataRecord,
    CompatibleDocumentRecord,
    DocDataRecord,
    DocumentPropertiesRecord,
    FaceNameRecord,
    FormObjectRecord,
    IdMappingsRecord,
    LayoutCompatibilityRecord,
    ListHeaderRecord,
    MemoListRecord,
    MemoShapeRecord,
    NumberingRecord,
    ParaShapeRecord,
    StyleRecord,
    TabDefRecord,
    TrackChangeAuthorRecord,
    TrackChangeRecord,
    _build_document_properties_payload,
    _parse_document_properties_payload,
)
from ._hwp_graphic_payloads import (  # noqa: E402, F401 - compatibility re-exports
    _build_arc_specific_payload,
    _build_control_native_metadata_payload,
    _build_ellipse_specific_payload,
    _build_filled_shape_common_payload,
    _build_graphic_control_header_payload,
    _build_line_attribute_flags,
    _build_memo_shape_native_metadata_payload,
    _build_minimal_shape_component_payload,
    _build_ole_attribute_flags,
    _build_ole_specific_payload,
    _build_polygon_specific_payload,
    _build_prefixed_metadata_payload,
    _build_shape_component_payload,
    _build_shape_fill_payload,
    _build_shape_line_info_payload,
    _build_shape_native_metadata_payload,
    _build_shape_specific_payload,
    _build_textart_specific_payload,
    _graphic_control_out_margin_values,
    _parse_graphic_control_layout,
    _parse_graphic_control_out_margins,
    _parse_picture_crop_payload,
    _parse_picture_image_adjustment_payload,
    _parse_shape_component_rotation_payload,
    _picture_crop_values,
    _picture_default_clip_size,
    _picture_effect_code,
    _picture_effect_name,
    _picture_image_adjustment_values,
    _set_graphic_control_layout_payload,
    _set_graphic_control_out_margins_payload,
    _set_picture_crop_payload,
    _set_picture_image_adjustment_payload,
    _set_picture_line_payload,
    _set_shape_component_rotation_payload,
    _shape_component_default_rotation_center,
    _shape_component_rotation_values,
    _shape_control_header_template,
)
from ._hwp_graphic_records import (  # noqa: F401 - compatibility re-exports
    EquationControlRecord,
    GraphicControlRecord,
    OleControlRecord,
    PictureControlRecord,
    ShapeControlRecord,
    TableControlRecord,
)
from ._hwp_page_payloads import (  # noqa: E402, F401 - compatibility re-exports
    _HEADER_FOOTER_APPLY_PAGE_TYPE_CODES,
    _HEADER_FOOTER_APPLY_PAGE_TYPE_VALUES,
    _NOTE_LINE_TYPE_CODES,
    _NOTE_LINE_TYPE_VALUES,
    _NOTE_LINE_WIDTH_CODES,
    _NOTE_LINE_WIDTH_VALUES,
    _NOTE_NUMBERING_CODES,
    _NOTE_NUMBERING_VALUES,
    _NOTE_PLACEMENT_CODES,
    _NUMBER_TYPE_CODES,
    _NUMBER_TYPE_VALUES,
    _PAGE_BORDER_FILL_BORDER_AREA_FLAG,
    _PAGE_BORDER_FILL_FOOTER_INSIDE_FLAG,
    _PAGE_BORDER_FILL_HEADER_INSIDE_FLAG,
    _PAGE_BORDER_FILL_TEXT_BORDER_FLAG,
    _PAGE_HIDING_FLAGS,
    _PAGE_NUM_POSITION_CODES,
    _PAGE_NUM_POSITION_VALUES,
    _SECTION_DEF_BORDER_FIRST_FLAG,
    _SECTION_DEF_FILL_FIRST_FLAG,
    _SECTION_DEF_HIDE_EMPTY_LINE_FLAG,
    _SECTION_DEF_HIDE_FOOTER_FLAG,
    _SECTION_DEF_HIDE_HEADER_FLAG,
    _SECTION_DEF_HIDE_MASTER_PAGE_FLAG,
    _SECTION_DEF_HIDE_PAGE_NUM_FLAG,
    _SECTION_DEF_WONGGOJI_FORMAT_FLAG,
    _SECTION_DEFINITION_STARTS_ON_CODES,
    _SECTION_DEFINITION_STARTS_ON_MASK,
    _SECTION_DEFINITION_STARTS_ON_SHIFT,
    _SECTION_DEFINITION_STARTS_ON_VALUES,
    _SECTION_PAGE_BORDER_FILL_TYPES,
    _SECTION_PAGE_STARTS_ON_CODES,
    _SECTION_PAGE_STARTS_ON_MASK,
    _SECTION_PAGE_STARTS_ON_VALUES,
    _build_colorref,
    _build_note_shape_payload,
    _build_page_border_fill_payload,
    _build_page_def_payload,
    _build_page_hiding_payload,
    _build_page_num_payload,
    _build_page_starts_on_payload,
    _build_rgb_color,
    _build_section_definition_payload,
    _parse_note_shape_payload,
    _parse_page_border_fill_payload,
    _parse_page_def_payload,
    _parse_page_hiding_payload,
    _parse_page_num_payload,
    _parse_page_starts_on_payload,
    _parse_rgb_color,
    _parse_section_definition_payload,
)
from ._hwp_paragraph_records import (  # noqa: F401 - compatibility re-exports
    ParagraphHeaderRecord,
    ParagraphTextRecord,
    _replace_with_limit,
)
from ._hwp_payload_codecs import (  # noqa: F401 - compatibility re-exports
    _bit_flags,
    _parse_arc_shape_payload,
    _parse_border_fill_payload,
    _parse_border_line_payload,
    _parse_bullet_payload,
    _parse_character_shape_payload,
    _parse_chart_data_payload,
    _parse_colorref,
    _parse_connectline_shape_payload,
    _parse_container_shape_payload,
    _parse_curve_shape_payload,
    _parse_ellipse_shape_payload,
    _parse_face_name_payload,
    _parse_form_object_payload,
    _parse_hwp_colorref_or_none,
    _parse_length_prefixed_utf16,
    _parse_line_attribute_flags,
    _parse_line_shape_payload,
    _parse_list_header_payload,
    _parse_memo_list_payload,
    _parse_memo_shape_native_metadata,
    _parse_memo_shape_payload,
    _parse_numbering_payload,
    _parse_paragraph_style_payload,
    _parse_polygon_shape_payload,
    _parse_rectangle_shape_payload,
    _parse_shape_component_common_payload,
    _parse_shape_fill_payload,
    _parse_shape_line_info_payload,
    _parse_style_definition_payload,
    _parse_tab_def_payload,
    _parse_textart_payload_text,
    _parse_track_change_author_payload,
    _parse_track_change_payload,
    _payload_u16_values,
    _payload_u32_values,
    _shape_common_payload_size,
    _shape_specific_payload_size,
)
from ._hwp_record_base import (  # noqa: F401 - compatibility re-exports
    HwpBinaryFileHeader,
    HwpDocumentProperties,
    HwpParagraph,
    HwpRecord,
    HwpStreamCapacity,
    OpaquePayloadRecord,
    RecordNode,
    TypedRecord,
    UnknownRecord,
    hwp_tag_name,
)
from ._hwp_section_helpers import (  # noqa: F401 - compatibility re-exports
    _build_blank_paragraph_header,
    _build_blank_paragraph_raw_text,
    _build_blank_section_model,
    _build_shape_text_records,
    _collect_text_from_node,
    _find_child_record,
    _find_paragraph_control_node,
    _find_section_definition_control_node,
    _insert_blank_like_paragraph,
    _insert_paragraph_control_node,
    _is_blank_preserved_control_id,
    _iter_paragraph_raw_text_tokens,
    _normalize_section_definition_paragraph_for_multi_section,
)
from ._hwp_table_payloads import (  # noqa: F401 - compatibility re-exports
    _build_table_cell_list_header_payload,
    _build_table_cell_specs,
    _build_table_record_payload,
    _normalize_table_cell_texts,
    _normalize_table_measurements,
    _normalize_table_spans,
    _TableCellSpec,
)
from .exceptions import HwpBinaryEditError, InvalidHwpFileError


def _normalize_stream_path(path: str | tuple[str, ...] | list[str]) -> str:
    if isinstance(path, str):
        return path.strip("/").replace("\\", "/")
    return "/".join(path)


def _stream_components(path: str) -> list[str]:
    normalized = _normalize_stream_path(path)
    return [part for part in normalized.split("/") if part]




from ._hwp_binary_authoring import HwpBinaryAuthoringMixin  # noqa: E402
from ._hwp_binary_sections import HwpBinarySectionMixin  # noqa: E402
from ._hwp_record_tree import (  # noqa: E402, F401 - compatibility re-exports
    DocInfoModel,
    SectionModel,
    SectionParagraphModel,
    _iter_records,
    _specialize_record_node,
    _specialize_record_tree,
    _strip_para_text_controls,
    build_record_tree,
    flatten_record_tree,
    record_node_from_record,
)


class HwpBinaryDocument(HwpBinaryAuthoringMixin, HwpBinarySectionMixin):
    def __init__(
        self,
        source_path: Path,
        streams: dict[str, bytes],
        file_header: HwpBinaryFileHeader,
        *,
        compound_layout: CfbLayout | None = None,
        original_file_bytes: bytes | None = None,
    ) -> None:
        self.source_path = source_path
        self._streams = dict(streams)
        self._original_sizes = {path: len(data) for path, data in streams.items()}
        self._file_header = file_header
        self._compound_layout = compound_layout
        self._original_file_bytes = original_file_bytes
        self._dirty = False
        self._section_model_batch_depth = 0
        self._section_model_cache: dict[int, SectionModel] = {}
        self._section_model_dirty: set[int] = set()
        self._docinfo_model_batch_depth = 0
        self._docinfo_model_cache: DocInfoModel | None = None
        self._docinfo_model_cache_dirty = False

    @classmethod
    def open(cls, path: str | Path) -> "HwpBinaryDocument":
        input_path = Path(path).expanduser().resolve()
        if not input_path.exists():
            raise FileNotFoundError(input_path)
        if not olefile.isOleFile(str(input_path)):
            raise InvalidHwpFileError(f"{input_path} is not a valid OLE Compound File HWP document.")

        original_file_bytes = input_path.read_bytes()

        with olefile.OleFileIO(str(input_path)) as ole:
            streams: dict[str, bytes] = {}
            for entry in ole.listdir(streams=True, storages=False):
                stream_path = _normalize_stream_path(entry)
                streams[stream_path] = ole.openstream(entry).read()

        file_header_stream = streams.get("FileHeader")
        if file_header_stream is None:
            raise InvalidHwpFileError(f"{input_path} does not contain a FileHeader stream.")

        file_header = HwpBinaryFileHeader.from_bytes(file_header_stream)
        if file_header.signature != FILE_HEADER_SIGNATURE.decode("ascii"):
            raise InvalidHwpFileError(f"{input_path} does not have the expected HWP file signature.")

        return cls(
            source_path=input_path,
            streams=streams,
            file_header=file_header,
            compound_layout=capture_compound_file_layout(ole, original_file_bytes=original_file_bytes),
            original_file_bytes=original_file_bytes,
        )

    def file_header(self) -> HwpBinaryFileHeader:
        return self._file_header

    def list_stream_paths(self) -> list[str]:
        return sorted(self._streams)

    def section_stream_paths(self) -> list[str]:
        return sorted(path for path in self._streams if path.startswith("BodyText/Section"))

    def bindata_stream_paths(self) -> list[str]:
        return sorted(path for path in self._streams if path.startswith("BinData/"))

    def has_stream(self, path: str | tuple[str, ...] | list[str]) -> bool:
        return _normalize_stream_path(path) in self._streams

    def stream_size(self, path: str | tuple[str, ...] | list[str]) -> int:
        normalized = _normalize_stream_path(path)
        return self._original_sizes[normalized]

    def stream_capacity(self, path: str | tuple[str, ...] | list[str], data: bytes, *, compress: bool | None = None) -> HwpStreamCapacity:
        normalized = _normalize_stream_path(path)
        if normalized not in self._streams:
            raise KeyError(normalized)
        if compress is None:
            compress = self._stream_is_compressed(normalized)
        original_size = self._original_sizes[normalized]
        if compress:
            compressor = zlib.compressobj(level=9, wbits=-15)
            current_size = len(compressor.compress(data) + compressor.flush())
        else:
            current_size = len(data)
        return HwpStreamCapacity(
            path=normalized,
            compressed=compress,
            original_size=original_size,
            current_size=current_size,
            free_bytes=original_size - current_size,
            fits=current_size <= original_size,
        )

    def section_capacity(self, section_index: int, records: list[HwpRecord] | None = None) -> HwpStreamCapacity:
        path = self.section_stream_paths()[section_index]
        if records is None:
            data = self.read_stream(path)
        else:
            data = b"".join(record.to_bytes() for record in records)
        return self.stream_capacity(path, data, compress=True)

    def docinfo_capacity(self, records: list[HwpRecord] | None = None) -> HwpStreamCapacity:
        data = self.read_stream("DocInfo") if records is None else b"".join(record.to_bytes() for record in records)
        return self.stream_capacity("DocInfo", data, compress=True)

    def read_stream(
        self,
        path: str | tuple[str, ...] | list[str],
        *,
        decompress: bool | None = None,
    ) -> bytes:
        normalized = _normalize_stream_path(path)
        if normalized not in self._streams:
            raise KeyError(normalized)

        data = self._streams[normalized]
        if decompress is None:
            decompress = self._stream_is_compressed(normalized)
        if not decompress:
            return data
        return zlib.decompress(data, -15)

    def add_stream(self, path: str | tuple[str, ...] | list[str], data: bytes) -> None:
        normalized = _normalize_stream_path(path)
        previous = self._streams.get(normalized)
        self._streams[normalized] = data
        self._original_sizes.setdefault(normalized, len(data))
        if previous != data:
            self._dirty = True

    def remove_stream(self, path: str | tuple[str, ...] | list[str]) -> None:
        normalized = _normalize_stream_path(path)
        removed = self._streams.pop(normalized, None)
        self._original_sizes.pop(normalized, None)
        if removed is not None:
            self._dirty = True

    def write_stream(
        self,
        path: str | tuple[str, ...] | list[str],
        data: bytes,
        *,
        compress: bool | None = None,
    ) -> None:
        normalized = _normalize_stream_path(path)
        if normalized not in self._streams:
            raise KeyError(normalized)

        if compress is None:
            compress = self._stream_is_compressed(normalized)

        if compress:
            compressor = zlib.compressobj(level=9, wbits=-15)
            payload = compressor.compress(data) + compressor.flush()
        else:
            payload = data

        if self._streams[normalized] != payload:
            self._streams[normalized] = payload
            self._dirty = True

    def preview_text(self) -> str:
        data = self.read_stream("PrvText", decompress=False)
        return data.decode("utf-16-le", errors="ignore").rstrip("\x00")

    def set_preview_text(self, value: str) -> None:
        encoded = value.encode("utf-16-le")
        self.write_stream("PrvText", encoded, compress=False)

    def docinfo_records(self) -> list[HwpRecord]:
        return list(_iter_records(self.read_stream("DocInfo")))

    def document_properties(self) -> HwpDocumentProperties:
        records = self.docinfo_records()
        if not records or records[0].tag_id != TAG_DOCUMENT_PROPERTIES:
            raise InvalidHwpFileError("DocInfo does not start with a document properties record.")
        return _parse_document_properties_payload(records[0].payload)

    def docinfo_model(self) -> DocInfoModel:
        if self._docinfo_model_batch_depth > 0:
            if self._docinfo_model_cache is None:
                self._docinfo_model_cache = DocInfoModel.from_records(self.docinfo_records())
            return self._docinfo_model_cache
        return DocInfoModel.from_records(self.docinfo_records())

    def replace_docinfo_model(self, model: DocInfoModel) -> None:
        if self._docinfo_model_batch_depth > 0:
            self._docinfo_model_cache = model
            self._docinfo_model_cache_dirty = True
            self._dirty = True
            return
        self.write_stream("DocInfo", model.to_bytes(), compress=True)

    def begin_docinfo_model_batch(self) -> None:
        self._docinfo_model_batch_depth += 1

    def end_docinfo_model_batch(self) -> None:
        if self._docinfo_model_batch_depth <= 0:
            return
        self._docinfo_model_batch_depth -= 1
        if self._docinfo_model_batch_depth > 0:
            return
        model = self._docinfo_model_cache
        dirty = self._docinfo_model_cache_dirty
        self._docinfo_model_cache = None
        self._docinfo_model_cache_dirty = False
        if dirty and model is not None:
            self.write_stream("DocInfo", model.to_bytes(), compress=True)

    @contextmanager
    def docinfo_model_batch(self) -> Iterator[None]:
        self.begin_docinfo_model_batch()
        try:
            yield
        finally:
            self.end_docinfo_model_batch()

    def section_records(self, section_index: int) -> list[HwpRecord]:
        path = self.section_stream_paths()[section_index]
        return list(_iter_records(self.read_stream(path)))

    def replace_section_records(self, section_index: int, records: list[HwpRecord]) -> None:
        path = self.section_stream_paths()[section_index]
        raw = b"".join(record.to_bytes() for record in records)
        self.write_stream(path, raw, compress=True)

    def section_model(self, section_index: int = 0) -> SectionModel:
        if self._section_model_batch_depth > 0:
            cached = self._section_model_cache.get(section_index)
            if cached is not None:
                return cached
            model = SectionModel.from_records(section_index, self.section_records(section_index))
            self._section_model_cache[section_index] = model
            return model
        return SectionModel.from_records(section_index, self.section_records(section_index))

    def replace_section_model(self, section_index: int, model: SectionModel) -> None:
        if self._section_model_batch_depth > 0:
            self._section_model_cache[section_index] = model
            self._section_model_dirty.add(section_index)
            self._dirty = True
            return
        self.replace_section_records(section_index, model.to_records())

    def begin_section_model_batch(self) -> None:
        self._section_model_batch_depth += 1

    def end_section_model_batch(self) -> None:
        if self._section_model_batch_depth <= 0:
            return
        self._section_model_batch_depth -= 1
        if self._section_model_batch_depth > 0:
            return
        cached_models = self._section_model_cache
        dirty_indexes = self._section_model_dirty
        self._section_model_cache = {}
        self._section_model_dirty = set()
        for section_index in dirty_indexes:
            self.replace_section_records(section_index, cached_models[section_index].to_records())

    @contextmanager
    def section_model_batch(self) -> Iterator[None]:
        self.begin_section_model_batch()
        try:
            yield
        finally:
            self.end_section_model_batch()


    def paragraphs(self, section_index: int = 0) -> list[HwpParagraph]:
        paragraphs: list[HwpParagraph] = []
        current: dict[str, int | str] | None = None
        paragraph_index = -1
        for record in self.section_records(section_index):
            if record.tag_id == TAG_PARA_HEADER:
                if current is not None:
                    paragraphs.append(
                        HwpParagraph(
                            index=paragraph_index,
                            section_index=section_index,
                            char_count=int(current["char_count"]),
                            control_mask=int(current["control_mask"]),
                            para_shape_id=int(current["para_shape_id"]),
                            style_id=int(current["style_id"]),
                            split_flags=int(current["split_flags"]),
                            raw_text=str(current["raw_text"]),
                            text=str(current["text"]),
                            text_record_offset=int(current["text_record_offset"]),
                            text_record_size=int(current["text_record_size"]),
                        )
                    )
                paragraph_index += 1
                current = {
                    "char_count": int.from_bytes(record.payload[0:4], "little") if len(record.payload) >= 4 else 0,
                    "control_mask": int.from_bytes(record.payload[4:8], "little") if len(record.payload) >= 8 else 0,
                    "para_shape_id": int.from_bytes(record.payload[8:10], "little") if len(record.payload) >= 10 else 0,
                    "style_id": record.payload[10] if len(record.payload) >= 11 else 0,
                    "split_flags": record.payload[11] if len(record.payload) >= 12 else 0,
                    "raw_text": "",
                    "text": "",
                    "text_record_offset": -1,
                    "text_record_size": 0,
                }
                continue

            if record.tag_id == TAG_PARA_TEXT and current is not None:
                raw_text = record.payload.decode("utf-16-le", errors="ignore")
                current["raw_text"] = raw_text
                current["text"] = _strip_para_text_controls(raw_text)
                current["text_record_offset"] = record.offset
                current["text_record_size"] = record.size

        if current is not None:
            paragraphs.append(
                HwpParagraph(
                    index=paragraph_index,
                    section_index=section_index,
                    char_count=int(current["char_count"]),
                    control_mask=int(current["control_mask"]),
                    para_shape_id=int(current["para_shape_id"]),
                    style_id=int(current["style_id"]),
                    split_flags=int(current["split_flags"]),
                    raw_text=str(current["raw_text"]),
                    text=str(current["text"]),
                    text_record_offset=int(current["text_record_offset"]),
                    text_record_size=int(current["text_record_size"]),
                )
            )
        return paragraphs

    def get_document_text(self) -> str:
        fragments: list[str] = []
        for index, _path in enumerate(self.section_stream_paths()):
            for paragraph in self.paragraphs(index):
                cleaned = paragraph.text.replace("\r", "\n").strip("\n")
                if cleaned:
                    fragments.append(cleaned)
        return "\n".join(fragments)

    def set_paragraph_text_same_length(
        self,
        section_index: int,
        paragraph_index: int,
        new_text: str,
    ) -> None:
        paragraph = self.paragraphs(section_index)[paragraph_index]
        if paragraph.raw_text != paragraph.text:
            raise HwpBinaryEditError(
                "set_paragraph_text_same_length is only supported for paragraphs without hidden control characters."
            )
        if len(new_text.encode("utf-16-le")) != len(paragraph.raw_text.encode("utf-16-le")):
            raise HwpBinaryEditError(
                "set_paragraph_text_same_length requires the new text to match the paragraph's UTF-16 byte length."
            )
        self._rewrite_paragraph_text_record(section_index, paragraph_index, lambda _current: new_text)

    def replace_paragraph_text_same_length(
        self,
        section_index: int,
        paragraph_index: int,
        old: str,
        new: str,
        *,
        count: int = -1,
    ) -> int:
        if not old:
            raise ValueError("old must be a non-empty string.")
        if len(old.encode("utf-16-le")) != len(new.encode("utf-16-le")):
            raise HwpBinaryEditError(
                "replace_paragraph_text_same_length requires old and new to have the same UTF-16 byte length."
            )

        replaced = 0

        def updater(current: str) -> str:
            nonlocal replaced
            updated, changed = _replace_with_limit(current, old, new, count)
            replaced = changed
            return updated

        self._rewrite_paragraph_text_record(section_index, paragraph_index, updater)
        return replaced

    def replace_text_same_length(
        self,
        old: str,
        new: str,
        *,
        section_index: int | None = None,
        count: int = -1,
    ) -> int:
        if not old:
            raise ValueError("old must be a non-empty string.")
        if len(old.encode("utf-16-le")) != len(new.encode("utf-16-le")):
            raise HwpBinaryEditError("replace_text_same_length requires old and new to have the same UTF-16 byte length.")

        remaining = count
        replacements = 0
        target_paths = self.section_stream_paths()
        if section_index is not None:
            target_paths = [target_paths[section_index]]

        for path in target_paths:
            raw = self.read_stream(path)
            chunks: list[bytes] = []
            changed_in_stream = 0

            for record in _iter_records(raw):
                header_bytes = raw[record.offset : record.offset + record.header_size]
                payload = record.payload
                if record.tag_id == TAG_PARA_TEXT:
                    current_text = payload.decode("utf-16-le", errors="ignore")
                    limit = remaining if remaining >= 0 else -1
                    updated_text, changed = _replace_with_limit(current_text, old, new, limit)
                    if changed:
                        payload = updated_text.encode("utf-16-le")
                        changed_in_stream += changed
                        replacements += changed
                        if remaining >= 0:
                            remaining -= changed
                chunks.append(header_bytes)
                chunks.append(payload)
                if remaining == 0:
                    # Keep the rest of the stream untouched when the requested replacement count is exhausted.
                    tail_offset = record.offset + record.header_size + len(record.payload)
                    chunks.append(raw[tail_offset:])
                    break
            else:
                tail_offset = len(raw)

            if changed_in_stream:
                updated_raw = b"".join(chunks)
                if len(updated_raw) != len(raw):
                    raise HwpBinaryEditError("Section rewrite changed the decompressed stream length unexpectedly.")
                self.write_stream(path, updated_raw, compress=True)

            if remaining == 0:
                break

        return replacements

    def _rewrite_paragraph_text_record(
        self,
        section_index: int,
        paragraph_index: int,
        update_text: Callable[[str], str],
    ) -> None:
        path = self.section_stream_paths()[section_index]
        raw = self.read_stream(path)
        chunks: list[bytes] = []
        current_paragraph_index = -1
        changed = False

        for record in _iter_records(raw):
            header_bytes = raw[record.offset : record.offset + record.header_size]
            payload = record.payload

            if record.tag_id == TAG_PARA_HEADER:
                current_paragraph_index += 1
            elif record.tag_id == TAG_PARA_TEXT and current_paragraph_index == paragraph_index:
                current_text = payload.decode("utf-16-le", errors="ignore")
                updated_text = update_text(current_text)
                updated_payload = updated_text.encode("utf-16-le")
                if len(updated_payload) != len(payload):
                    raise HwpBinaryEditError("Paragraph text rewrite changed the paragraph text record size.")
                payload = updated_payload
                changed = True

            chunks.append(header_bytes)
            chunks.append(payload)

        if not changed:
            raise IndexError(
                f"Paragraph {paragraph_index} in section {section_index} does not contain a writable PARA_TEXT record."
            )

        updated_raw = b"".join(chunks)
        if len(updated_raw) != len(raw):
            raise HwpBinaryEditError("Section rewrite changed the decompressed stream length unexpectedly.")
        self.write_stream(path, updated_raw, compress=True)

    def save(self, path: str | Path | None = None, *, preserve_original_bytes: bool = True) -> Path:
        target_path = self.source_path if path is None else Path(path).expanduser().resolve()
        self._write_to_path(target_path, preserve_original_bytes=preserve_original_bytes)
        if path is not None:
            self.source_path = target_path
        return target_path

    def save_copy(self, path: str | Path, *, preserve_original_bytes: bool = True) -> Path:
        target_path = Path(path).expanduser().resolve()
        self._write_to_path(target_path, preserve_original_bytes=preserve_original_bytes)
        return target_path

    def _stream_is_compressed(self, path: str) -> bool:
        if not self._file_header.compressed:
            return False
        return path == "DocInfo" or path.startswith("BodyText/")

    def _compress_stream_preserving_capacity(self, path: str, data: bytes) -> bytes:
        capacity = self.stream_capacity(path, data, compress=True)
        original_size = capacity.original_size
        compressor = zlib.compressobj(level=9, wbits=-15)
        compressed = compressor.compress(data) + compressor.flush()
        if not capacity.fits:
            raise HwpBinaryEditError(
                f"{path} compressed to {capacity.current_size} bytes and no longer fits in its original {original_size}-byte stream."
            )
        return compressed + (b"\x00" * (original_size - len(compressed)))

    def _write_to_path(self, target_path: Path, *, preserve_original_bytes: bool) -> None:
        target_path.parent.mkdir(parents=True, exist_ok=True)
        if preserve_original_bytes and not self._dirty and self._original_file_bytes is not None:
            target_path.write_bytes(self._original_file_bytes)
            return
        write_compound_file(target_path, self._streams, layout=self._compound_layout)
        self._original_file_bytes = target_path.read_bytes()
        self._dirty = False


from ._hwp_nested_controls import (  # noqa: E402, F401 - compatibility re-exports
    _append_control_marker_to_paragraph,
    _apply_section_show_line_numbers,
    _build_default_paragraph_line_seg_payload,
    _build_header_footer_control_node,
    _build_note_control_node,
    _ensure_paragraph_line_seg,
    _paragraph_line_seg_node,
    _paragraph_uses_synthetic_line_number_layout,
    _remove_synthetic_paragraph_line_seg,
    _section_uses_synthetic_line_number_layout,
    _set_header_footer_control_text,
    _set_paragraph_line_number_flag,
    _top_level_section_paragraphs,
)
