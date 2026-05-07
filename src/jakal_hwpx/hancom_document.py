from __future__ import annotations

import json
import struct
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

from lxml import etree

from .document import DocumentMetadata, HwpxDocument
from .hwp_binary import (
    TAG_BULLET,
    TAG_CHAR_SHAPE,
    TAG_CTRL_HEADER,
    TAG_CTRL_DATA,
    TAG_MEMO_SHAPE,
    TAG_NUMBERING,
    TAG_PARA_SHAPE,
    TAG_STYLE,
    ParagraphHeaderRecord,
    SectionParagraphModel,
    TypedRecord,
    _build_memo_shape_native_metadata_payload,
)
from .hwp_document import (
    HwpArcShapeObject,
    HwpConnectLineShapeObject,
    HwpContainerShapeObject,
    HwpCurveShapeObject,
    HwpDocument,
    HwpEllipseShapeObject,
    HwpLineShapeObject,
    HwpPolygonShapeObject,
    HwpRectangleShapeObject,
    HwpTextArtShapeObject,
)
from .namespaces import NS, qname

HancomConverter = Callable[[str | Path, str | Path, str], Path]

_HWP_STYLE_TYPE_CODES = {
    "PARA": 0,
    "CHAR": 1,
}
_HWP_STYLE_TYPE_NAMES = {value: key for key, value in _HWP_STYLE_TYPE_CODES.items()}
_HWP_DOCINFO_COUNT_INDEX_BY_TAG = {
    TAG_CHAR_SHAPE: 9,
    TAG_NUMBERING: 11,
    TAG_BULLET: 12,
    TAG_PARA_SHAPE: 13,
    TAG_STYLE: 14,
    TAG_MEMO_SHAPE: 15,
}
_DEFAULT_HWP_NUMBERING_PAYLOAD = bytes.fromhex(
    "0c00000000003200ffffffff03005e0031002e000c01000000003200ffffffff03005e0032002e000c00000000003200ffffffff"
    "03005e00330029000c01000000003200ffffffff03005e00340029000c00000000003200ffffffff040028005e00350029000c01"
    "000000003200ffffffff040028005e00360029002c00000000003200ffffffff02005e0037000000010000000100000001000000"
    "01000000010000000100000001000000"
)
_DEFAULT_HWP_BULLET_PAYLOAD = bytes.fromhex("0800000000003200ffffffffa1f0000000000000000000")
_DEFAULT_HWP_PARA_SHAPE_PAYLOAD = bytes.fromhex(
    "800100000000000000000000000000000000000000000000a000000000000000050000000000000000000000000000000000a0000000"
)
_DEFAULT_HWP_CHAR_SHAPE_PAYLOAD = bytes.fromhex(
    "0100010001000100000002000000646464646464640000000000000064646464646464000000000000004c040000020000000a0a0000000000000000ffffffffc0c0c000040000000000"
)

_HWP_PARA_ALIGN_CODES = {
    "JUSTIFY": 0,
    "LEFT": 1,
    "RIGHT": 2,
    "CENTER": 3,
    "DISTRIBUTE": 4,
    "DISTRIBUTE_SPACE": 5,
}
_HWP_PARA_ALIGN_NAMES = {value: key for key, value in _HWP_PARA_ALIGN_CODES.items()}
_HWP_NUMBERING_LEVEL_COUNT = 7
_HWP_CHAR_FONT_ORDER = ("hangul", "latin", "hanja", "japanese", "other", "symbol", "user")
_HWP_NATIVE_FIELD_TYPE_BY_SEMANTIC_TYPE = {
    "DOCPROPERTY": "%doc",
    "DATE": "%dat",
    "MAILMERGE": "%mai",
    "FORMULA": "%for",
    "CROSSREF": "%ref",
}
_HWPX_FORM_FIELD_TYPE = "JAKAL_FORM"
_HWPX_MEMO_FIELD_TYPE = "JAKAL_MEMO"
_HWPX_CHART_CARRIER_PREFIX = "JAKAL_CHART:"


def _semantic_field_type_from_native(native_field_type: str, parameters: dict[str, str | int] | None = None) -> str:
    normalized = native_field_type.strip().upper()
    normalized_parameters = {str(key).upper(): str(value) for key, value in (parameters or {}).items()}
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


@dataclass
class HancomMetadata:
    title: str | None = None
    language: str | None = None
    creator: str | None = None
    subject: str | None = None
    description: str | None = None
    lastsaveby: str | None = None
    created: str | None = None
    modified: str | None = None
    date: str | None = None
    keyword: str | None = None
    extra: dict[str, str] = field(default_factory=dict)

    @classmethod
    def from_hwpx_metadata(cls, metadata: DocumentMetadata) -> "HancomMetadata":
        return cls(
            title=metadata.title,
            language=metadata.language,
            creator=metadata.creator,
            subject=metadata.subject,
            description=metadata.description,
            lastsaveby=metadata.lastsaveby,
            created=metadata.created,
            modified=metadata.modified,
            date=metadata.date,
            keyword=metadata.keyword,
            extra=dict(metadata.extra),
        )

    def apply_to_hwpx_document(self, document: HwpxDocument) -> None:
        document.set_metadata(
            title=self.title,
            language=self.language,
            creator=self.creator,
            subject=self.subject,
            description=self.description,
            lastsaveby=self.lastsaveby,
            created=self.created,
            modified=self.modified,
            date=self.date,
            keyword=self.keyword,
        )


@dataclass
class Paragraph:
    text: str
    style_id: str | None = None
    para_pr_id: str | None = None
    char_pr_id: str | None = None
    hwp_para_shape_id: int | None = None
    hwp_style_id: int | None = None
    hwp_split_flags: int | None = None
    hwp_control_mask: int | None = None


@dataclass
class Table:
    rows: int
    cols: int
    cell_texts: list[list[str]] = field(default_factory=list)
    row_heights: list[int] | None = None
    col_widths: list[int] | None = None
    cell_spans: dict[tuple[int, int], tuple[int, int]] | None = None
    cell_border_fill_ids: dict[tuple[int, int], int] | None = None
    cell_margins: dict[tuple[int, int], dict[str, int]] | None = None
    cell_vertical_aligns: dict[tuple[int, int], str] | None = None
    table_border_fill_id: int = 1
    cell_spacing: int | None = None
    table_margins: dict[str, int] | None = None
    layout: dict[str, str] = field(default_factory=dict)
    out_margins: dict[str, int] = field(default_factory=dict)
    page_break: str | None = None
    repeat_header: bool | None = None


@dataclass
class Picture:
    name: str
    data: bytes
    extension: str | None = None
    width: int = 7200
    height: int = 7200
    shape_comment: str | None = None
    layout: dict[str, str] = field(default_factory=dict)
    out_margins: dict[str, int] = field(default_factory=dict)
    rotation: dict[str, str] = field(default_factory=dict)
    image_adjustment: dict[str, str] = field(default_factory=dict)
    crop: dict[str, int] = field(default_factory=dict)
    line_color: str = "#000000"
    line_width: int = 0
    line_style: dict[str, str] = field(default_factory=dict)
    original_width: int | None = None
    original_height: int | None = None
    current_width: int | None = None
    current_height: int | None = None
    has_size_node: bool = True
    size_attributes: dict[str, str] = field(default_factory=dict)
    has_line_node: bool = False
    has_position_node: bool = True
    has_out_margin_node: bool = True


@dataclass
class Hyperlink:
    target: str
    display_text: str | None = None
    metadata_fields: list[str | int] = field(default_factory=list)


@dataclass
class Bookmark:
    name: str


@dataclass
class Field:
    field_type: str
    display_text: str | None = None
    name: str | None = None
    parameters: dict[str, str | int] = field(default_factory=dict)
    editable: bool = False
    dirty: bool = False
    native_field_type: str | None = None

    @property
    def semantic_field_type(self) -> str:
        return _semantic_field_type_from_native(self.native_field_type or self.field_type, self.parameters)

    @property
    def effective_native_field_type(self) -> str:
        return self.native_field_type or _HWP_NATIVE_FIELD_TYPE_BY_SEMANTIC_TYPE.get(self.semantic_field_type, self.field_type)

    @property
    def is_mail_merge(self) -> bool:
        return self.semantic_field_type == "MAILMERGE"

    @property
    def is_calculation(self) -> bool:
        return self.semantic_field_type == "FORMULA"

    @property
    def is_cross_reference(self) -> bool:
        return self.semantic_field_type == "CROSSREF"

    @property
    def is_doc_property(self) -> bool:
        return self.semantic_field_type == "DOCPROPERTY"

    @property
    def is_date(self) -> bool:
        return self.semantic_field_type == "DATE"

    def get_parameter(self, name: str) -> str | None:
        value = self.parameters.get(name)
        return None if value is None else str(value)

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

    def set_parameter(self, name: str, value: str | int) -> None:
        self.parameters[str(name)] = value

    def set_name(self, value: str) -> None:
        self.name = value

    def set_display_text(self, value: str) -> None:
        self.display_text = value

    def set_field_type(self, value: str) -> None:
        normalized = value.strip()
        semantic = _semantic_field_type_from_native(normalized, self.parameters)
        if normalized.startswith("%"):
            self.native_field_type = normalized
            self.field_type = semantic
            return
        self.field_type = semantic if semantic != normalized else normalized
        if self.native_field_type is None or _semantic_field_type_from_native(self.native_field_type, self.parameters) != self.field_type:
            self.native_field_type = _HWP_NATIVE_FIELD_TYPE_BY_SEMANTIC_TYPE.get(self.field_type)

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


@dataclass
class AutoNumber:
    kind: str = "newNum"
    number: int | str = 1
    number_type: str = "PAGE"

    def set_number(self, value: int | str) -> None:
        self.number = value

    def set_number_type(self, value: str) -> None:
        self.number_type = value


@dataclass
class Note:
    kind: str
    text: str
    number: int | None = None
    nested_blocks: list[object] = field(default_factory=list)

    def set_text(self, value: str) -> None:
        self.text = value

    def set_number(self, value: int | None) -> None:
        self.number = value


@dataclass
class Equation:
    script: str
    width: int = 4800
    height: int = 2300
    shape_comment: str | None = None
    text_color: str = "#000000"
    base_unit: int = 1100
    font: str = "HYhwpEQ"
    layout: dict[str, str] = field(default_factory=dict)
    out_margins: dict[str, int] = field(default_factory=dict)
    rotation: dict[str, str] = field(default_factory=dict)


@dataclass
class Shape:
    kind: str
    text: str = ""
    width: int = 12000
    height: int = 3200
    fill_color: str = "#FFFFFF"
    line_color: str = "#000000"
    shape_comment: str | None = None
    layout: dict[str, str] = field(default_factory=dict)
    out_margins: dict[str, int] = field(default_factory=dict)
    rotation: dict[str, str] = field(default_factory=dict)
    text_margins: dict[str, int] = field(default_factory=dict)
    specific_fields: dict[str, object] = field(default_factory=dict)
    line_width: int = 33
    line_style: dict[str, str] = field(default_factory=dict)
    has_line_node: bool = True
    original_width: int | None = None
    original_height: int | None = None
    current_width: int | None = None
    current_height: int | None = None
    has_size_node: bool = True
    size_attributes: dict[str, str] = field(default_factory=dict)
    has_position_node: bool = True
    has_out_margin_node: bool = True
    has_fill_node: bool = True
    has_text_margin_node: bool = True
    nested_blocks: list[object] = field(default_factory=list)


@dataclass
class Ole:
    name: str
    data: bytes
    width: int = 42001
    height: int = 13501
    shape_comment: str | None = None
    object_type: str = "EMBEDDED"
    draw_aspect: str = "CONTENT"
    has_moniker: bool = False
    eq_baseline: int = 0
    line_color: str = "#000000"
    line_width: int = 0
    line_style: dict[str, str] = field(default_factory=dict)
    layout: dict[str, str] = field(default_factory=dict)
    out_margins: dict[str, int] = field(default_factory=dict)
    rotation: dict[str, str] = field(default_factory=dict)
    extent: dict[str, int] = field(default_factory=dict)
    original_width: int | None = None
    original_height: int | None = None
    current_width: int | None = None
    current_height: int | None = None
    has_size_node: bool = True
    size_attributes: dict[str, str] = field(default_factory=dict)
    has_position_node: bool = True
    has_out_margin_node: bool = True


@dataclass
class Form:
    label: str = ""
    form_type: str = "INPUT"
    name: str | None = None
    value: str | None = None
    checked: bool = False
    items: list[str] = field(default_factory=list)
    editable: bool = True
    locked: bool = False
    placeholder: str | None = None

    @property
    def text(self) -> str:
        return self.label

    def set_label(self, value: str) -> None:
        self.label = value

    def set_text(self, value: str) -> None:
        self.label = value

    def set_form_type(self, value: str) -> None:
        self.form_type = value

    def set_name(self, value: str | None) -> None:
        self.name = value

    def set_value(self, value: str | None) -> None:
        self.value = value

    def set_checked(self, value: bool) -> None:
        self.checked = value

    def set_items(self, values: list[str]) -> None:
        self.items = list(values)

    def set_editable(self, value: bool) -> None:
        self.editable = value

    def set_locked(self, value: bool) -> None:
        self.locked = value

    def set_placeholder(self, value: str | None) -> None:
        self.placeholder = value


@dataclass
class Memo:
    text: str
    author: str | None = None
    memo_id: str | None = None
    anchor_id: str | None = None
    order: int | None = None
    visible: bool = True

    def set_text(self, value: str) -> None:
        self.text = value

    def set_author(self, value: str | None) -> None:
        self.author = value

    def set_memo_id(self, value: str | None) -> None:
        self.memo_id = value

    def set_anchor_id(self, value: str | None) -> None:
        self.anchor_id = value

    def set_order(self, value: int | None) -> None:
        self.order = value

    def set_visible(self, value: bool) -> None:
        self.visible = value


@dataclass
class Chart:
    title: str = ""
    chart_type: str = "BAR"
    categories: list[str] = field(default_factory=list)
    series: list[dict[str, object]] = field(default_factory=list)
    data_ref: str | None = None
    legend_visible: bool = True
    width: int = 12000
    height: int = 3200
    shape_comment: str | None = None
    layout: dict[str, str] = field(default_factory=dict)
    out_margins: dict[str, int] = field(default_factory=dict)
    rotation: dict[str, str] = field(default_factory=dict)

    def set_title(self, value: str) -> None:
        self.title = value

    def set_chart_type(self, value: str) -> None:
        self.chart_type = value

    def set_categories(self, values: list[str]) -> None:
        self.categories = list(values)

    def set_series(self, values: list[dict[str, object]]) -> None:
        self.series = [dict(value) for value in values]

    def set_data_ref(self, value: str | None) -> None:
        self.data_ref = value

    def set_legend_visible(self, value: bool) -> None:
        self.legend_visible = value


HancomBlock = (
    Paragraph
    | Table
    | Picture
    | Hyperlink
    | Bookmark
    | Field
    | AutoNumber
    | Note
    | Equation
    | Shape
    | Ole
    | Form
    | Memo
    | Chart
)


@dataclass
class HeaderFooter:
    kind: str
    text: str
    apply_page_type: str = "BOTH"
    nested_blocks: list[object] = field(default_factory=list)

    def set_text(self, value: str) -> None:
        self.text = value

    def set_apply_page_type(self, value: str) -> None:
        self.apply_page_type = value


@dataclass
class StyleDefinition:
    name: str
    style_id: str | None = None
    english_name: str | None = None
    style_type: str = "PARA"
    para_pr_id: str | None = None
    char_pr_id: str | None = None
    next_style_id: str | None = None
    lang_id: str = "1042"
    lock_form: bool = False
    native_hwp_payload: bytes | None = None


@dataclass
class ParagraphStyle:
    style_id: str | None = None
    alignment_horizontal: str | None = None
    alignment_vertical: str | None = None
    line_spacing: int | str | None = None
    line_spacing_type: str | None = None
    margins: dict[str, str] = field(default_factory=dict)
    tab_pr_id: str | None = None
    snap_to_grid: str | None = None
    condense: str | None = None
    font_line_height: str | None = None
    suppress_line_numbers: str | None = None
    checked: str | None = None
    heading: dict[str, str] = field(default_factory=dict)
    break_setting: dict[str, str] = field(default_factory=dict)
    auto_spacing: dict[str, str] = field(default_factory=dict)
    native_hwp_payload: bytes | None = None


@dataclass
class CharacterStyle:
    style_id: str | None = None
    text_color: str | None = None
    shade_color: str | None = None
    height: int | str | None = None
    use_font_space: str | None = None
    use_kerning: str | None = None
    sym_mark: str | None = None
    font_refs: dict[str, str] = field(default_factory=dict)
    native_hwp_payload: bytes | None = None


@dataclass
class NumberingDefinition:
    style_id: str | None = None
    formats: list[str] = field(default_factory=list)
    para_heads: list[int] = field(default_factory=list)
    para_head_flags: list[dict[str, bool]] = field(default_factory=list)
    width_adjusts: list[int] = field(default_factory=list)
    text_offsets: list[int] = field(default_factory=list)
    char_pr_ids: list[int | None] = field(default_factory=list)
    start_numbers: list[int] = field(default_factory=list)
    unknown_short: int = 0
    unknown_short_bits: dict[str, bool] = field(default_factory=dict)
    native_hwp_payload: bytes | None = None


@dataclass
class BulletDefinition:
    style_id: str | None = None
    flags: int | None = None
    flag_bits: dict[str, bool] = field(default_factory=dict)
    bullet_char: str = "\uf0a1"
    width_adjust: int = 0
    text_offset: int = 50
    char_pr_id: int | None = None
    unknown_tail: bytes | None = None
    native_hwp_payload: bytes | None = None


@dataclass
class MemoShapeDefinition:
    memo_shape_id: str | None = None
    width: int | None = None
    line_width: int | None = None
    line_type: str | None = None
    line_color: str | None = None
    fill_color: str | None = None
    active_color: str | None = None
    memo_type: str | None = None
    native_hwp_payload: bytes | None = None

    def configure(
        self,
        *,
        memo_shape_id: str | None = None,
        width: int | None = None,
        line_width: int | None = None,
        line_type: str | None = None,
        line_color: str | None = None,
        fill_color: str | None = None,
        active_color: str | None = None,
        memo_type: str | None = None,
    ) -> None:
        if memo_shape_id is not None:
            self.memo_shape_id = memo_shape_id
        if width is not None:
            self.width = width
        if line_width is not None:
            self.line_width = line_width
        if line_type is not None:
            self.line_type = line_type
        if line_color is not None:
            self.line_color = line_color
        if fill_color is not None:
            self.fill_color = fill_color
        if active_color is not None:
            self.active_color = active_color
        if memo_type is not None:
            self.memo_type = memo_type


@dataclass
class SectionSettings:
    page_width: int | None = None
    page_height: int | None = None
    landscape: str | None = None
    margins: dict[str, int] = field(default_factory=dict)
    page_border_fills: list[dict[str, str | int]] = field(default_factory=list)
    visibility: dict[str, str] = field(default_factory=dict)
    grid: dict[str, int] = field(default_factory=dict)
    start_numbers: dict[str, str] = field(default_factory=dict)
    page_numbers: list[dict[str, str]] = field(default_factory=list)
    page_number_positions: list[tuple[int | None, int | None]] = field(default_factory=list)
    footnote_pr: dict[str, object] = field(default_factory=dict)
    endnote_pr: dict[str, object] = field(default_factory=dict)
    line_number_shape: dict[str, str] = field(default_factory=dict)
    numbering_shape_id: str | None = None
    memo_shape_id: str | None = None

    def set_page_size(
        self,
        *,
        width: int | None = None,
        height: int | None = None,
        landscape: str | None = None,
    ) -> None:
        if width is not None:
            self.page_width = width
        if height is not None:
            self.page_height = height
        if landscape is not None:
            self.landscape = landscape

    def set_margins(self, **values: int) -> None:
        self.margins.update({key: int(value) for key, value in values.items()})

    def set_visibility(
        self,
        *,
        hide_first_header: bool | None = None,
        hide_first_footer: bool | None = None,
        hide_first_master_page: bool | None = None,
        border: str | None = None,
        fill: str | None = None,
        hide_first_page_num: bool | None = None,
        hide_first_empty_line: bool | None = None,
        show_line_number: bool | None = None,
    ) -> None:
        values = {
            "hideFirstHeader": hide_first_header,
            "hideFirstFooter": hide_first_footer,
            "hideFirstMasterPage": hide_first_master_page,
            "border": border,
            "fill": fill,
            "hideFirstPageNum": hide_first_page_num,
            "hideFirstEmptyLine": hide_first_empty_line,
            "showLineNumber": show_line_number,
        }
        for key, value in values.items():
            if value is None:
                continue
            if isinstance(value, bool):
                self.visibility[key] = "1" if value else "0"
            else:
                self.visibility[key] = str(value)

    def set_grid(
        self,
        *,
        line_grid: int | None = None,
        char_grid: int | None = None,
        wonggoji_format: bool | None = None,
    ) -> None:
        if line_grid is not None:
            self.grid["lineGrid"] = int(line_grid)
        if char_grid is not None:
            self.grid["charGrid"] = int(char_grid)
        if wonggoji_format is not None:
            self.grid["wonggojiFormat"] = 1 if wonggoji_format else 0

    def set_start_numbers(
        self,
        *,
        page_starts_on: str | None = None,
        page: int | str | None = None,
        pic: int | str | None = None,
        tbl: int | str | None = None,
        equation: int | str | None = None,
    ) -> None:
        values = {
            "pageStartsOn": page_starts_on,
            "page": page,
            "pic": pic,
            "tbl": tbl,
            "equation": equation,
        }
        for key, value in values.items():
            if value is not None:
                self.start_numbers[key] = str(value)

    def set_page_numbers(self, page_numbers: list[dict[str, str]]) -> None:
        self.page_numbers = [dict(page_number) for page_number in page_numbers]
        self.page_number_positions = []

    def set_note_settings(
        self,
        *,
        footnote_pr: dict[str, object] | None = None,
        endnote_pr: dict[str, object] | None = None,
    ) -> None:
        if footnote_pr is not None:
            self.footnote_pr = dict(footnote_pr)
        if endnote_pr is not None:
            self.endnote_pr = dict(endnote_pr)

    def set_memo_shape_id(self, value: str | None) -> None:
        self.memo_shape_id = value


ParagraphNode = Paragraph
TableNode = Table
PictureNode = Picture
HyperlinkNode = Hyperlink
BookmarkNode = Bookmark
FieldNode = Field
AutoNumberNode = AutoNumber
NoteNode = Note
EquationNode = Equation
ShapeNode = Shape
OleNode = Ole
HeaderFooterNode = HeaderFooter
StyleDefinitionNode = StyleDefinition
ParagraphStyleNode = ParagraphStyle
CharacterStyleNode = CharacterStyle
NumberingDefinitionNode = NumberingDefinition
BulletDefinitionNode = BulletDefinition
MemoShapeDefinitionNode = MemoShapeDefinition
SectionSettingsNode = SectionSettings


@dataclass
class HancomSection:
    settings: SectionSettings = field(default_factory=SectionSettings)
    header_footer_blocks: list[HeaderFooter] = field(default_factory=list)
    blocks: list[HancomBlock] = field(default_factory=list)


class HancomDocument:
    def __init__(
        self,
        *,
        metadata: HancomMetadata | None = None,
        sections: list[HancomSection] | None = None,
        source_format: str | None = None,
        converter: HancomConverter | None = None,
    ) -> None:
        self.metadata = metadata or HancomMetadata()
        self.sections = sections if sections is not None else [HancomSection()]
        self.source_format = source_format
        self._converter = converter
        self._temp_dir: Path | None = None
        self.style_definitions: list[StyleDefinition] = []
        self.paragraph_styles: list[ParagraphStyle] = []
        self.character_styles: list[CharacterStyle] = []
        self.numbering_definitions: list[NumberingDefinition] = []
        self.bullet_definitions: list[BulletDefinition] = []
        self.memo_shape_definitions: list[MemoShapeDefinition] = []

    @classmethod
    def blank(cls, *, converter: HancomConverter | None = None) -> "HancomDocument":
        return cls(converter=converter)

    def bridge(self, *, converter: HancomConverter | None = None):
        from .bridge import HwpHwpxBridge

        return HwpHwpxBridge.from_hancom(self, converter=converter or self._converter)

    @classmethod
    def read_hwpx(cls, path: str | Path, *, converter: HancomConverter | None = None) -> "HancomDocument":
        return cls.from_hwpx_document(HwpxDocument.open(path), converter=converter)

    @classmethod
    def read_hwp(cls, path: str | Path, *, converter: HancomConverter | None = None) -> "HancomDocument":
        return cls.from_hwp_document(HwpDocument.open(path, converter=converter), converter=converter)

    @classmethod
    def from_hwpx_document(
        cls,
        document: HwpxDocument,
        *,
        converter: HancomConverter | None = None,
    ) -> "HancomDocument":
        instance = cls(
            metadata=HancomMetadata.from_hwpx_metadata(document.metadata()),
            sections=[],
            source_format="hwpx",
            converter=converter,
        )
        instance.style_definitions = [_extract_style_definition(style) for style in document.styles()]
        instance.paragraph_styles = [_extract_paragraph_style(style) for style in document.paragraph_styles()]
        instance.character_styles = [_extract_character_style(style) for style in document.character_styles()]
        instance.memo_shape_definitions = [_extract_hwpx_memo_shape(memo_shape) for memo_shape in document.memo_shapes()]
        for section in document.sections:
            instance.sections.append(_extract_hwpx_section(section))
        if not instance.sections:
            instance.sections.append(HancomSection())
        return instance

    @classmethod
    def from_hwp_document(
        cls,
        document: HwpDocument,
        *,
        converter: HancomConverter | None = None,
    ) -> "HancomDocument":
        effective_converter = converter or getattr(document, "_converter", None)
        docinfo_model = document.docinfo_model()
        style_definitions = _extract_hwp_style_definitions(docinfo_model)
        sections = [_extract_hwp_section(section) for section in document.sections()]
        if not sections:
            sections = [HancomSection()]
        preview_text = document.preview_text().strip()
        title = preview_text or (document.source_path.stem if document.source_path else None)
        metadata = HancomMetadata(title=title)
        instance = cls(metadata=metadata, sections=sections, source_format="hwp", converter=effective_converter)
        instance.style_definitions = style_definitions
        instance.paragraph_styles = _build_hwp_placeholder_paragraph_styles(docinfo_model, sections, style_definitions)
        instance.character_styles = _build_hwp_placeholder_character_styles(docinfo_model, style_definitions)
        instance.numbering_definitions = _extract_hwp_numbering_definitions(docinfo_model)
        instance.bullet_definitions = _extract_hwp_bullet_definitions(docinfo_model)
        instance.memo_shape_definitions = _extract_hwp_memo_shape_definitions(docinfo_model)
        _merge_hwp_style_defaults_into_paragraphs(instance)
        return instance

    def append_section(self) -> HancomSection:
        section = HancomSection()
        self.sections.append(section)
        return section

    def append_paragraph(self, text: str, *, section_index: int = 0) -> Paragraph:
        section = self._ensure_section(section_index)
        block = Paragraph(text)
        section.blocks.append(block)
        return block

    def append_table(
        self,
        *,
        rows: int,
        cols: int,
        cell_texts: list[list[str]] | None = None,
        row_heights: list[int] | None = None,
        col_widths: list[int] | None = None,
        cell_spans: dict[tuple[int, int], tuple[int, int]] | None = None,
        cell_border_fill_ids: dict[tuple[int, int], int] | None = None,
        table_border_fill_id: int = 1,
        section_index: int = 0,
    ) -> Table:
        section = self._ensure_section(section_index)
        block = Table(
            rows=rows,
            cols=cols,
            cell_texts=cell_texts or _blank_cell_texts(rows, cols),
            row_heights=row_heights,
            col_widths=col_widths,
            cell_spans=cell_spans,
            cell_border_fill_ids=cell_border_fill_ids,
            table_border_fill_id=table_border_fill_id,
        )
        section.blocks.append(block)
        return block

    def append_picture(
        self,
        name: str,
        data: bytes,
        *,
        extension: str | None = None,
        width: int = 7200,
        height: int = 7200,
        original_width: int | None = None,
        original_height: int | None = None,
        current_width: int | None = None,
        current_height: int | None = None,
        section_index: int = 0,
    ) -> Picture:
        section = self._ensure_section(section_index)
        block = Picture(
            name=name,
            data=data,
            extension=extension,
            width=width,
            height=height,
            original_width=original_width,
            original_height=original_height,
            current_width=current_width,
            current_height=current_height,
        )
        section.blocks.append(block)
        return block

    def append_hyperlink(
        self,
        target: str,
        *,
        display_text: str | None = None,
        metadata_fields: list[str | int] | None = None,
        section_index: int = 0,
    ) -> Hyperlink:
        section = self._ensure_section(section_index)
        block = Hyperlink(target=target, display_text=display_text, metadata_fields=list(metadata_fields or []))
        section.blocks.append(block)
        return block

    def append_bookmark(self, name: str, *, section_index: int = 0) -> Bookmark:
        section = self._ensure_section(section_index)
        block = Bookmark(name=name)
        section.blocks.append(block)
        return block

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
    ) -> Field:
        section = self._ensure_section(section_index)
        normalized_field_type = field_type.strip()
        native_field_type = normalized_field_type if normalized_field_type.startswith("%") else None
        block = Field(
            field_type=_semantic_field_type_from_native(normalized_field_type, parameters)
            if native_field_type is not None
            else normalized_field_type,
            display_text=display_text,
            name=name,
            parameters=dict(parameters or {}),
            editable=editable,
            dirty=dirty,
            native_field_type=native_field_type,
        )
        section.blocks.append(block)
        return block

    def append_mail_merge_field(
        self,
        field_name: str,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        section_index: int = 0,
    ) -> Field:
        block = self.append_field(
            field_type="MAILMERGE",
            display_text=display_text,
            name=field_name,
            parameters={"FieldName": field_name, "MergeField": field_name},
            editable=editable,
            dirty=dirty,
            section_index=section_index,
        )
        block.configure_mail_merge(field_name, display_text=display_text)
        return block

    def append_calculation_field(
        self,
        expression: str,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        section_index: int = 0,
    ) -> Field:
        block = self.append_field(
            field_type="FORMULA",
            display_text=display_text,
            parameters={"Expression": expression, "Command": expression},
            editable=editable,
            dirty=dirty,
            section_index=section_index,
        )
        block.configure_calculation(expression, display_text=display_text)
        return block

    def append_cross_reference(
        self,
        bookmark_name: str,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        section_index: int = 0,
    ) -> Field:
        block = self.append_field(
            field_type="CROSSREF",
            display_text=display_text,
            name=bookmark_name,
            parameters={"BookmarkName": bookmark_name, "Path": bookmark_name},
            editable=editable,
            dirty=dirty,
            section_index=section_index,
        )
        block.configure_cross_reference(bookmark_name, display_text=display_text)
        return block

    def append_doc_property_field(
        self,
        property_name: str,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        section_index: int = 0,
    ) -> Field:
        block = self.append_field(
            field_type="DOCPROPERTY",
            display_text=display_text,
            name=property_name,
            parameters={"FieldName": property_name},
            editable=editable,
            dirty=dirty,
            section_index=section_index,
        )
        block.configure_doc_property(property_name, display_text=display_text)
        return block

    def append_date_field(
        self,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        section_index: int = 0,
    ) -> Field:
        block = self.append_field(
            field_type="DATE",
            display_text=display_text,
            editable=editable,
            dirty=dirty,
            section_index=section_index,
        )
        block.configure_date(display_text=display_text)
        return block

    def append_auto_number(
        self,
        *,
        number: int | str = 1,
        number_type: str = "PAGE",
        kind: str = "newNum",
        section_index: int = 0,
    ) -> AutoNumber:
        section = self._ensure_section(section_index)
        block = AutoNumber(kind=kind, number=number, number_type=number_type)
        section.blocks.append(block)
        return block

    def append_note(
        self,
        text: str,
        *,
        kind: str = "footNote",
        number: int | None = None,
        section_index: int = 0,
    ) -> Note:
        section = self._ensure_section(section_index)
        block = Note(kind=kind, text=text, number=number)
        section.blocks.append(block)
        return block

    def append_form(
        self,
        label: str = "",
        *,
        form_type: str = "INPUT",
        name: str | None = None,
        value: str | None = None,
        checked: bool = False,
        items: list[str] | None = None,
        editable: bool = True,
        locked: bool = False,
        placeholder: str | None = None,
        section_index: int = 0,
    ) -> Form:
        section = self._ensure_section(section_index)
        block = Form(
            label=label,
            form_type=form_type,
            name=name,
            value=value,
            checked=checked,
            items=list(items or []),
            editable=editable,
            locked=locked,
            placeholder=placeholder,
        )
        section.blocks.append(block)
        return block

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
    ) -> Memo:
        section = self._ensure_section(section_index)
        block = Memo(
            text=text,
            author=author,
            memo_id=memo_id,
            anchor_id=anchor_id,
            order=order,
            visible=visible,
        )
        section.blocks.append(block)
        return block

    def append_chart(
        self,
        title: str = "",
        *,
        chart_type: str = "BAR",
        categories: list[str] | None = None,
        series: list[dict[str, object]] | None = None,
        data_ref: str | None = None,
        legend_visible: bool = True,
        width: int = 12000,
        height: int = 3200,
        shape_comment: str | None = None,
        section_index: int = 0,
    ) -> Chart:
        section = self._ensure_section(section_index)
        block = Chart(
            title=title,
            chart_type=chart_type,
            categories=list(categories or []),
            series=[dict(value) for value in (series or [])],
            data_ref=data_ref,
            legend_visible=legend_visible,
            width=width,
            height=height,
            shape_comment=shape_comment,
        )
        section.blocks.append(block)
        return block

    def append_footnote(
        self,
        text: str,
        *,
        number: int | None = None,
        section_index: int = 0,
    ) -> Note:
        return self.append_note(text, kind="footNote", number=number, section_index=section_index)

    def append_endnote(
        self,
        text: str,
        *,
        number: int | None = None,
        section_index: int = 0,
    ) -> Note:
        return self.append_note(text, kind="endNote", number=number, section_index=section_index)

    def append_equation(
        self,
        script: str,
        *,
        width: int = 4800,
        height: int = 2300,
        shape_comment: str | None = None,
        text_color: str = "#000000",
        base_unit: int = 1100,
        font: str = "HYhwpEQ",
        section_index: int = 0,
    ) -> Equation:
        section = self._ensure_section(section_index)
        block = Equation(
            script=script,
            width=width,
            height=height,
            shape_comment=shape_comment,
            text_color=text_color,
            base_unit=base_unit,
            font=font,
        )
        section.blocks.append(block)
        return block

    def append_shape(
        self,
        *,
        kind: str = "rect",
        text: str = "",
        width: int = 12000,
        height: int = 3200,
        fill_color: str = "#FFFFFF",
        line_color: str = "#000000",
        line_width: int = 33,
        shape_comment: str | None = None,
        original_width: int | None = None,
        original_height: int | None = None,
        current_width: int | None = None,
        current_height: int | None = None,
        section_index: int = 0,
    ) -> Shape:
        section = self._ensure_section(section_index)
        block = Shape(
            kind=kind,
            text=text,
            width=width,
            height=height,
            fill_color=fill_color,
            line_color=line_color,
            shape_comment=shape_comment,
            line_width=line_width,
            original_width=original_width,
            original_height=original_height,
            current_width=current_width,
            current_height=current_height,
        )
        section.blocks.append(block)
        return block

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
        original_width: int | None = None,
        original_height: int | None = None,
        current_width: int | None = None,
        current_height: int | None = None,
        section_index: int = 0,
    ) -> Ole:
        section = self._ensure_section(section_index)
        block = Ole(
            name=name,
            data=data,
            width=width,
            height=height,
            shape_comment=shape_comment,
            object_type=object_type,
            draw_aspect=draw_aspect,
            has_moniker=has_moniker,
            eq_baseline=eq_baseline,
            line_color=line_color,
            line_width=line_width,
            original_width=original_width,
            original_height=original_height,
            current_width=current_width,
            current_height=current_height,
        )
        section.blocks.append(block)
        return block

    def append_header(
        self,
        text: str,
        *,
        apply_page_type: str = "BOTH",
        section_index: int = 0,
    ) -> HeaderFooter:
        section = self._ensure_section(section_index)
        block = HeaderFooter(kind="header", text=text, apply_page_type=apply_page_type)
        section.header_footer_blocks.append(block)
        return block

    def append_footer(
        self,
        text: str,
        *,
        apply_page_type: str = "BOTH",
        section_index: int = 0,
    ) -> HeaderFooter:
        section = self._ensure_section(section_index)
        block = HeaderFooter(kind="footer", text=text, apply_page_type=apply_page_type)
        section.header_footer_blocks.append(block)
        return block

    def append_style(
        self,
        name: str,
        *,
        style_id: str | None = None,
        english_name: str | None = None,
        style_type: str = "PARA",
        para_pr_id: str | None = None,
        char_pr_id: str | None = None,
        next_style_id: str | None = None,
        lang_id: str = "1042",
        lock_form: bool = False,
        native_hwp_payload: bytes | None = None,
    ) -> StyleDefinition:
        node = StyleDefinition(
            name=name,
            style_id=style_id,
            english_name=english_name,
            style_type=style_type,
            para_pr_id=para_pr_id,
            char_pr_id=char_pr_id,
            next_style_id=next_style_id,
            lang_id=lang_id,
            lock_form=lock_form,
            native_hwp_payload=native_hwp_payload,
        )
        self.style_definitions.append(node)
        return node

    def append_paragraph_style(
        self,
        *,
        style_id: str | None = None,
        alignment_horizontal: str | None = None,
        alignment_vertical: str | None = None,
        line_spacing: int | str | None = None,
        line_spacing_type: str | None = None,
        margins: dict[str, str] | None = None,
        tab_pr_id: str | None = None,
        snap_to_grid: str | None = None,
        condense: str | None = None,
        font_line_height: str | None = None,
        suppress_line_numbers: str | None = None,
        checked: str | None = None,
        heading: dict[str, str] | None = None,
        break_setting: dict[str, str] | None = None,
        auto_spacing: dict[str, str] | None = None,
    ) -> ParagraphStyle:
        node = ParagraphStyle(
            style_id=style_id,
            alignment_horizontal=alignment_horizontal,
            alignment_vertical=alignment_vertical,
            line_spacing=line_spacing,
            line_spacing_type=line_spacing_type,
            margins=dict(margins or {}),
            tab_pr_id=tab_pr_id,
            snap_to_grid=snap_to_grid,
            condense=condense,
            font_line_height=font_line_height,
            suppress_line_numbers=suppress_line_numbers,
            checked=checked,
            heading=dict(heading or {}),
            break_setting=dict(break_setting or {}),
            auto_spacing=dict(auto_spacing or {}),
        )
        self.paragraph_styles.append(node)
        return node

    def append_character_style(
        self,
        *,
        style_id: str | None = None,
        text_color: str | None = None,
        shade_color: str | None = None,
        height: int | str | None = None,
        use_font_space: str | None = None,
        use_kerning: str | None = None,
        sym_mark: str | None = None,
        font_refs: dict[str, str] | None = None,
    ) -> CharacterStyle:
        node = CharacterStyle(
            style_id=style_id,
            text_color=text_color,
            shade_color=shade_color,
            height=height,
            use_font_space=use_font_space,
            use_kerning=use_kerning,
            sym_mark=sym_mark,
            font_refs=dict(font_refs or {}),
        )
        self.character_styles.append(node)
        return node

    def append_numbering_definition(
        self,
        *,
        style_id: str | None = None,
        formats: list[str] | None = None,
        para_heads: list[int] | None = None,
        para_head_flags: list[dict[str, bool]] | None = None,
        width_adjusts: list[int] | None = None,
        text_offsets: list[int] | None = None,
        char_pr_ids: list[int | None] | None = None,
        start_numbers: list[int] | None = None,
        unknown_short: int = 0,
        unknown_short_bits: dict[str, bool] | None = None,
        native_hwp_payload: bytes | None = None,
    ) -> NumberingDefinition:
        node = NumberingDefinition(
            style_id=style_id,
            formats=list(formats or []),
            para_heads=list(para_heads or []),
            para_head_flags=[dict(value) for value in (para_head_flags or [])],
            width_adjusts=list(width_adjusts or []),
            text_offsets=list(text_offsets or []),
            char_pr_ids=list(char_pr_ids or []),
            start_numbers=list(start_numbers or []),
            unknown_short=unknown_short,
            unknown_short_bits=dict(unknown_short_bits or {}),
            native_hwp_payload=native_hwp_payload,
        )
        self.numbering_definitions.append(node)
        return node

    def append_bullet_definition(
        self,
        *,
        style_id: str | None = None,
        flags: int | None = None,
        flag_bits: dict[str, bool] | None = None,
        bullet_char: str = "\uf0a1",
        width_adjust: int = 0,
        text_offset: int = 50,
        char_pr_id: int | None = None,
        unknown_tail: bytes | None = None,
        native_hwp_payload: bytes | None = None,
    ) -> BulletDefinition:
        node = BulletDefinition(
            style_id=style_id,
            flags=flags,
            flag_bits=dict(flag_bits or {}),
            bullet_char=bullet_char,
            width_adjust=width_adjust,
            text_offset=text_offset,
            char_pr_id=char_pr_id,
            unknown_tail=unknown_tail,
            native_hwp_payload=native_hwp_payload,
        )
        self.bullet_definitions.append(node)
        return node

    def append_memo_shape_definition(
        self,
        *,
        memo_shape_id: str | None = None,
        width: int | None = None,
        line_width: int | None = None,
        line_type: str | None = None,
        line_color: str | None = None,
        fill_color: str | None = None,
        active_color: str | None = None,
        memo_type: str | None = None,
    ) -> MemoShapeDefinition:
        if memo_shape_id is None:
            used_ids = {
                int(node.memo_shape_id)
                for node in self.memo_shape_definitions
                if node.memo_shape_id is not None and str(node.memo_shape_id).isdigit()
            }
            next_id = 0
            while next_id in used_ids:
                next_id += 1
            memo_shape_id = str(next_id)
        node = MemoShapeDefinition(
            memo_shape_id=memo_shape_id,
            width=width,
            line_width=line_width,
            line_type=line_type,
            line_color=line_color,
            fill_color=fill_color,
            active_color=active_color,
            memo_type=memo_type,
        )
        self.memo_shape_definitions.append(node)
        return node

    def to_hwpx_document(self) -> HwpxDocument:
        document = HwpxDocument.blank()
        self.metadata.apply_to_hwpx_document(document)
        _apply_styles(document, self)
        while len(document.sections) < len(self.sections):
            document.add_section(text="")
        for section_index in range(len(self.sections)):
            _reset_hwpx_section(document, section_index)
        for section_index, section in enumerate(self.sections):
            uses_source_paragraphs = _section_uses_source_paragraphs(section)
            _apply_section_settings(
                document,
                section_index,
                section.settings,
                apply_page_numbers=not uses_source_paragraphs,
            )
            if not uses_source_paragraphs:
                _apply_header_footer_blocks(document, section_index, section.header_footer_blocks)
            _write_section_to_hwpx(document, section_index, section)
            if uses_source_paragraphs and section.settings.page_numbers:
                _apply_page_numbers_to_source_paragraphs(document, section_index, section.settings)
        return document

    def write_to_hwpx(self, path: str | Path, *, validate: bool = True) -> Path:
        return self.to_hwpx_document().save(path, validate=validate)

    def to_hwp_document(self, *, converter: HancomConverter | None = None) -> HwpDocument:
        effective_converter = converter or self._converter
        document = HwpDocument.blank(converter=effective_converter)
        document.ensure_section_count(max(len(self.sections), 1))
        preview_text = self.metadata.title or self.metadata.subject or self.metadata.description
        if preview_text:
            document.set_preview_text(preview_text)
        _sync_hwp_docinfo_styles(document, self)
        docinfo = document.docinfo_model()
        paragraph_attribute_limits = (len(docinfo.para_shape_records()), len(docinfo.style_records()))
        for section_index, section in enumerate(self.sections):
            document.apply_section_settings(
                section_index=section_index,
                page_width=section.settings.page_width,
                page_height=section.settings.page_height,
                landscape=section.settings.landscape,
                margins=section.settings.margins or None,
                visibility=section.settings.visibility or None,
                grid=section.settings.grid or None,
                start_numbers=section.settings.start_numbers or None,
                numbering_shape_id=section.settings.numbering_shape_id,
                memo_shape_id=section.settings.memo_shape_id,
            )
            if section.settings.page_border_fills:
                document.apply_section_page_border_fills(
                    section.settings.page_border_fills,
                    section_index=section_index,
                )
            if section.settings.page_numbers:
                document.apply_section_page_numbers(
                    section.settings.page_numbers,
                    section_index=section_index,
                )
            if section.settings.footnote_pr or section.settings.endnote_pr:
                document.apply_section_note_settings(
                    section_index=section_index,
                    footnote_pr=section.settings.footnote_pr or None,
                    endnote_pr=section.settings.endnote_pr or None,
                )
            header_written = False
            footer_written = False
            for block in section.header_footer_blocks:
                if block.kind == "header" and not header_written:
                    header = document.append_header(block.text, apply_page_type=block.apply_page_type, section_index=section_index)
                    if header is not None and block.nested_blocks and hasattr(header, "set_nested_blocks"):
                        header.set_nested_blocks(block.nested_blocks)
                    header_written = True
                elif block.kind == "footer" and not footer_written:
                    footer = document.append_footer(block.text, apply_page_type=block.apply_page_type, section_index=section_index)
                    if footer is not None and block.nested_blocks and hasattr(footer, "set_nested_blocks"):
                        footer.set_nested_blocks(block.nested_blocks)
                    footer_written = True
            binary_document = document.binary_document()
            batch_active = False
            try:
                for block in section.blocks:
                    uses_direct_section_records = isinstance(block, (Table, Picture, Hyperlink))
                    if uses_direct_section_records and batch_active:
                        binary_document.end_section_model_batch()
                        batch_active = False
                    elif not uses_direct_section_records and not batch_active:
                        binary_document.begin_section_model_batch()
                        batch_active = True
                    _append_hancom_block_to_hwp(
                        document,
                        block,
                        section_index=section_index,
                        paragraph_attribute_limits=paragraph_attribute_limits,
                    )
            finally:
                if batch_active:
                    binary_document.end_section_model_batch()
        return document

    def write_to_hwp(self, path: str | Path, *, converter: HancomConverter | None = None) -> Path:
        document = self.to_hwp_document(converter=converter)
        return document.save(path)

    def _ensure_section(self, section_index: int) -> HancomSection:
        while len(self.sections) <= section_index:
            self.sections.append(HancomSection())
        return self.sections[section_index]

    def _ensure_temp_dir(self) -> Path:
        if self._temp_dir is None:
            self._temp_dir = Path(tempfile.mkdtemp(prefix="jakal_hancom_document_"))
        return self._temp_dir


def _blank_cell_texts(rows: int, cols: int) -> list[list[str]]:
    return [["" for _ in range(cols)] for _ in range(rows)]


def _extract_hwp_section(section) -> HancomSection:
    controls_by_paragraph: dict[int, list[object]] = {}
    header_footer_blocks: list[HeaderFooter] = []
    for control in section.controls():
        if getattr(control, "control_id", "") in {"head", "foot"}:
            block = _extract_hwp_header_footer(control)
            if block is not None:
                header_footer_blocks.append(block)
            continue
        controls_by_paragraph.setdefault(control.paragraph_index, []).append(control)

    blocks: list[HancomBlock] = []
    for paragraph in section.model().paragraphs():
        paragraph_controls = controls_by_paragraph.get(paragraph.index, [])
        if paragraph.header.level != 0:
            blocks.extend(_extract_supported_nested_hwp_blocks(paragraph_controls))
            continue
        blocks.extend(_extract_hwp_paragraph_blocks(paragraph, paragraph_controls))
    return HancomSection(
        settings=_extract_hwp_section_settings(section),
        header_footer_blocks=header_footer_blocks,
        blocks=blocks,
    )


def _extract_hwp_paragraph_blocks(paragraph, controls: list[object]) -> list[HancomBlock]:
    if not controls:
        text = _normalize_hwp_text(paragraph.text)
        if not text:
            return []
        return [_mark_hwpx_source_paragraph(_build_hwp_paragraph_block(paragraph, text), paragraph.index, order=0)]

    blocks: list[HancomBlock] = []
    pending_text_parts: list[str] = []
    control_index = 0
    tokens = _tokenize_hwp_paragraph_raw_text(paragraph.raw_text)

    for token_kind, token_value in tokens:
        if token_kind == "text":
            pending_text_parts.append(token_value)
            continue

        match_index = _find_matching_control_index(controls, control_index, token_value)
        if match_index is None:
            continue
        control = controls[match_index]
        control_index = match_index + 1
        block = _extract_hwp_control(control)
        pending_text = "".join(pending_text_parts)
        pending_text_parts.clear()
        normalized_pending = _normalize_hwp_text(pending_text)
        owned_text = _block_owned_text(block)
        if normalized_pending and (owned_text is None or normalized_pending != _normalize_hwp_text(owned_text)):
            blocks.append(_build_hwp_paragraph_block(paragraph, normalized_pending))
        if block is not None:
            blocks.append(block)

    trailing_text = _normalize_hwp_text("".join(pending_text_parts))
    if trailing_text:
        blocks.append(_build_hwp_paragraph_block(paragraph, trailing_text))

    for control in controls[control_index:]:
        block = _extract_hwp_control(control)
        if block is not None:
            blocks.append(block)
    return [
        _mark_hwpx_source_paragraph(block, paragraph.index, order=index)
        for index, block in enumerate(blocks)
    ]


def _extract_supported_nested_hwp_blocks(controls: list[object]) -> list[HancomBlock]:
    blocks: list[HancomBlock] = []
    for control in controls:
        block = _extract_hwp_control(control)
        if isinstance(block, (AutoNumber, Bookmark, Note)):
            blocks.append(block)
    return blocks


def _extract_hwp_nested_paragraph_blocks(paragraph, controls: list[object], paragraph_index: int) -> list[HancomBlock]:
    blocks = _extract_hwp_paragraph_blocks(paragraph, controls)
    return [
        _mark_hwpx_nested_paragraph(block, paragraph_index, order=index)
        for index, block in enumerate(blocks)
    ]


def _extract_hwp_section_settings(section) -> SectionSettings:
    settings = section.document.binary_document().section_page_settings(section.index)
    section_definition = section.document.binary_document().section_definition_settings(section.index)
    page_border_fills = section.document.binary_document().section_page_border_fills(section.index)
    page_numbers = section.document.binary_document().section_page_numbers(section.index)
    note_settings = section.document.binary_document().section_note_settings(section.index)
    page_width = settings.get("page_width")
    page_height = settings.get("page_height")
    margins = {
        key: value
        for key, value in settings.items()
        if key in {"left", "right", "top", "bottom", "header", "footer", "gutter"}
    }
    return SectionSettings(
        page_width=page_width or None,
        page_height=page_height or None,
        landscape="LANDSCAPE" if (page_width or 0) > (page_height or 0) else "WIDELY",
        margins=margins,
        page_border_fills=page_border_fills,
        visibility=dict(section_definition.get("visibility", {})),
        grid=dict(section_definition.get("grid", {})),
        start_numbers=dict(section_definition.get("start_numbers", {})),
        page_numbers=page_numbers,
        footnote_pr=dict(note_settings.get("footNotePr", {})),
        endnote_pr=dict(note_settings.get("endNotePr", {})),
        numbering_shape_id=_normalize_hwp_setting_scalar(section_definition.get("numbering_shape_id")),
        memo_shape_id=_normalize_hwp_setting_scalar(section_definition.get("memo_shape_id")),
    )


def _normalize_hwp_text(value: str) -> str:
    return _xml_safe_scalar(value.replace("\r", "")).strip()


def _drop_surrogate_codepoints(value: str) -> str:
    return "".join(character for character in value if not 0xD800 <= ord(character) <= 0xDFFF)


def _xml_safe_scalar(value: object | None) -> str:
    if value is None:
        return ""
    text = str(value)
    def _is_xml_char(character: str) -> bool:
        codepoint = ord(character)
        return (
            codepoint in {0x09, 0x0A, 0x0D}
            or 0x20 <= codepoint <= 0xD7FF
            or 0xE000 <= codepoint <= 0xFFFD
            or 0x10000 <= codepoint <= 0x10FFFF
        )
    return "".join(
        character
        for character in text
        if _is_xml_char(character)
    )


def _build_hwp_paragraph_block(paragraph, text: str) -> Paragraph:
    header = getattr(paragraph, "header", None)
    style_id = getattr(header, "style_id", None) if header is not None else None
    para_shape_id = getattr(header, "para_shape_id", None) if header is not None else None
    return Paragraph(
        text=text,
        style_id=str(style_id) if style_id is not None else None,
        para_pr_id=str(para_shape_id) if para_shape_id is not None else None,
        hwp_para_shape_id=para_shape_id,
        hwp_style_id=style_id,
        hwp_split_flags=getattr(header, "split_flags", None),
        hwp_control_mask=getattr(header, "control_mask", None),
    )


def _tokenize_hwp_paragraph_raw_text(raw_text: str) -> list[tuple[str, str]]:
    encoded = _drop_surrogate_codepoints(raw_text).encode("utf-16-le")
    if not encoded:
        return []
    units = list(struct.unpack("<" + "H" * (len(encoded) // 2), encoded))
    tokens: list[tuple[str, str]] = []
    text_buffer: list[str] = []
    index = 0
    while index < len(units):
        value = units[index]
        if value >= 32:
            text_buffer.append(chr(value))
            index += 1
            continue

        if value in (9, 10, 13):
            text_buffer.append(chr(value))

        if (
            index + 7 < len(units)
            and units[index + 3] == 0
            and units[index + 4] == 0
            and units[index + 5] == 0
            and units[index + 6] == 0
        ):
            if text_buffer:
                tokens.append(("text", "".join(text_buffer)))
                text_buffer = []
            token_bytes = struct.pack("<8H", *units[index : index + 8])
            control_id = token_bytes[2:6][::-1].decode("latin1", errors="replace")
            tokens.append(("control", control_id))
            index += 8
            continue

        index += 1

    if text_buffer:
        tokens.append(("text", "".join(text_buffer)))
    return tokens


def _find_matching_control_index(controls: list[object], start_index: int, control_id: str) -> int | None:
    for index in range(start_index, len(controls)):
        if getattr(controls[index], "control_id", "") == control_id:
            return index
    return None


def _extract_hwp_control(control) -> HancomBlock | None:
    from .hwp_document import (
        HwpAutoNumberObject,
        HwpBookmarkObject,
        HwpChartObject,
        HwpEquationObject,
        HwpFieldObject,
        HwpFormObject,
        HwpHyperlinkObject,
        HwpMemoObject,
        HwpNoteObject,
        HwpOleObject,
        HwpPictureObject,
        HwpShapeObject,
        HwpTableObject,
    )

    if isinstance(control, HwpTableObject):
        cells = control.cells()
        cell_texts = control.cell_text_matrix()
        cell_spans = {
            (cell.row, cell.column): (cell.row_span, cell.col_span)
            for cell in cells
            if cell.row_span != 1 or cell.col_span != 1
        }
        cell_border_fill_ids = {
            (cell.row, cell.column): cell.border_fill_id
            for cell in cells
        }
        cell_margins = {
            (cell.row, cell.column): {
                "left": cell.margins[0],
                "right": cell.margins[1],
                "top": cell.margins[2],
                "bottom": cell.margins[3],
            }
            for cell in cells
        }
        return Table(
            rows=control.row_count,
            cols=control.column_count,
            cell_texts=cell_texts,
            row_heights=list(control.row_heights),
            col_widths=_extract_hwp_table_column_widths(control),
            cell_spans=cell_spans or None,
            cell_border_fill_ids=cell_border_fill_ids or None,
            cell_margins=cell_margins or None,
            table_border_fill_id=control.table_border_fill_id,
            cell_spacing=control.cell_spacing,
            table_margins={
                "left": control.table_margins[0],
                "right": control.table_margins[1],
                "top": control.table_margins[2],
                "bottom": control.table_margins[3],
            },
        )
    if isinstance(control, HwpPictureObject):
        size = control.size() if hasattr(control, "size") else {"width": 7200, "height": 7200}
        path = control.bindata_path()
        line_style = dict(control.line_style()) if hasattr(control, "line_style") else {}
        return Picture(
            name=Path(path).name,
            data=control.binary_data(),
            extension=control.extension,
            width=size.get("width", 7200),
            height=size.get("height", 7200),
            shape_comment=getattr(control, "shape_comment", "") or None,
            layout=dict(control.layout()) if hasattr(control, "layout") else {},
            out_margins=dict(control.out_margins()) if hasattr(control, "out_margins") else {},
            rotation=dict(control.rotation()) if hasattr(control, "rotation") else {},
            image_adjustment=dict(control.image_adjustment()) if hasattr(control, "image_adjustment") else {},
            crop=dict(control.crop()) if hasattr(control, "crop") else {},
            line_color=getattr(control, "line_color", "#000000"),
            line_width=int(getattr(control, "line_width", 0)),
            line_style=line_style,
        )
    if isinstance(control, HwpHyperlinkObject):
        return Hyperlink(
            target=control.url,
            display_text=control.display_text or None,
            metadata_fields=list(control.metadata_fields),
        )
    if isinstance(control, HwpBookmarkObject):
        return Bookmark(name=control.name or "")
    if isinstance(control, HwpAutoNumberObject):
        return AutoNumber(
            kind=control.kind,
            number=control.number or 1,
            number_type=control.number_type or "PAGE",
        )
    if isinstance(control, HwpFormObject):
        return Form(
            label=control.label,
            form_type=control.form_type,
            name=control.name,
            value=control.value,
            checked=control.checked,
            items=list(control.items),
            editable=control.editable,
            locked=control.locked,
            placeholder=control.placeholder,
        )
    if isinstance(control, HwpMemoObject):
        return Memo(
            text=control.text,
            author=control.author,
            memo_id=control.memo_id,
            anchor_id=control.anchor_id,
            order=control.order,
            visible=control.visible,
        )
    if isinstance(control, HwpNoteObject):
        number = int(control.number) if control.number and str(control.number).isdigit() else None
        return Note(kind=control.kind, text=control.text, number=number)
    if isinstance(control, HwpOleObject):
        size = control.size()
        line_style = dict(control.line_style()) if hasattr(control, "line_style") else {}
        return Ole(
            name=control.text or "embedded.ole",
            data=control.binary_data(),
            width=size.get("width", 42001),
            height=size.get("height", 13501),
            shape_comment=getattr(control, "shape_comment", "") or None,
            object_type=getattr(control, "object_type", "EMBEDDED"),
            draw_aspect=getattr(control, "draw_aspect", "CONTENT"),
            has_moniker=bool(getattr(control, "has_moniker", False)),
            eq_baseline=int(getattr(control, "eq_baseline", 0)),
            line_color=getattr(control, "line_color", "#000000"),
            line_width=int(getattr(control, "line_width", 0)),
            line_style=line_style,
            layout=dict(control.layout()) if hasattr(control, "layout") else {},
            out_margins=dict(control.out_margins()) if hasattr(control, "out_margins") else {},
            rotation=dict(control.rotation()) if hasattr(control, "rotation") else {},
            extent=dict(control.extent()) if hasattr(control, "extent") else {},
        )
    if isinstance(control, HwpChartObject):
        size = control.size()
        return Chart(
            title=control.title,
            chart_type=control.chart_type,
            categories=list(control.categories),
            series=[dict(value) for value in control.series],
            data_ref=control.data_ref,
            legend_visible=control.legend_visible,
            width=size.get("width", 12000),
            height=size.get("height", 3200),
            shape_comment=getattr(control, "shape_comment", "") or None,
            layout=dict(control.layout()) if hasattr(control, "layout") else {},
            out_margins=dict(control.out_margins()) if hasattr(control, "out_margins") else {},
            rotation=dict(control.rotation()) if hasattr(control, "rotation") else {},
        )
    if isinstance(control, HwpShapeObject):
        size = control.size()
        line_style = dict(control.line_style()) if hasattr(control, "line_style") else {}
        return Shape(
            kind=control.kind,
            text=control.text,
            width=size.get("width", 12000),
            height=size.get("height", 3200),
            fill_color=getattr(control, "fill_color", "#FFFFFF"),
            line_color=getattr(control, "line_color", "#000000"),
            shape_comment=getattr(control, "shape_comment", "") or None,
            layout=dict(control.layout()) if hasattr(control, "layout") else {},
            out_margins=dict(control.out_margins()) if hasattr(control, "out_margins") else {},
            rotation=dict(control.rotation()) if hasattr(control, "rotation") else {},
            text_margins=dict(control.text_margins()) if hasattr(control, "text_margins") else {},
            specific_fields=dict(control.specific_fields()) if hasattr(control, "specific_fields") else {},
            line_width=int(getattr(control, "line_width", 33)),
            line_style=line_style,
        )
    if isinstance(control, HwpEquationObject):
        size = control.size() if hasattr(control, "size") else {"width": 4800, "height": 2300}
        return Equation(
            script=control.script,
            width=size.get("width", 4800),
            height=size.get("height", 2300),
            shape_comment=getattr(control, "shape_comment", "") or None,
            font=getattr(control, "font", "HYhwpEQ"),
            layout=dict(control.layout()) if hasattr(control, "layout") else {},
            out_margins=dict(control.out_margins()) if hasattr(control, "out_margins") else {},
            rotation=dict(control.rotation()) if hasattr(control, "rotation") else {},
        )
    if isinstance(control, HwpFieldObject):
        return Field(
            field_type=control.field_type,
            display_text=_normalize_hwp_text(control.display_text) or None,
            name=control.name,
            parameters=dict(control.parameters),
            editable=control.editable,
            dirty=control.dirty,
            native_field_type=control.native_field_type,
        )
    return None


def _extract_hwp_table_column_widths(table) -> list[int]:
    widths = [0] * table.column_count
    for cell in table.cells():
        if cell.col_span == 1 and widths[cell.column] == 0:
            widths[cell.column] = cell.width
    fallback = max((value for value in widths if value > 0), default=1000)
    for cell in table.cells():
        if cell.col_span <= 1:
            continue
        missing = [index for index in range(cell.column, cell.column + cell.col_span) if widths[index] == 0]
        if not missing:
            continue
        distributed = max(cell.width // cell.col_span, 1)
        for index in missing:
            widths[index] = distributed
    return [value or fallback for value in widths]


def _normalize_str_map(values: dict[str, object] | None) -> dict[str, str]:
    if not values:
        return {}
    return {str(key): str(value) for key, value in values.items() if value is not None}


def _normalize_int_map(values: dict[str, object] | None) -> dict[str, int]:
    if not values:
        return {}
    result: dict[str, int] = {}
    for key, value in values.items():
        if value is None:
            continue
        try:
            result[str(key)] = int(value)
        except (TypeError, ValueError):
            continue
    return result


def _optional_hwpx_int_attr(node: etree._Element | None, attr_name: str) -> int | None:
    if node is None:
        return None
    value = node.get(attr_name)
    if value in (None, ""):
        return None
    try:
        return int(value)
    except ValueError:
        return None


def _extract_hwpx_graphic_size_fields(element: etree._Element, *, width: int, height: int) -> dict[str, int | None]:
    original_size = element.xpath("./hp:orgSz", namespaces=NS)
    current_size = element.xpath("./hp:curSz", namespaces=NS)
    original = original_size[0] if original_size else None
    current = current_size[0] if current_size else None
    original_width = _optional_hwpx_int_attr(original, "width")
    original_height = _optional_hwpx_int_attr(original, "height")
    current_width = _optional_hwpx_int_attr(current, "width")
    current_height = _optional_hwpx_int_attr(current, "height")
    return {
        "original_width": original_width if original_width != width else None,
        "original_height": original_height if original_height != height else None,
        "current_width": current_width if current_width != width else None,
        "current_height": current_height if current_height != height else None,
    }


def _extract_hwpx_graphic_size_attributes(element: etree._Element) -> dict[str, str]:
    nodes = element.xpath("./hp:sz", namespaces=NS)
    if not nodes:
        return {}
    return {str(key): str(value) for key, value in nodes[0].attrib.items()}


def _apply_hwpx_graphic_size_node(wrapper, *, has_size_node: bool, size_attributes: dict[str, str]) -> None:
    nodes = wrapper.element.xpath("./hp:sz", namespaces=NS)
    if not has_size_node:
        for node in nodes:
            parent = node.getparent()
            if parent is not None:
                parent.remove(node)
        if nodes:
            wrapper.section.mark_modified()
        return
    if not size_attributes:
        return
    node = nodes[0] if nodes else etree.SubElement(wrapper.element, qname("hp", "sz"))
    for key, value in size_attributes.items():
        node.set(str(key), str(value))
    wrapper.section.mark_modified()


def _remove_hwpx_child_nodes(wrapper, xpath: str) -> None:
    nodes = wrapper.element.xpath(xpath, namespaces=NS)
    for node in nodes:
        parent = node.getparent()
        if parent is not None:
            parent.remove(node)
    if nodes:
        wrapper.section.mark_modified()


def _apply_hwpx_optional_graphic_nodes(
    wrapper,
    *,
    has_position_node: bool,
    has_out_margin_node: bool,
) -> None:
    if not has_position_node:
        _remove_hwpx_child_nodes(wrapper, "./hp:pos")
    if not has_out_margin_node:
        _remove_hwpx_child_nodes(wrapper, "./hp:outMargin")


def _hwpx_layout_kwargs(values: dict[str, object] | None) -> dict[str, object]:
    if not values:
        return {}
    mapping = {
        "textWrap": "text_wrap",
        "textFlow": "text_flow",
        "treatAsChar": "treat_as_char",
        "affectLSpacing": "affect_line_spacing",
        "flowWithText": "flow_with_text",
        "allowOverlap": "allow_overlap",
        "holdAnchorAndSO": "hold_anchor_and_so",
        "vertRelTo": "vert_rel_to",
        "horzRelTo": "horz_rel_to",
        "vertAlign": "vert_align",
        "horzAlign": "horz_align",
        "vertOffset": "vert_offset",
        "horzOffset": "horz_offset",
    }
    bool_fields = {
        "treatAsChar",
        "affectLSpacing",
        "flowWithText",
        "allowOverlap",
        "holdAnchorAndSO",
    }
    result: dict[str, object] = {}
    for source, target in mapping.items():
        value = values.get(source)
        if value in (None, ""):
            continue
        if source in bool_fields:
            coerced = _coerce_bool(value)
            if coerced is None:
                continue
            result[target] = coerced
            continue
        result[target] = value
    return result


def _hwpx_rotation_kwargs(values: dict[str, object] | None) -> dict[str, object]:
    if not values:
        return {}
    mapping = {
        "angle": "angle",
        "centerX": "center_x",
        "centerY": "center_y",
        "rotateimage": "rotate_image",
    }
    return {
        target: value
        for source, target in mapping.items()
        if (value := values.get(source)) not in (None, "")
    }


def _hwpx_line_style_kwargs(values: dict[str, object] | None) -> dict[str, object]:
    if not values:
        return {}
    mapping = {
        "color": "color",
        "width": "width",
        "style": "style",
        "endCap": "end_cap",
        "headStyle": "head_style",
        "tailStyle": "tail_style",
        "headfill": "head_fill",
        "tailfill": "tail_fill",
        "headSz": "head_size",
        "tailSz": "tail_size",
        "outlineStyle": "outline_style",
        "alpha": "alpha",
    }
    bool_fields = {"headfill", "tailfill"}
    result: dict[str, object] = {}
    for source, target in mapping.items():
        value = values.get(source)
        if value in (None, ""):
            continue
        if source in bool_fields:
            coerced = _coerce_bool(value)
            if coerced is None:
                continue
            result[target] = coerced
            continue
        result[target] = value
    return result


def _extract_hwpx_shape_points(node: etree._Element) -> list[dict[str, int]]:
    points: list[dict[str, int]] = []
    for child in node:
        if etree.QName(child).localname.startswith("pt"):
            try:
                points.append(
                    {
                        "x": int(child.get("x", "0")),
                        "y": int(child.get("y", "0")),
                    }
                )
            except ValueError:
                continue
    return points


def _extract_hwpx_shape_specific_fields(shape) -> dict[str, object]:
    points = _extract_hwpx_shape_points(shape.element)
    if shape.kind in {"line", "connectLine"} and len(points) >= 2:
        return {
            "start": points[0],
            "end": points[1],
        }
    if shape.kind == "polygon" and points:
        return {
            "point_count": len(points),
            "points": points,
        }
    return {}


_NESTED_TEXT_CONTROL_NAMES = {
    "tbl",
    "pic",
    "ole",
    "equation",
    "chart",
    "rect",
    "line",
    "ellipse",
    "arc",
    "polygon",
    "curve",
    "connectLine",
    "textart",
    "container",
    "fieldBegin",
    "fieldEnd",
    "header",
    "footer",
    "footNote",
    "endNote",
    "hiddenComment",
    "autoNum",
    "newNum",
    "bookmark",
}


def _extract_hwpx_direct_text(scope: etree._Element) -> str:
    chunks: list[str] = []
    for text_node in scope.xpath(".//hp:t", namespaces=NS):
        parent = text_node.getparent()
        skip = False
        while parent is not None and parent is not scope:
            if etree.QName(parent).localname in _NESTED_TEXT_CONTROL_NAMES:
                skip = True
                break
            parent = parent.getparent()
        if not skip:
            chunks.append(text_node.text or "")
    return "".join(chunks)


def _extract_hwpx_shape_direct_text(shape) -> str:
    direct_text = shape.element.get("text")
    if direct_text is not None:
        return direct_text
    draw_text_nodes = shape.element.xpath("./hp:drawText", namespaces=NS)
    if not draw_text_nodes:
        return ""
    return _extract_hwpx_direct_text(draw_text_nodes[0])


def _set_hwpx_shape_points(node: etree._Element, points: list[dict[str, int]]) -> None:
    for child in list(node):
        if etree.QName(child).localname.startswith("pt"):
            node.remove(child)
    for index, point in enumerate(points):
        child = etree.SubElement(node, qname("hc", f"pt{index}"))
        child.set("x", str(int(point.get("x", 0))))
        child.set("y", str(int(point.get("y", 0))))


def _apply_hwp_shape_specific_fields(shape, block: Shape) -> None:
    fields = dict(block.specific_fields)
    if not fields:
        return
    if isinstance(shape, (HwpLineShapeObject, HwpConnectLineShapeObject)):
        start = fields.get("start")
        end = fields.get("end")
        attributes = fields.get("attributes")
        if start is not None or end is not None or attributes is not None:
            shape.set_endpoints(start=start, end=end, attributes=int(attributes) if attributes is not None else None)
        return
    if isinstance(shape, HwpRectangleShapeObject) and "corner_radius" in fields:
        shape.set_corner_radius(int(fields["corner_radius"]))
        return
    if isinstance(shape, HwpEllipseShapeObject):
        kwargs = {key: fields[key] for key in ("arc_flags", "center", "axis1", "axis2", "start", "end", "start_2", "end_2") if key in fields}
        if kwargs:
            shape.set_geometry(**kwargs)
        return
    if isinstance(shape, HwpArcShapeObject):
        kwargs = {key: fields[key] for key in ("arc_type", "center", "axis1", "axis2") if key in fields}
        if kwargs:
            shape.set_geometry(**kwargs)
        return
    if isinstance(shape, HwpPolygonShapeObject) and isinstance(fields.get("points"), list):
        shape.set_points(fields["points"], trailing_payload=fields.get("trailing_payload") if isinstance(fields.get("trailing_payload"), bytes) else None)
        return
    if isinstance(shape, HwpCurveShapeObject) and isinstance(fields.get("raw_payload"), (bytes, bytearray)):
        shape.set_raw_payload(bytes(fields["raw_payload"]))
        return
    if isinstance(shape, HwpContainerShapeObject) and isinstance(fields.get("raw_payload"), (bytes, bytearray)):
        shape.set_raw_payload(bytes(fields["raw_payload"]))
        return
    if isinstance(shape, HwpTextArtShapeObject):
        if "text" in fields:
            shape.set_text(str(fields["text"]))
        if "font_name" in fields or "font_style" in fields:
            shape.set_style(
                font_name=str(fields["font_name"]) if "font_name" in fields else None,
                font_style=str(fields["font_style"]) if "font_style" in fields else None,
            )


def _apply_hwpx_shape_specific_fields(shape, block: Shape) -> None:
    fields = dict(block.specific_fields)
    if not fields:
        return
    if block.kind in {"line", "connectLine"}:
        start = fields.get("start", {"x": 0, "y": 0})
        end = fields.get("end", {"x": 0, "y": 0})
        if isinstance(start, dict) and isinstance(end, dict):
            _set_hwpx_shape_points(shape.element, [start, end])
            shape.section.mark_modified()
        return
    if block.kind == "polygon" and isinstance(fields.get("points"), list):
        points = [point for point in fields["points"] if isinstance(point, dict)]
        if points:
            _set_hwpx_shape_points(shape.element, points)
            shape.section.mark_modified()


def _apply_hwpx_table_block(table, block: Table) -> None:
    if block.layout:
        table.set_layout(**_hwpx_layout_kwargs(block.layout))
    if block.out_margins:
        table.set_out_margins(**block.out_margins)
    if block.table_margins:
        table.set_in_margins(**block.table_margins)
    if block.cell_spacing is not None:
        table.set_cell_spacing(block.cell_spacing)
    table.set_table_border_fill_id(block.table_border_fill_id)
    if block.page_break:
        table.set_page_break(block.page_break)
    if block.repeat_header is not None:
        table.set_repeat_header(block.repeat_header)
    for (row, column), border_fill_id in (block.cell_border_fill_ids or {}).items():
        table.cell(row, column).set_border_fill_id(border_fill_id)
    for (row, column), margins in (block.cell_margins or {}).items():
        table.cell(row, column).set_margins(**_normalize_int_map(margins))
    for (row, column), vertical_align in (block.cell_vertical_aligns or {}).items():
        table.cell(row, column).set_vertical_align(vertical_align)


def _apply_hwpx_picture_block(picture, block: Picture) -> None:
    if any(
        value is not None
        for value in (block.original_width, block.original_height, block.current_width, block.current_height)
    ):
        picture.set_size(
            width=block.width,
            height=block.height,
            original_width=block.original_width,
            original_height=block.original_height,
            current_width=block.current_width,
            current_height=block.current_height,
        )
    _apply_hwpx_graphic_size_node(
        picture,
        has_size_node=block.has_size_node,
        size_attributes=block.size_attributes,
    )
    if block.shape_comment:
        picture.shape_comment = block.shape_comment
    if block.layout:
        picture.set_layout(**_hwpx_layout_kwargs(block.layout))
    if block.out_margins:
        picture.set_out_margins(**block.out_margins)
    if block.rotation:
        picture.set_rotation(**_hwpx_rotation_kwargs(block.rotation))
    if block.image_adjustment:
        picture.set_image_adjustment(**block.image_adjustment)
    if block.crop:
        picture.set_crop(**block.crop)
    line_style_kwargs = _hwpx_line_style_kwargs(block.line_style)
    if line_style_kwargs:
        picture.set_line_style(**line_style_kwargs)
    elif block.has_line_node or block.line_width > 0 or block.line_color != "#000000":
        picture.set_line_style(color=block.line_color, width=block.line_width, style="SOLID" if block.line_width > 0 else "NONE")
    _apply_hwpx_optional_graphic_nodes(
        picture,
        has_position_node=block.has_position_node,
        has_out_margin_node=block.has_out_margin_node,
    )


def _apply_hwpx_equation_block(equation, block: Equation) -> None:
    if block.layout:
        equation.set_layout(**_hwpx_layout_kwargs(block.layout))
    if block.out_margins:
        equation.set_out_margins(**block.out_margins)
    if block.rotation:
        equation.set_rotation(**_hwpx_rotation_kwargs(block.rotation))


def _apply_hwpx_shape_block(shape, block: Shape) -> None:
    if any(
        value is not None
        for value in (block.original_width, block.original_height, block.current_width, block.current_height)
    ):
        shape.set_size(
            width=block.width,
            height=block.height,
            original_width=block.original_width,
            original_height=block.original_height,
            current_width=block.current_width,
            current_height=block.current_height,
        )
    _apply_hwpx_graphic_size_node(
        shape,
        has_size_node=block.has_size_node,
        size_attributes=block.size_attributes,
    )
    if block.kind == "textart" and block.has_text_margin_node and block.text:
        _ensure_hwpx_shape_draw_text(shape, block.text)
    if block.shape_comment:
        shape.shape_comment = block.shape_comment
    line_style_kwargs = _hwpx_line_style_kwargs(block.line_style)
    if line_style_kwargs:
        shape.set_line_style(**line_style_kwargs)
    else:
        shape.set_line_style(color=block.line_color, width=block.line_width, style="SOLID" if block.line_width > 0 else "NONE")
    if block.layout:
        shape.set_layout(**_hwpx_layout_kwargs(block.layout))
    if block.out_margins:
        shape.set_out_margins(**block.out_margins)
    if block.rotation:
        shape.set_rotation(**_hwpx_rotation_kwargs(block.rotation))
    if block.text_margins:
        shape.set_text_margins(**block.text_margins)
    _apply_hwpx_shape_specific_fields(shape, block)
    if not block.has_fill_node:
        _remove_hwpx_child_nodes(shape, "./hc:fillBrush")
    if not block.has_line_node:
        _remove_hwpx_child_nodes(shape, "./hp:lineShape")
    if not block.has_text_margin_node:
        _remove_hwpx_child_nodes(shape, "./hp:drawText/hp:textMargin")
    _apply_hwpx_optional_graphic_nodes(
        shape,
        has_position_node=block.has_position_node,
        has_out_margin_node=block.has_out_margin_node,
    )


def _ensure_hwpx_shape_draw_text(shape, text: str) -> None:
    if shape.element.xpath("./hp:drawText", namespaces=NS):
        return
    draw_text = etree.SubElement(shape.element, qname("hp", "drawText"))
    draw_text.set("lastWidth", str(shape.size().get("width", 12000)))
    draw_text.set("name", "")
    draw_text.set("editable", "0")

    sub_list = etree.SubElement(draw_text, qname("hp", "subList"))
    sub_list.set("id", "")
    sub_list.set("textDirection", "HORIZONTAL")
    sub_list.set("lineWrap", "BREAK")
    sub_list.set("vertAlign", "CENTER")
    sub_list.set("linkListIDRef", "0")
    sub_list.set("linkListNextIDRef", "0")
    sub_list.set("textWidth", "0")
    sub_list.set("textHeight", "0")
    sub_list.set("hasTextRef", "0")
    sub_list.set("hasNumRef", "0")

    paragraph = etree.SubElement(sub_list, qname("hp", "p"))
    paragraph.set("id", "0")
    paragraph.set("paraPrIDRef", "0")
    paragraph.set("styleIDRef", "0")
    paragraph.set("pageBreak", "0")
    paragraph.set("columnBreak", "0")
    paragraph.set("merged", "0")
    run = etree.SubElement(paragraph, qname("hp", "run"))
    run.set("charPrIDRef", "0")
    text_node = etree.SubElement(run, qname("hp", "t"))
    text_node.text = text

    text_margin = etree.SubElement(draw_text, qname("hp", "textMargin"))
    text_margin.set("left", "283")
    text_margin.set("right", "283")
    text_margin.set("top", "283")
    text_margin.set("bottom", "283")
    shape.section.mark_modified()


def _apply_hwpx_ole_block(ole, block: Ole) -> None:
    if any(
        value is not None
        for value in (block.original_width, block.original_height, block.current_width, block.current_height)
    ):
        ole.set_size(
            width=block.width,
            height=block.height,
            original_width=block.original_width,
            original_height=block.original_height,
            current_width=block.current_width,
            current_height=block.current_height,
        )
    _apply_hwpx_graphic_size_node(
        ole,
        has_size_node=block.has_size_node,
        size_attributes=block.size_attributes,
    )
    if block.shape_comment:
        ole.shape_comment = block.shape_comment
    if block.layout:
        ole.set_layout(**_hwpx_layout_kwargs(block.layout))
    if block.out_margins:
        ole.set_out_margins(**block.out_margins)
    if block.rotation:
        ole.set_rotation(**_hwpx_rotation_kwargs(block.rotation))
    if block.extent:
        ole.set_extent(**block.extent)
    line_style_kwargs = _hwpx_line_style_kwargs(block.line_style)
    if line_style_kwargs:
        ole.set_line_style(**line_style_kwargs)
    else:
        ole.set_line_style(color=block.line_color, width=block.line_width, style="SOLID" if block.line_width > 0 else "NONE")
    _apply_hwpx_optional_graphic_nodes(
        ole,
        has_position_node=block.has_position_node,
        has_out_margin_node=block.has_out_margin_node,
    )


def _parse_hwp_field_parameters(control) -> dict[str, str | int]:
    data_node = next((child for child in control.control_node.children if child.tag_id == TAG_CTRL_DATA), None)
    if data_node is None:
        return {}
    decoded = data_node.payload.decode("utf-16-le", errors="ignore")
    parameters: dict[str, str | int] = {}
    for entry in decoded.split(";"):
        if not entry or "=" not in entry:
            continue
        key, value = entry.split("=", 1)
        if value.isdigit():
            parameters[key] = int(value)
        else:
            parameters[key] = value
    return parameters


def _block_owned_text(block: HancomBlock | None) -> str | None:
    if block is None:
        return None
    if isinstance(block, Hyperlink):
        return block.display_text
    if isinstance(block, Field):
        return block.display_text
    if isinstance(block, Form):
        return block.label
    if isinstance(block, Memo):
        return block.text
    if isinstance(block, Chart):
        return block.title
    if isinstance(block, Shape):
        return block.text
    if isinstance(block, Ole):
        return block.name
    return None


def _block_owned_text_for_write(block: HancomBlock | HeaderFooter | None) -> str | None:
    if block is None or getattr(block, "source_text_owned_by_container", False):
        return None
    if isinstance(block, Hyperlink):
        return block.display_text
    if isinstance(block, Field):
        return block.display_text
    return None


def _is_owned_text_shadow_paragraph(block: object, next_block: object | None) -> bool:
    if not isinstance(block, Paragraph):
        return False
    if not block.text:
        return False
    owned_text = _block_owned_text_for_write(next_block)  # type: ignore[arg-type]
    if block.text != owned_text:
        return False
    return (
        getattr(block, "source_paragraph_index", None) == getattr(next_block, "source_paragraph_index", None)
        and getattr(block, "source_order", None) == getattr(next_block, "source_order", None)
    )


def _append_hancom_block_to_hwp(
    document: HwpDocument,
    block: HancomBlock,
    *,
    section_index: int,
    paragraph_attribute_limits: tuple[int, int] | None = None,
) -> None:
    if isinstance(block, Paragraph):
        if block.text:
            para_shape_id, style_id, split_flags, control_mask = _resolve_hwp_paragraph_attributes(
                document,
                block,
                paragraph_attribute_limits=paragraph_attribute_limits,
            )
            document.append_paragraph(
                block.text,
                section_index=section_index,
                para_shape_id=para_shape_id,
                style_id=style_id,
                split_flags=split_flags,
                control_mask=control_mask,
            )
        return
    if isinstance(block, Table):
        table = document.append_table(
            rows=block.rows,
            cols=block.cols,
            cell_texts=block.cell_texts,
            row_heights=block.row_heights,
            col_widths=block.col_widths,
            cell_spans=block.cell_spans,
            cell_border_fill_ids=block.cell_border_fill_ids,
            table_border_fill_id=block.table_border_fill_id,
            section_index=section_index,
        )
        if table is None:
            table = document.section(section_index).tables()[-1]
        if block.col_widths:
            table.set_column_widths(block.col_widths)
        if block.cell_spacing is not None:
            table.set_cell_spacing(block.cell_spacing)
        if block.table_margins:
            table.set_table_margins(**block.table_margins)
        for (row, column), border_fill_id in (block.cell_border_fill_ids or {}).items():
            table.set_cell_border_fill_id(row, column, border_fill_id)
        for (row, column), margins in (block.cell_margins or {}).items():
            table.set_cell_margins(row, column, **_normalize_int_map(margins))
        return
    if isinstance(block, Picture):
        picture = document.append_picture(
            block.data,
            extension=block.extension,
            width=block.width,
            height=block.height,
            section_index=section_index,
        )
        if picture is None:
            picture = document.section(section_index).pictures()[-1]
        if block.shape_comment:
            picture.set_shape_comment(block.shape_comment)
        if block.layout:
            picture.set_layout(**_hwpx_layout_kwargs(block.layout))
        if block.out_margins:
            picture.set_out_margins(**block.out_margins)
        if block.rotation:
            picture.set_rotation(**_hwpx_rotation_kwargs(block.rotation))
        if block.image_adjustment:
            picture.set_image_adjustment(**block.image_adjustment)
        if block.crop:
            picture.set_crop(**block.crop)
        line_style_kwargs = _hwpx_line_style_kwargs(block.line_style)
        if line_style_kwargs:
            picture.set_line_style(**line_style_kwargs)
        elif block.line_color or block.line_width:
            picture.set_line_style(color=block.line_color, width=block.line_width)
        return
    if isinstance(block, Hyperlink):
        document.append_hyperlink(
            block.target,
            text=_block_owned_text_for_write(block) or None,
            metadata_fields=block.metadata_fields or None,
            section_index=section_index,
        )
        return
    if isinstance(block, Bookmark):
        if block.name:
            document.append_bookmark(block.name, section_index=section_index)
        return
    if isinstance(block, Field):
        document.append_field(
            field_type=block.effective_native_field_type,
            display_text=_block_owned_text_for_write(block),
            name=block.name,
            parameters=block.parameters,
            editable=block.editable,
            dirty=block.dirty,
            section_index=section_index,
        )
        return
    if isinstance(block, Form):
        document.append_form(
            block.label,
            form_type=block.form_type,
            name=block.name,
            value=block.value,
            checked=block.checked,
            items=block.items,
            editable=block.editable,
            locked=block.locked,
            placeholder=block.placeholder,
            section_index=section_index,
        )
        return
    if isinstance(block, Memo):
        document.append_memo(
            block.text,
            author=block.author,
            memo_id=block.memo_id,
            anchor_id=block.anchor_id,
            order=block.order,
            visible=block.visible,
            section_index=section_index,
        )
        return
    if isinstance(block, AutoNumber):
        document.append_auto_number(number=block.number, number_type=block.number_type, kind=block.kind, section_index=section_index)
        return
    if isinstance(block, Note):
        document.append_note(block.text, kind=block.kind, number=block.number, section_index=section_index)
        return
    if isinstance(block, Equation):
        equation = document.append_equation(
            block.script,
            width=block.width,
            height=block.height,
            font=block.font,
            shape_comment=block.shape_comment,
            section_index=section_index,
        )
        if equation is None:
            equation = document.section(section_index).equations()[-1]
        if block.layout:
            equation.set_layout(**_hwpx_layout_kwargs(block.layout))
        if block.out_margins:
            equation.set_out_margins(**block.out_margins)
        if block.rotation:
            equation.set_rotation(**_hwpx_rotation_kwargs(block.rotation))
        return
    if isinstance(block, Shape):
        kind = block.kind if block.kind in {"line", "connectLine", "rect", "ellipse", "arc", "polygon", "curve", "container", "textart"} else "rect"
        shape = document.append_shape(
            kind=kind,
            text=block.text,
            width=block.width,
            height=block.height,
            fill_color=block.fill_color,
            line_color=block.line_color,
            shape_comment=block.shape_comment,
            section_index=section_index,
        )
        if shape is None:
            shape = document.section(section_index).shapes()[-1]
        _apply_hwp_shape_specific_fields(shape, block)
        if block.layout:
            shape.set_layout(**_hwpx_layout_kwargs(block.layout))
        if block.out_margins:
            shape.set_out_margins(**block.out_margins)
        if block.rotation:
            shape.set_rotation(**_hwpx_rotation_kwargs(block.rotation))
        line_style_kwargs = _hwpx_line_style_kwargs(block.line_style)
        if line_style_kwargs:
            shape.set_line_style(**line_style_kwargs)
        if block.text_margins and hasattr(shape, "set_text_margins"):
            shape.set_text_margins(**block.text_margins)
        return
    if isinstance(block, Chart):
        chart = document.append_chart(
            block.title,
            chart_type=block.chart_type,
            categories=block.categories,
            series=block.series,
            data_ref=block.data_ref,
            legend_visible=block.legend_visible,
            width=block.width,
            height=block.height,
            shape_comment=block.shape_comment,
            section_index=section_index,
        )
        if chart is None:
            chart = document.section(section_index).charts()[-1]
        if block.layout:
            chart.set_layout(**_hwpx_layout_kwargs(block.layout))
        if block.out_margins:
            chart.set_out_margins(**block.out_margins)
        if block.rotation:
            chart.set_rotation(**_hwpx_rotation_kwargs(block.rotation))
        return
    if isinstance(block, Ole):
        ole = document.append_ole(
            block.name,
            block.data,
            width=block.width,
            height=block.height,
            shape_comment=block.shape_comment,
            object_type=block.object_type,
            draw_aspect=block.draw_aspect,
            has_moniker=block.has_moniker,
            eq_baseline=block.eq_baseline,
            line_color=block.line_color,
            line_width=block.line_width,
            section_index=section_index,
        )
        if ole is None:
            ole = document.section(section_index).oles()[-1]
        if block.layout:
            ole.set_layout(**_hwpx_layout_kwargs(block.layout))
        if block.out_margins:
            ole.set_out_margins(**block.out_margins)
        if block.rotation:
            ole.set_rotation(**_hwpx_rotation_kwargs(block.rotation))
        if block.extent:
            ole.set_extent(**block.extent)
        line_style_kwargs = _hwpx_line_style_kwargs(block.line_style)
        if line_style_kwargs:
            ole.set_line_style(**line_style_kwargs)
        elif block.line_color or block.line_width:
            ole.set_line_style(color=block.line_color, width=block.line_width)
        return


def _resolve_hwp_paragraph_attributes(
    document: HwpDocument,
    block: Paragraph,
    *,
    paragraph_attribute_limits: tuple[int, int] | None = None,
) -> tuple[int, int, int, int]:
    if paragraph_attribute_limits is None:
        docinfo = document.docinfo_model()
        para_shape_count = len(docinfo.para_shape_records())
        style_count = len(docinfo.style_records())
    else:
        para_shape_count, style_count = paragraph_attribute_limits
    para_shape_id = block.hwp_para_shape_id
    style_id = block.hwp_style_id
    if para_shape_id is None:
        para_shape_id = _parse_optional_int(block.para_pr_id)
    if style_id is None:
        style_id = _parse_optional_int(block.style_id)
    resolved_para_shape_id = (
        para_shape_id
        if para_shape_id is not None and 0 <= para_shape_id < para_shape_count and para_shape_id <= 0xFFFF
        else 0
    )
    resolved_style_id = (
        style_id
        if style_id is not None and 0 <= style_id < style_count and style_id <= 0xFF
        else 0
    )
    return (
        resolved_para_shape_id,
        resolved_style_id,
        block.hwp_split_flags or 0,
        block.hwp_control_mask or 0,
    )


def _parse_optional_int(value: str | None) -> int | None:
    if value is None:
        return None
    stripped = value.strip()
    return int(stripped) if stripped.isdigit() else None


def _normalize_hwp_setting_scalar(value: object | None) -> str | None:
    if value is None:
        return None
    return str(value)


def _extract_hwp_header_footer(control) -> HeaderFooter | None:
    text = getattr(control, "text", "")
    if not text:
        text = ""
    kind = getattr(control, "kind", "")
    if kind not in {"header", "footer"}:
        return None
    nested_blocks = _extract_hwp_header_footer_nested_blocks(control)
    if not text and not nested_blocks:
        return None
    return HeaderFooter(
        kind=kind,
        text=text,
        apply_page_type=getattr(control, "apply_page_type", "BOTH") or "BOTH",
        nested_blocks=nested_blocks,
    )


def _extract_hwp_header_footer_nested_blocks(control) -> list[HancomBlock]:
    from .hwp_document import _build_control_wrapper

    blocks: list[HancomBlock] = []
    paragraph_index = 0
    section_model = control.document.section_model(control.section_index)
    for child in control.control_node.children:
        if not isinstance(child, ParagraphHeaderRecord):
            continue
        paragraph = SectionParagraphModel(
            section_index=control.section_index,
            index=paragraph_index,
            header=child,
        )
        control_ordinal = 0
        controls: list[object] = []
        for nested_child in child.children:
            if nested_child.tag_id != TAG_CTRL_HEADER:
                continue
            wrapper = _build_control_wrapper(control.document, section_model, paragraph, nested_child, control_ordinal)
            if wrapper is not None:
                controls.append(wrapper)
                control_ordinal += 1
        blocks.extend(_extract_hwp_nested_paragraph_blocks(paragraph, controls, paragraph_index))
        paragraph_index += 1
    return blocks


def _extract_hwpx_section(section) -> HancomSection:
    blocks: list[HancomBlock] = []
    section_settings = _extract_section_settings(section)
    header_footer_blocks = _extract_header_footer_blocks(section)
    paragraphs = section.root_element.xpath("./hp:p", namespaces=NS)
    for paragraph_index, paragraph in enumerate(paragraphs):
        paragraph_blocks = _extract_hwpx_paragraph_blocks(section, paragraph, paragraph_index)
        if not paragraph_blocks and not _hwpx_paragraph_has_header_footer(paragraph):
            paragraph_blocks.append(_mark_hwpx_source_paragraph(_build_hwpx_paragraph_block(paragraph, ""), paragraph_index))
        blocks.extend(paragraph_blocks)
    return HancomSection(settings=section_settings, header_footer_blocks=header_footer_blocks, blocks=blocks)


def _hwpx_paragraph_has_header_footer(paragraph: etree._Element) -> bool:
    return bool(paragraph.xpath("./hp:run/hp:ctrl/hp:header | ./hp:run/hp:ctrl/hp:footer", namespaces=NS))


def _extract_hwpx_paragraph_blocks(section, paragraph: etree._Element, paragraph_index: int) -> list[HancomBlock]:
    blocks: list[HancomBlock] = []
    memo_blocks: list[Memo] = []
    memo_carriers: list[Memo] = []
    skipped_field_id: str | None = None
    seen_nodes: set[str] = set()

    for run in paragraph.xpath("./hp:run", namespaces=NS):
        for child in list(run):
            local_name = etree.QName(child).localname
            if local_name == "t":
                if skipped_field_id is None:
                    text = (child.text or "").replace("\r", "")
                    if text:
                        blocks.append(
                            _mark_hwpx_source_paragraph(
                                _build_hwpx_paragraph_block(paragraph, text, char_pr_id=run.get("charPrIDRef")),
                                paragraph_index,
                                order=_source_order_for_element(paragraph, child),
                            )
                        )
                continue

            candidates = list(child) if local_name == "ctrl" else [child]
            for node in candidates:
                node_local_name = etree.QName(node).localname
                if skipped_field_id is not None:
                    if node_local_name == "fieldEnd" and node.get("beginIDRef") == skipped_field_id:
                        skipped_field_id = None
                    continue

                extracted = _extract_hwpx_run_node(section, node)
                if extracted is None:
                    continue
                seen_nodes.add(_hwpx_node_key(node))
                caption_text = _extract_hwpx_caption_text(node)
                if caption_text:
                    blocks.append(
                        _mark_hwpx_source_paragraph(
                            _build_hwpx_paragraph_block(paragraph, caption_text, char_pr_id=run.get("charPrIDRef")),
                            paragraph_index,
                            order=_source_order_for_element(paragraph, node),
                        )
                    )
                extracted_blocks = extracted if isinstance(extracted, list) else [extracted]
                for block in extracted_blocks:
                    owned_text = _hwpx_owned_inline_text(block)
                    if owned_text:
                        blocks.append(
                            _mark_hwpx_source_paragraph(
                                _build_hwpx_paragraph_block(paragraph, owned_text, char_pr_id=run.get("charPrIDRef")),
                                paragraph_index,
                                order=_source_order_for_element(paragraph, node),
                            )
                        )
                    _mark_hwpx_source_paragraph(block, paragraph_index, order=_source_order_for_element(paragraph, node))
                    if isinstance(block, Memo) and _memo_requires_hwpx_carrier(block):
                        memo_carriers.append(block)
                    elif isinstance(block, Memo):
                        memo_blocks.append(block)
                    else:
                        blocks.append(block)

                if node_local_name == "fieldBegin":
                    skipped_field_id = node.get("id") or node.get("fieldid")

    if memo_blocks or memo_carriers:
        for memo in _merge_hwpx_memos_with_carriers(memo_blocks, memo_carriers):
            blocks.append(_mark_hwpx_source_paragraph(memo, paragraph_index))
    blocks.extend(_extract_missing_hwpx_descendant_blocks(section, paragraph, paragraph_index, seen_nodes))
    return _sort_hwpx_blocks_by_source_order(blocks)


def _sort_hwpx_blocks_by_source_order(blocks: list[HancomBlock]) -> list[HancomBlock]:
    return [
        block
        for _index, block in sorted(
            enumerate(blocks),
            key=lambda item: (
                getattr(item[1], "source_order", 10**9),
                item[0],
            ),
        )
    ]


def _sort_hwpx_nested_blocks_by_source_order(blocks: list[HancomBlock]) -> list[HancomBlock]:
    return [
        block
        for _index, block in sorted(
            enumerate(blocks),
            key=lambda item: (
                getattr(item[1], "nested_paragraph_index", 10**9),
                getattr(item[1], "nested_source_order", 10**9),
                item[0],
            ),
        )
    ]


def _mark_hwpx_nested_paragraph(block, paragraph_index: int, *, order: int | None = None):
    setattr(block, "nested_paragraph_index", paragraph_index)
    if order is not None:
        setattr(block, "nested_source_order", order)
    return block


def _extract_hwpx_text_container_blocks(section, element: etree._Element) -> list[HancomBlock]:
    local_name = etree.QName(element).localname
    if local_name in {"rect", "line", "ellipse", "arc", "polygon", "curve", "connectLine", "container"}:
        paragraphs = element.xpath("./hp:drawText/hp:subList/hp:p", namespaces=NS)
    elif local_name in {"header", "footer", "footNote", "endNote", "hiddenComment"}:
        paragraphs = element.xpath("./hp:subList/hp:p", namespaces=NS)
    else:
        return []

    blocks: list[HancomBlock] = []
    for paragraph_index, paragraph in enumerate(paragraphs):
        blocks.extend(_extract_hwpx_nested_paragraph_blocks(section, paragraph, paragraph_index))
    return _sort_hwpx_nested_blocks_by_source_order(blocks)


def _extract_hwpx_nested_paragraph_blocks(section, paragraph: etree._Element, paragraph_index: int) -> list[HancomBlock]:
    blocks: list[HancomBlock] = []
    skipped_field_id: str | None = None
    for run in paragraph.xpath("./hp:run", namespaces=NS):
        for child in list(run):
            local_name = etree.QName(child).localname
            if local_name == "t":
                if skipped_field_id is None:
                    text = (child.text or "").replace("\r", "")
                    if text:
                        blocks.append(
                            _mark_hwpx_nested_paragraph(
                                _build_hwpx_paragraph_block(paragraph, text, char_pr_id=run.get("charPrIDRef")),
                                paragraph_index,
                                order=_source_order_for_element(paragraph, child),
                            )
                        )
                continue

            candidates = list(child) if local_name == "ctrl" else [child]
            for node in candidates:
                node_local_name = etree.QName(node).localname
                if skipped_field_id is not None:
                    if node_local_name == "fieldEnd" and node.get("beginIDRef") == skipped_field_id:
                        skipped_field_id = None
                    continue
                if node_local_name not in _HWPX_NESTED_INLINE_BLOCK_NAMES:
                    continue
                extracted = _extract_hwpx_run_node(section, node)
                if extracted is None:
                    continue
                extracted_blocks = extracted if isinstance(extracted, list) else [extracted]
                for block in extracted_blocks:
                    blocks.append(
                        _mark_hwpx_nested_paragraph(
                            block,
                            paragraph_index,
                            order=_source_order_for_element(paragraph, node),
                        )
                    )
                if node_local_name == "fieldBegin":
                    skipped_field_id = node.get("id") or node.get("fieldid")
    return _sort_hwpx_nested_blocks_by_source_order(blocks)


def _extract_missing_hwpx_descendant_blocks(
    section,
    paragraph: etree._Element,
    paragraph_index: int,
    seen_nodes: set[str],
) -> list[HancomBlock]:
    expression = (
        ".//hp:bookmark | .//hp:fieldBegin | .//hp:checkBtn | .//hp:radioBtn | .//hp:btn | .//hp:edit | "
        ".//hp:comboBox | .//hp:listBox | .//hp:scrollBar | .//hp:autoNum | .//hp:newNum | .//hp:tbl | "
        ".//hp:pic | .//hp:footNote | .//hp:endNote | .//hp:hiddenComment | .//hp:equation | .//hp:chart | "
        ".//hp:rect | .//hp:line | .//hp:ellipse | .//hp:arc | .//hp:polygon | .//hp:curve | "
        ".//hp:connectLine | .//hp:textart | .//hp:container | .//hp:ole"
    )
    blocks: list[HancomBlock] = []
    memo_blocks: list[Memo] = []
    memo_carriers: list[Memo] = []
    for node in paragraph.xpath(expression, namespaces=NS):
        if _hwpx_node_is_inside_text_container(node) and etree.QName(node).localname in _HWPX_NESTED_INLINE_BLOCK_NAMES:
            continue
        node_key = _hwpx_node_key(node)
        if node_key in seen_nodes:
            continue
        extracted = _extract_hwpx_run_node(section, node)
        if extracted is None:
            continue
        seen_nodes.add(node_key)
        if not _hwpx_node_is_inside_table_cell(node):
            caption_text = _extract_hwpx_caption_text(node)
            if caption_text:
                blocks.append(
                    _mark_hwpx_source_paragraph(
                        _build_hwpx_paragraph_block(paragraph, caption_text),
                        paragraph_index,
                        order=_source_order_for_element(paragraph, node),
                    )
                )
        extracted_blocks = extracted if isinstance(extracted, list) else [extracted]
        for block in extracted_blocks:
            if _hwpx_node_is_inside_table_cell(node):
                _clear_nested_owned_text(block)
            elif _hwpx_node_is_inside_text_container(node):
                _mark_hwpx_text_owned_by_container(block)
            owned_text = _hwpx_owned_inline_text(block)
            if owned_text:
                blocks.append(
                    _mark_hwpx_source_paragraph(
                        _build_hwpx_paragraph_block(paragraph, owned_text),
                        paragraph_index,
                        order=_source_order_for_element(paragraph, node),
                    )
                )
            _mark_hwpx_source_paragraph(block, paragraph_index, order=_source_order_for_element(paragraph, node))
            if isinstance(block, Memo) and _memo_requires_hwpx_carrier(block):
                memo_carriers.append(block)
            elif isinstance(block, Memo):
                memo_blocks.append(block)
            else:
                blocks.append(block)
    if memo_blocks or memo_carriers:
        for memo in _merge_hwpx_memos_with_carriers(memo_blocks, memo_carriers):
            blocks.append(_mark_hwpx_source_paragraph(memo, paragraph_index))
    return blocks


def _extract_hwpx_caption_text(node: etree._Element) -> str:
    if etree.QName(node).localname not in _CAPTION_TEXT_BLOCK_NAMES:
        return ""
    caption_nodes = node.xpath("./hp:caption", namespaces=NS)
    if not caption_nodes:
        return ""
    return "".join(text_node.text or "" for text_node in caption_nodes[0].xpath(".//hp:t", namespaces=NS))


_CAPTION_TEXT_BLOCK_NAMES = {
    "tbl",
    "pic",
    "ole",
    "rect",
    "line",
    "ellipse",
    "arc",
    "polygon",
    "curve",
    "connectLine",
    "textart",
    "container",
}


def _hwpx_node_is_inside_table_cell(node: etree._Element) -> bool:
    return any(etree.QName(ancestor).localname == "tc" for ancestor in node.iterancestors())


_HWPX_TEXT_CONTAINER_NAMES = {
    "header",
    "footer",
    "footNote",
    "endNote",
    "hiddenComment",
    "rect",
    "line",
    "ellipse",
    "arc",
    "polygon",
    "curve",
    "connectLine",
    "textart",
    "container",
}


_HWPX_NESTED_INLINE_BLOCK_NAMES = {
    "bookmark",
    "fieldBegin",
    "autoNum",
    "newNum",
}


def _hwpx_node_is_inside_text_container(node: etree._Element) -> bool:
    return any(etree.QName(ancestor).localname in _HWPX_TEXT_CONTAINER_NAMES for ancestor in node.iterancestors())


def _clear_nested_owned_text(block: HancomBlock) -> None:
    if isinstance(block, Shape):
        block.text = ""
    elif isinstance(block, Chart):
        block.title = ""
    elif isinstance(block, Hyperlink):
        block.display_text = None
    elif isinstance(block, Field):
        block.display_text = None
    elif isinstance(block, Form):
        block.label = ""
    elif isinstance(block, Memo):
        block.text = ""


def _mark_hwpx_text_owned_by_container(block: HancomBlock) -> None:
    if isinstance(block, (Hyperlink, Field)):
        setattr(block, "source_text_owned_by_container", True)


def _hwpx_owned_inline_text(block: HancomBlock) -> str | None:
    if getattr(block, "source_text_owned_by_container", False):
        return None
    if isinstance(block, (Hyperlink, Field)):
        return block.display_text
    return None


def _hwpx_node_key(node: etree._Element) -> str:
    return node.getroottree().getpath(node)


def _extract_hwpx_run_node(section, node: etree._Element) -> HancomBlock | list[HancomBlock] | None:
    local_name = etree.QName(node).localname
    if local_name == "bookmark":
        return _extract_bookmark(_build_bookmark_wrapper(section, node))
    if local_name == "fieldBegin":
        field = _build_field_wrapper(section, node)
        if (field.field_type or "").upper() == "HYPERLINK":
            return Hyperlink(target=field.hyperlink_target or "", display_text=field.display_text or None)
        return _extract_field(field)
    if local_name in {"checkBtn", "radioBtn", "btn", "edit", "comboBox", "listBox", "scrollBar"}:
        return _extract_hwpx_form(_build_form_wrapper(section, node))
    if local_name in {"autoNum", "newNum"}:
        return _extract_auto_number(_build_auto_number_wrapper(section, node))
    if local_name == "tbl":
        return _extract_hwpx_table(_build_table_wrapper(section, node))
    if local_name == "pic":
        return _extract_hwpx_picture(_build_picture_wrapper(section, node))
    if local_name in {"footNote", "endNote"}:
        return _extract_hwpx_note(_build_note_wrapper(section, node))
    if local_name == "hiddenComment":
        return _extract_hwpx_memo(_build_memo_wrapper(section, node))
    if local_name == "equation":
        return _extract_hwpx_equation(_build_equation_wrapper(section, node))
    if local_name == "chart":
        if _hwpx_chart_is_covered_by_carrier_shape(node):
            return None
        return _extract_hwpx_chart(_build_chart_wrapper(section, node))
    if local_name in {"rect", "line", "ellipse", "arc", "polygon", "curve", "connectLine", "textart", "container"}:
        return _extract_hwpx_shape(_build_shape_wrapper(section, node))
    if local_name == "ole":
        if node.xpath("ancestor::hp:default[parent::hp:switch[hp:case/hp:chart]]", namespaces=NS):
            return None
        return _extract_hwpx_ole(_build_ole_wrapper(section, node))
    return None


def _hwpx_chart_is_covered_by_carrier_shape(node: etree._Element) -> bool:
    for ancestor in node.iterancestors():
        if etree.QName(ancestor).localname not in {
            "rect",
            "line",
            "ellipse",
            "arc",
            "polygon",
            "curve",
            "connectLine",
            "textart",
            "container",
        }:
            continue
        comments = ancestor.xpath("./hp:shapeComment/text()", namespaces=NS)
        if comments and str(comments[0]).startswith(_HWPX_CHART_CARRIER_PREFIX):
            return True
    return False


def _source_order_for_element(paragraph: etree._Element, element: etree._Element) -> int:
    for index, candidate in enumerate(paragraph.iter()):
        if candidate is element:
            return index
    return 10**9


def _mark_hwpx_source_paragraph(block, paragraph_index: int, *, order: int | None = None):
    setattr(block, "source_paragraph_index", paragraph_index)
    if order is not None:
        setattr(block, "source_order", order)
    return block


def _build_hwpx_paragraph_block(paragraph, text: str, *, char_pr_id: str | None = None) -> Paragraph:
    return Paragraph(
        text=text,
        style_id=paragraph.get("styleIDRef"),
        para_pr_id=paragraph.get("paraPrIDRef"),
        char_pr_id=char_pr_id or _extract_hwpx_paragraph_char_pr_id(paragraph),
    )


def _extract_hwpx_paragraph_char_pr_id(paragraph) -> str | None:
    run_ids = [
        value
        for value in paragraph.xpath("./hp:run/@charPrIDRef", namespaces=NS)
        if isinstance(value, str) and value
    ]
    if not run_ids:
        return None
    first = run_ids[0]
    return first if all(value == first for value in run_ids) else None


def _extract_hwpx_table(table) -> Table:
    cell_texts = _blank_cell_texts(table.row_count, table.column_count)
    cell_spans: dict[tuple[int, int], tuple[int, int]] = {}
    cells = table.cells()
    for cell in cells:
        if cell.row < table.row_count and cell.column < table.column_count:
            cell_texts[cell.row][cell.column] = cell.text
            if cell.row_span != 1 or cell.col_span != 1:
                cell_spans[(cell.row, cell.column)] = (cell.row_span, cell.col_span)
    return Table(
        rows=table.row_count,
        cols=table.column_count,
        cell_texts=cell_texts,
        cell_border_fill_ids={(cell.row, cell.column): cell.border_fill_id for cell in cells} or None,
        cell_margins={(cell.row, cell.column): cell.margins for cell in cells} or None,
        cell_vertical_aligns={
            (cell.row, cell.column): cell.vertical_align for cell in cells if cell.vertical_align != "CENTER"
        }
        or None,
        cell_spans=cell_spans or None,
        table_border_fill_id=table.table_border_fill_id,
        cell_spacing=table.cell_spacing,
        table_margins=table.in_margins() or None,
        layout=_normalize_str_map(table.layout()),
        out_margins=_normalize_int_map(table.out_margins()),
        page_break=table.page_break,
        repeat_header=table.repeat_header,
    )


def _extract_hwpx_picture(picture) -> Picture:
    path = picture.binary_part_path()
    extension = Path(path).suffix.lstrip(".") or None
    size = picture.size()
    line_style = picture.line_style()
    size_nodes = picture.element.xpath("./hp:sz", namespaces=NS)
    width = size.get("width", 7200)
    height = size.get("height", 7200)
    return Picture(
        name=Path(path).name,
        data=picture.binary_data(),
        extension=extension,
        width=width,
        height=height,
        shape_comment=picture.shape_comment or None,
        layout=_normalize_str_map(picture.layout()),
        out_margins=_normalize_int_map(picture.out_margins()),
        rotation=_normalize_str_map(picture.rotation()),
        image_adjustment=_normalize_str_map(picture.image_adjustment()),
        crop=_normalize_int_map(picture.crop()) if hasattr(picture, "crop") else {},
        line_color=line_style.get("color", "#000000"),
        line_width=int(line_style.get("width") or "0"),
        line_style=_normalize_str_map(line_style),
        has_size_node=bool(size_nodes),
        size_attributes=_extract_hwpx_graphic_size_attributes(picture.element),
        has_line_node=bool(picture.element.xpath("./hp:lineShape", namespaces=NS)),
        has_position_node=bool(picture.element.xpath("./hp:pos", namespaces=NS)),
        has_out_margin_node=bool(picture.element.xpath("./hp:outMargin", namespaces=NS)),
        **_extract_hwpx_graphic_size_fields(picture.element, width=width, height=height),
    )


def _extract_hwpx_note(note) -> Note:
    number = int(note.number) if note.number and str(note.number).isdigit() else None
    return Note(
        kind=note.kind,
        text=note.text,
        number=number,
        nested_blocks=_extract_hwpx_text_container_blocks(note.section, note.element),
    )


def _extract_hwpx_memo(memo) -> Memo:
    return Memo(text=memo.text)


def _extract_hwpx_memo_shape(memo_shape) -> MemoShapeDefinition:
    return MemoShapeDefinition(
        memo_shape_id=memo_shape.memo_shape_id,
        width=memo_shape.width,
        line_width=memo_shape.line_width,
        line_type=memo_shape.line_type,
        line_color=memo_shape.line_color,
        fill_color=memo_shape.fill_color,
        active_color=memo_shape.active_color,
        memo_type=memo_shape.memo_type,
    )


def _extract_hwpx_equation(equation) -> Equation:
    size = equation.size()
    return Equation(
        script=equation.script,
        width=size.get("width", 4800),
        height=size.get("height", 2300),
        shape_comment=equation.shape_comment or None,
        text_color=equation.element.get("textColor", "#000000"),
        base_unit=int(equation.element.get("baseUnit", "1100")),
        font=equation.element.get("font", "HYhwpEQ"),
        layout=_normalize_str_map(equation.layout()),
        out_margins=_normalize_int_map(equation.out_margins()),
        rotation=_normalize_str_map(equation.rotation()),
    )


def _extract_hwpx_chart(chart) -> Chart:
    size = chart.size()
    return Chart(
        title=chart.title,
        chart_type=chart.chart_type,
        categories=chart.categories,
        series=chart.series,
        data_ref=chart.data_ref,
        legend_visible=chart.legend_visible,
        width=size.get("width", 12000),
        height=size.get("height", 3200),
        shape_comment=chart.shape_comment or None,
        layout=_normalize_str_map(chart.layout()),
        out_margins=_normalize_int_map(chart.out_margins()),
        rotation=_normalize_str_map(chart.rotation()),
    )


def _extract_hwpx_shape(shape) -> Shape | Chart:
    carrier = _extract_hwpx_chart_carrier(shape)
    if carrier is not None:
        return carrier
    size = shape.size()
    fill_style = shape.fill_style()
    line_style = shape.line_style()
    size_nodes = shape.element.xpath("./hp:sz", namespaces=NS)
    width = size.get("width", 12000)
    height = size.get("height", 3200)
    return Shape(
        kind=shape.kind,
        text=_extract_hwpx_shape_direct_text(shape),
        width=width,
        height=height,
        fill_color=fill_style.get("faceColor", "#FFFFFF"),
        line_color=line_style.get("color", "#000000"),
        shape_comment=shape.shape_comment or None,
        layout=_normalize_str_map(shape.layout()),
        out_margins=_normalize_int_map(shape.out_margins()),
        rotation=_normalize_str_map(shape.rotation()),
        text_margins=_normalize_int_map(shape.text_margins()),
        specific_fields=_extract_hwpx_shape_specific_fields(shape),
        line_width=int(line_style.get("width", "33") or "33"),
        line_style=_normalize_str_map(line_style),
        has_line_node=bool(shape.element.xpath("./hp:lineShape", namespaces=NS)),
        has_size_node=bool(size_nodes),
        size_attributes=_extract_hwpx_graphic_size_attributes(shape.element),
        has_position_node=bool(shape.element.xpath("./hp:pos", namespaces=NS)),
        has_out_margin_node=bool(shape.element.xpath("./hp:outMargin", namespaces=NS)),
        has_fill_node=bool(shape.element.xpath("./hc:fillBrush", namespaces=NS)),
        has_text_margin_node=bool(shape.element.xpath("./hp:drawText/hp:textMargin", namespaces=NS)),
        nested_blocks=_extract_hwpx_text_container_blocks(shape.section, shape.element),
        **_extract_hwpx_graphic_size_fields(shape.element, width=width, height=height),
    )


def _extract_hwpx_ole(ole) -> Ole:
    size = ole.size()
    line_style = ole.line_style()
    size_nodes = ole.element.xpath("./hp:sz", namespaces=NS)
    width = size.get("width", 42001)
    height = size.get("height", 13501)
    return Ole(
        name=Path(ole.binary_part_path()).name,
        data=ole.binary_data(),
        width=width,
        height=height,
        shape_comment=ole.shape_comment or None,
        object_type=ole.object_type or "EMBEDDED",
        draw_aspect=ole.draw_aspect or "CONTENT",
        has_moniker=ole.has_moniker,
        eq_baseline=int(ole.element.get("eqBaseLine", "0")),
        line_color=line_style.get("color", "#000000"),
        line_width=int(line_style.get("width") or "0"),
        line_style=_normalize_str_map(line_style),
        layout=_normalize_str_map(ole.layout()),
        out_margins=_normalize_int_map(ole.out_margins()),
        rotation=_normalize_str_map(ole.rotation()),
        extent=_normalize_int_map(ole.extent()),
        has_size_node=bool(size_nodes),
        size_attributes=_extract_hwpx_graphic_size_attributes(ole.element),
        has_position_node=bool(ole.element.xpath("./hp:pos", namespaces=NS)),
        has_out_margin_node=bool(ole.element.xpath("./hp:outMargin", namespaces=NS)),
        **_extract_hwpx_graphic_size_fields(ole.element, width=width, height=height),
    )


def _extract_bookmark(bookmark) -> Bookmark:
    return Bookmark(name=bookmark.name or "")


def _extract_field(field) -> Field | Form | Memo:
    if (field.field_type or "").upper() == _HWPX_FORM_FIELD_TYPE:
        return _extract_hwpx_form_field(field)
    if (field.field_type or "").upper() == _HWPX_MEMO_FIELD_TYPE:
        return _extract_hwpx_memo_field(field)
    return Field(
        field_type=field.field_type or "",
        display_text=field.display_text or None,
        name=field.name,
        parameters=field.parameter_map(),
        editable=field.element.get("editable") == "1",
        dirty=field.element.get("dirty") == "1",
    )


def _extract_hwpx_form_field(field) -> Form:
    parameters = field.parameter_map()
    items = _json_list_parameter(parameters.get("Items"))
    return Form(
        label=field.display_text or "",
        form_type=parameters.get("FormType", "INPUT") or "INPUT",
        name=field.name or None,
        value=parameters.get("Value"),
        checked=_coerce_bool(parameters.get("Checked")) or False,
        items=items,
        editable=_coerce_bool(parameters.get("Editable")) is not False,
        locked=_coerce_bool(parameters.get("Locked")) or False,
        placeholder=parameters.get("Placeholder") or None,
    )


def _extract_hwpx_form(form) -> Form:
    return Form(
        label=form.label,
        form_type=form.form_type,
        name=form.name,
        value=form.value,
        checked=form.checked,
        items=form.items,
        editable=form.editable,
        locked=form.locked,
        placeholder=form.placeholder,
    )


def _extract_hwpx_memo_field(field) -> Memo:
    parameters = field.parameter_map()
    order_value = parameters.get("Order")
    try:
        order = int(order_value) if order_value not in (None, "") else None
    except (TypeError, ValueError):
        order = None
    return Memo(
        text=field.display_text or parameters.get("Text") or "",
        author=parameters.get("Author") or None,
        memo_id=parameters.get("MemoId") or field.name or None,
        anchor_id=parameters.get("AnchorId") or None,
        order=order,
        visible=_coerce_bool(parameters.get("Visible")) is not False,
    )


def _memo_requires_hwpx_carrier(block: Memo) -> bool:
    return any(
        value is not None
        for value in (block.author, block.memo_id, block.anchor_id, block.order)
    ) or block.visible is False


def _merge_hwpx_memo_with_carrier(memo: Memo, carrier: Memo) -> Memo:
    return Memo(
        text=memo.text or carrier.text,
        author=carrier.author,
        memo_id=carrier.memo_id,
        anchor_id=carrier.anchor_id,
        order=carrier.order,
        visible=carrier.visible,
    )


def _merge_hwpx_memos_with_carriers(memos: list[Memo], carriers: list[Memo]) -> list[Memo]:
    if not carriers:
        return list(memos)
    unused_carriers = list(range(len(carriers)))
    merged: list[Memo] = []

    for memo in memos:
        carrier_index = _find_matching_hwpx_memo_carrier(memo, carriers, unused_carriers)
        if carrier_index is None:
            merged.append(memo)
            continue
        unused_carriers.remove(carrier_index)
        merged.append(_merge_hwpx_memo_with_carrier(memo, carriers[carrier_index]))

    for carrier_index in unused_carriers:
        merged.append(carriers[carrier_index])
    return merged


def _find_matching_hwpx_memo_carrier(memo: Memo, carriers: list[Memo], unused_carriers: list[int]) -> int | None:
    memo_text = _normalize_hwpx_memo_pairing_key(memo.text)
    if memo_text:
        for carrier_index in unused_carriers:
            if _normalize_hwpx_memo_pairing_key(carriers[carrier_index].text) == memo_text:
                return carrier_index
    return unused_carriers[0] if unused_carriers else None


def _normalize_hwpx_memo_pairing_key(value: str | None) -> str:
    return " ".join((value or "").split())


def _extract_hwpx_chart_carrier(shape) -> Chart | None:
    comment = shape.shape_comment or ""
    if not comment.startswith(_HWPX_CHART_CARRIER_PREFIX):
        return None
    try:
        payload = json.loads(comment[len(_HWPX_CHART_CARRIER_PREFIX) :])
    except json.JSONDecodeError:
        return None
    if not isinstance(payload, dict):
        return None
    size = shape.size()
    return Chart(
        title=shape.text or str(payload.get("title", "")),
        chart_type=str(payload.get("chartType", "BAR") or "BAR"),
        categories=[str(value) for value in payload.get("categories", []) if value is not None],
        series=[dict(value) for value in payload.get("series", []) if isinstance(value, dict)],
        data_ref=str(payload["dataRef"]) if payload.get("dataRef") not in (None, "") else None,
        legend_visible=_coerce_bool(payload.get("legendVisible")) is not False,
        width=size.get("width", 12000),
        height=size.get("height", 3200),
        shape_comment=str(payload["shapeComment"]) if payload.get("shapeComment") not in (None, "") else None,
        layout=_normalize_str_map(shape.layout()),
        out_margins=_normalize_int_map(shape.out_margins()),
        rotation=_normalize_str_map(shape.rotation()),
    )


def _build_hwpx_form_parameters(block: Form) -> dict[str, str]:
    parameters: dict[str, str] = {
        "FormType": block.form_type,
        "Checked": "1" if block.checked else "0",
        "Editable": "1" if block.editable else "0",
        "Locked": "1" if block.locked else "0",
    }
    if block.value is not None:
        parameters["Value"] = block.value
    if block.items:
        parameters["Items"] = json.dumps([str(value) for value in block.items], ensure_ascii=True, separators=(",", ":"))
    if block.placeholder:
        parameters["Placeholder"] = block.placeholder
    return parameters


def _build_hwpx_memo_parameters(block: Memo) -> dict[str, str]:
    parameters: dict[str, str] = {
        "Visible": "1" if block.visible else "0",
    }
    if block.text:
        parameters["Text"] = block.text
    if block.author:
        parameters["Author"] = block.author
    if block.memo_id:
        parameters["MemoId"] = block.memo_id
    if block.anchor_id:
        parameters["AnchorId"] = block.anchor_id
    if block.order is not None:
        parameters["Order"] = str(block.order)
    return parameters


def _build_hwpx_chart_carrier_comment(block: Chart) -> str:
    payload = {
        "title": block.title,
        "chartType": block.chart_type,
        "categories": [str(value) for value in block.categories],
        "series": [dict(value) for value in block.series],
        "dataRef": block.data_ref,
        "legendVisible": bool(block.legend_visible),
        "shapeComment": block.shape_comment,
    }
    return _HWPX_CHART_CARRIER_PREFIX + json.dumps(payload, ensure_ascii=True, separators=(",", ":"))


def _json_list_parameter(value: str | None) -> list[str]:
    if not value:
        return []
    try:
        decoded = json.loads(value)
    except json.JSONDecodeError:
        return []
    if not isinstance(decoded, list):
        return []
    return [str(item) for item in decoded if item is not None]


def _extract_auto_number(auto_number) -> AutoNumber:
    return AutoNumber(
        kind=auto_number.kind,
        number=auto_number.number or 1,
        number_type=auto_number.number_type or "PAGE",
    )


def _extract_page_border_fills(section) -> list[dict[str, str | int]]:
    fills: list[dict[str, str | int]] = []
    for node in section.findall(".//hp:pageBorderFill"):
        offset = node.find("./hp:offset")
        fills.append(
            {
                "type": node.get_attr("type") or "BOTH",
                "borderFillIDRef": node.get_attr("borderFillIDRef") or "0",
                "textBorder": node.get_attr("textBorder") or "PAPER",
                "headerInside": node.get_attr("headerInside") or "0",
                "footerInside": node.get_attr("footerInside") or "0",
                "fillArea": node.get_attr("fillArea") or "PAPER",
                "left": int((offset.get_attr("left") if offset is not None else "0") or 0),
                "right": int((offset.get_attr("right") if offset is not None else "0") or 0),
                "top": int((offset.get_attr("top") if offset is not None else "0") or 0),
                "bottom": int((offset.get_attr("bottom") if offset is not None else "0") or 0),
            }
        )
    return fills


def _extract_page_numbers_with_positions(section) -> tuple[list[dict[str, str]], list[tuple[int | None, int | None]]]:
    page_numbers: list[dict[str, str]] = []
    positions: list[tuple[int | None, int | None]] = []
    for node in section.findall(".//hp:pageNum"):
        if not _is_direct_section_paragraph_control(node.element):
            continue
        page_numbers.append(
            {
                "pos": node.get_attr("pos") or "BOTTOM_CENTER",
                "formatType": node.get_attr("formatType") or "DIGIT",
                "sideChar": node.get_attr("sideChar") or "-",
            }
        )
        positions.append(_source_paragraph_position_for_element(section, node.element))
    return page_numbers, positions


def _extract_page_numbers(section) -> list[dict[str, str]]:
    page_numbers, _positions = _extract_page_numbers_with_positions(section)
    return page_numbers


def _is_direct_section_paragraph_control(element: etree._Element) -> bool:
    parent = element.getparent()
    if parent is None or etree.QName(parent).localname != "ctrl":
        return False
    run = parent.getparent()
    if run is None or etree.QName(run).localname != "run":
        return False
    paragraph = run.getparent()
    if paragraph is None or etree.QName(paragraph).localname != "p":
        return False
    section_root = paragraph.getparent()
    return section_root is not None and etree.QName(section_root).localname == "sec"


def _extract_note_pr(section, kind: str) -> dict[str, object]:
    node = section.find(f".//hp:{kind}")
    if node is None:
        return {}
    auto_num = node.find("./hp:autoNumFormat")
    note_line = node.find("./hp:noteLine")
    note_spacing = node.find("./hp:noteSpacing")
    numbering = node.find("./hp:numbering")
    placement = node.find("./hp:placement")
    return {
        "autoNumFormat": {
            "type": auto_num.get_attr("type") if auto_num is not None else "DIGIT",
            "userChar": auto_num.get_attr("userChar") if auto_num is not None else "",
            "prefixChar": auto_num.get_attr("prefixChar") if auto_num is not None else "",
            "suffixChar": auto_num.get_attr("suffixChar") if auto_num is not None else "",
            "supscript": auto_num.get_attr("supscript") if auto_num is not None else "0",
        },
        "noteLine": {
            "length": note_line.get_attr("length") if note_line is not None else "-1",
            "type": note_line.get_attr("type") if note_line is not None else "SOLID",
            "width": note_line.get_attr("width") if note_line is not None else "0.12 mm",
            "color": note_line.get_attr("color") if note_line is not None else "#000000",
        },
        "noteSpacing": {
            "betweenNotes": note_spacing.get_attr("betweenNotes") if note_spacing is not None else "283",
            "belowLine": note_spacing.get_attr("belowLine") if note_spacing is not None else "567",
            "aboveLine": note_spacing.get_attr("aboveLine") if note_spacing is not None else "850",
        },
        "numbering": {
            "type": numbering.get_attr("type") if numbering is not None else "CONTINUOUS",
            "newNum": numbering.get_attr("newNum") if numbering is not None else "1",
        },
        "placement": {
            "place": placement.get_attr("place") if placement is not None else ("END_OF_DOCUMENT" if kind == "endNotePr" else "EACH_COLUMN"),
            "beneathText": placement.get_attr("beneathText") if placement is not None else "0",
        },
    }


def _extract_line_number_shape(section) -> dict[str, str]:
    node = section.find(".//hp:lineNumberShape")
    if node is None:
        return {}
    return {
        key: node.get_attr(key) or "0"
        for key in ("restartType", "countBy", "distance", "startNumber")
    }


def _extract_section_settings(section) -> SectionSettings:
    settings = section.document.section_settings(section.section_index())
    memo_shape_ids = {
        memo_shape.memo_shape_id
        for memo_shape in section.document.memo_shapes()
        if memo_shape.memo_shape_id is not None
    }
    memo_shape_id = settings.memo_shape_id if settings.memo_shape_id in memo_shape_ids else None
    page_numbers, page_number_positions = _extract_page_numbers_with_positions(section)
    return SectionSettings(
        page_width=settings.page_width,
        page_height=settings.page_height,
        landscape=settings.landscape,
        margins=settings.margins(),
        page_border_fills=_extract_page_border_fills(section),
        visibility=settings.visibility(),
        grid=settings.grid(),
        start_numbers=settings.start_numbers(),
        page_numbers=page_numbers,
        page_number_positions=page_number_positions,
        footnote_pr=_extract_note_pr(section, "footNotePr"),
        endnote_pr=_extract_note_pr(section, "endNotePr"),
        line_number_shape=_extract_line_number_shape(section),
        memo_shape_id=memo_shape_id,
    )


def _extract_header_footer_blocks(section) -> list[HeaderFooter]:
    section_index = section.section_index()
    blocks: list[HeaderFooter] = []
    for header in section.document.headers(section_index=section_index):
        paragraph_index, source_order = _source_paragraph_position_for_element(section, header.element)
        blocks.append(
            _mark_hwpx_source_paragraph(
                HeaderFooter(
                    kind=header.kind,
                    text=_extract_hwpx_direct_text(header.element),
                    apply_page_type=header.apply_page_type or "BOTH",
                    nested_blocks=_extract_hwpx_text_container_blocks(header.section, header.element),
                ),
                paragraph_index or 0,
                order=source_order,
            )
        )
    for footer in section.document.footers(section_index=section_index):
        paragraph_index, source_order = _source_paragraph_position_for_element(section, footer.element)
        blocks.append(
            _mark_hwpx_source_paragraph(
                HeaderFooter(
                    kind=footer.kind,
                    text=_extract_hwpx_direct_text(footer.element),
                    apply_page_type=footer.apply_page_type or "BOTH",
                    nested_blocks=_extract_hwpx_text_container_blocks(footer.section, footer.element),
                ),
                paragraph_index or 0,
                order=source_order,
            )
        )
    return blocks


def _source_paragraph_position_for_element(section, element: etree._Element) -> tuple[int | None, int | None]:
    paragraphs = section.root_element.xpath("./hp:p", namespaces=NS)
    for index, paragraph in enumerate(paragraphs):
        if paragraph is element or element in paragraph.iterdescendants():
            return index, _source_order_for_element(paragraph, element)
    return None, None


def _source_paragraph_index_for_element(section, element: etree._Element) -> int | None:
    paragraphs = section.root_element.xpath("./hp:p", namespaces=NS)
    for index, paragraph in enumerate(paragraphs):
        if paragraph is element or element in paragraph.iterdescendants():
            return index
    return None


def _extract_style_definition(style) -> StyleDefinition:
    return StyleDefinition(
        name=style.name or "",
        style_id=style.style_id,
        english_name=style.english_name,
        para_pr_id=style.para_pr_id,
        char_pr_id=style.char_pr_id,
        style_type=style.element.get("type", "PARA"),
        next_style_id=style.element.get("nextStyleIDRef"),
        lang_id=style.element.get("langID", "1042"),
        lock_form=style.element.get("lockForm") == "1",
    )


def _extract_paragraph_style(style) -> ParagraphStyle:
    line_spacing_types = style.element.xpath(".//hh:lineSpacing/@type", namespaces=NS)
    return ParagraphStyle(
        style_id=style.style_id,
        alignment_horizontal=style.alignment_horizontal,
        alignment_vertical=style.element.xpath("./hh:align/@vertical", namespaces=NS)[0]
        if style.element.xpath("./hh:align/@vertical", namespaces=NS)
        else None,
        line_spacing=style.line_spacing,
        line_spacing_type=line_spacing_types[0] if line_spacing_types else None,
        margins=_extract_hwpx_margin_map(style.element),
        tab_pr_id=style.element.get("tabPrIDRef"),
        snap_to_grid=style.element.get("snapToGrid"),
        condense=style.element.get("condense"),
        font_line_height=style.element.get("fontLineHeight"),
        suppress_line_numbers=style.element.get("suppressLineNumbers"),
        checked=style.element.get("checked"),
        heading=_extract_hwpx_attribute_map(
            style.element,
            "./hh:heading",
            "type",
            "idRef",
            "level",
        ),
        break_setting=_extract_hwpx_attribute_map(
            style.element,
            "./hh:breakSetting",
            "breakLatinWord",
            "breakNonLatinWord",
            "widowOrphan",
            "keepWithNext",
            "keepLines",
            "pageBreakBefore",
            "lineWrap",
        ),
        auto_spacing=_extract_hwpx_attribute_map(
            style.element,
            "./hh:autoSpacing",
            "eAsianEng",
            "eAsianNum",
        ),
    )


def _extract_character_style(style) -> CharacterStyle:
    return CharacterStyle(
        style_id=style.style_id,
        text_color=style.text_color,
        shade_color=style.element.get("shadeColor"),
        height=style.height,
        use_font_space=style.element.get("useFontSpace"),
        use_kerning=style.element.get("useKerning"),
        sym_mark=style.element.get("symMark"),
        font_refs=_extract_hwpx_attribute_map(
            style.element,
            "./hh:fontRef",
            "hangul",
            "latin",
            "hanja",
            "japanese",
            "other",
            "symbol",
            "user",
        ),
    )


def _extract_hwpx_margin_map(element: etree._Element) -> dict[str, str]:
    values: dict[str, str] = {}
    for key in ("intent", "left", "right", "prev", "next"):
        nodes = element.xpath(f"./hh:margin/hc:{key}", namespaces=NS)
        if not nodes:
            continue
        value = nodes[0].get("value")
        if value is not None:
            values[key] = value
        unit = nodes[0].get("unit")
        if unit is not None and "unit" not in values:
            values["unit"] = unit
    return values


def _extract_hwpx_attribute_map(element: etree._Element, xpath: str, *attributes: str) -> dict[str, str]:
    nodes = element.xpath(xpath, namespaces=NS)
    if not nodes:
        return {}
    return {
        attribute: value
        for attribute in attributes
        if (value := nodes[0].get(attribute)) is not None
    }


def _extract_hwp_style_definitions(docinfo_model) -> list[StyleDefinition]:
    definitions: list[StyleDefinition] = []
    for style_id, record in enumerate(docinfo_model.style_records()):
        definition = _parse_hwp_style_definition_payload(record.payload, style_id=style_id)
        if definition is not None:
            definitions.append(definition)
    return definitions


def _parse_hwp_style_definition_payload(payload: bytes, *, style_id: int | None = None) -> StyleDefinition | None:
    if len(payload) < 4:
        return None
    cursor = 0
    name_length = int.from_bytes(payload[cursor : cursor + 2], "little", signed=False)
    cursor += 2
    name_end = cursor + (name_length * 2)
    if name_end > len(payload):
        return None
    name = payload[cursor:name_end].decode("utf-16-le", errors="ignore")
    cursor = name_end
    if cursor + 2 > len(payload):
        return None
    english_name_length = int.from_bytes(payload[cursor : cursor + 2], "little", signed=False)
    cursor += 2
    english_name_end = cursor + (english_name_length * 2)
    if english_name_end > len(payload):
        return None
    english_name = payload[cursor:english_name_end].decode("utf-16-le", errors="ignore")
    cursor = english_name_end
    if cursor + 10 > len(payload):
        return None
    style_word = int.from_bytes(payload[cursor : cursor + 2], "little", signed=False)
    cursor += 2
    lang_id = int.from_bytes(payload[cursor : cursor + 2], "little", signed=False)
    cursor += 2
    next_style_id = int.from_bytes(payload[cursor : cursor + 2], "little", signed=False)
    cursor += 2
    para_pr_id = int.from_bytes(payload[cursor : cursor + 2], "little", signed=False)
    cursor += 2
    char_pr_id = int.from_bytes(payload[cursor : cursor + 2], "little", signed=False)
    style_type_code = style_word & 0x00FF
    resolved_style_id = style_word >> 8 if style_id is None else style_id
    return StyleDefinition(
        name=name or f"Style {resolved_style_id}",
        style_id=str(resolved_style_id),
        english_name=english_name or None,
        style_type=_HWP_STYLE_TYPE_NAMES.get(style_type_code, "PARA"),
        para_pr_id=str(para_pr_id),
        char_pr_id=str(char_pr_id),
        next_style_id=str(next_style_id),
        lang_id=str(lang_id),
        lock_form=False,
        native_hwp_payload=bytes(payload),
    )


def _build_hwp_placeholder_paragraph_styles(
    docinfo_model,
    sections: list[HancomSection],
    style_definitions: list[StyleDefinition],
) -> list[ParagraphStyle]:
    style_ids = {
        value
        for value in (_parse_optional_int(style.para_pr_id) for style in style_definitions)
        if value is not None
    }
    for section in sections:
        for block in section.blocks:
            if not isinstance(block, Paragraph):
                continue
            candidate = block.hwp_para_shape_id
            if candidate is None:
                candidate = _parse_optional_int(block.para_pr_id)
            if candidate is not None:
                style_ids.add(candidate)
    payloads = docinfo_model.para_shape_records()
    styles: list[ParagraphStyle] = []
    for style_id in sorted(style_ids):
        payload = bytes(payloads[style_id].payload) if 0 <= style_id < len(payloads) else None
        style = ParagraphStyle(style_id=str(style_id), native_hwp_payload=payload)
        if payload is not None:
            _apply_hwp_paragraph_style_payload(style, payload)
        styles.append(style)
    return styles


def _build_hwp_placeholder_character_styles(docinfo_model, style_definitions: list[StyleDefinition]) -> list[CharacterStyle]:
    style_ids = {
        value
        for value in (_parse_optional_int(style.char_pr_id) for style in style_definitions)
        if value is not None
    }
    payloads = docinfo_model.char_shape_records()
    styles: list[CharacterStyle] = []
    for style_id in sorted(style_ids):
        payload = bytes(payloads[style_id].payload) if 0 <= style_id < len(payloads) else None
        style = CharacterStyle(style_id=str(style_id), native_hwp_payload=payload)
        if payload is not None:
            _apply_hwp_character_style_payload(style, payload)
        styles.append(style)
    return styles


def _extract_hwp_numbering_definitions(docinfo_model) -> list[NumberingDefinition]:
    definitions: list[NumberingDefinition] = []
    for index, record in enumerate(docinfo_model.numbering_records()):
        definition = NumberingDefinition(style_id=str(index), native_hwp_payload=bytes(record.payload))
        _apply_hwp_numbering_payload(definition, record.payload)
        fields = record.fields()
        definition.para_head_flags = [dict(level.get("para_head_flags", {})) for level in fields.get("levels", [])]
        definition.unknown_short_bits = dict(fields.get("unknown_short_bits", {}))
        definitions.append(definition)
    return definitions


def _extract_hwp_bullet_definitions(docinfo_model) -> list[BulletDefinition]:
    definitions: list[BulletDefinition] = []
    for index, record in enumerate(docinfo_model.bullet_records()):
        definition = BulletDefinition(style_id=str(index), native_hwp_payload=bytes(record.payload))
        _apply_hwp_bullet_payload(definition, record.payload)
        fields = record.fields()
        definition.flags = int(fields.get("flags", 0))
        definition.flag_bits = dict(fields.get("flag_bits", {}))
        definitions.append(definition)
    return definitions


def _extract_hwp_memo_shape_definitions(docinfo_model) -> list[MemoShapeDefinition]:
    definitions: list[MemoShapeDefinition] = []
    for index, record in enumerate(docinfo_model.memo_shape_records()):
        fields = record.fields()
        definitions.append(
            MemoShapeDefinition(
                memo_shape_id=str(index),
                width=_coerce_int(fields.get("width")),
                line_width=_coerce_int(fields.get("line_width")),
                line_type=str(fields["line_type"]) if fields.get("line_type") is not None else None,
                line_color=str(fields["line_color"]) if fields.get("line_color") is not None else None,
                fill_color=str(fields["fill_color"]) if fields.get("fill_color") is not None else None,
                active_color=str(fields["active_color"]) if fields.get("active_color") is not None else None,
                memo_type=str(fields["memo_type"]) if fields.get("memo_type") is not None else None,
                native_hwp_payload=bytes(record.payload),
            )
        )
    return definitions


def _coerce_int(value: object | None, *, default: int | None = None) -> int | None:
    if value is None:
        return default
    if isinstance(value, int):
        return value
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return default


def _seed_hwp_payload(
    native_payload: bytes | None,
    fallback_payload: bytes | None,
    default_payload: bytes,
    *,
    min_len: int | None = None,
) -> bytearray:
    resolved = bytes(native_payload or fallback_payload or default_payload)
    target_length = max(len(default_payload), min_len or 0)
    if len(resolved) < target_length:
        resolved += b"\x00" * (target_length - len(resolved))
    return bytearray(resolved)


def _parse_hwp_colorref(payload: bytes) -> str:
    values = payload[:4].ljust(4, b"\x00")
    return f"#{values[2]:02X}{values[1]:02X}{values[0]:02X}"


def _build_hwp_colorref(value: object | None, *, default: str = "#000000") -> bytes:
    normalized = str(value or default).strip().lstrip("#")
    if len(normalized) != 6:
        normalized = default.lstrip("#")
    try:
        red = int(normalized[0:2], 16)
        green = int(normalized[2:4], 16)
        blue = int(normalized[4:6], 16)
    except ValueError:
        red = green = blue = 0
    return bytes((blue, green, red, 0))


def _apply_hwp_paragraph_style_payload(style: ParagraphStyle, payload: bytes) -> None:
    values = bytes(payload)
    style.native_hwp_payload = values
    if len(values) >= 4:
        attributes = int.from_bytes(values[0:4], "little", signed=False)
        style.alignment_horizontal = _HWP_PARA_ALIGN_NAMES.get(attributes & 0x07)
        style.snap_to_grid = "1" if attributes & (1 << 7) else "0"
        style.break_setting = {
            "breakLatinWord": "KEEP_WORD" if attributes & (1 << 8) else "BREAK_WORD",
            "breakNonLatinWord": "KEEP_WORD" if attributes & (1 << 8) else "BREAK_WORD",
            "widowOrphan": "1" if attributes & (1 << 9) else "0",
            "keepWithNext": "1" if attributes & (1 << 10) else "0",
            "keepLines": "1" if attributes & (1 << 11) else "0",
            "pageBreakBefore": "1" if attributes & (1 << 12) else "0",
            "lineWrap": "SQUEEZE" if attributes & (1 << 15) else "BREAK",
        }
        style.auto_spacing = {
            "eAsianEng": "1" if attributes & (1 << 23) else "0",
            "eAsianNum": "1" if attributes & (1 << 24) else "0",
        }
    if len(values) >= 24:
        margins = {
            "intent": str(int.from_bytes(values[4:8], "little", signed=True)),
            "left": str(int.from_bytes(values[8:12], "little", signed=True)),
            "right": str(int.from_bytes(values[12:16], "little", signed=True)),
            "prev": str(int.from_bytes(values[16:20], "little", signed=True)),
            "next": str(int.from_bytes(values[20:24], "little", signed=True)),
            "unit": "HWPUNIT",
        }
        style.margins = margins
    if len(values) >= 30:
        style.tab_pr_id = str(int.from_bytes(values[28:30], "little", signed=False))
    if len(values) >= 32:
        numbering_id = int.from_bytes(values[30:32], "little", signed=False)
        style.heading = {
            "type": "NUMBER" if numbering_id > 0 else "NONE",
            "idRef": str(numbering_id),
            "level": "0",
        }
    if len(values) >= 54:
        style.line_spacing = str(int.from_bytes(values[50:54], "little", signed=True))


def _build_hwp_paragraph_style_payload(style: ParagraphStyle, fallback_payload: bytes | None = None) -> bytes:
    payload = _seed_hwp_payload(style.native_hwp_payload, fallback_payload, _DEFAULT_HWP_PARA_SHAPE_PAYLOAD)
    attributes = int.from_bytes(payload[0:4], "little", signed=False)
    alignment_code = _HWP_PARA_ALIGN_CODES.get(str(style.alignment_horizontal or "").upper())
    if alignment_code is not None:
        attributes = (attributes & ~0x07) | alignment_code
    snap_to_grid = _coerce_bool(style.snap_to_grid)
    if snap_to_grid is not None:
        attributes = (attributes | (1 << 7)) if snap_to_grid else (attributes & ~(1 << 7))
    keep_word = str(style.break_setting.get("breakLatinWord", "")).upper() == "KEEP_WORD" or str(style.break_setting.get("breakNonLatinWord", "")).upper() == "KEEP_WORD"
    attributes = (attributes | (1 << 8)) if keep_word else (attributes & ~(1 << 8))
    for key, bit in {
        "widowOrphan": 9,
        "keepWithNext": 10,
        "keepLines": 11,
        "pageBreakBefore": 12,
    }.items():
        enabled = _coerce_bool(style.break_setting.get(key))
        if enabled is None:
            continue
        attributes = (attributes | (1 << bit)) if enabled else (attributes & ~(1 << bit))
    line_wrap = str(style.break_setting.get("lineWrap", "")).upper()
    if line_wrap:
        attributes = (attributes | (1 << 15)) if line_wrap == "SQUEEZE" else (attributes & ~(1 << 15))
    for key, bit in {
        "eAsianEng": 23,
        "eAsianNum": 24,
    }.items():
        enabled = _coerce_bool(style.auto_spacing.get(key))
        if enabled is None:
            continue
        attributes = (attributes | (1 << bit)) if enabled else (attributes & ~(1 << bit))
    payload[0:4] = attributes.to_bytes(4, "little", signed=False)
    for key, offset in {
        "intent": 4,
        "left": 8,
        "right": 12,
        "prev": 16,
        "next": 20,
    }.items():
        if key not in style.margins:
            continue
        value = _coerce_int(style.margins.get(key))
        if value is None:
            continue
        payload[offset : offset + 4] = int(value).to_bytes(4, "little", signed=True)
    tab_pr_id = _coerce_int(style.tab_pr_id)
    if tab_pr_id is not None:
        payload[28:30] = int(tab_pr_id).to_bytes(2, "little", signed=False)
    heading_id = _coerce_int(style.heading.get("idRef")) if style.heading else None
    if heading_id is not None:
        payload[30:32] = int(heading_id).to_bytes(2, "little", signed=False)
    line_spacing = _coerce_int(style.line_spacing)
    if line_spacing is not None:
        payload[50:54] = int(line_spacing).to_bytes(4, "little", signed=True)
    return bytes(payload)


def _apply_hwp_character_style_payload(style: CharacterStyle, payload: bytes) -> None:
    values = bytes(payload)
    style.native_hwp_payload = values
    if len(values) >= 14:
        style.font_refs = {
            key: str(int.from_bytes(values[index * 2 : (index + 1) * 2], "little", signed=False))
            for index, key in enumerate(_HWP_CHAR_FONT_ORDER)
        }
    if len(values) >= 46:
        style.height = str(int.from_bytes(values[42:46], "little", signed=False))
    if len(values) >= 50:
        attributes = int.from_bytes(values[46:50], "little", signed=False)
        style.use_font_space = "1" if attributes & 0x20000000 else "0"
        style.use_kerning = "1" if attributes & 0x40000000 else "0"
    if len(values) >= 58:
        style.text_color = _parse_hwp_colorref(values[54:58])
    if len(values) >= 64:
        style.shade_color = "none" if values[60:64] == b"\xff\xff\xff\xff" else _parse_hwp_colorref(values[60:64])


def _build_hwp_character_style_payload(style: CharacterStyle, fallback_payload: bytes | None = None) -> bytes:
    payload = _seed_hwp_payload(style.native_hwp_payload, fallback_payload, _DEFAULT_HWP_CHAR_SHAPE_PAYLOAD)
    for index, key in enumerate(_HWP_CHAR_FONT_ORDER):
        if key not in style.font_refs:
            continue
        value = _coerce_int(style.font_refs.get(key))
        if value is None:
            continue
        payload[index * 2 : (index + 1) * 2] = int(value).to_bytes(2, "little", signed=False)
    height = _coerce_int(style.height)
    if height is not None:
        payload[42:46] = int(height).to_bytes(4, "little", signed=False)
    attributes = int.from_bytes(payload[46:50], "little", signed=False)
    use_font_space = _coerce_bool(style.use_font_space)
    if use_font_space is not None:
        attributes = (attributes | 0x20000000) if use_font_space else (attributes & ~0x20000000)
    use_kerning = _coerce_bool(style.use_kerning)
    if use_kerning is not None:
        attributes = (attributes | 0x40000000) if use_kerning else (attributes & ~0x40000000)
    payload[46:50] = int(attributes).to_bytes(4, "little", signed=False)
    if style.text_color is not None:
        payload[54:58] = _build_hwp_colorref(style.text_color)
    if style.shade_color is not None:
        payload[60:64] = b"\xff\xff\xff\xff" if str(style.shade_color).lower() == "none" else _build_hwp_colorref(style.shade_color)
    return bytes(payload)


def _apply_hwp_numbering_payload(definition: NumberingDefinition, payload: bytes) -> None:
    values = bytes(payload)
    definition.native_hwp_payload = values
    definition.formats = []
    definition.para_heads = []
    definition.para_head_flags = []
    definition.width_adjusts = []
    definition.text_offsets = []
    definition.char_pr_ids = []
    definition.start_numbers = []
    definition.unknown_short = 0
    definition.unknown_short_bits = {}
    cursor = 0
    for _ in range(_HWP_NUMBERING_LEVEL_COUNT):
        if cursor + 14 > len(values):
            break
        para_head = int.from_bytes(values[cursor : cursor + 4], "little", signed=False)
        definition.para_heads.append(para_head)
        definition.para_head_flags.append({f"bit_{bit}": bool(para_head & (1 << bit)) for bit in range(32)})
        cursor += 4
        definition.width_adjusts.append(int.from_bytes(values[cursor : cursor + 2], "little", signed=True))
        cursor += 2
        definition.text_offsets.append(int.from_bytes(values[cursor : cursor + 2], "little", signed=True))
        cursor += 2
        char_pr_id = int.from_bytes(values[cursor : cursor + 4], "little", signed=True)
        definition.char_pr_ids.append(None if char_pr_id < 0 else char_pr_id)
        cursor += 4
        format_length = int.from_bytes(values[cursor : cursor + 2], "little", signed=False)
        cursor += 2
        format_byte_length = min(format_length * 2, max(len(values) - cursor, 0))
        format_bytes = values[cursor : cursor + format_byte_length]
        if len(format_bytes) % 2:
            format_bytes = format_bytes[:-1]
        definition.formats.append(format_bytes.decode("utf-16-le", errors="ignore"))
        cursor += format_byte_length
    if cursor + 2 <= len(values):
        definition.unknown_short = int.from_bytes(values[cursor : cursor + 2], "little", signed=False)
        definition.unknown_short_bits = {f"bit_{bit}": bool(definition.unknown_short & (1 << bit)) for bit in range(16)}
        cursor += 2
    remaining_levels = min(_HWP_NUMBERING_LEVEL_COUNT, (len(values) - cursor) // 4)
    for _ in range(remaining_levels):
        definition.start_numbers.append(int.from_bytes(values[cursor : cursor + 4], "little", signed=False))
        cursor += 4


def _build_hwp_numbering_payload(definition: NumberingDefinition, fallback_payload: bytes | None = None) -> bytes:
    template = NumberingDefinition()
    _apply_hwp_numbering_payload(template, bytes(definition.native_hwp_payload or fallback_payload or _DEFAULT_HWP_NUMBERING_PAYLOAD))
    payload = bytearray()
    for level in range(_HWP_NUMBERING_LEVEL_COUNT):
        flag_value = _bit_dict_to_int(definition.para_head_flags[level]) if level < len(definition.para_head_flags) else 0
        para_head = (
            definition.para_heads[level]
            if level < len(definition.para_heads)
            else (flag_value if flag_value else (template.para_heads[level] if level < len(template.para_heads) else 12))
        )
        width_adjust = (
            definition.width_adjusts[level]
            if level < len(definition.width_adjusts)
            else (template.width_adjusts[level] if level < len(template.width_adjusts) else 0)
        )
        text_offset = (
            definition.text_offsets[level]
            if level < len(definition.text_offsets)
            else (template.text_offsets[level] if level < len(template.text_offsets) else 50)
        )
        if level < len(definition.char_pr_ids):
            char_pr_id = definition.char_pr_ids[level]
        else:
            char_pr_id = template.char_pr_ids[level] if level < len(template.char_pr_ids) else None
        if level < len(definition.formats) and definition.formats[level]:
            format_text = definition.formats[level]
        else:
            format_text = template.formats[level] if level < len(template.formats) else f"^{level + 1}."
        encoded_format = str(format_text).encode("utf-16-le")
        payload.extend(int(para_head).to_bytes(4, "little", signed=False))
        payload.extend(int(width_adjust).to_bytes(2, "little", signed=True))
        payload.extend(int(text_offset).to_bytes(2, "little", signed=True))
        payload.extend(int(-1 if char_pr_id is None else char_pr_id).to_bytes(4, "little", signed=True))
        payload.extend((len(encoded_format) // 2).to_bytes(2, "little", signed=False))
        payload.extend(encoded_format)
    unknown_short = definition.unknown_short if not definition.unknown_short_bits else _bit_dict_to_int(definition.unknown_short_bits)
    payload.extend(int(unknown_short).to_bytes(2, "little", signed=False))
    for level in range(_HWP_NUMBERING_LEVEL_COUNT):
        start_number = (
            definition.start_numbers[level]
            if level < len(definition.start_numbers)
            else (template.start_numbers[level] if level < len(template.start_numbers) else 1)
        )
        payload.extend(int(start_number).to_bytes(4, "little", signed=False))
    return bytes(payload)


def _apply_hwp_bullet_payload(definition: BulletDefinition, payload: bytes) -> None:
    values = bytes(payload)
    definition.native_hwp_payload = values
    if len(values) >= 4:
        definition.flags = int.from_bytes(values[0:4], "little", signed=False)
        definition.flag_bits = {f"bit_{bit}": bool(definition.flags & (1 << bit)) for bit in range(32)}
    if len(values) >= 6:
        definition.width_adjust = int.from_bytes(values[4:6], "little", signed=True)
    if len(values) >= 8:
        definition.text_offset = int.from_bytes(values[6:8], "little", signed=True)
    if len(values) >= 12:
        char_pr_id = int.from_bytes(values[8:12], "little", signed=True)
        definition.char_pr_id = None if char_pr_id < 0 else char_pr_id
    if len(values) >= 14:
        definition.bullet_char = values[12:14].decode("utf-16-le", errors="ignore") or "\uf0a1"
    definition.unknown_tail = bytes(values[14:])


def _build_hwp_bullet_payload(definition: BulletDefinition, fallback_payload: bytes | None = None) -> bytes:
    payload = _seed_hwp_payload(definition.native_hwp_payload, fallback_payload, _DEFAULT_HWP_BULLET_PAYLOAD, min_len=14)
    template = BulletDefinition()
    _apply_hwp_bullet_payload(template, bytes(payload))
    flags = definition.flags
    if flags is None and definition.flag_bits:
        flags = _bit_dict_to_int(definition.flag_bits)
    if flags is None:
        flags = template.flags if template.flags is not None else int.from_bytes(payload[0:4], "little", signed=False)
    payload[0:4] = int(flags).to_bytes(4, "little", signed=False)
    payload[4:6] = int(definition.width_adjust).to_bytes(2, "little", signed=True)
    payload[6:8] = int(definition.text_offset).to_bytes(2, "little", signed=True)
    char_pr_id = template.char_pr_id if definition.char_pr_id is None else definition.char_pr_id
    payload[8:12] = int(-1 if char_pr_id is None else char_pr_id).to_bytes(4, "little", signed=True)
    bullet_char = definition.bullet_char or template.bullet_char or "\uf0a1"
    encoded_char = bullet_char[:1].encode("utf-16-le", errors="ignore")[:2].ljust(2, b"\x00")
    payload[12:14] = encoded_char
    if definition.unknown_tail is not None:
        payload[14:] = definition.unknown_tail
    return bytes(payload)


def _build_hwp_memo_shape_payload(definition: MemoShapeDefinition, fallback_payload: bytes | None = None) -> bytes:
    semantic_fields = {
        "width": definition.width,
        "lineWidth": definition.line_width,
        "lineType": definition.line_type,
        "lineColor": definition.line_color,
        "fillColor": definition.fill_color,
        "activeColor": definition.active_color,
        "memoType": definition.memo_type,
    }
    if definition.native_hwp_payload is not None and not any(value is not None for value in semantic_fields.values()):
        return bytes(definition.native_hwp_payload)
    return _build_memo_shape_native_metadata_payload(**semantic_fields)


def _sync_hwp_memo_shape_payloads(records, nodes: list[MemoShapeDefinition]) -> None:
    for node in nodes:
        memo_shape_id = _parse_optional_int(node.memo_shape_id)
        if memo_shape_id is None or not 0 <= memo_shape_id < len(records):
            continue
        fallback_payload = bytes(records[memo_shape_id].payload)
        records[memo_shape_id].payload = bytes(_build_hwp_memo_shape_payload(node, fallback_payload))


def _merge_hwp_style_defaults_into_paragraphs(document: HancomDocument) -> None:
    style_by_id = {
        style.style_id: style
        for style in document.style_definitions
        if style.style_id is not None
    }
    for section in document.sections:
        for block in section.blocks:
            if not isinstance(block, Paragraph) or block.style_id is None:
                continue
            style = style_by_id.get(block.style_id)
            if style is None:
                continue
            if block.para_pr_id is None and style.para_pr_id is not None:
                block.para_pr_id = style.para_pr_id
            if block.char_pr_id is None and style.char_pr_id is not None:
                block.char_pr_id = style.char_pr_id


def _sync_hwp_docinfo_styles(document: HwpDocument, hancom_document: HancomDocument) -> None:
    docinfo_model = document.docinfo_model()
    initial_style_count = len(docinfo_model.style_records())
    _assign_missing_hwp_style_ids(
        hancom_document.paragraph_styles,
        existing_ids={
            value
            for value in (
                _parse_optional_int(style.style_id)
                for style in hancom_document.paragraph_styles
            )
            if value is not None
        },
    )
    _assign_missing_hwp_style_ids(
        hancom_document.character_styles,
        existing_ids={
            value
            for value in (
                _parse_optional_int(style.style_id)
                for style in hancom_document.character_styles
            )
            if value is not None
        },
    )
    _assign_missing_hwp_style_ids(
        hancom_document.style_definitions,
        existing_ids={
            value
            for value in (
                _parse_optional_int(style.style_id)
                for style in hancom_document.style_definitions
            )
            if value is not None and 0 <= value <= 0xFF
        },
        max_id=0xFF,
    )
    _assign_missing_hwp_style_ids(
        hancom_document.numbering_definitions,
        existing_ids={
            value
            for value in (
                _parse_optional_int(definition.style_id)
                for definition in hancom_document.numbering_definitions
            )
            if value is not None
        },
    )
    _assign_missing_hwp_style_ids(
        hancom_document.bullet_definitions,
        existing_ids={
            value
            for value in (
                _parse_optional_int(definition.style_id)
                for definition in hancom_document.bullet_definitions
            )
            if value is not None
        },
    )

    required_char_shape_ids = {
        value
        for value in (
            _parse_optional_int(style.style_id)
            for style in hancom_document.character_styles
        )
        if value is not None
    }
    required_para_shape_ids = {
        value
        for value in (
            _parse_optional_int(style.style_id)
            for style in hancom_document.paragraph_styles
        )
        if value is not None
    }
    required_style_ids = {
        value
        for value in (
            _parse_optional_int(style.style_id)
            for style in hancom_document.style_definitions
        )
        if value is not None and 0 <= value <= 0xFF
    }
    required_numbering_ids = {
        value
        for value in (
            _parse_optional_int(definition.style_id)
            for definition in hancom_document.numbering_definitions
        )
        if value is not None
    }
    required_bullet_ids = {
        value
        for value in (
            _parse_optional_int(definition.style_id)
            for definition in hancom_document.bullet_definitions
        )
        if value is not None
    }
    required_memo_shape_ids = {
        value
        for value in (
            _parse_optional_int(definition.memo_shape_id)
            for definition in hancom_document.memo_shape_definitions
        )
        if value is not None
    }

    for style in hancom_document.style_definitions:
        para_pr_id = _parse_optional_int(style.para_pr_id)
        char_pr_id = _parse_optional_int(style.char_pr_id)
        if para_pr_id is not None:
            required_para_shape_ids.add(para_pr_id)
        if char_pr_id is not None:
            required_char_shape_ids.add(char_pr_id)
    for definition in hancom_document.numbering_definitions:
        for char_pr_id in definition.char_pr_ids:
            if char_pr_id is not None:
                required_char_shape_ids.add(char_pr_id)
    for definition in hancom_document.bullet_definitions:
        if definition.char_pr_id is not None:
            required_char_shape_ids.add(definition.char_pr_id)

    for section in hancom_document.sections:
        for block in section.blocks:
            if not isinstance(block, Paragraph):
                continue
            style_id = block.hwp_style_id
            if style_id is None:
                style_id = _parse_optional_int(block.style_id)
            if style_id is not None and 0 <= style_id <= 0xFF:
                required_style_ids.add(style_id)
            para_shape_id = block.hwp_para_shape_id
            if para_shape_id is None:
                para_shape_id = _parse_optional_int(block.para_pr_id)
            if para_shape_id is not None:
                required_para_shape_ids.add(para_shape_id)
            char_pr_id = _parse_optional_int(block.char_pr_id)
            if char_pr_id is not None:
                required_char_shape_ids.add(char_pr_id)
        numbering_shape_id = _parse_optional_int(section.settings.numbering_shape_id)
        if numbering_shape_id is not None:
            required_numbering_ids.add(numbering_shape_id)
    if required_char_shape_ids:
        _ensure_hwp_docinfo_record_capacity(docinfo_model, TAG_CHAR_SHAPE, max(required_char_shape_ids) + 1)
    if required_para_shape_ids:
        _ensure_hwp_docinfo_record_capacity(docinfo_model, TAG_PARA_SHAPE, max(required_para_shape_ids) + 1)
    if required_style_ids:
        _ensure_hwp_docinfo_record_capacity(docinfo_model, TAG_STYLE, max(required_style_ids) + 1)
    if required_numbering_ids:
        _ensure_hwp_docinfo_record_capacity(
            docinfo_model,
            TAG_NUMBERING,
            max(required_numbering_ids) + 1,
            default_payload=_DEFAULT_HWP_NUMBERING_PAYLOAD,
        )
    if required_bullet_ids:
        _ensure_hwp_docinfo_record_capacity(
            docinfo_model,
            TAG_BULLET,
            max(required_bullet_ids) + 1,
            default_payload=_DEFAULT_HWP_BULLET_PAYLOAD,
        )
    if required_memo_shape_ids:
        _ensure_hwp_docinfo_record_capacity(
            docinfo_model,
            TAG_MEMO_SHAPE,
            max(required_memo_shape_ids) + 1,
            default_payload=_build_memo_shape_native_metadata_payload(),
        )

    _sync_hwp_docinfo_slot_payloads(
        docinfo_model.char_shape_records(),
        hancom_document.character_styles,
        builder=_build_hwp_character_style_payload,
    )
    _sync_hwp_docinfo_slot_payloads(
        docinfo_model.para_shape_records(),
        hancom_document.paragraph_styles,
        builder=_build_hwp_paragraph_style_payload,
    )
    _sync_hwp_docinfo_slot_payloads(
        docinfo_model.records_by_tag_id(TAG_NUMBERING),
        hancom_document.numbering_definitions,
        builder=_build_hwp_numbering_payload,
    )
    _sync_hwp_docinfo_slot_payloads(
        docinfo_model.records_by_tag_id(TAG_BULLET),
        hancom_document.bullet_definitions,
        builder=_build_hwp_bullet_payload,
    )
    _sync_hwp_memo_shape_payloads(
        docinfo_model.memo_shape_records(),
        hancom_document.memo_shape_definitions,
    )

    # Explicit style definitions still need direct HWP style record authoring.
    explicit_style_ids = {
        style_id
        for style_id in (_parse_optional_int(style.style_id) for style in hancom_document.style_definitions)
        if style_id is not None and 0 <= style_id <= 0xFF
    }
    for style in hancom_document.style_definitions:
        style_id = _parse_optional_int(style.style_id)
        if style_id is None or not 0 <= style_id <= 0xFF:
            continue
        docinfo_model.style_records()[style_id].payload = _build_hwp_style_definition_payload(style, style_id=style_id)

    for style_id in range(initial_style_count, len(docinfo_model.style_records())):
        if style_id in explicit_style_ids or style_id > 0xFF:
            continue
        docinfo_model.style_records()[style_id].payload = _build_hwp_style_definition_payload(
            StyleDefinition(name=f"Style {style_id}", style_id=str(style_id)),
            style_id=style_id,
        )

    document.binary_document().replace_docinfo_model(docinfo_model)


def _assign_missing_hwp_style_ids(nodes, *, existing_ids: set[int], max_id: int | None = None) -> None:
    next_id = max(existing_ids, default=-1) + 1
    for node in nodes:
        if getattr(node, "style_id", None) is not None:
            continue
        while next_id in existing_ids:
            next_id += 1
        if max_id is not None and next_id > max_id:
            break
        node.style_id = str(next_id)
        existing_ids.add(next_id)
        next_id += 1


def _sync_hwp_docinfo_slot_payloads(records, nodes, *, builder=None) -> None:
    for node in nodes:
        style_id = _parse_optional_int(getattr(node, "style_id", None))
        if style_id is None:
            continue
        if not 0 <= style_id < len(records):
            continue
        fallback_payload = bytes(records[style_id].payload)
        if builder is None:
            payload = getattr(node, "native_hwp_payload", None)
            if payload is None:
                continue
            records[style_id].payload = bytes(payload)
            continue
        records[style_id].payload = bytes(builder(node, fallback_payload))


def _ensure_hwp_docinfo_record_capacity(
    docinfo_model,
    tag_id: int,
    count: int,
    *,
    default_payload: bytes | None = None,
) -> None:
    if count <= 0:
        return
    records = docinfo_model.records_by_tag_id(tag_id)
    parent = docinfo_model.id_mappings_record()
    if records:
        template_level = records[-1].level
        template_payload = bytes(records[-1].payload)
        insert_at = max(index for index, child in enumerate(parent.children) if child.tag_id == tag_id)
    else:
        if default_payload is None:
            return
        template_level = 1
        template_payload = bytes(default_payload)
        insert_at = max((index for index, child in enumerate(parent.children) if child.tag_id <= tag_id), default=-1)
    while len(records) < count:
        clone = TypedRecord(tag_id=tag_id, level=template_level, payload=template_payload)
        insert_at += 1
        parent.children.insert(insert_at, clone)
        records.append(clone)
    count_index = _HWP_DOCINFO_COUNT_INDEX_BY_TAG.get(tag_id)
    if count_index is not None:
        docinfo_model.id_mappings_record().set_count(count_index, len(records))


def _build_hwp_style_definition_payload(style: StyleDefinition, *, style_id: int) -> bytes:
    name = style.name or f"Style {style_id}"
    english_name = style.english_name or name
    payload = bytearray()
    payload.extend(len(name).to_bytes(2, "little", signed=False))
    payload.extend(name.encode("utf-16-le"))
    payload.extend(len(english_name).to_bytes(2, "little", signed=False))
    payload.extend(english_name.encode("utf-16-le"))
    style_type_code = _HWP_STYLE_TYPE_CODES.get(str(style.style_type or "PARA").upper(), 0)
    payload.extend((((style_id & 0xFF) << 8) | (style_type_code & 0xFF)).to_bytes(2, "little", signed=False))
    payload.extend(_coerce_hwp_docinfo_word(style.lang_id, default=1042).to_bytes(2, "little", signed=False))
    payload.extend(_coerce_hwp_docinfo_word(style.next_style_id).to_bytes(2, "little", signed=False))
    payload.extend(_coerce_hwp_docinfo_word(style.para_pr_id).to_bytes(2, "little", signed=False))
    payload.extend(_coerce_hwp_docinfo_word(style.char_pr_id).to_bytes(2, "little", signed=False))
    trailing_payload = _extract_hwp_style_definition_trailing_payload(style.native_hwp_payload)
    if trailing_payload:
        payload.extend(trailing_payload)
    return bytes(payload)


def _extract_hwp_style_definition_trailing_payload(payload: bytes | None) -> bytes:
    if not payload or len(payload) < 4:
        return b""
    cursor = 0
    name_length = int.from_bytes(payload[cursor : cursor + 2], "little", signed=False)
    cursor += 2 + name_length * 2
    if cursor + 2 > len(payload):
        return b""
    english_name_length = int.from_bytes(payload[cursor : cursor + 2], "little", signed=False)
    cursor += 2 + english_name_length * 2 + 10
    if cursor >= len(payload):
        return b""
    return bytes(payload[cursor:])


def _coerce_hwp_docinfo_word(value: str | None, *, default: int = 0) -> int:
    parsed = _parse_optional_int(value)
    resolved = default if parsed is None else parsed
    return max(0, min(resolved, 0xFFFF))


def _bit_dict_to_int(bits: dict[str, bool] | None) -> int:
    if not bits:
        return 0
    value = 0
    for key, enabled in bits.items():
        if not enabled or not key.startswith("bit_"):
            continue
        try:
            bit = int(key[4:])
        except ValueError:
            continue
        if bit < 0:
            continue
        value |= 1 << bit
    return value


def _build_table_wrapper(section, node):
    from .elements import TableXml

    return TableXml(section.document, section, node)


def _build_picture_wrapper(section, node):
    from .elements import PictureXml

    return PictureXml(section.document, section, node)


def _build_field_wrapper(section, node):
    from .elements import FieldXml

    return FieldXml(section.document, section, node)


def _build_bookmark_wrapper(section, node):
    from .elements import BookmarkXml

    return BookmarkXml(section.document, section, node)


def _build_auto_number_wrapper(section, node):
    from .elements import AutoNumberXml

    return AutoNumberXml(section.document, section, node)


def _build_note_wrapper(section, node):
    from .elements import NoteXml

    return NoteXml(section.document, section, node)


def _build_memo_wrapper(section, node):
    from .elements import MemoXml

    return MemoXml(section.document, section, node)


def _build_form_wrapper(section, node):
    from .elements import FormXml

    return FormXml(section.document, section, node)


def _build_equation_wrapper(section, node):
    from .elements import EquationXml

    return EquationXml(section.document, section, node)


def _build_shape_wrapper(section, node):
    from .elements import ShapeXml

    return ShapeXml(section.document, section, node)


def _build_chart_wrapper(section, node):
    from .elements import ChartXml

    return ChartXml(section.document, section, node)


def _build_ole_wrapper(section, node):
    from .elements import OleXml

    return OleXml(section.document, section, node)


def _reset_hwpx_section(document: HwpxDocument, section_index: int) -> None:
    paragraphs = document.sections[section_index].paragraphs()
    while len(paragraphs) > 1:
        document.delete_paragraph(section_index, 1)
        paragraphs = document.sections[section_index].paragraphs()
    document.set_paragraph_text(section_index, 0, "")


def _write_section_to_hwpx(document: HwpxDocument, section_index: int, section: HancomSection) -> None:
    if _section_uses_source_paragraphs(section):
        _write_grouped_section_to_hwpx(document, section_index, section)
        return

    blocks_written = 0
    first_paragraph_consumed = False
    for block in section.blocks:
        if isinstance(block, Paragraph):
            if not first_paragraph_consumed and blocks_written == 0:
                document.set_paragraph_text(section_index, 0, block.text, char_pr_id=_resolve_hwpx_char_pr_id(document, block))
                _apply_hwpx_paragraph_style_block(document, section_index, 0, block)
                first_paragraph_consumed = True
            else:
                document.append_paragraph(
                    block.text,
                    section_index=section_index,
                    para_pr_id=_resolve_hwpx_para_pr_id(document, block),
                    style_id=_resolve_hwpx_style_id(document, block),
                    char_pr_id=_resolve_hwpx_char_pr_id(document, block),
                )
            blocks_written += 1
            continue

        paragraph_index = 0 if blocks_written == 0 else _append_control_host_paragraph(document, section_index)

        if isinstance(block, Table):
            appended = document.append_table(
                block.rows,
                block.cols,
                section_index=section_index,
                paragraph_index=paragraph_index,
                cell_texts=block.cell_texts,
            )
            _apply_hwpx_table_block(appended, block)
        elif isinstance(block, Picture):
            appended = document.append_picture(
                block.name,
                block.data,
                section_index=section_index,
                paragraph_index=paragraph_index,
                width=block.width,
                height=block.height,
                original_width=block.original_width,
                original_height=block.original_height,
                current_width=block.current_width,
                current_height=block.current_height,
                media_type=_media_type_for_extension(block.extension),
            )
            _apply_hwpx_picture_block(appended, block)
        elif isinstance(block, Hyperlink):
            document.append_hyperlink(
                block.target,
                display_text=_block_owned_text_for_write(block),
                section_index=section_index,
                paragraph_index=paragraph_index,
            )
        elif isinstance(block, Bookmark):
            document.append_bookmark(
                block.name,
                section_index=section_index,
                paragraph_index=paragraph_index,
            )
        elif isinstance(block, Field):
            document.append_field(
                field_type=block.effective_native_field_type,
                display_text=_block_owned_text_for_write(block),
                name=block.name,
                parameters=block.parameters,
                editable=block.editable,
                dirty=block.dirty,
                section_index=section_index,
                paragraph_index=paragraph_index,
            )
        elif isinstance(block, Form):
            document.append_form(
                block.label,
                form_type=block.form_type,
                name=block.name,
                value=block.value,
                checked=block.checked,
                items=block.items,
                editable=block.editable,
                locked=block.locked,
                placeholder=block.placeholder,
                section_index=section_index,
                paragraph_index=paragraph_index,
            )
        elif isinstance(block, Memo):
            document.append_memo(
                block.text,
                section_index=section_index,
                paragraph_index=paragraph_index,
            )
            if _memo_requires_hwpx_carrier(block):
                document.append_field(
                    field_type=_HWPX_MEMO_FIELD_TYPE,
                    display_text=None,
                    name=block.memo_id,
                    parameters=_build_hwpx_memo_parameters(block),
                    section_index=section_index,
                    paragraph_index=paragraph_index,
                )
        elif isinstance(block, AutoNumber):
            document.append_auto_number(
                number=block.number,
                number_type=block.number_type,
                kind=block.kind,
                section_index=section_index,
                paragraph_index=paragraph_index,
            )
        elif isinstance(block, Note):
            appended = document.append_note(
                "" if _has_nested_blocks(block) else block.text,
                kind=block.kind,
                number=block.number,
                section_index=section_index,
                paragraph_index=paragraph_index,
            )
            _apply_hwpx_nested_text_blocks(document, appended, block, vertical_align="TOP")
        elif isinstance(block, Equation):
            appended = document.append_equation(
                block.script,
                section_index=section_index,
                paragraph_index=paragraph_index,
                width=block.width,
                height=block.height,
                shape_comment=block.shape_comment,
                text_color=block.text_color,
                base_unit=block.base_unit,
                font=block.font,
            )
            _apply_hwpx_equation_block(appended, block)
        elif isinstance(block, Shape):
            appended = document.append_shape(
                kind=block.kind,
                text="" if _has_nested_blocks(block) else block.text,
                width=block.width,
                height=block.height,
                fill_color=block.fill_color,
                line_color=block.line_color,
                line_width=block.line_width,
                original_width=block.original_width,
                original_height=block.original_height,
                current_width=block.current_width,
                current_height=block.current_height,
                shape_comment=block.shape_comment,
                section_index=section_index,
                paragraph_index=paragraph_index,
            )
            _apply_hwpx_shape_block(appended, block)
            _apply_hwpx_nested_text_blocks(document, appended, block)
        elif isinstance(block, Chart):
            appended = document.append_chart(
                block.title,
                chart_type=block.chart_type,
                categories=block.categories,
                series=block.series,
                data_ref=block.data_ref,
                legend_visible=block.legend_visible,
                width=block.width,
                height=block.height,
                shape_comment=block.shape_comment,
                section_index=section_index,
                paragraph_index=paragraph_index,
            )
            if block.layout:
                appended.set_layout(**_hwpx_layout_kwargs(block.layout))
            if block.out_margins:
                appended.set_out_margins(**block.out_margins)
            if block.rotation:
                appended.set_rotation(**_hwpx_rotation_kwargs(block.rotation))
        elif isinstance(block, Ole):
            appended = document.append_ole(
                block.name,
                block.data,
                width=block.width,
                height=block.height,
                shape_comment=block.shape_comment,
                object_type=block.object_type,
                draw_aspect=block.draw_aspect,
                has_moniker=block.has_moniker,
                eq_baseline=block.eq_baseline,
                original_width=block.original_width,
                original_height=block.original_height,
                current_width=block.current_width if block.current_width is not None else 0,
                current_height=block.current_height if block.current_height is not None else 0,
                section_index=section_index,
                paragraph_index=paragraph_index,
            )
            _apply_hwpx_ole_block(appended, block)
        blocks_written += 1


def _section_uses_source_paragraphs(section: HancomSection) -> bool:
    return any(
        getattr(block, "source_paragraph_index", None) is not None
        for block in [*section.header_footer_blocks, *section.blocks]
    )


def _write_grouped_section_to_hwpx(document: HwpxDocument, section_index: int, section: HancomSection) -> None:
    paragraph_map: dict[int, int] = {}
    ordered_items = [
        block
        for _, block in sorted(
            enumerate([*section.header_footer_blocks, *section.blocks]),
            key=lambda item: (
                getattr(item[1], "source_paragraph_index", 10**9),
                getattr(item[1], "source_order", item[0]),
                item[0],
            ),
        )
    ]

    index = 0
    while index < len(ordered_items):
        block = ordered_items[index]
        next_block = ordered_items[index + 1] if index + 1 < len(ordered_items) else None
        if _is_owned_text_shadow_paragraph(block, next_block):
            index += 1
            continue

        source_paragraph_index = getattr(block, "source_paragraph_index", None)
        if source_paragraph_index is None:
            paragraph_index = _append_control_host_paragraph(document, section_index)
        else:
            paragraph_index = _ensure_hwpx_source_paragraph(
                document,
                section_index,
                paragraph_map,
                int(source_paragraph_index),
            )

        if isinstance(block, HeaderFooter):
            _append_header_footer_block_to_hwpx(document, section_index, paragraph_index, block)
        elif isinstance(block, Paragraph):
            _apply_hwpx_paragraph_style_block(document, section_index, paragraph_index, block)
            if block.text:
                _append_text_to_hwpx_paragraph(
                    document,
                    section_index,
                    paragraph_index,
                    block.text,
                    char_pr_id=_resolve_hwpx_char_pr_id(document, block),
            )
            index += 1
            continue
        else:
            _append_non_paragraph_block_to_hwpx(document, section_index, paragraph_index, block)
        index += 1


def _ensure_hwpx_source_paragraph(
    document: HwpxDocument,
    section_index: int,
    paragraph_map: dict[int, int],
    source_paragraph_index: int,
) -> int:
    if source_paragraph_index in paragraph_map:
        return paragraph_map[source_paragraph_index]
    if not paragraph_map:
        paragraph_map[source_paragraph_index] = 0
        return 0
    target_index = _append_control_host_paragraph(document, section_index)
    paragraph_map[source_paragraph_index] = target_index
    return target_index


def _append_text_to_hwpx_paragraph(
    document: HwpxDocument,
    section_index: int,
    paragraph_index: int,
    text: str,
    *,
    char_pr_id: str | None = None,
) -> None:
    paragraph = document._paragraph_element(section_index, paragraph_index)
    run = document._append_run(paragraph, char_pr_id=char_pr_id)
    text_node = etree.SubElement(run, qname("hp", "t"))
    text_node.text = text
    document.sections[section_index].mark_modified()


def _nested_blocks(block: object) -> list[HancomBlock]:
    values = getattr(block, "nested_blocks", [])
    return [value for value in values if _is_supported_nested_hwpx_block(value)]


def _has_nested_blocks(block: object) -> bool:
    return bool(_nested_blocks(block))


def _is_supported_nested_hwpx_block(value: object) -> bool:
    return isinstance(value, (Paragraph, Bookmark, Hyperlink, Field, AutoNumber))


def _build_hwpx_nested_sublist(document: HwpxDocument, blocks: list[HancomBlock], *, vertical_align: str = "CENTER") -> etree._Element:
    sublist = etree.Element(qname("hp", "subList"))
    sublist.set("id", "")
    sublist.set("textDirection", "HORIZONTAL")
    sublist.set("lineWrap", "BREAK")
    sublist.set("vertAlign", vertical_align)
    sublist.set("linkListIDRef", "0")
    sublist.set("linkListNextIDRef", "0")
    sublist.set("textWidth", "0")
    sublist.set("textHeight", "0")
    sublist.set("hasTextRef", "0")
    sublist.set("hasNumRef", "0")

    grouped: dict[int, list[HancomBlock]] = {}
    for block in blocks:
        grouped.setdefault(int(getattr(block, "nested_paragraph_index", 0)), []).append(block)
    if not grouped:
        grouped[0] = []

    for _paragraph_index, paragraph_blocks in sorted(grouped.items()):
        first_paragraph = next((block for block in paragraph_blocks if isinstance(block, Paragraph)), None)
        paragraph = etree.SubElement(sublist, qname("hp", "p"))
        paragraph.set("id", document._next_control_number(".//hp:p/@id"))
        paragraph.set("paraPrIDRef", first_paragraph.para_pr_id if isinstance(first_paragraph, Paragraph) and first_paragraph.para_pr_id else "0")
        paragraph.set("styleIDRef", first_paragraph.style_id if isinstance(first_paragraph, Paragraph) and first_paragraph.style_id else "0")
        paragraph.set("pageBreak", "0")
        paragraph.set("columnBreak", "0")
        paragraph.set("merged", "0")
        for block in _sort_hwpx_nested_blocks_by_source_order(paragraph_blocks):
            _append_nested_block_to_hwpx_paragraph(document, paragraph, block)
    return sublist


def _append_nested_block_to_hwpx_paragraph(document: HwpxDocument, paragraph: etree._Element, block: HancomBlock) -> None:
    if isinstance(block, Paragraph):
        if block.text:
            run = etree.SubElement(paragraph, qname("hp", "run"))
            run.set("charPrIDRef", block.char_pr_id or "0")
            etree.SubElement(run, qname("hp", "t")).text = block.text
        return
    if isinstance(block, Bookmark):
        run = etree.SubElement(paragraph, qname("hp", "run"))
        run.set("charPrIDRef", "0")
        ctrl = etree.SubElement(run, qname("hp", "ctrl"))
        etree.SubElement(ctrl, qname("hp", "bookmark")).set("name", block.name)
        return
    if isinstance(block, Hyperlink):
        _append_nested_field_to_hwpx_paragraph(
            document,
            paragraph,
            field_type="HYPERLINK",
            display_text=block.display_text,
            parameters=_hwpx_hyperlink_parameters(block.target),
        )
        return
    if isinstance(block, Field):
        _append_nested_field_to_hwpx_paragraph(
            document,
            paragraph,
            field_type=block.effective_native_field_type,
            display_text=block.display_text,
            name=block.name,
            parameters=block.parameters,
            editable=block.editable,
            dirty=block.dirty,
        )
        return
    if isinstance(block, AutoNumber):
        run = etree.SubElement(paragraph, qname("hp", "run"))
        run.set("charPrIDRef", "0")
        ctrl = etree.SubElement(run, qname("hp", "ctrl"))
        node = etree.SubElement(ctrl, qname("hp", block.kind))
        node.set("num", str(block.number))
        node.set("numType", str(block.number_type))


def _append_nested_field_to_hwpx_paragraph(
    document: HwpxDocument,
    paragraph: etree._Element,
    *,
    field_type: str,
    display_text: str | None = None,
    name: str | None = None,
    parameters: dict[str, int | str] | None = None,
    editable: bool = False,
    dirty: bool = False,
) -> None:
    begin_id = document._next_control_number(".//hp:fieldBegin/@id")
    field_id = document._next_control_number(".//hp:fieldBegin/@fieldid | .//hp:fieldEnd/@fieldid")

    begin_run = etree.SubElement(paragraph, qname("hp", "run"))
    begin_run.set("charPrIDRef", "0")
    begin_ctrl = etree.SubElement(begin_run, qname("hp", "ctrl"))
    field_begin = etree.SubElement(begin_ctrl, qname("hp", "fieldBegin"))
    field_begin.set("id", begin_id)
    field_begin.set("type", field_type)
    field_begin.set("name", name or "")
    field_begin.set("editable", "1" if editable else "0")
    field_begin.set("dirty", "1" if dirty else "0")
    field_begin.set("zorder", "0")
    field_begin.set("fieldid", field_id)

    if parameters:
        params = etree.SubElement(field_begin, qname("hp", "parameters"))
        params.set("cnt", str(len(parameters)))
        params.set("name", "")
        for key, value in parameters.items():
            tag_name = "integerParam" if isinstance(value, int) else "stringParam"
            parameter = etree.SubElement(params, qname("hp", tag_name))
            parameter.set("name", str(key))
            parameter.text = str(value)

    if display_text is not None:
        display_run = etree.SubElement(paragraph, qname("hp", "run"))
        display_run.set("charPrIDRef", "0")
        etree.SubElement(display_run, qname("hp", "t")).text = display_text

    end_run = etree.SubElement(paragraph, qname("hp", "run"))
    end_run.set("charPrIDRef", "0")
    end_ctrl = etree.SubElement(end_run, qname("hp", "ctrl"))
    field_end = etree.SubElement(end_ctrl, qname("hp", "fieldEnd"))
    field_end.set("beginIDRef", begin_id)
    field_end.set("fieldid", field_id)


def _hwpx_hyperlink_parameters(target: str) -> dict[str, int | str]:
    parameters: dict[str, int | str] = {"Command": target, "Path": target}
    if target.startswith("mailto:"):
        parameters["Category"] = "HWPHYPERLINK_TYPE_EMAIL"
    elif target.startswith(("http://", "https://")):
        parameters["Category"] = "HWPHYPERLINK_TYPE_WEB"
    else:
        parameters["Category"] = "HWPHYPERLINK_TYPE_PATH"
    return parameters


def _apply_hwpx_nested_text_blocks(document: HwpxDocument, wrapper, block: object, *, vertical_align: str = "CENTER") -> None:
    nested = _nested_blocks(block)
    if not nested:
        return
    sublist = _build_hwpx_nested_sublist(document, nested, vertical_align=vertical_align)
    element = wrapper.element
    local_name = etree.QName(element).localname
    if local_name in {"header", "footer", "footNote", "endNote", "hiddenComment"}:
        old_nodes = element.xpath("./hp:subList", namespaces=NS)
        if old_nodes:
            element.replace(old_nodes[0], sublist)
        else:
            element.append(sublist)
        wrapper.section.mark_modified()
        return

    draw_text_nodes = element.xpath("./hp:drawText", namespaces=NS)
    draw_text = draw_text_nodes[0] if draw_text_nodes else etree.SubElement(element, qname("hp", "drawText"))
    if not draw_text_nodes:
        draw_text.set("lastWidth", str(getattr(block, "width", 12000)))
        draw_text.set("name", "")
        draw_text.set("editable", "0")
    old_nodes = draw_text.xpath("./hp:subList", namespaces=NS)
    if old_nodes:
        draw_text.replace(old_nodes[0], sublist)
    else:
        draw_text.insert(0, sublist)
    wrapper.section.mark_modified()


def _append_non_paragraph_block_to_hwpx(
    document: HwpxDocument,
    section_index: int,
    paragraph_index: int,
    block: HancomBlock,
) -> None:
    if isinstance(block, Table):
        appended = document.append_table(
            block.rows,
            block.cols,
            section_index=section_index,
            paragraph_index=paragraph_index,
            cell_texts=block.cell_texts,
        )
        _apply_hwpx_table_block(appended, block)
    elif isinstance(block, Picture):
        appended = document.append_picture(
            block.name,
            block.data,
            section_index=section_index,
            paragraph_index=paragraph_index,
            width=block.width,
            height=block.height,
            original_width=block.original_width,
            original_height=block.original_height,
            current_width=block.current_width,
            current_height=block.current_height,
            media_type=_media_type_for_extension(block.extension),
        )
        _apply_hwpx_picture_block(appended, block)
    elif isinstance(block, Hyperlink):
        document.append_hyperlink(
            block.target,
            display_text=_block_owned_text_for_write(block),
            section_index=section_index,
            paragraph_index=paragraph_index,
        )
    elif isinstance(block, Bookmark):
        document.append_bookmark(
            block.name,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )
    elif isinstance(block, Field):
        document.append_field(
            field_type=block.effective_native_field_type,
            display_text=_block_owned_text_for_write(block),
            name=block.name,
            parameters=block.parameters,
            editable=block.editable,
            dirty=block.dirty,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )
    elif isinstance(block, Form):
        document.append_form(
            block.label,
            form_type=block.form_type,
            name=block.name,
            value=block.value,
            checked=block.checked,
            items=block.items,
            editable=block.editable,
            locked=block.locked,
            placeholder=block.placeholder,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )
    elif isinstance(block, Memo):
        document.append_memo(
            block.text,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )
        if _memo_requires_hwpx_carrier(block):
            document.append_field(
                field_type=_HWPX_MEMO_FIELD_TYPE,
                display_text=None,
                name=block.memo_id,
                parameters=_build_hwpx_memo_parameters(block),
                section_index=section_index,
                paragraph_index=paragraph_index,
            )
    elif isinstance(block, AutoNumber):
        document.append_auto_number(
            number=block.number,
            number_type=block.number_type,
            kind=block.kind,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )
    elif isinstance(block, Note):
        appended = document.append_note(
            "" if _has_nested_blocks(block) else block.text,
            kind=block.kind,
            number=block.number,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )
        _apply_hwpx_nested_text_blocks(document, appended, block, vertical_align="TOP")
    elif isinstance(block, Equation):
        appended = document.append_equation(
            block.script,
            section_index=section_index,
            paragraph_index=paragraph_index,
            width=block.width,
            height=block.height,
            shape_comment=block.shape_comment,
            text_color=block.text_color,
            base_unit=block.base_unit,
            font=block.font,
        )
        _apply_hwpx_equation_block(appended, block)
    elif isinstance(block, Shape):
        appended = document.append_shape(
            kind=block.kind,
            text="" if _has_nested_blocks(block) else block.text,
            width=block.width,
            height=block.height,
            fill_color=block.fill_color,
            line_color=block.line_color,
            line_width=block.line_width,
            original_width=block.original_width,
            original_height=block.original_height,
            current_width=block.current_width,
            current_height=block.current_height,
            shape_comment=block.shape_comment,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )
        _apply_hwpx_shape_block(appended, block)
        _apply_hwpx_nested_text_blocks(document, appended, block)
    elif isinstance(block, Chart):
        appended = document.append_chart(
            block.title,
            chart_type=block.chart_type,
            categories=block.categories,
            series=block.series,
            data_ref=block.data_ref,
            legend_visible=block.legend_visible,
            width=block.width,
            height=block.height,
            shape_comment=block.shape_comment,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )
        if block.layout:
            appended.set_layout(**_hwpx_layout_kwargs(block.layout))
        if block.out_margins:
            appended.set_out_margins(**block.out_margins)
        if block.rotation:
            appended.set_rotation(**_hwpx_rotation_kwargs(block.rotation))
    elif isinstance(block, Ole):
        appended = document.append_ole(
            block.name,
            block.data,
            width=block.width,
            height=block.height,
            shape_comment=block.shape_comment,
            object_type=block.object_type,
            draw_aspect=block.draw_aspect,
            has_moniker=block.has_moniker,
            eq_baseline=block.eq_baseline,
            original_width=block.original_width,
            original_height=block.original_height,
            current_width=block.current_width if block.current_width is not None else 0,
            current_height=block.current_height if block.current_height is not None else 0,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )
        _apply_hwpx_ole_block(appended, block)


def _append_header_footer_block_to_hwpx(
    document: HwpxDocument,
    section_index: int,
    paragraph_index: int,
    block: HeaderFooter,
) -> None:
    if block.kind == "header":
        appended = document.append_header(
            "" if _has_nested_blocks(block) else block.text,
            apply_page_type=block.apply_page_type,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )
        _apply_hwpx_nested_text_blocks(document, appended, block, vertical_align="TOP")
    elif block.kind == "footer":
        appended = document.append_footer(
            "" if _has_nested_blocks(block) else block.text,
            apply_page_type=block.apply_page_type,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )
        _apply_hwpx_nested_text_blocks(document, appended, block, vertical_align="TOP")


def _append_control_host_paragraph(document: HwpxDocument, section_index: int) -> int:
    document.append_paragraph("", section_index=section_index)
    return len(document.sections[section_index].paragraphs()) - 1


def _apply_hwpx_paragraph_style_block(document: HwpxDocument, section_index: int, paragraph_index: int, block: Paragraph) -> None:
    style_id = _resolve_hwpx_style_id(document, block)
    para_pr_id = _resolve_hwpx_para_pr_id(document, block)
    char_pr_id = _resolve_hwpx_char_pr_id(document, block)
    if style_id is None and para_pr_id is None and char_pr_id is None:
        return
    document.apply_style_to_paragraph(
        section_index,
        paragraph_index,
        style_id=style_id,
        para_pr_id=para_pr_id,
        char_pr_id=char_pr_id,
    )


def _resolve_hwpx_style_id(document: HwpxDocument, block: Paragraph) -> str | None:
    if block.style_id is None:
        return None
    known = {style.style_id for style in document.styles() if style.style_id is not None}
    return block.style_id if block.style_id in known else None


def _resolve_hwpx_para_pr_id(document: HwpxDocument, block: Paragraph) -> str | None:
    candidate = block.para_pr_id
    if candidate is None and block.hwp_para_shape_id is not None:
        candidate = str(block.hwp_para_shape_id)
    if candidate is None:
        return None
    known = {style.style_id for style in document.paragraph_styles() if style.style_id is not None}
    return candidate if candidate in known else None


def _resolve_hwpx_char_pr_id(document: HwpxDocument, block: Paragraph) -> str | None:
    if block.char_pr_id is None:
        return None
    known = {style.style_id for style in document.character_styles() if style.style_id is not None}
    return block.char_pr_id if block.char_pr_id in known else None


def _media_type_for_extension(extension: str | None) -> str | None:
    if extension is None:
        return None
    normalized = extension.lower().lstrip(".")
    if normalized in {"jpg", "jpeg"}:
        return "image/jpeg"
    if normalized == "png":
        return "image/png"
    if normalized == "gif":
        return "image/gif"
    if normalized == "bmp":
        return "image/bmp"
    return None


def _apply_section_settings(
    document: HwpxDocument,
    section_index: int,
    settings: SectionSettings,
    *,
    apply_page_numbers: bool = True,
) -> None:
    target = document.section_settings(section_index)
    if settings.page_width is not None or settings.page_height is not None or settings.landscape is not None:
        target.set_page_size(
            width=settings.page_width,
            height=settings.page_height,
            landscape=settings.landscape,
        )
    if settings.margins:
        target.set_margins(**settings.margins)
    if settings.page_border_fills:
        _apply_page_border_fills(document, section_index, settings.page_border_fills)
    if settings.visibility:
        target.set_visibility(
            hide_first_header=_coerce_bool(settings.visibility.get("hideFirstHeader")),
            hide_first_footer=_coerce_bool(settings.visibility.get("hideFirstFooter")),
            hide_first_master_page=_coerce_bool(settings.visibility.get("hideFirstMasterPage")),
            border=settings.visibility.get("border"),
            fill=settings.visibility.get("fill"),
            hide_first_page_num=_coerce_bool(settings.visibility.get("hideFirstPageNum")),
            hide_first_empty_line=_coerce_bool(settings.visibility.get("hideFirstEmptyLine")),
            show_line_number=_coerce_bool(settings.visibility.get("showLineNumber")),
        )
    if settings.grid:
        target.set_grid(
            line_grid=settings.grid.get("lineGrid"),
            char_grid=settings.grid.get("charGrid"),
            wonggoji_format=_coerce_bool(settings.grid.get("wonggojiFormat")),
        )
    if settings.start_numbers:
        target.set_start_numbers(
            page_starts_on=settings.start_numbers.get("pageStartsOn"),
            page=settings.start_numbers.get("page"),
            pic=settings.start_numbers.get("pic"),
            tbl=settings.start_numbers.get("tbl"),
            equation=settings.start_numbers.get("equation"),
        )
    if settings.memo_shape_id is not None:
        target.set_memo_shape_id(settings.memo_shape_id)
    if settings.line_number_shape:
        _apply_line_number_shape(document, section_index, settings.line_number_shape)
    if settings.footnote_pr or settings.endnote_pr:
        _apply_note_settings(
            document,
            section_index,
            footnote_pr=settings.footnote_pr or None,
            endnote_pr=settings.endnote_pr or None,
        )
    if apply_page_numbers and settings.page_numbers:
        _apply_page_numbers(document, section_index, settings.page_numbers)


def _apply_page_border_fills(
    document: HwpxDocument,
    section_index: int,
    page_border_fills: list[dict[str, str | int]],
) -> None:
    sec_pr = document.section_xml(section_index).find(".//hp:secPr")
    if sec_pr is None:
        return
    for node in list(sec_pr.element.xpath("./hp:pageBorderFill", namespaces=NS)):
        node.getparent().remove(node)
    for border_fill in page_border_fills:
        border_type = str(border_fill.get("type", "BOTH"))
        node = etree.SubElement(sec_pr.element, f"{{{NS['hp']}}}pageBorderFill")
        node.set("type", border_type)
        node.set("borderFillIDRef", str(border_fill.get("borderFillIDRef", "0")))
        node.set("textBorder", str(border_fill.get("textBorder", "PAPER")))
        node.set("headerInside", str(border_fill.get("headerInside", "0")))
        node.set("footerInside", str(border_fill.get("footerInside", "0")))
        node.set("fillArea", str(border_fill.get("fillArea", "PAPER")))
        offset = etree.SubElement(node, f"{{{NS['hp']}}}offset")
        offset.set("left", str(border_fill.get("left", 0)))
        offset.set("right", str(border_fill.get("right", 0)))
        offset.set("top", str(border_fill.get("top", 0)))
        offset.set("bottom", str(border_fill.get("bottom", 0)))
    document.sections[section_index].mark_modified()


def _apply_line_number_shape(
    document: HwpxDocument,
    section_index: int,
    line_number_shape: dict[str, str],
) -> None:
    node = document.section_xml(section_index).find(".//hp:lineNumberShape")
    if node is None:
        return
    for key in ("restartType", "countBy", "distance", "startNumber"):
        value = line_number_shape.get(key)
        if value is not None:
            node.set_attr(key, value)


def _apply_note_settings(
    document: HwpxDocument,
    section_index: int,
    *,
    footnote_pr: dict[str, object] | None = None,
    endnote_pr: dict[str, object] | None = None,
) -> None:
    sec_pr = document.section_xml(section_index).find(".//hp:secPr")
    if sec_pr is None:
        return
    for kind in ("footNotePr", "endNotePr"):
        for node in list(sec_pr.element.xpath(f"./hp:{kind}", namespaces=NS)):
            node.getparent().remove(node)
    if footnote_pr:
        _append_note_pr(sec_pr.element, "footNotePr", footnote_pr)
    if endnote_pr:
        _append_note_pr(sec_pr.element, "endNotePr", endnote_pr)
    document.sections[section_index].mark_modified()


def _append_note_pr(sec_pr: etree._Element, kind: str, note_pr: dict[str, object]) -> None:
    note_node = etree.SubElement(sec_pr, f"{{{NS['hp']}}}{kind}")
    auto_num = dict(note_pr.get("autoNumFormat", {}))
    auto_num_node = etree.SubElement(note_node, f"{{{NS['hp']}}}autoNumFormat")
    auto_num_node.set("type", _xml_safe_scalar(auto_num.get("type", "DIGIT")))
    auto_num_node.set("userChar", _xml_safe_scalar(auto_num.get("userChar", "")))
    auto_num_node.set("prefixChar", _xml_safe_scalar(auto_num.get("prefixChar", "")))
    auto_num_node.set("suffixChar", _xml_safe_scalar(auto_num.get("suffixChar", "")))
    auto_num_node.set("supscript", _xml_safe_scalar(auto_num.get("supscript", "0")))
    note_line = dict(note_pr.get("noteLine", {}))
    note_line_node = etree.SubElement(note_node, f"{{{NS['hp']}}}noteLine")
    note_line_node.set("length", str(note_line.get("length", "-1")))
    note_line_node.set("type", str(note_line.get("type", "SOLID")))
    note_line_node.set("width", str(note_line.get("width", "0.12 mm")))
    note_line_node.set("color", str(note_line.get("color", "#000000")))
    note_spacing = dict(note_pr.get("noteSpacing", {}))
    note_spacing_node = etree.SubElement(note_node, f"{{{NS['hp']}}}noteSpacing")
    note_spacing_node.set("betweenNotes", str(note_spacing.get("betweenNotes", "283")))
    note_spacing_node.set("belowLine", str(note_spacing.get("belowLine", "567")))
    note_spacing_node.set("aboveLine", str(note_spacing.get("aboveLine", "850")))
    numbering = dict(note_pr.get("numbering", {}))
    numbering_node = etree.SubElement(note_node, f"{{{NS['hp']}}}numbering")
    numbering_node.set("type", str(numbering.get("type", "CONTINUOUS")))
    numbering_node.set("newNum", str(numbering.get("newNum", "1")))
    placement = dict(note_pr.get("placement", {}))
    placement_node = etree.SubElement(note_node, f"{{{NS['hp']}}}placement")
    placement_node.set("place", str(placement.get("place", "END_OF_DOCUMENT" if kind == "endNotePr" else "EACH_COLUMN")))
    placement_node.set("beneathText", str(placement.get("beneathText", "0")))


def _apply_page_numbers(
    document: HwpxDocument,
    section_index: int,
    page_numbers: list[dict[str, str]],
) -> None:
    section_xml = document.section_xml(section_index)
    for node in section_xml.findall(".//hp:pageNum"):
        parent = node.element.getparent()
        if parent is not None:
            parent.remove(node.element)
    for page_number in page_numbers:
        paragraph_index = _append_control_host_paragraph(document, section_index)
        document.append_control_xml(_page_number_xml(page_number), section_index=section_index, paragraph_index=paragraph_index)


def _apply_page_numbers_to_source_paragraphs(
    document: HwpxDocument,
    section_index: int,
    settings: SectionSettings,
) -> None:
    section_xml = document.section_xml(section_index)
    for node in section_xml.findall(".//hp:pageNum"):
        parent = node.element.getparent()
        if parent is not None:
            parent.remove(node.element)

    paragraph_count = len(document.sections[section_index].paragraphs())
    for index, page_number in enumerate(settings.page_numbers):
        paragraph_index: int | None = None
        if index < len(settings.page_number_positions):
            paragraph_index = settings.page_number_positions[index][0]
        if paragraph_index is None or paragraph_index >= paragraph_count:
            paragraph_index = _append_control_host_paragraph(document, section_index)
            paragraph_count += 1
        document.append_control_xml(_page_number_xml(page_number), section_index=section_index, paragraph_index=paragraph_index)


def _page_number_xml(page_number: dict[str, str]) -> str:
    return (
        '<hp:pageNum xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" '
        f'pos="{page_number.get("pos", "BOTTOM_CENTER")}" '
        f'formatType="{page_number.get("formatType", "DIGIT")}" '
        f'sideChar="{page_number.get("sideChar", "-")}"/>'
    )


def _apply_header_footer_blocks(document: HwpxDocument, section_index: int, blocks: list[HeaderFooter]) -> None:
    for block in blocks:
        paragraph_index = _append_control_host_paragraph(document, section_index)
        _append_header_footer_block_to_hwpx(document, section_index, paragraph_index, block)


def _apply_styles(document: HwpxDocument, hancom_document: HancomDocument) -> None:
    for style in hancom_document.paragraph_styles:
        appended = document.append_paragraph_style(
            style_id=style.style_id,
            alignment_horizontal=style.alignment_horizontal,
            alignment_vertical=style.alignment_vertical,
            line_spacing=None,
        )
        for key, value in {
            "tabPrIDRef": style.tab_pr_id,
            "snapToGrid": style.snap_to_grid,
            "condense": style.condense,
            "fontLineHeight": style.font_line_height,
            "suppressLineNumbers": style.suppress_line_numbers,
            "checked": style.checked,
        }.items():
            if value is not None:
                appended.element.set(key, str(value))
        if any(value is not None for value in (style.tab_pr_id, style.snap_to_grid, style.condense, style.font_line_height, style.suppress_line_numbers, style.checked)):
            appended.header_part.mark_modified()
        if style.line_spacing is not None:
            appended.set_line_spacing(style.line_spacing, spacing_type=style.line_spacing_type)
        if style.margins:
            appended.set_margin(
                intent=style.margins.get("intent"),
                left=style.margins.get("left"),
                right=style.margins.get("right"),
                prev=style.margins.get("prev"),
                next=style.margins.get("next"),
                unit=style.margins.get("unit"),
            )
        if style.heading:
            heading_nodes = appended.element.xpath("./hh:heading", namespaces=NS)
            heading = heading_nodes[0] if heading_nodes else etree.SubElement(appended.element, qname("hh", "heading"))
            for attribute in ("type", "idRef", "level"):
                value = style.heading.get(attribute)
                if value is not None:
                    heading.set(attribute, str(value))
            appended.header_part.mark_modified()
        if style.break_setting:
            appended.set_break_setting(
                break_latin_word=style.break_setting.get("breakLatinWord"),
                break_non_latin_word=style.break_setting.get("breakNonLatinWord"),
                widow_orphan=_coerce_bool(style.break_setting.get("widowOrphan")),
                keep_with_next=_coerce_bool(style.break_setting.get("keepWithNext")),
                keep_lines=_coerce_bool(style.break_setting.get("keepLines")),
                page_break_before=_coerce_bool(style.break_setting.get("pageBreakBefore")),
                line_wrap=style.break_setting.get("lineWrap"),
            )
        if style.auto_spacing:
            appended.set_auto_spacing(
                e_asian_eng=_coerce_bool(style.auto_spacing.get("eAsianEng")),
                e_asian_num=_coerce_bool(style.auto_spacing.get("eAsianNum")),
            )
        if style.style_id is None:
            style.style_id = appended.style_id

    for style in hancom_document.character_styles:
        appended = document.append_character_style(
            style_id=style.style_id,
            text_color=style.text_color,
            height=style.height,
        )
        if any(value is not None for value in (style.shade_color, style.use_font_space, style.use_kerning, style.sym_mark)):
            appended.set_effects(
                shade_color=style.shade_color,
                use_font_space=_coerce_bool(style.use_font_space),
                use_kerning=_coerce_bool(style.use_kerning),
                sym_mark=style.sym_mark,
            )
        if style.font_refs:
            appended.set_font_refs(
                hangul=style.font_refs.get("hangul"),
                latin=style.font_refs.get("latin"),
                hanja=style.font_refs.get("hanja"),
                japanese=style.font_refs.get("japanese"),
                other=style.font_refs.get("other"),
                symbol=style.font_refs.get("symbol"),
                user=style.font_refs.get("user"),
            )
        if style.style_id is None:
            style.style_id = appended.style_id

    for style in hancom_document.style_definitions:
        document.append_style(
            style.name,
            english_name=style.english_name,
            style_id=style.style_id,
            style_type=style.style_type,
            para_pr_id=style.para_pr_id,
            char_pr_id=style.char_pr_id,
            next_style_id=style.next_style_id,
            lang_id=style.lang_id,
            lock_form=style.lock_form,
        )

    for memo_shape in hancom_document.memo_shape_definitions:
        appended = document.append_memo_shape(
            memo_shape_id=memo_shape.memo_shape_id,
            width=memo_shape.width if memo_shape.width is not None else 12000,
            line_width=memo_shape.line_width if memo_shape.line_width is not None else 12,
            line_type=memo_shape.line_type or "SOLID",
            line_color=memo_shape.line_color or "#808080",
            fill_color=memo_shape.fill_color or "#FFF2CC",
            active_color=memo_shape.active_color or "#FFC000",
            memo_type=memo_shape.memo_type or "NOMAL",
        )
        if memo_shape.memo_shape_id is None:
            memo_shape.memo_shape_id = appended.memo_shape_id


def _coerce_bool(value) -> bool | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, int):
        return bool(value)
    text = str(value).strip().lower()
    if text in {"1", "true", "yes"}:
        return True
    if text in {"0", "false", "no"}:
        return False
    return None
