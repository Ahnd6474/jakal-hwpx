from __future__ import annotations

from dataclasses import dataclass, field

from .document import DocumentMetadata, HwpxDocument

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


# Keep repr/pickle paths stable while the implementation lives behind the public façade.
_PUBLIC_MODEL_TYPES = (
    HancomMetadata,
    Paragraph,
    Table,
    Picture,
    Hyperlink,
    Bookmark,
    Field,
    AutoNumber,
    Note,
    Equation,
    Shape,
    Ole,
    Form,
    Memo,
    Chart,
    HeaderFooter,
    StyleDefinition,
    ParagraphStyle,
    CharacterStyle,
    NumberingDefinition,
    BulletDefinition,
    MemoShapeDefinition,
    SectionSettings,
    HancomSection,
)
for _model_type in _PUBLIC_MODEL_TYPES:
    _model_type.__module__ = "jakal_hwpx.hancom_document"
