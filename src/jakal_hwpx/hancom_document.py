from __future__ import annotations

import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

from lxml import etree

from ._hancom import convert_document
from .document import DocumentMetadata, HwpxDocument
from .exceptions import HancomInteropError
from .hwp_document import HwpDocument
from .namespaces import NS

HancomConverter = Callable[[str | Path, str | Path, str], Path]


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


@dataclass
class Table:
    rows: int
    cols: int
    cell_texts: list[list[str]] = field(default_factory=list)
    row_heights: list[int] | None = None
    col_widths: list[int] | None = None
    cell_spans: dict[tuple[int, int], tuple[int, int]] | None = None
    cell_border_fill_ids: dict[tuple[int, int], int] | None = None
    table_border_fill_id: int = 1


@dataclass
class Picture:
    name: str
    data: bytes
    extension: str | None = None
    width: int = 7200
    height: int = 7200


@dataclass
class Hyperlink:
    target: str
    display_text: str | None = None


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


@dataclass
class AutoNumber:
    kind: str = "newNum"
    number: int | str = 1
    number_type: str = "PAGE"


@dataclass
class Note:
    kind: str
    text: str
    number: int | None = None


@dataclass
class Equation:
    script: str
    width: int = 4800
    height: int = 2300


@dataclass
class Shape:
    kind: str
    text: str = ""
    width: int = 12000
    height: int = 3200
    fill_color: str = "#FFFFFF"
    line_color: str = "#000000"


@dataclass
class Ole:
    name: str
    data: bytes
    width: int = 42001
    height: int = 13501


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
)


@dataclass
class HeaderFooter:
    kind: str
    text: str
    apply_page_type: str = "BOTH"


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


@dataclass
class ParagraphStyle:
    style_id: str | None = None
    alignment_horizontal: str | None = None
    alignment_vertical: str | None = None
    line_spacing: int | str | None = None


@dataclass
class CharacterStyle:
    style_id: str | None = None
    text_color: str | None = None
    height: int | str | None = None


@dataclass
class SectionSettings:
    page_width: int | None = None
    page_height: int | None = None
    landscape: str | None = None
    margins: dict[str, int] = field(default_factory=dict)
    visibility: dict[str, str] = field(default_factory=dict)
    grid: dict[str, int] = field(default_factory=dict)
    start_numbers: dict[str, str] = field(default_factory=dict)


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
        self._converter = converter or convert_document
        self._temp_dir: Path | None = None
        self.style_definitions: list[StyleDefinition] = []
        self.paragraph_styles: list[ParagraphStyle] = []
        self.character_styles: list[CharacterStyle] = []

    @classmethod
    def blank(cls, *, converter: HancomConverter | None = None) -> "HancomDocument":
        return cls(converter=converter)

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
        effective_converter = converter or getattr(document, "_converter", None) or convert_document
        fallback_sections = [
            HancomSection(blocks=[Paragraph(paragraph.text.strip()) for paragraph in section.paragraphs() if paragraph.text.strip()])
            for section in document.sections()
        ]
        fallback_sections = fallback_sections or [HancomSection()]
        try:
            bridged = document.to_hwpx_document(force_refresh=False)
        except HancomInteropError:
            metadata = HancomMetadata(title=document.source_path.stem if document.source_path else None)
            return cls(metadata=metadata, sections=fallback_sections, source_format="hwp", converter=effective_converter)

        instance = cls.from_hwpx_document(bridged, converter=effective_converter)
        instance.source_format = "hwp"
        if instance.metadata.title is None and document.source_path is not None:
            instance.metadata.title = document.source_path.stem
        if not any(section.blocks for section in instance.sections) and any(section.blocks for section in fallback_sections):
            instance.sections = fallback_sections
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
        section_index: int = 0,
    ) -> Picture:
        section = self._ensure_section(section_index)
        block = Picture(name=name, data=data, extension=extension, width=width, height=height)
        section.blocks.append(block)
        return block

    def append_hyperlink(
        self,
        target: str,
        *,
        display_text: str | None = None,
        section_index: int = 0,
    ) -> Hyperlink:
        section = self._ensure_section(section_index)
        block = Hyperlink(target=target, display_text=display_text)
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
        block = Field(
            field_type=field_type,
            display_text=display_text,
            name=name,
            parameters=dict(parameters or {}),
            editable=editable,
            dirty=dirty,
        )
        section.blocks.append(block)
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

    def append_equation(
        self,
        script: str,
        *,
        width: int = 4800,
        height: int = 2300,
        section_index: int = 0,
    ) -> Equation:
        section = self._ensure_section(section_index)
        block = Equation(script=script, width=width, height=height)
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
        section_index: int = 0,
    ) -> Ole:
        section = self._ensure_section(section_index)
        block = Ole(name=name, data=data, width=width, height=height)
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
    ) -> ParagraphStyle:
        node = ParagraphStyle(
            style_id=style_id,
            alignment_horizontal=alignment_horizontal,
            alignment_vertical=alignment_vertical,
            line_spacing=line_spacing,
        )
        self.paragraph_styles.append(node)
        return node

    def append_character_style(
        self,
        *,
        style_id: str | None = None,
        text_color: str | None = None,
        height: int | str | None = None,
    ) -> CharacterStyle:
        node = CharacterStyle(style_id=style_id, text_color=text_color, height=height)
        self.character_styles.append(node)
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
            _apply_section_settings(document, section_index, section.settings)
            _apply_header_footer_blocks(document, section_index, section.header_footer_blocks)
            _write_section_to_hwpx(document, section_index, section)
        return document

    def write_to_hwpx(self, path: str | Path, *, validate: bool = True) -> Path:
        return self.to_hwpx_document().save(path, validate=validate)

    def to_hwp_document(self, *, converter: HancomConverter | None = None) -> HwpDocument:
        effective_converter = converter or self._converter
        workspace = self._ensure_temp_dir()
        hwpx_path = workspace / "hancom_document.hwpx"
        hwp_path = workspace / "hancom_document.hwp"
        self.write_to_hwpx(hwpx_path)
        effective_converter(hwpx_path, hwp_path, "HWP")
        return HwpDocument.open(hwp_path, converter=effective_converter)

    def write_to_hwp(self, path: str | Path, *, converter: HancomConverter | None = None) -> Path:
        effective_converter = converter or self._converter
        return self.to_hwpx_document().save_as_hwp(path, converter=effective_converter)

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


def _extract_hwpx_section(section) -> HancomSection:
    blocks: list[HancomBlock] = []
    section_settings = _extract_section_settings(section)
    header_footer_blocks = _extract_header_footer_blocks(section)
    for paragraph in section.root_element.xpath("./hp:p", namespaces=NS):
        tables = paragraph.xpath(".//hp:tbl", namespaces=NS)
        pictures = paragraph.xpath(".//hp:pic", namespaces=NS)
        hyperlinks = paragraph.xpath(".//hp:fieldBegin[@type='HYPERLINK']", namespaces=NS)
        fields = paragraph.xpath(".//hp:fieldBegin[not(@type='HYPERLINK')]", namespaces=NS)
        bookmarks = paragraph.xpath(".//hp:bookmark", namespaces=NS)
        auto_numbers = paragraph.xpath(".//hp:autoNum | .//hp:newNum", namespaces=NS)
        notes = paragraph.xpath(".//hp:footNote | .//hp:endNote", namespaces=NS)
        equations = paragraph.xpath(".//hp:equation", namespaces=NS)
        shapes = paragraph.xpath(
            ".//hp:rect | .//hp:line | .//hp:ellipse | .//hp:arc | .//hp:polygon | .//hp:curve | .//hp:connectLine | .//hp:textart | .//hp:container",
            namespaces=NS,
        )
        oles = paragraph.xpath(".//hp:ole", namespaces=NS)

        text = "".join(paragraph.xpath("./hp:run/hp:t/text()", namespaces=NS)).replace("\r", "").strip()
        if text:
            blocks.append(Paragraph(text))
        for node in bookmarks:
            blocks.append(_extract_bookmark(_build_bookmark_wrapper(section, node)))
        for node in hyperlinks:
            field = _build_field_wrapper(section, node)
            blocks.append(Hyperlink(target=field.hyperlink_target or "", display_text=field.display_text or None))
        for node in fields:
            blocks.append(_extract_field(_build_field_wrapper(section, node)))
        for node in auto_numbers:
            blocks.append(_extract_auto_number(_build_auto_number_wrapper(section, node)))
        for node in tables:
            blocks.append(_extract_hwpx_table(_build_table_wrapper(section, node)))
        for node in pictures:
            blocks.append(_extract_hwpx_picture(_build_picture_wrapper(section, node)))
        for node in notes:
            blocks.append(_extract_hwpx_note(_build_note_wrapper(section, node)))
        for node in equations:
            blocks.append(_extract_hwpx_equation(_build_equation_wrapper(section, node)))
        for node in shapes:
            blocks.append(_extract_hwpx_shape(_build_shape_wrapper(section, node)))
        for node in oles:
            blocks.append(_extract_hwpx_ole(_build_ole_wrapper(section, node)))
    return HancomSection(settings=section_settings, header_footer_blocks=header_footer_blocks, blocks=blocks)


def _extract_hwpx_table(table) -> Table:
    cell_texts = _blank_cell_texts(table.row_count, table.column_count)
    cell_spans: dict[tuple[int, int], tuple[int, int]] = {}
    for cell in table.cells():
        if cell.row < table.row_count and cell.column < table.column_count:
            cell_texts[cell.row][cell.column] = cell.text
            if cell.row_span != 1 or cell.col_span != 1:
                cell_spans[(cell.row, cell.column)] = (cell.row_span, cell.col_span)
    return Table(
        rows=table.row_count,
        cols=table.column_count,
        cell_texts=cell_texts,
        cell_spans=cell_spans or None,
    )


def _extract_hwpx_picture(picture) -> Picture:
    path = picture.binary_part_path()
    extension = Path(path).suffix.lstrip(".") or None
    size = picture.size()
    return Picture(
        name=Path(path).name,
        data=picture.binary_data(),
        extension=extension,
        width=size.get("width", 7200),
        height=size.get("height", 7200),
    )


def _extract_hwpx_note(note) -> Note:
    number = int(note.number) if note.number and str(note.number).isdigit() else None
    return Note(kind=note.kind, text=note.text, number=number)


def _extract_hwpx_equation(equation) -> Equation:
    size = equation.size()
    return Equation(
        script=equation.script,
        width=size.get("width", 4800),
        height=size.get("height", 2300),
    )


def _extract_hwpx_shape(shape) -> Shape:
    size = shape.size()
    fill_style = shape.fill_style()
    line_style = shape.line_style()
    return Shape(
        kind=shape.kind,
        text=shape.text,
        width=size.get("width", 12000),
        height=size.get("height", 3200),
        fill_color=fill_style.get("faceColor", "#FFFFFF"),
        line_color=line_style.get("color", "#000000"),
    )


def _extract_hwpx_ole(ole) -> Ole:
    size = ole.size()
    return Ole(
        name=Path(ole.binary_part_path()).name,
        data=ole.binary_data(),
        width=size.get("width", 42001),
        height=size.get("height", 13501),
    )


def _extract_bookmark(bookmark) -> Bookmark:
    return Bookmark(name=bookmark.name or "")


def _extract_field(field) -> Field:
    return Field(
        field_type=field.field_type or "",
        display_text=field.display_text or None,
        name=field.name,
        parameters=field.parameter_map(),
        editable=field.element.get("editable") == "1",
        dirty=field.element.get("dirty") == "1",
    )


def _extract_auto_number(auto_number) -> AutoNumber:
    return AutoNumber(
        kind=auto_number.kind,
        number=auto_number.number or 1,
        number_type=auto_number.number_type or "PAGE",
    )


def _extract_section_settings(section) -> SectionSettings:
    settings = section.document.section_settings(section.section_index())
    return SectionSettings(
        page_width=settings.page_width,
        page_height=settings.page_height,
        landscape=settings.landscape,
        margins=settings.margins(),
        visibility=settings.visibility(),
        grid=settings.grid(),
        start_numbers=settings.start_numbers(),
    )


def _extract_header_footer_blocks(section) -> list[HeaderFooter]:
    section_index = section.section_index()
    blocks: list[HeaderFooter] = []
    for header in section.document.headers(section_index=section_index):
        blocks.append(
            HeaderFooter(
                kind=header.kind,
                text=header.text,
                apply_page_type=header.apply_page_type or "BOTH",
            )
        )
    for footer in section.document.footers(section_index=section_index):
        blocks.append(
            HeaderFooter(
                kind=footer.kind,
                text=footer.text,
                apply_page_type=footer.apply_page_type or "BOTH",
            )
        )
    return blocks


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
    return ParagraphStyle(
        style_id=style.style_id,
        alignment_horizontal=style.alignment_horizontal,
        alignment_vertical=style.element.xpath("./hh:align/@vertical", namespaces=NS)[0]
        if style.element.xpath("./hh:align/@vertical", namespaces=NS)
        else None,
        line_spacing=style.line_spacing,
    )


def _extract_character_style(style) -> CharacterStyle:
    return CharacterStyle(
        style_id=style.style_id,
        text_color=style.text_color,
        height=style.height,
    )


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


def _build_equation_wrapper(section, node):
    from .elements import EquationXml

    return EquationXml(section.document, section, node)


def _build_shape_wrapper(section, node):
    from .elements import ShapeXml

    return ShapeXml(section.document, section, node)


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
    blocks_written = 0
    first_paragraph_consumed = False
    for block in section.blocks:
        if isinstance(block, Paragraph):
            if not first_paragraph_consumed and blocks_written == 0:
                document.set_paragraph_text(section_index, 0, block.text)
                first_paragraph_consumed = True
            else:
                document.append_paragraph(block.text, section_index=section_index)
            blocks_written += 1
            continue

        paragraph_index = 0 if blocks_written == 0 else _append_control_host_paragraph(document, section_index)

        if isinstance(block, Table):
            document.append_table(
                block.rows,
                block.cols,
                section_index=section_index,
                paragraph_index=paragraph_index,
                cell_texts=block.cell_texts,
            )
        elif isinstance(block, Picture):
            document.append_picture(
                block.name,
                block.data,
                section_index=section_index,
                paragraph_index=paragraph_index,
                width=block.width,
                height=block.height,
                media_type=_media_type_for_extension(block.extension),
            )
        elif isinstance(block, Hyperlink):
            document.append_hyperlink(
                block.target,
                display_text=block.display_text,
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
                field_type=block.field_type,
                display_text=block.display_text,
                name=block.name,
                parameters=block.parameters,
                editable=block.editable,
                dirty=block.dirty,
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
            document.append_note(
                block.text,
                kind=block.kind,
                number=block.number,
                section_index=section_index,
                paragraph_index=paragraph_index,
            )
        elif isinstance(block, Equation):
            document.append_equation(
                block.script,
                section_index=section_index,
                paragraph_index=paragraph_index,
                width=block.width,
                height=block.height,
            )
        elif isinstance(block, Shape):
            document.append_shape(
                kind=block.kind,
                text=block.text,
                width=block.width,
                height=block.height,
                fill_color=block.fill_color,
                line_color=block.line_color,
                section_index=section_index,
                paragraph_index=paragraph_index,
            )
        elif isinstance(block, Ole):
            document.append_ole(
                block.name,
                block.data,
                width=block.width,
                height=block.height,
                section_index=section_index,
                paragraph_index=paragraph_index,
            )
        blocks_written += 1


def _append_control_host_paragraph(document: HwpxDocument, section_index: int) -> int:
    document.append_paragraph("", section_index=section_index)
    return len(document.sections[section_index].paragraphs()) - 1


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


def _apply_section_settings(document: HwpxDocument, section_index: int, settings: SectionSettings) -> None:
    target = document.section_settings(section_index)
    if settings.page_width is not None or settings.page_height is not None or settings.landscape is not None:
        target.set_page_size(
            width=settings.page_width,
            height=settings.page_height,
            landscape=settings.landscape,
        )
    if settings.margins:
        target.set_margins(**settings.margins)
    if settings.visibility:
        target.set_visibility(
            hide_first_header=_coerce_bool(settings.visibility.get("hideFirstHeader")),
            hide_first_footer=_coerce_bool(settings.visibility.get("hideFirstFooter")),
            hide_first_master_page=_coerce_bool(settings.visibility.get("hideFirstMasterPage")),
            border=settings.visibility.get("border"),
            fill=settings.visibility.get("fill"),
            hide_first_page_num=_coerce_bool(settings.visibility.get("hideFirstPageNum")),
            hide_first_empty_line=_coerce_bool(settings.visibility.get("hideFirstEmptyLine")),
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


def _apply_header_footer_blocks(document: HwpxDocument, section_index: int, blocks: list[HeaderFooter]) -> None:
    for block in blocks:
        if block.kind == "header":
            document.append_header(
                block.text,
                apply_page_type=block.apply_page_type,
                section_index=section_index,
            )
        elif block.kind == "footer":
            document.append_footer(
                block.text,
                apply_page_type=block.apply_page_type,
                section_index=section_index,
            )


def _apply_styles(document: HwpxDocument, hancom_document: HancomDocument) -> None:
    for style in hancom_document.paragraph_styles:
        appended = document.append_paragraph_style(
            style_id=style.style_id,
            alignment_horizontal=style.alignment_horizontal,
            alignment_vertical=style.alignment_vertical,
            line_spacing=style.line_spacing,
        )
        if style.style_id is None:
            style.style_id = appended.style_id

    for style in hancom_document.character_styles:
        appended = document.append_character_style(
            style_id=style.style_id,
            text_color=style.text_color,
            height=style.height,
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
