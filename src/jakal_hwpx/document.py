from __future__ import annotations

import io
import os
import re
import shutil
import tempfile
import time
import zipfile
from collections import Counter
from dataclasses import replace
from pathlib import Path, PurePosixPath

from lxml import etree

from .elements import (
    _missing_preserved_tokens,
    _preserved_structure_signature,
    AutoNumber,
    Bookmark,
    CharacterStyle,
    Equation,
    Field,
    HeaderFooterBlock,
    Note,
    ParagraphStyle,
    Picture,
    SectionSettings,
    ShapeObject,
    StyleDefinition,
    Table,
)
from .exceptions import HwpxValidationError, InvalidHwpxFileError, ValidationIssue
from .namespaces import NS, SECTION_PATTERN, qname
from .parts import (
    BinaryDataPart,
    ContainerPart,
    ContentHpfPart,
    DocumentMetadata,
    HeaderPart,
    HwpxPart,
    MimetypePart,
    PreviewTextPart,
    SectionPart,
    SettingsPart,
    VersionPart,
    XmlPart,
    infer_part_class,
)


def _normalize_path(path: str) -> str:
    return PurePosixPath(path).as_posix()


def _issue(
    kind: str,
    message: str,
    *,
    part_path: str | None = None,
    section_index: int | None = None,
    paragraph_index: int | None = None,
    cell_row: int | None = None,
    cell_column: int | None = None,
    xpath: str | None = None,
    line: int | None = None,
    column: int | None = None,
    context: str | None = None,
) -> ValidationIssue:
    return ValidationIssue(
        kind=kind,
        message=message,
        part_path=part_path,
        section_index=section_index,
        paragraph_index=paragraph_index,
        cell_row=cell_row,
        cell_column=cell_column,
        xpath=xpath,
        line=line,
        column=column,
        context=context,
    )


def _clone_zip_info(source: zipfile.ZipInfo, filename: str) -> zipfile.ZipInfo:
    clone = zipfile.ZipInfo(filename=filename, date_time=source.date_time)
    clone.comment = source.comment
    clone.extra = source.extra
    clone.create_system = source.create_system
    clone.create_version = source.create_version
    clone.extract_version = source.extract_version
    clone.reserved = source.reserved
    clone.flag_bits = source.flag_bits
    clone.volume = source.volume
    clone.internal_attr = source.internal_attr
    clone.external_attr = source.external_attr
    clone.compress_type = source.compress_type
    return clone


def _full_nsmap() -> dict[str, str]:
    return dict(NS)


def _build_blank_version_part() -> VersionPart:
    root = etree.Element(qname("hv", "HCFVersion"), nsmap={"hv": NS["hv"]})
    root.set("tagetApplication", "WORDPROCESSOR")
    root.set("major", "5")
    root.set("minor", "1")
    root.set("micro", "1")
    root.set("buildNumber", "0")
    root.set("os", "1")
    root.set("xmlVersion", "1.5")
    root.set("application", "Hancom Office Hangul")
    root.set("appVersion", "12, 0, 0, 535 WIN32LEWindows_10")
    return VersionPart.from_root("version.xml", root)


def _build_blank_container_part() -> ContainerPart:
    root = etree.Element(
        qname("ocf", "container"),
        nsmap={"ocf": NS["ocf"], "hpf": NS["hpf"]},
    )
    rootfiles = etree.SubElement(root, qname("ocf", "rootfiles"))
    rootfile = etree.SubElement(rootfiles, qname("ocf", "rootfile"))
    rootfile.set("full-path", "Contents/content.hpf")
    rootfile.set("media-type", "application/hwpml-package+xml")
    return ContainerPart.from_root("META-INF/container.xml", root)


def _build_blank_content_hpf_part() -> ContentHpfPart:
    root = etree.Element(qname("opf", "package"), nsmap=_full_nsmap())
    root.set("version", "")
    root.set("unique-identifier", "")
    root.set("id", "")

    metadata = etree.SubElement(root, qname("opf", "metadata"))
    etree.SubElement(metadata, qname("opf", "title"))
    language = etree.SubElement(metadata, qname("opf", "language"))
    language.text = "ko"

    manifest = etree.SubElement(root, qname("opf", "manifest"))
    for item_id, href in (
        ("header", "Contents/header.xml"),
        ("section0", "Contents/section0.xml"),
        ("settings", "settings.xml"),
    ):
        item = etree.SubElement(manifest, qname("opf", "item"))
        item.set("id", item_id)
        item.set("href", href)
        item.set("media-type", "application/xml")

    spine = etree.SubElement(root, qname("opf", "spine"))
    for item_id in ("header", "section0"):
        itemref = etree.SubElement(spine, qname("opf", "itemref"))
        itemref.set("idref", item_id)
        itemref.set("linear", "yes")
    return ContentHpfPart.from_root("Contents/content.hpf", root)


def _build_blank_header_part() -> HeaderPart:
    root = etree.Element(qname("hh", "head"), nsmap=_full_nsmap())
    root.set("version", "1.5")
    root.set("secCnt", "1")

    begin_num = etree.SubElement(root, qname("hh", "beginNum"))
    begin_num.set("page", "1")
    begin_num.set("footnote", "1")
    begin_num.set("endnote", "1")
    begin_num.set("pic", "1")
    begin_num.set("tbl", "1")
    begin_num.set("equation", "1")

    ref_list = etree.SubElement(root, qname("hh", "refList"))

    char_properties = etree.SubElement(ref_list, qname("hh", "charProperties"))
    char_properties.set("itemCnt", "1")
    char_pr = etree.SubElement(char_properties, qname("hh", "charPr"))
    char_pr.set("id", "0")
    char_pr.set("height", "1000")
    char_pr.set("textColor", "#000000")
    char_pr.set("shadeColor", "none")
    char_pr.set("useFontSpace", "0")
    char_pr.set("useKerning", "0")
    char_pr.set("symMark", "NONE")
    font_ref = etree.SubElement(char_pr, qname("hh", "fontRef"))
    for key in ("hangul", "latin", "hanja", "japanese", "other", "symbol", "user"):
        font_ref.set(key, "0")
    ratio = etree.SubElement(char_pr, qname("hh", "ratio"))
    spacing = etree.SubElement(char_pr, qname("hh", "spacing"))
    rel_sz = etree.SubElement(char_pr, qname("hh", "relSz"))
    offset = etree.SubElement(char_pr, qname("hh", "offset"))
    for node in (ratio, rel_sz):
        for key in ("hangul", "latin", "hanja", "japanese", "other", "symbol", "user"):
            node.set(key, "100")
    for node in (spacing, offset):
        for key in ("hangul", "latin", "hanja", "japanese", "other", "symbol", "user"):
            node.set(key, "0")
    underline = etree.SubElement(char_pr, qname("hh", "underline"))
    underline.set("type", "NONE")
    underline.set("shape", "SOLID")
    underline.set("color", "#000000")
    strikeout = etree.SubElement(char_pr, qname("hh", "strikeout"))
    strikeout.set("shape", "NONE")
    strikeout.set("color", "#000000")
    outline = etree.SubElement(char_pr, qname("hh", "outline"))
    outline.set("type", "NONE")
    shadow = etree.SubElement(char_pr, qname("hh", "shadow"))
    shadow.set("type", "NONE")
    shadow.set("color", "#C0C0C0")
    shadow.set("offsetX", "10")
    shadow.set("offsetY", "10")

    para_properties = etree.SubElement(ref_list, qname("hh", "paraProperties"))
    para_properties.set("itemCnt", "1")
    para_pr = etree.SubElement(para_properties, qname("hh", "paraPr"))
    para_pr.set("id", "0")
    para_pr.set("tabPrIDRef", "0")
    para_pr.set("condense", "0")
    para_pr.set("fontLineHeight", "0")
    para_pr.set("snapToGrid", "1")
    para_pr.set("suppressLineNumbers", "0")
    para_pr.set("checked", "0")
    align = etree.SubElement(para_pr, qname("hh", "align"))
    align.set("horizontal", "JUSTIFY")
    align.set("vertical", "BASELINE")
    heading = etree.SubElement(para_pr, qname("hh", "heading"))
    heading.set("type", "NONE")
    heading.set("idRef", "0")
    heading.set("level", "0")
    break_setting = etree.SubElement(para_pr, qname("hh", "breakSetting"))
    break_setting.set("breakLatinWord", "KEEP_WORD")
    break_setting.set("breakNonLatinWord", "KEEP_WORD")
    break_setting.set("widowOrphan", "0")
    break_setting.set("keepWithNext", "0")
    break_setting.set("keepLines", "0")
    break_setting.set("pageBreakBefore", "0")
    break_setting.set("lineWrap", "BREAK")
    auto_spacing = etree.SubElement(para_pr, qname("hh", "autoSpacing"))
    auto_spacing.set("eAsianEng", "0")
    auto_spacing.set("eAsianNum", "0")
    margin = etree.SubElement(para_pr, qname("hh", "margin"))
    for key in ("intent", "left", "right", "prev", "next"):
        node = etree.SubElement(margin, qname("hc", key))
        node.set("value", "0")
        node.set("unit", "HWPUNIT")
    line_spacing = etree.SubElement(para_pr, qname("hh", "lineSpacing"))
    line_spacing.set("type", "PERCENT")
    line_spacing.set("value", "160")
    line_spacing.set("unit", "HWPUNIT")

    styles = etree.SubElement(ref_list, qname("hh", "styles"))
    styles.set("itemCnt", "1")
    style = etree.SubElement(styles, qname("hh", "style"))
    style.set("id", "0")
    style.set("type", "PARA")
    style.set("name", "Normal")
    style.set("engName", "Normal")
    style.set("paraPrIDRef", "0")
    style.set("charPrIDRef", "0")
    style.set("nextStyleIDRef", "0")
    style.set("langID", "1042")
    style.set("lockForm", "0")
    return HeaderPart.from_root("Contents/header.xml", root)


def _build_blank_settings_part() -> SettingsPart:
    root = etree.Element(
        qname("ha", "HWPApplicationSetting"),
        nsmap={"ha": NS["ha"], "config": NS["config"]},
    )
    caret = etree.SubElement(root, qname("ha", "CaretPosition"))
    caret.set("listIDRef", "0")
    caret.set("paraIDRef", "0")
    caret.set("pos", "0")
    return SettingsPart.from_root("settings.xml", root)


def _build_blank_section_part() -> SectionPart:
    root = etree.Element(qname("hs", "sec"), nsmap=_full_nsmap())
    paragraph = etree.SubElement(root, qname("hp", "p"))
    paragraph.set("id", "0")
    paragraph.set("paraPrIDRef", "0")
    paragraph.set("styleIDRef", "0")
    paragraph.set("pageBreak", "0")
    paragraph.set("columnBreak", "0")
    paragraph.set("merged", "0")

    sec_run = etree.SubElement(paragraph, qname("hp", "run"))
    sec_run.set("charPrIDRef", "0")
    sec_pr = etree.SubElement(sec_run, qname("hp", "secPr"))
    sec_pr.set("id", "")
    sec_pr.set("textDirection", "HORIZONTAL")
    sec_pr.set("spaceColumns", "1134")
    sec_pr.set("tabStop", "8000")
    sec_pr.set("tabStopVal", "4000")
    sec_pr.set("tabStopUnit", "HWPUNIT")
    sec_pr.set("outlineShapeIDRef", "0")
    sec_pr.set("memoShapeIDRef", "0")
    sec_pr.set("textVerticalWidthHead", "0")
    sec_pr.set("masterPageCnt", "0")
    grid = etree.SubElement(sec_pr, qname("hp", "grid"))
    grid.set("lineGrid", "0")
    grid.set("charGrid", "0")
    grid.set("wonggojiFormat", "0")
    start_num = etree.SubElement(sec_pr, qname("hp", "startNum"))
    start_num.set("pageStartsOn", "BOTH")
    start_num.set("page", "0")
    start_num.set("pic", "0")
    start_num.set("tbl", "0")
    start_num.set("equation", "0")
    visibility = etree.SubElement(sec_pr, qname("hp", "visibility"))
    visibility.set("hideFirstHeader", "0")
    visibility.set("hideFirstFooter", "0")
    visibility.set("hideFirstMasterPage", "0")
    visibility.set("border", "SHOW_ALL")
    visibility.set("fill", "SHOW_ALL")
    visibility.set("hideFirstPageNum", "0")
    visibility.set("hideFirstEmptyLine", "0")
    visibility.set("showLineNumber", "0")
    line_number_shape = etree.SubElement(sec_pr, qname("hp", "lineNumberShape"))
    line_number_shape.set("restartType", "0")
    line_number_shape.set("countBy", "0")
    line_number_shape.set("distance", "0")
    line_number_shape.set("startNumber", "0")
    page_pr = etree.SubElement(sec_pr, qname("hp", "pagePr"))
    page_pr.set("landscape", "WIDELY")
    page_pr.set("width", "59528")
    page_pr.set("height", "84186")
    page_pr.set("gutterType", "LEFT_ONLY")
    page_margin = etree.SubElement(page_pr, qname("hp", "margin"))
    page_margin.set("header", "4252")
    page_margin.set("footer", "4252")
    page_margin.set("gutter", "0")
    page_margin.set("left", "8504")
    page_margin.set("right", "8504")
    page_margin.set("top", "5668")
    page_margin.set("bottom", "4252")

    text_run = etree.SubElement(paragraph, qname("hp", "run"))
    text_run.set("charPrIDRef", "0")
    etree.SubElement(text_run, qname("hp", "t"))
    return SectionPart.from_root("Contents/section0.xml", root)


def _build_blank_document_state() -> tuple[dict[str, HwpxPart], list[str], dict[str, zipfile.ZipInfo], list[str]]:
    parts: dict[str, HwpxPart] = {
        "mimetype": MimetypePart("mimetype", b"application/hwp+zip"),
        "version.xml": _build_blank_version_part(),
        "Contents/content.hpf": _build_blank_content_hpf_part(),
        "Contents/header.xml": _build_blank_header_part(),
        "Contents/section0.xml": _build_blank_section_part(),
        "settings.xml": _build_blank_settings_part(),
        "META-INF/container.xml": _build_blank_container_part(),
    }
    part_order = [
        "mimetype",
        "version.xml",
        "Contents/content.hpf",
        "Contents/header.xml",
        "Contents/section0.xml",
        "settings.xml",
        "META-INF/container.xml",
    ]
    return parts, part_order, {}, []


class HwpxDocument:
    def __init__(
        self,
        *,
        parts: dict[str, HwpxPart] | None = None,
        part_order: list[str] | None = None,
        zip_infos: dict[str, zipfile.ZipInfo] | None = None,
        source_path: Path | None = None,
        duplicate_entries: list[str] | None = None,
    ):
        if parts is None and part_order is None and zip_infos is None:
            parts, part_order, zip_infos, duplicate_entries = _build_blank_document_state()
        elif parts is None or part_order is None or zip_infos is None:
            raise TypeError("parts, part_order, and zip_infos must be provided together.")

        self._parts = parts
        self._part_order = part_order
        self._zip_infos = zip_infos
        self.source_path = source_path
        self.duplicate_entries = duplicate_entries or []
        self._control_preservation_expectations: dict[int, tuple[etree._Element, Counter[str], ValidationIssue]] = {}
        self._control_preservation_errors: list[ValidationIssue] = []
        for part in self._parts.values():
            part.document = self

    @classmethod
    def blank(cls) -> "HwpxDocument":
        return cls()

    @classmethod
    def open(cls, path: str | os.PathLike[str]) -> "HwpxDocument":
        input_path = Path(path)
        if not input_path.exists():
            raise FileNotFoundError(input_path)
        if not zipfile.is_zipfile(input_path):
            raise InvalidHwpxFileError(f"{input_path} is not a valid zip-based HWPX package.")
        with zipfile.ZipFile(input_path) as archive:
            return cls._from_zipfile(archive=archive, source_path=input_path)

    @classmethod
    def from_bytes(cls, raw_bytes: bytes) -> "HwpxDocument":
        with zipfile.ZipFile(io.BytesIO(raw_bytes)) as archive:
            return cls._from_zipfile(archive=archive, source_path=None)

    @classmethod
    def _from_zipfile(cls, archive: zipfile.ZipFile, source_path: Path | None) -> "HwpxDocument":
        parts, part_order, zip_infos, duplicates = cls._read_zipfile_state(archive)
        document = cls(
            parts=parts,
            part_order=part_order,
            zip_infos=zip_infos,
            source_path=source_path,
            duplicate_entries=duplicates,
        )
        document.validate()
        return document

    @staticmethod
    def _read_zipfile_state(
        archive: zipfile.ZipFile,
    ) -> tuple[dict[str, HwpxPart], list[str], dict[str, zipfile.ZipInfo], list[str]]:
        parts: dict[str, HwpxPart] = {}
        part_order: list[str] = []
        zip_infos: dict[str, zipfile.ZipInfo] = {}
        duplicates: list[str] = []

        for info in archive.infolist():
            normalized = _normalize_path(info.filename)
            raw_bytes = archive.read(info.filename)
            part_cls = infer_part_class(normalized, raw_bytes)
            part = part_cls(normalized, raw_bytes)
            if normalized in parts:
                duplicates.append(normalized)
            parts[normalized] = part
            zip_infos[normalized] = info
            if normalized not in part_order:
                part_order.append(normalized)
        return parts, part_order, zip_infos, duplicates

    @property
    def mimetype(self) -> MimetypePart:
        return self.get_part("mimetype", MimetypePart)

    @property
    def content_hpf(self) -> ContentHpfPart:
        return self.get_part("Contents/content.hpf", ContentHpfPart)

    @property
    def header(self) -> HeaderPart:
        return self.get_part("Contents/header.xml", HeaderPart)

    @property
    def container(self) -> ContainerPart:
        return self.get_part("META-INF/container.xml", ContainerPart)

    @property
    def preview_text(self) -> PreviewTextPart | None:
        part = self._parts.get("Preview/PrvText.txt")
        return part if isinstance(part, PreviewTextPart) else None

    @property
    def is_distribution_protected(self) -> bool:
        return self.content_hpf.is_distribution_protected

    @property
    def sections(self) -> list[SectionPart]:
        values = [
            part
            for path, part in self._parts.items()
            if re.match(SECTION_PATTERN, path) and isinstance(part, SectionPart)
        ]
        return sorted(values, key=lambda part: part.section_index())

    def get_part(self, path: str, expected_type: type[HwpxPart] | None = None) -> HwpxPart:
        normalized = _normalize_path(path)
        if normalized not in self._parts:
            raise KeyError(normalized)
        part = self._parts[normalized]
        if expected_type is not None and not isinstance(part, expected_type):
            raise TypeError(f"{normalized} is not a {expected_type.__name__}.")
        return part

    def list_part_paths(self) -> list[str]:
        seen = set()
        ordered = []
        for path in self._part_order:
            if path in self._parts and path not in seen:
                ordered.append(path)
                seen.add(path)
        for path in self._parts:
            if path not in seen:
                ordered.append(path)
        return ordered

    def add_part(self, path: str, raw_bytes: bytes) -> HwpxPart:
        normalized = _normalize_path(path)
        part_cls = infer_part_class(normalized, raw_bytes)
        part = part_cls(normalized, raw_bytes)
        part.document = self
        self._parts[normalized] = part
        if normalized not in self._part_order:
            if normalized == "mimetype":
                self._part_order.insert(0, normalized)
            else:
                self._part_order.append(normalized)
        return part

    def remove_part(self, path: str) -> None:
        normalized = _normalize_path(path)
        if normalized in self._parts:
            del self._parts[normalized]
        if normalized in self._part_order:
            self._part_order.remove(normalized)
        self._zip_infos.pop(normalized, None)

    def metadata(self) -> DocumentMetadata:
        return self.content_hpf.metadata()

    def set_metadata(self, **values: str | None) -> None:
        self.content_hpf.set_metadata(**values)

    def headers(self, section_index: int | None = None) -> list[HeaderFooterBlock]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        blocks: list[HeaderFooterBlock] = []
        for section in sections:
            for node in section.root_element.xpath(".//hp:header", namespaces=NS):
                blocks.append(HeaderFooterBlock(self, section, node))
        return blocks

    def footers(self, section_index: int | None = None) -> list[HeaderFooterBlock]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        blocks: list[HeaderFooterBlock] = []
        for section in sections:
            for node in section.root_element.xpath(".//hp:footer", namespaces=NS):
                blocks.append(HeaderFooterBlock(self, section, node))
        return blocks

    def tables(self, section_index: int | None = None) -> list[Table]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        tables: list[Table] = []
        for section in sections:
            for node in section.root_element.xpath(".//hp:tbl", namespaces=NS):
                tables.append(Table(self, section, node))
        return tables

    def pictures(self, section_index: int | None = None) -> list[Picture]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        pictures: list[Picture] = []
        for section in sections:
            for node in section.root_element.xpath(".//hp:pic", namespaces=NS):
                pictures.append(Picture(self, section, node))
        return pictures

    def section_settings(self, section_index: int = 0) -> SectionSettings:
        self._ensure_editable_sections()
        nodes = self.sections[section_index].root_element.xpath("./hp:p[1]//hp:secPr[1]", namespaces=NS)
        if not nodes:
            raise ValueError("Section does not contain hp:secPr.")
        return SectionSettings(self.sections[section_index], nodes[0])

    def notes(self, section_index: int | None = None) -> list[Note]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        notes: list[Note] = []
        for section in sections:
            for tag in (".//hp:footNote", ".//hp:endNote"):
                for node in section.root_element.xpath(tag, namespaces=NS):
                    notes.append(Note(self, section, node))
        return notes

    def bookmarks(self, section_index: int | None = None) -> list[Bookmark]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        bookmarks: list[Bookmark] = []
        for section in sections:
            for node in section.root_element.xpath(".//hp:bookmark", namespaces=NS):
                bookmarks.append(Bookmark(self, section, node))
        return bookmarks

    def fields(self, section_index: int | None = None) -> list[Field]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        fields: list[Field] = []
        for section in sections:
            for node in section.root_element.xpath(".//hp:fieldBegin", namespaces=NS):
                fields.append(Field(self, section, node))
        return fields

    def hyperlinks(self, section_index: int | None = None) -> list[Field]:
        return [field for field in self.fields(section_index=section_index) if field.is_hyperlink]

    def mail_merge_fields(self, section_index: int | None = None) -> list[Field]:
        return [field for field in self.fields(section_index=section_index) if field.is_mail_merge]

    def calculation_fields(self, section_index: int | None = None) -> list[Field]:
        return [field for field in self.fields(section_index=section_index) if field.is_calculation]

    def cross_references(self, section_index: int | None = None) -> list[Field]:
        return [field for field in self.fields(section_index=section_index) if field.is_cross_reference]

    def auto_numbers(self, section_index: int | None = None) -> list[AutoNumber]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        values: list[AutoNumber] = []
        for section in sections:
            for tag in (".//hp:autoNum", ".//hp:newNum"):
                for node in section.root_element.xpath(tag, namespaces=NS):
                    values.append(AutoNumber(self, section, node))
        return values

    def equations(self, section_index: int | None = None) -> list[Equation]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        values: list[Equation] = []
        for section in sections:
            for node in section.root_element.xpath(".//hp:equation", namespaces=NS):
                values.append(Equation(self, section, node))
        return values

    def shapes(self, section_index: int | None = None) -> list[ShapeObject]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        shape_tags = [
            "hp:rect",
            "hp:line",
            "hp:ellipse",
            "hp:arc",
            "hp:polygon",
            "hp:curve",
            "hp:connectLine",
            "hp:textart",
            "hp:container",
            "hp:ole",
        ]
        values: list[ShapeObject] = []
        for section in sections:
            for tag in shape_tags:
                for node in section.root_element.xpath(f".//{tag}", namespaces=NS):
                    values.append(ShapeObject(self, section, node))
        return values

    def styles(self) -> list[StyleDefinition]:
        return [StyleDefinition(self.header, node) for node in self.header.root_element.xpath(".//hh:style", namespaces=NS)]

    def paragraph_styles(self) -> list[ParagraphStyle]:
        return [ParagraphStyle(self.header, node) for node in self.header.root_element.xpath(".//hh:paraPr", namespaces=NS)]

    def character_styles(self) -> list[CharacterStyle]:
        return [CharacterStyle(self.header, node) for node in self.header.root_element.xpath(".//hh:charPr", namespaces=NS)]

    def get_style(self, style_id: str) -> StyleDefinition:
        for style in self.styles():
            if style.style_id == style_id:
                return style
        raise KeyError(style_id)

    def get_paragraph_style(self, style_id: str) -> ParagraphStyle:
        for style in self.paragraph_styles():
            if style.style_id == style_id:
                return style
        raise KeyError(style_id)

    def get_character_style(self, style_id: str) -> CharacterStyle:
        for style in self.character_styles():
            if style.style_id == style_id:
                return style
        raise KeyError(style_id)

    def get_document_text(self, section_separator: str = "\n\n") -> str:
        return section_separator.join(section.extract_text() for section in self.sections)

    def _track_control_preservation_targets(
        self,
        paragraphs: list[etree._Element],
        signatures: list[Counter[str]],
        *,
        issues: list[ValidationIssue],
    ) -> None:
        for paragraph, signature, issue in zip(paragraphs, signatures, issues):
            if not signature:
                continue
            self._control_preservation_expectations[id(paragraph)] = (paragraph, signature, issue)

    def _record_control_preservation_errors(self, errors: list[ValidationIssue]) -> None:
        for error in errors:
            if error not in self._control_preservation_errors:
                self._control_preservation_errors.append(error)

    def control_preservation_validation_errors(self) -> list[ValidationIssue]:
        errors = list(self._control_preservation_errors)
        for marker, (paragraph, expected_signature, issue) in self._control_preservation_expectations.items():
            if paragraph.getparent() is None:
                missing = list(expected_signature.elements())
            else:
                missing = _missing_preserved_tokens(expected_signature, _preserved_structure_signature(paragraph))
            if missing:
                errors.append(replace(issue, message=f"lost preserved nodes after edit: {', '.join(missing)}"))
        return errors

    def replace_text(self, old: str, new: str, count: int = -1, include_header: bool = True) -> int:
        targets: list[XmlPart] = list(self.sections)
        if include_header:
            targets.insert(0, self.header)
        remaining = count
        replacements = 0
        for part in targets:
            limit = remaining if remaining >= 0 else -1
            changed = part.replace_hp_text(old, new, count=limit)
            replacements += changed
            if remaining >= 0:
                remaining -= changed
                if remaining <= 0:
                    break
        return replacements

    def append_paragraph(
        self,
        text: str,
        *,
        section_index: int = 0,
        template_index: int | None = None,
        para_pr_id: str | None = None,
        style_id: str | None = None,
        char_pr_id: str | None = None,
    ):
        self._ensure_editable_sections()
        return self.sections[section_index].append_paragraph(
            text,
            template_index=template_index,
            para_pr_id=para_pr_id,
            style_id=style_id,
            char_pr_id=char_pr_id,
        )

    def insert_paragraph(
        self,
        section_index: int,
        paragraph_index: int,
        text: str,
        *,
        template_index: int | None = None,
        para_pr_id: str | None = None,
        style_id: str | None = None,
        char_pr_id: str | None = None,
    ):
        self._ensure_editable_sections()
        return self.sections[section_index].insert_paragraph(
            paragraph_index,
            text,
            template_index=template_index,
            para_pr_id=para_pr_id,
            style_id=style_id,
            char_pr_id=char_pr_id,
        )

    def set_paragraph_text(
        self,
        section_index: int,
        paragraph_index: int,
        text: str,
        *,
        char_pr_id: str | None = None,
    ):
        self._ensure_editable_sections()
        return self.sections[section_index].set_paragraph_text(paragraph_index, text, char_pr_id=char_pr_id)

    def delete_paragraph(self, section_index: int, paragraph_index: int) -> None:
        self._ensure_editable_sections()
        self.sections[section_index].delete_paragraph(paragraph_index)

    def apply_style_to_paragraph(
        self,
        section_index: int,
        paragraph_index: int,
        *,
        style_id: str | None = None,
        para_pr_id: str | None = None,
        char_pr_id: str | None = None,
    ):
        self._ensure_editable_sections()
        paragraph = self.sections[section_index].root_element.xpath("./hp:p", namespaces=NS)[paragraph_index]
        if style_id is not None:
            paragraph.set("styleIDRef", style_id)
        if para_pr_id is not None:
            paragraph.set("paraPrIDRef", para_pr_id)
        if char_pr_id is not None:
            for run in paragraph.xpath("./hp:run", namespaces=NS):
                run.set("charPrIDRef", char_pr_id)
        self.sections[section_index].mark_modified()

    def apply_style_batch(
        self,
        *,
        section_index: int | None = None,
        text_contains: str | None = None,
        regex: str | None = None,
        style_id: str | None = None,
        para_pr_id: str | None = None,
        char_pr_id: str | None = None,
    ) -> int:
        self._ensure_editable_sections()
        pattern = re.compile(regex) if regex else None
        sections = self.sections if section_index is None else [self.sections[section_index]]
        updated = 0
        for section in sections:
            paragraphs = section.root_element.xpath("./hp:p", namespaces=NS)
            for index, paragraph in enumerate(paragraphs):
                text = "".join(node.text or "" for node in paragraph.xpath(".//hp:t", namespaces=NS))
                if text_contains is not None and text_contains not in text:
                    continue
                if pattern is not None and not pattern.search(text):
                    continue
                if text_contains is None and pattern is None:
                    continue
                target_section_index = self.sections.index(section)
                self.apply_style_to_paragraph(
                    target_section_index,
                    index,
                    style_id=style_id,
                    para_pr_id=para_pr_id,
                    char_pr_id=char_pr_id,
                )
                updated += 1
        return updated

    def add_section(self, *, clone_from: int = 0, text: str | None = None) -> SectionPart:
        self._ensure_editable_sections()
        source = self.sections[clone_from]
        next_path = self.content_hpf.next_section_path()
        new_section = SectionPart.from_root(next_path, source.clone_root())
        new_section.document = self
        new_section.mark_modified()
        self._parts[next_path] = new_section
        if next_path not in self._part_order:
            self._part_order.append(next_path)

        item_id = self.content_hpf.next_manifest_id(PurePosixPath(next_path).stem)
        self.content_hpf.ensure_manifest_item(item_id, next_path, "application/xml")
        self.content_hpf.ensure_spine_itemref(item_id)
        self.header.set_section_count(len(self.sections))
        if text is not None:
            first_text_paragraph = next((index for index, value in enumerate(new_section.text_fragments()) if value), 0)
            new_section.set_paragraph_text(first_text_paragraph, text)
        return new_section

    def remove_section(self, section_index: int) -> None:
        self._ensure_editable_sections()
        sections = self.sections
        if len(sections) <= 1:
            raise ValueError("A document must keep at least one section.")
        target = sections[section_index]
        target_path = target.path
        item_id = None
        for item in self.content_hpf.manifest_items():
            if item.get("href") == target_path:
                item_id = item.get("id")
                break
        self.remove_part(target_path)
        self.content_hpf.remove_manifest_item(href=target_path)
        if item_id:
            self.content_hpf.remove_spine_itemref(item_id)
        self.header.set_section_count(len(self.sections))

    def set_preview_text(self, text: str) -> PreviewTextPart:
        part = self.preview_text
        if part is None:
            part = self.add_part("Preview/PrvText.txt", b"")
        assert isinstance(part, PreviewTextPart)
        part.text = text
        return part

    def add_or_replace_binary(
        self,
        name: str,
        data: bytes,
        *,
        media_type: str | None = None,
        manifest_id: str | None = None,
    ) -> BinaryDataPart:
        import mimetypes

        normalized_name = PurePosixPath(name).name
        part_path = f"BinData/{normalized_name}"
        if part_path in self._parts:
            part = self.get_part(part_path, BinaryDataPart)
            part.data = data
        else:
            part = self.add_part(part_path, data)
        assert isinstance(part, BinaryDataPart)

        if media_type is None:
            media_type = mimetypes.guess_type(normalized_name)[0] or "application/octet-stream"
        if manifest_id is None:
            manifest_id = self.content_hpf.next_manifest_id(PurePosixPath(normalized_name).stem or "bindata")
        self.content_hpf.ensure_manifest_item(
            manifest_id,
            part_path,
            media_type,
            isEmbeded="1",
        )
        return part

    def append_bookmark(
        self,
        name: str,
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> Bookmark:
        self._ensure_editable_sections()
        paragraph = self._resolve_paragraph_for_insert(section_index=section_index, paragraph_index=paragraph_index)
        run = self._append_run(paragraph, char_pr_id=char_pr_id)
        ctrl = etree.SubElement(run, qname("hp", "ctrl"))
        bookmark = etree.SubElement(ctrl, qname("hp", "bookmark"))
        bookmark.set("name", name)
        etree.SubElement(run, qname("hp", "t")).text = ""
        section = self.sections[section_index]
        section.mark_modified()
        return Bookmark(self, section, bookmark)

    def append_field(
        self,
        *,
        field_type: str,
        display_text: str | None = None,
        name: str | None = None,
        parameters: dict[str, int | str] | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
        editable: bool = False,
        dirty: bool = False,
    ) -> Field:
        self._ensure_editable_sections()
        paragraph = self._resolve_paragraph_for_insert(section_index=section_index, paragraph_index=paragraph_index)
        begin_id = self._next_control_number(".//hp:fieldBegin/@id")
        field_id = self._next_control_number(".//hp:fieldBegin/@fieldid | .//hp:fieldEnd/@fieldid")

        begin_run = self._append_run(paragraph, char_pr_id=char_pr_id)
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
                parameter.set("name", key)
                parameter.text = str(value)

        if display_text is not None:
            display_run = self._append_run(paragraph, char_pr_id=char_pr_id)
            etree.SubElement(display_run, qname("hp", "t")).text = display_text

        end_run = self._append_run(paragraph, char_pr_id=char_pr_id)
        end_ctrl = etree.SubElement(end_run, qname("hp", "ctrl"))
        field_end = etree.SubElement(end_ctrl, qname("hp", "fieldEnd"))
        field_end.set("beginIDRef", begin_id)
        field_end.set("fieldid", field_id)

        section = self.sections[section_index]
        section.mark_modified()
        return Field(self, section, field_begin)

    def append_hyperlink(
        self,
        target: str,
        *,
        display_text: str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> Field:
        parameters: dict[str, int | str] = {"Command": target, "Path": target}
        if target.startswith("mailto:"):
            parameters["Category"] = "HWPHYPERLINK_TYPE_EMAIL"
        elif target.startswith(("http://", "https://")):
            parameters["Category"] = "HWPHYPERLINK_TYPE_WEB"
        else:
            parameters["Category"] = "HWPHYPERLINK_TYPE_PATH"
        return self.append_field(
            field_type="HYPERLINK",
            display_text=display_text,
            parameters=parameters,
            section_index=section_index,
            paragraph_index=paragraph_index,
            char_pr_id=char_pr_id,
        )

    def append_mail_merge_field(
        self,
        field_name: str,
        *,
        display_text: str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> Field:
        field = self.append_field(
            field_type="MAILMERGE",
            display_text=display_text,
            name=field_name,
            parameters={"FieldName": field_name, "MergeField": field_name},
            section_index=section_index,
            paragraph_index=paragraph_index,
            char_pr_id=char_pr_id,
        )
        field.configure_mail_merge(field_name, display_text=display_text)
        return field

    def append_calculation_field(
        self,
        expression: str,
        *,
        display_text: str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> Field:
        field = self.append_field(
            field_type="FORMULA",
            display_text=display_text,
            parameters={"Expression": expression, "Command": expression},
            section_index=section_index,
            paragraph_index=paragraph_index,
            char_pr_id=char_pr_id,
        )
        field.configure_calculation(expression, display_text=display_text)
        return field

    def append_cross_reference(
        self,
        bookmark_name: str,
        *,
        display_text: str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> Field:
        field = self.append_field(
            field_type="CROSSREF",
            display_text=display_text,
            name=bookmark_name,
            parameters={"BookmarkName": bookmark_name, "Path": bookmark_name},
            section_index=section_index,
            paragraph_index=paragraph_index,
            char_pr_id=char_pr_id,
        )
        field.configure_cross_reference(bookmark_name, display_text=display_text)
        return field

    def compile(self, validate: bool = True) -> bytes:
        if validate:
            self.validate()

        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, mode="w") as archive:
            for path in self.list_part_paths():
                part = self._parts[path]
                data = part.to_bytes()
                source_info = self._zip_infos.get(path)
                if source_info is None:
                    info = zipfile.ZipInfo(filename=path)
                    info.compress_type = zipfile.ZIP_STORED if path == "mimetype" else zipfile.ZIP_DEFLATED
                else:
                    info = _clone_zip_info(source_info, path)
                    if path == "mimetype":
                        info.compress_type = zipfile.ZIP_STORED
                archive.writestr(info, data)
        return buffer.getvalue()

    def save(self, path: str | os.PathLike[str], validate: bool = True) -> Path:
        output_path = Path(path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_bytes(self.compile(validate=validate))
        return output_path

    def validation_errors(self) -> list[ValidationIssue]:
        errors: list[ValidationIssue] = []
        content_hpf = self._parts.get("Contents/content.hpf")
        is_distribution_protected = isinstance(content_hpf, ContentHpfPart) and content_hpf.is_distribution_protected
        required = [
            "mimetype",
            "version.xml",
            "Contents/content.hpf",
            "META-INF/container.xml",
        ]
        if not is_distribution_protected:
            required.append("Contents/header.xml")
        for path in required:
            if path not in self._parts:
                errors.append(_issue("missing_part", f"Missing required part: {path}", part_path=path))

        if self.duplicate_entries:
            errors.append(
                _issue(
                    "duplicate_zip_entry",
                    f"Duplicate zip entries found: {', '.join(sorted(set(self.duplicate_entries)))}",
                )
            )

        mimetype = self._parts.get("mimetype")
        if isinstance(mimetype, MimetypePart) and mimetype.mime.strip() != "application/hwp+zip":
            errors.append(
                _issue("mimetype", "mimetype must be application/hwp+zip", part_path="mimetype")
            )

        xml_parts = [part for part in self._parts.values() if isinstance(part, XmlPart)]
        for part in xml_parts:
            try:
                part.to_bytes()
            except Exception as exc:  # noqa: BLE001
                errors.append(
                    _issue("xml_serialization", f"Invalid XML part serialization: {exc}", part_path=part.path)
                )

        container = self._parts.get("META-INF/container.xml")
        if isinstance(container, ContainerPart):
            for rootfile_path in container.rootfile_paths():
                normalized = _normalize_path(rootfile_path)
                if normalized not in self._parts:
                    if normalized == "Preview/PrvText.txt":
                        continue
                    errors.append(
                        _issue(
                            "container_reference",
                            f"container.xml references missing rootfile: {normalized}",
                            part_path="META-INF/container.xml",
                        )
                    )

        if isinstance(content_hpf, ContentHpfPart):
            manifest_ids: set[str] = set()
            for item in content_hpf.manifest_items():
                item_id = item.get("id")
                href = item.get("href")
                if not item_id:
                    errors.append(
                        _issue(
                            "manifest_item",
                            "content.hpf manifest item is missing id",
                            part_path="Contents/content.hpf",
                        )
                    )
                elif item_id in manifest_ids:
                    errors.append(
                        _issue(
                            "manifest_item",
                            f"Duplicate content.hpf manifest id: {item_id}",
                            part_path="Contents/content.hpf",
                        )
                    )
                else:
                    manifest_ids.add(item_id)
                if not href:
                    errors.append(
                        _issue(
                            "manifest_item",
                            "content.hpf manifest item is missing href",
                            part_path="Contents/content.hpf",
                        )
                    )
                    continue
                if _normalize_path(href) not in self._parts:
                    errors.append(
                        _issue(
                            "manifest_reference",
                            f"content.hpf manifest references missing part: {href}",
                            part_path="Contents/content.hpf",
                        )
                    )

            for itemref in content_hpf.spine_itemrefs():
                item_id = itemref.get("idref")
                if not item_id:
                    errors.append(
                        _issue(
                            "spine_itemref",
                            "content.hpf spine itemref is missing idref",
                            part_path="Contents/content.hpf",
                        )
                    )
                elif item_id not in manifest_ids:
                    errors.append(
                        _issue(
                            "spine_itemref",
                            f"content.hpf spine references unknown manifest id: {item_id}",
                            part_path="Contents/content.hpf",
                        )
                    )

        header = self._parts.get("Contents/header.xml")
        if isinstance(header, HeaderPart):
            if header.section_count and header.section_count != len(self.sections):
                errors.append(
                    _issue(
                        "header_section_count",
                        f"header.xml secCnt={header.section_count} does not match section count={len(self.sections)}",
                        part_path="Contents/header.xml",
                    )
                )

        if not is_distribution_protected and not self.sections:
            errors.append(_issue("document_structure", "Document must contain at least one section part"))

        errors.extend(self.control_preservation_validation_errors())
        return errors

    def validate(self) -> None:
        errors = self.validation_errors()
        if errors:
            raise HwpxValidationError(errors)

    def roundtrip_validate(self) -> None:
        temp_dir = Path(tempfile.mkdtemp(prefix="jakal_hwpx_"))
        try:
            temp_path = temp_dir / "roundtrip.hwpx"
            self.save(temp_path)
            reopened = HwpxDocument.open(temp_path)
            reopened.validate()
        finally:
            self._cleanup_temp_dir(temp_dir)

    def xml_validation_errors(self) -> list[ValidationIssue]:
        errors: list[ValidationIssue] = []
        expected_roots = {
            "version.xml": "HCFVersion",
            "Contents/content.hpf": "package",
            "Contents/header.xml": "head",
            "settings.xml": "HWPApplicationSetting",
            "META-INF/container.xml": "container",
            "META-INF/manifest.xml": "manifest",
            "META-INF/container.rdf": "RDF",
        }
        for path, expected_local_name in expected_roots.items():
            part = self._parts.get(path)
            if part is None:
                continue
            if not isinstance(part, XmlPart):
                errors.append(
                    _issue(
                        "xml_part_type",
                        f"{path} is expected to be XML but is loaded as {type(part).__name__}",
                        part_path=path,
                    )
                )
                continue
            actual = part.root.local_name
            if actual != expected_local_name:
                errors.append(
                    _issue(
                        "xml_root",
                        f"{path} root element must be {expected_local_name}, got {actual}",
                        part_path=path,
                    )
                )
        for section_index, section in enumerate(self.sections):
            if section.root.local_name != "sec":
                errors.append(
                    _issue(
                        "xml_root",
                        f"{section.path} root element must be sec, got {section.root.local_name}",
                        part_path=section.path,
                        section_index=section_index,
                    )
                )
            if not section.root_element.xpath("./hp:p", namespaces=NS):
                errors.append(
                    _issue(
                        "section_structure",
                        f"{section.path} must contain at least one paragraph",
                        part_path=section.path,
                        section_index=section_index,
                    )
                )
            for table in self.tables(section_index=section_index):
                if table.row_count < 1 or table.column_count < 1:
                    errors.append(
                        _issue(
                            "table_structure",
                            f"{section.path} contains a table with invalid row/column count",
                            part_path=section.path,
                            section_index=section_index,
                        )
                    )
                for cell in table.cells():
                    if cell.row_span < 1 or cell.col_span < 1:
                        errors.append(
                            _issue(
                                "table_cell_span",
                                f"{section.path} table cell has invalid span",
                                part_path=section.path,
                                section_index=section_index,
                                cell_row=cell.row,
                                cell_column=cell.column,
                            )
                        )
        return errors

    def schema_validation_errors(self, schema_map: dict[str, str | os.PathLike[str]]) -> list[ValidationIssue]:
        validators: dict[Path, etree.XMLSchema] = {}
        errors: list[ValidationIssue] = []
        for part_path, schema_path in schema_map.items():
            part = self._parts.get(_normalize_path(part_path))
            if part is None:
                errors.append(
                    _issue("schema_validation", f"Schema configured for missing part: {part_path}", part_path=part_path)
                )
                continue
            if not isinstance(part, XmlPart):
                errors.append(
                    _issue("schema_validation", f"Schema validation requires XML part: {part_path}", part_path=part_path)
                )
                continue
            resolved_schema_path = Path(schema_path)
            if resolved_schema_path not in validators:
                schema_doc = etree.parse(str(resolved_schema_path))
                validators[resolved_schema_path] = etree.XMLSchema(schema_doc)
            validator = validators[resolved_schema_path]
            if not validator.validate(part._root):
                for entry in validator.error_log:
                    errors.append(
                        _issue(
                            "schema_validation",
                            entry.message,
                            part_path=part_path,
                            line=entry.line,
                            column=entry.column,
                        )
                    )
        return errors

    def reference_validation_errors(self) -> list[ValidationIssue]:
        errors = self.validation_errors()
        style_ids = {style.style_id for style in self.styles() if style.style_id is not None} if not self.is_distribution_protected else set()
        para_pr_ids = {
            style.style_id for style in self.paragraph_styles() if style.style_id is not None
        } if not self.is_distribution_protected else set()
        char_pr_ids = {
            style.style_id for style in self.character_styles() if style.style_id is not None
        } if not self.is_distribution_protected else set()
        manifest_ids = {item.get("id") for item in self.content_hpf.manifest_items() if item.get("id")}
        bookmark_names: set[str] = set()
        field_begin_ids: dict[str, str] = {}
        field_id_map: dict[str, str] = {}

        for bookmark in self.bookmarks():
            if not bookmark.name:
                errors.append(_issue("bookmark", "bookmark is missing name"))
                continue
            if bookmark.name in bookmark_names:
                errors.append(_issue("bookmark", f"duplicate bookmark name: {bookmark.name}"))
            bookmark_names.add(bookmark.name)

        for field in self.fields():
            if field.control_id:
                if field.control_id in field_begin_ids:
                    errors.append(_issue("field_reference", f"duplicate fieldBegin id: {field.control_id}"))
                field_begin_ids[field.control_id] = field.field_id or ""
            if field.field_id:
                if field.field_id in field_id_map:
                    errors.append(_issue("field_reference", f"duplicate fieldid: {field.field_id}"))
                field_id_map[field.field_id] = field.control_id or ""

        for section_index, section in enumerate(self.sections):
            for paragraph_index, paragraph in enumerate(section.root_element.xpath("./hp:p", namespaces=NS)):
                style_id = paragraph.get("styleIDRef")
                para_pr_id = paragraph.get("paraPrIDRef")
                if style_id and style_ids and style_id not in style_ids:
                    errors.append(
                        _issue(
                            "style_reference",
                            f"paragraph references unknown styleIDRef={style_id}",
                            part_path=section.path,
                            section_index=section_index,
                            paragraph_index=paragraph_index,
                        )
                    )
                if para_pr_id and para_pr_ids and para_pr_id not in para_pr_ids:
                    errors.append(
                        _issue(
                            "style_reference",
                            f"paragraph references unknown paraPrIDRef={para_pr_id}",
                            part_path=section.path,
                            section_index=section_index,
                            paragraph_index=paragraph_index,
                        )
                    )
                for run in paragraph.xpath("./hp:run", namespaces=NS):
                    char_pr_id = run.get("charPrIDRef")
                    if char_pr_id and char_pr_ids and char_pr_id not in char_pr_ids:
                        errors.append(
                            _issue(
                                "style_reference",
                                f"run references unknown charPrIDRef={char_pr_id}",
                                part_path=section.path,
                                section_index=section_index,
                                paragraph_index=paragraph_index,
                            )
                        )
            for picture in self.pictures(section_index=section_index):
                if picture.binary_item_id and picture.binary_item_id not in manifest_ids:
                    errors.append(
                        _issue(
                            "binary_reference",
                            f"picture references unknown manifest id {picture.binary_item_id}",
                            part_path=section.path,
                            section_index=section_index,
                        )
                    )
            for shape in self.shapes(section_index=section_index):
                binary_refs = shape.element.xpath("./@binaryItemIDRef", namespaces=NS)
                for binary_ref in binary_refs:
                    if binary_ref not in manifest_ids:
                        errors.append(
                            _issue(
                                "binary_reference",
                                f"shape references unknown manifest id {binary_ref}",
                                part_path=section.path,
                                section_index=section_index,
                            )
                        )
            for field_end in section.root_element.xpath(".//hp:fieldEnd", namespaces=NS):
                begin_id = field_end.get("beginIDRef")
                field_id = field_end.get("fieldid")
                if not begin_id:
                    errors.append(
                        _issue(
                            "field_reference",
                            "fieldEnd is missing beginIDRef",
                            part_path=section.path,
                            section_index=section_index,
                            xpath=".//hp:fieldEnd",
                        )
                    )
                    continue
                if begin_id not in field_begin_ids:
                    errors.append(
                        _issue(
                            "field_reference",
                            f"fieldEnd references unknown fieldBegin id {begin_id}",
                            part_path=section.path,
                            section_index=section_index,
                            xpath=".//hp:fieldEnd",
                        )
                    )
                    continue
                expected_field_id = field_begin_ids[begin_id]
                if field_id and expected_field_id and field_id != expected_field_id:
                    errors.append(
                        _issue(
                            "field_reference",
                            f"fieldEnd fieldid={field_id} does not match fieldBegin fieldid={expected_field_id}",
                            part_path=section.path,
                            section_index=section_index,
                            xpath=".//hp:fieldEnd",
                        )
                    )
        for field in self.cross_references():
            target = field.get_parameter("BookmarkName") or field.get_parameter("Path") or field.name
            if target and target not in bookmark_names:
                errors.append(_issue("cross_reference", f"cross reference points to unknown bookmark: {target}"))
        return errors

    def save_reopen_validation_errors(self) -> list[ValidationIssue]:
        temp_dir = Path(tempfile.mkdtemp(prefix="jakal_hwpx_reopen_"))
        try:
            temp_path = temp_dir / "validation.hwpx"
            self.save(temp_path)
            reopened = HwpxDocument.open(temp_path)
            return reopened.validation_errors()
        except Exception as exc:  # noqa: BLE001
            if isinstance(exc, HwpxValidationError):
                return exc.errors
            return [_issue("save_reopen_validation", str(exc))]
        finally:
            self._cleanup_temp_dir(temp_dir)

    def _ensure_editable_sections(self) -> None:
        if self.sections:
            return
        if self.is_distribution_protected:
            raise RuntimeError(
                "This HWPX package is distribution protected. The encrypted section/header parts can be preserved, "
                "but high-level editing is not available."
            )
        raise RuntimeError("This HWPX package does not contain editable section XML parts.")

    def _resolve_paragraph_for_insert(self, *, section_index: int, paragraph_index: int | None) -> etree._Element:
        if paragraph_index is None:
            return self.append_paragraph("", section_index=section_index).element
        paragraphs = self.sections[section_index].root_element.xpath("./hp:p", namespaces=NS)
        return paragraphs[paragraph_index]

    def _append_run(self, paragraph: etree._Element, *, char_pr_id: str | None = None) -> etree._Element:
        resolved_char_pr_id = char_pr_id or self._default_char_pr_id(paragraph)
        run = etree.Element(qname("hp", "run"))
        run.set("charPrIDRef", resolved_char_pr_id)
        line_seg = paragraph.xpath("./hp:linesegarray[1]", namespaces=NS)
        if line_seg:
            line_seg[0].addprevious(run)
        else:
            paragraph.append(run)
        return run

    def _default_char_pr_id(self, paragraph: etree._Element) -> str:
        runs = paragraph.xpath("./hp:run", namespaces=NS)
        for run in runs:
            value = run.get("charPrIDRef")
            if value is not None:
                return value
        return "0"

    def _next_control_number(self, expression: str) -> str:
        highest = 0
        for section in self.sections:
            for value in section.root_element.xpath(expression, namespaces=NS):
                try:
                    highest = max(highest, int(value))
                except (TypeError, ValueError):
                    continue
        return str(highest + 1)

    @staticmethod
    def _cleanup_temp_dir(temp_dir: Path, retries: int = 10, delay_seconds: float = 0.5) -> None:
        last_error: Exception | None = None
        for _ in range(retries):
            try:
                for child in temp_dir.glob("*"):
                    if child.is_file():
                        child.unlink(missing_ok=True)
                    elif child.is_dir():
                        shutil.rmtree(child, ignore_errors=True)
                temp_dir.rmdir()
                return
            except Exception as exc:  # noqa: BLE001
                last_error = exc
                time.sleep(delay_seconds)
        if temp_dir.exists() and last_error is not None:
            raise last_error
