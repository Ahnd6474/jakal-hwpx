from __future__ import annotations

import io
import json
import mimetypes
import os
import re
import shutil
import tempfile
import time
import zipfile
from collections import Counter
from dataclasses import replace
from pathlib import Path, PurePosixPath
from uuid import uuid4

from lxml import etree

from .elements import (
    _invalidate_paragraph_layout,
    _missing_preserved_tokens,
    _preserved_structure_signature,
    _build_chart_part_root,
    _set_form_caption_text,
    _set_graphic_layout,
    _set_margin_values,
    AutoNumberXml,
    BookmarkXml,
    CharacterStyleXml,
    ChartXml,
    EquationXml,
    FieldXml,
    FormXml,
    HeaderFooterXml,
    MemoXml,
    MemoShapeXml,
    NoteXml,
    OleXml,
    ParagraphStyleXml,
    PictureXml,
    SectionSettingsXml,
    ShapeXml,
    StyleDefinitionXml,
    TableXml,
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
from .xmlnode import HwpxXmlNode


def _normalize_path(path: str) -> str:
    return PurePosixPath(path).as_posix()


def _bool_like_literal_is_valid(value: str | None) -> bool:
    if value is None:
        return True
    return value.strip().lower() in {"0", "1", "false", "true", "no", "yes", "n", "y", "off", "on"}


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


_GRAPHIC_CONTROL_TAGS = (
    "tbl",
    "pic",
    "rect",
    "line",
    "ellipse",
    "arc",
    "polygon",
    "curve",
    "connectLine",
    "textart",
    "container",
    "ole",
)

_FORM_TAG_BY_TYPE = {
    "CHECKBOX": "checkBtn",
    "RADIOBUTTON": "radioBtn",
    "BUTTON": "btn",
    "INPUT": "edit",
    "EDIT": "edit",
    "COMBOBOX": "comboBox",
    "LISTBOX": "listBox",
    "SCROLLBAR": "scrollBar",
}
_FORM_TAGS = tuple(_FORM_TAG_BY_TYPE.values())
_FORM_METADATA_COMMAND_PREFIX = "JAKAL_FORM_META:"
_MEMO_CARRIER_FIELD_TYPE = "JAKAL_MEMO"
_BUTTON_FORM_TAGS = {"checkBtn", "radioBtn", "btn"}
_DEFAULT_CHART_FALLBACK_OLE_FIXTURE = "tests/fixtures/charts/chart_01_single_column.hwpx"
_DEFAULT_CHART_FALLBACK_OLE_BYTES: bytes | None = None


def _graphic_attribute_xpath(attribute: str) -> str:
    return " | ".join(f".//hp:{tag}/@{attribute}" for tag in _GRAPHIC_CONTROL_TAGS)


def _is_chart_fallback_ole(node: etree._Element) -> bool:
    return bool(node.xpath("ancestor::hp:default[parent::hp:switch[hp:case/hp:chart]]", namespaces=NS))


def _default_chart_fallback_ole_bytes() -> bytes:
    global _DEFAULT_CHART_FALLBACK_OLE_BYTES
    if _DEFAULT_CHART_FALLBACK_OLE_BYTES is not None:
        return _DEFAULT_CHART_FALLBACK_OLE_BYTES

    candidate = Path(__file__).resolve().parents[2] / ".codex-temp" / "HwpForge" / _DEFAULT_CHART_FALLBACK_OLE_FIXTURE
    if candidate.exists():
        with zipfile.ZipFile(candidate) as archive:
            _DEFAULT_CHART_FALLBACK_OLE_BYTES = archive.read("BinData/ole1.ole")
            return _DEFAULT_CHART_FALLBACK_OLE_BYTES

    # Fallback stub: enough to keep the package structurally valid when the probe fixture is absent.
    _DEFAULT_CHART_FALLBACK_OLE_BYTES = bytes.fromhex("d0cf11e0a1b11ae1") + (b"\x00" * 504)
    return _DEFAULT_CHART_FALLBACK_OLE_BYTES


def _expected_binary_media_type(path: str) -> str | None:
    suffix = PurePosixPath(path).suffix.lower()
    if suffix == ".ole":
        return "application/ole"
    return mimetypes.guess_type(PurePosixPath(path).name)[0]


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


def _new_temp_dir(prefix: str) -> Path:
    root = Path(
        os.environ.get("JAKAL_HWPX_TEMP_ROOT")
        or os.environ.get("TMPDIR")
        or os.environ.get("TEMP")
        or tempfile.gettempdir()
    )
    root.mkdir(parents=True, exist_ok=True)
    for _ in range(32):
        candidate = root / f"{prefix}{uuid4().hex}"
        try:
            candidate.mkdir(parents=True, exist_ok=False)
            return candidate
        except FileExistsError:
            continue
    return Path(tempfile.mkdtemp(prefix=prefix, dir=str(root)))


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

    def xml_part(self, path: str) -> HwpxXmlNode:
        part = self.get_part(path, XmlPart)
        return part.root

    def section_xml(self, section_index: int = 0) -> HwpxXmlNode:
        return self.sections[section_index].root

    def header_xml(self) -> HwpxXmlNode:
        return self.header.root

    def content_hpf_xml(self) -> HwpxXmlNode:
        return self.content_hpf.root

    def settings_xml(self) -> HwpxXmlNode:
        return self.get_part("settings.xml", SettingsPart).root

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

    def headers(self, section_index: int | None = None) -> list[HeaderFooterXml]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        blocks: list[HeaderFooterXml] = []
        for section in sections:
            for node in section.root_element.xpath(".//hp:header", namespaces=NS):
                blocks.append(HeaderFooterXml(self, section, node))
        return blocks

    def footers(self, section_index: int | None = None) -> list[HeaderFooterXml]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        blocks: list[HeaderFooterXml] = []
        for section in sections:
            for node in section.root_element.xpath(".//hp:footer", namespaces=NS):
                blocks.append(HeaderFooterXml(self, section, node))
        return blocks

    def tables(self, section_index: int | None = None) -> list[TableXml]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        tables: list[TableXml] = []
        for section in sections:
            for node in section.root_element.xpath(".//hp:tbl", namespaces=NS):
                tables.append(TableXml(self, section, node))
        return tables

    def pictures(self, section_index: int | None = None) -> list[PictureXml]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        pictures: list[PictureXml] = []
        for section in sections:
            for node in section.root_element.xpath(".//hp:pic", namespaces=NS):
                pictures.append(PictureXml(self, section, node))
        return pictures

    def oles(self, section_index: int | None = None) -> list[OleXml]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        values: list[OleXml] = []
        for section in sections:
            for node in section.root_element.xpath(".//hp:ole", namespaces=NS):
                if _is_chart_fallback_ole(node):
                    continue
                values.append(OleXml(self, section, node))
        return values

    def section_settings(self, section_index: int = 0) -> SectionSettingsXml:
        self._ensure_editable_sections()
        nodes = self.sections[section_index].root_element.xpath("./hp:p[1]//hp:secPr[1]", namespaces=NS)
        if not nodes:
            raise ValueError("Section does not contain hp:secPr.")
        return SectionSettingsXml(self.sections[section_index], nodes[0])

    def notes(self, section_index: int | None = None) -> list[NoteXml]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        notes: list[NoteXml] = []
        for section in sections:
            for tag in (".//hp:footNote", ".//hp:endNote"):
                for node in section.root_element.xpath(tag, namespaces=NS):
                    notes.append(NoteXml(self, section, node))
        return notes

    def memos(self, section_index: int | None = None) -> list[MemoXml]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        values: list[MemoXml] = []
        for section in sections:
            for node in section.root_element.xpath(".//hp:hiddenComment", namespaces=NS):
                values.append(MemoXml(self, section, node))
        return values

    def forms(self, section_index: int | None = None) -> list[FormXml]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        values: list[FormXml] = []
        xpath = " | ".join(f".//hp:{tag}" for tag in _FORM_TAGS)
        for section in sections:
            for node in section.root_element.xpath(xpath, namespaces=NS):
                values.append(FormXml(self, section, node))
        return values

    def charts(self, section_index: int | None = None) -> list[ChartXml]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        values: list[ChartXml] = []
        for section in sections:
            for node in section.root_element.xpath(".//hp:chart", namespaces=NS):
                values.append(ChartXml(self, section, node))
        return values

    def bookmarks(self, section_index: int | None = None) -> list[BookmarkXml]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        bookmarks: list[BookmarkXml] = []
        for section in sections:
            for node in section.root_element.xpath(".//hp:bookmark", namespaces=NS):
                bookmarks.append(BookmarkXml(self, section, node))
        return bookmarks

    def fields(self, section_index: int | None = None) -> list[FieldXml]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        fields: list[FieldXml] = []
        for section in sections:
            for node in section.root_element.xpath(".//hp:fieldBegin", namespaces=NS):
                fields.append(FieldXml(self, section, node))
        return fields

    def hyperlinks(self, section_index: int | None = None) -> list[FieldXml]:
        return [field for field in self.fields(section_index=section_index) if field.is_hyperlink]

    def mail_merge_fields(self, section_index: int | None = None) -> list[FieldXml]:
        return [field for field in self.fields(section_index=section_index) if field.is_mail_merge]

    def calculation_fields(self, section_index: int | None = None) -> list[FieldXml]:
        return [field for field in self.fields(section_index=section_index) if field.is_calculation]

    def cross_references(self, section_index: int | None = None) -> list[FieldXml]:
        return [field for field in self.fields(section_index=section_index) if field.is_cross_reference]

    def auto_numbers(self, section_index: int | None = None) -> list[AutoNumberXml]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        values: list[AutoNumberXml] = []
        for section in sections:
            for tag in (".//hp:autoNum", ".//hp:newNum"):
                for node in section.root_element.xpath(tag, namespaces=NS):
                    values.append(AutoNumberXml(self, section, node))
        return values

    def equations(self, section_index: int | None = None) -> list[EquationXml]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        values: list[EquationXml] = []
        for section in sections:
            for node in section.root_element.xpath(".//hp:equation", namespaces=NS):
                values.append(EquationXml(self, section, node))
        return values

    def shapes(self, section_index: int | None = None) -> list[ShapeXml]:
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
        values: list[ShapeXml] = []
        for section in sections:
            for tag in shape_tags:
                for node in section.root_element.xpath(f".//{tag}", namespaces=NS):
                    if etree.QName(node).localname == "ole":
                        if _is_chart_fallback_ole(node):
                            continue
                        values.append(OleXml(self, section, node))
                    else:
                        values.append(ShapeXml(self, section, node))
        return values

    def styles(self) -> list[StyleDefinitionXml]:
        return [StyleDefinitionXml(self.header, node) for node in self.header.root_element.xpath(".//hh:style", namespaces=NS)]

    def paragraph_styles(self) -> list[ParagraphStyleXml]:
        return [ParagraphStyleXml(self.header, node) for node in self.header.root_element.xpath(".//hh:paraPr", namespaces=NS)]

    def character_styles(self) -> list[CharacterStyleXml]:
        return [CharacterStyleXml(self.header, node) for node in self.header.root_element.xpath(".//hh:charPr", namespaces=NS)]

    def memo_shapes(self) -> list[MemoShapeXml]:
        return [MemoShapeXml(self.header, node) for node in self.header.root_element.xpath(".//hh:memoPr", namespaces=NS)]

    def get_style(self, style_id: str) -> StyleDefinitionXml:
        for style in self.styles():
            if style.style_id == style_id:
                return style
        raise KeyError(style_id)

    def get_paragraph_style(self, style_id: str) -> ParagraphStyleXml:
        for style in self.paragraph_styles():
            if style.style_id == style_id:
                return style
        raise KeyError(style_id)

    def get_character_style(self, style_id: str) -> CharacterStyleXml:
        for style in self.character_styles():
            if style.style_id == style_id:
                return style
        raise KeyError(style_id)

    def get_memo_shape(self, memo_shape_id: str) -> MemoShapeXml:
        for memo_shape in self.memo_shapes():
            if memo_shape.memo_shape_id == memo_shape_id:
                return memo_shape
        raise KeyError(memo_shape_id)

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
        is_embedded: bool | int | str = True,
    ) -> BinaryDataPart:
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
        if isinstance(is_embedded, bool):
            is_embedded_value: str | int | str = "1" if is_embedded else "0"
        else:
            is_embedded_value = is_embedded
        self.content_hpf.ensure_manifest_item(
            manifest_id,
            part_path,
            media_type,
            isEmbeded=is_embedded_value,
        )
        return part

    def append_control_xml(
        self,
        xml: str | bytes,
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> HwpxXmlNode:
        self._ensure_editable_sections()
        paragraph = self._resolve_paragraph_for_insert(section_index=section_index, paragraph_index=paragraph_index)
        run = self._append_run(paragraph, char_pr_id=char_pr_id)
        ctrl = etree.SubElement(run, qname("hp", "ctrl"))
        node = HwpxXmlNode(ctrl, self.sections[section_index]).append_xml(xml)
        _invalidate_paragraph_layout(paragraph)
        return node

    def append_run_xml(
        self,
        xml: str | bytes,
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> HwpxXmlNode:
        self._ensure_editable_sections()
        paragraph = self._resolve_paragraph_for_insert(section_index=section_index, paragraph_index=paragraph_index)
        run = self._append_run(paragraph, char_pr_id=char_pr_id)
        node = HwpxXmlNode(run, self.sections[section_index]).append_xml(xml)
        _invalidate_paragraph_layout(paragraph)
        return node

    def append_header(
        self,
        text: str = "",
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
        apply_page_type: str = "BOTH",
        hide_first: bool | None = None,
    ) -> HeaderFooterXml:
        self._ensure_editable_sections()
        block = self._build_header_footer_block(
            "header",
            text=text,
            char_pr_id=char_pr_id,
            apply_page_type=apply_page_type,
        )
        section = self._append_control(section_index, paragraph_index, block, char_pr_id=char_pr_id)
        if hide_first is not None:
            self.section_settings(section_index).set_visibility(hide_first_header=hide_first)
        return HeaderFooterXml(self, section, block)

    def append_footer(
        self,
        text: str = "",
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
        apply_page_type: str = "BOTH",
        hide_first: bool | None = None,
    ) -> HeaderFooterXml:
        self._ensure_editable_sections()
        block = self._build_header_footer_block(
            "footer",
            text=text,
            char_pr_id=char_pr_id,
            apply_page_type=apply_page_type,
        )
        section = self._append_control(section_index, paragraph_index, block, char_pr_id=char_pr_id)
        if hide_first is not None:
            self.section_settings(section_index).set_visibility(hide_first_footer=hide_first)
        return HeaderFooterXml(self, section, block)

    def append_note(
        self,
        text: str = "",
        *,
        kind: str = "footNote",
        number: int | str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> NoteXml:
        self._ensure_editable_sections()
        if kind not in {"footNote", "endNote"}:
            raise ValueError("append_note() kind must be 'footNote' or 'endNote'.")
        note = etree.Element(qname("hp", kind))
        note.set("id", self._next_control_number(".//hp:footNote/@id | .//hp:endNote/@id"))
        note.set("number", str(number if number is not None else self._next_control_number(".//hp:footNote/@number | .//hp:endNote/@number")))
        note.append(self._build_sublist_with_text(text, char_pr_id=char_pr_id, vertical_align="TOP"))
        section = self._append_control(section_index, paragraph_index, note, char_pr_id=char_pr_id)
        return NoteXml(self, section, note)

    def append_footnote(
        self,
        text: str = "",
        *,
        number: int | str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> NoteXml:
        return self.append_note(
            text,
            kind="footNote",
            number=number,
            section_index=section_index,
            paragraph_index=paragraph_index,
            char_pr_id=char_pr_id,
        )

    def append_endnote(
        self,
        text: str = "",
        *,
        number: int | str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> NoteXml:
        return self.append_note(
            text,
            kind="endNote",
            number=number,
            section_index=section_index,
            paragraph_index=paragraph_index,
            char_pr_id=char_pr_id,
        )

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
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> FormXml:
        self._ensure_editable_sections()
        normalized_form_type = form_type.strip().upper()
        tag = _FORM_TAG_BY_TYPE.get(normalized_form_type)
        if tag is None:
            raise ValueError(f"Unsupported HWPX form_type: {form_type}")

        form = etree.Element(qname("hp", tag))
        form.set("name", name or "")
        form.set("foreColor", "#000000")
        form.set("backColor", "#7FFFFFFF")
        form.set("groupName", "")
        form.set("tabStop", "1")
        form.set("editable", "1" if editable else "0")
        form.set("tabOrder", "1")
        form.set("enabled", "0" if locked else "1")
        form.set("borderTypeIDRef", "0")
        form.set("drawFrame", "1")
        form.set("printable", "1")
        form.set("command", "")

        extra_metadata: dict[str, object] = {}
        if tag == "checkBtn":
            form.set("caption", label)
            form.set("value", value or ("CHECKED" if checked else "UNCHECKED"))
            form.set("radioGroupName", "")
            form.set("triState", "0")
            form.set("backStyle", "TRANSPARENT")
        elif tag == "radioBtn":
            form.set("caption", label)
            form.set("value", value or ("CHECKED" if checked else "UNCHECKED"))
            form.set("radioGroupName", name or "")
            form.set("triState", "0")
            form.set("backStyle", "TRANSPARENT")
        elif tag == "btn":
            form.set("caption", label)
            form.set("radioGroupName", "")
            form.set("triState", "0")
        elif tag == "edit":
            form.set("multiLine", "0")
            form.set("passwordChar", "")
            form.set("maxLength", "32")
            form.set("scrollBars", "NONE")
            form.set("tabKeyBehavior", "NEXT_OBJECT")
            form.set("numOnly", "0")
            form.set("readOnly", "0")
            form.set("alignText", "LEFT")
            text_node = etree.SubElement(form, qname("hp", "text"))
            text_node.text = value or ""
        elif tag == "comboBox":
            form.set("listBoxRows", "4")
            form.set("listBoxWidth", "2400")
            form.set("editEnable", "1" if editable else "0")
            form.set("selectedValue", value or "")
            for item_value in items or ([value] if value else []):
                item = etree.SubElement(form, qname("hp", "listItem"))
                item.set("displayText", "")
                item.set("value", item_value or "")
        elif tag == "listBox":
            form.set("itemHeight", "300")
            form.set("topIdx", "4294967295")
            form.set("selectedValue", value or "")
            for item_value in items or ([value] if value else []):
                item = etree.SubElement(form, qname("hp", "listItem"))
                item.set("displayText", "")
                item.set("value", item_value or "")
        elif tag == "scrollBar":
            form.set("delay", "0")
            form.set("largeChange", "10")
            form.set("smallChange", "1")
            form.set("min", "0")
            form.set("max", "100")
            form.set("page", "10")
            form.set("value", value or "0")
            form.set("type", "HORIZONTAL")

        if tag != "scrollBar":
            form_char_pr = etree.SubElement(form, qname("hp", "formCharPr"))
            form_char_pr.set("charPrIDRef", char_pr_id or "0")
            form_char_pr.set("followContext", "0")
            form_char_pr.set("autoSz", "0")
            form_char_pr.set("wordWrap", "0")
        elif label:
            _set_form_caption_text(form, label, char_pr_id=char_pr_id or "0")

        size = etree.SubElement(form, qname("hp", "sz"))
        if tag == "btn":
            size.set("width", "3200")
            size.set("height", "1400")
        elif tag == "edit":
            size.set("width", "4200")
            size.set("height", "1400")
        elif tag == "listBox":
            size.set("width", "4200")
            size.set("height", "1800")
        elif tag == "scrollBar":
            size.set("width", "4200")
            size.set("height", "900")
        else:
            size.set("width", "2400" if tag in {"checkBtn", "radioBtn"} else "4200")
            size.set("height", "1200" if tag in {"checkBtn", "radioBtn"} else "1400")
        size.set("widthRelTo", "ABSOLUTE")
        size.set("heightRelTo", "ABSOLUTE")
        size.set("protect", "0")

        pos = etree.SubElement(form, qname("hp", "pos"))
        pos.set("treatAsChar", "1")
        pos.set("affectLSpacing", "0")
        pos.set("flowWithText", "1")
        pos.set("allowOverlap", "0")
        pos.set("holdAnchorAndSO", "0")
        pos.set("vertRelTo", "PARA")
        pos.set("horzRelTo", "COLUMN")
        pos.set("vertAlign", "TOP")
        pos.set("horzAlign", "LEFT")
        pos.set("vertOffset", "0")
        pos.set("horzOffset", "0")

        out_margin = etree.SubElement(form, qname("hp", "outMargin"))
        out_margin.set("left", "0")
        out_margin.set("right", "0")
        out_margin.set("top", "0")
        out_margin.set("bottom", "0")

        if label and tag not in _BUTTON_FORM_TAGS and tag != "scrollBar":
            _set_form_caption_text(form, label, char_pr_id=char_pr_id or "0")

        if placeholder:
            extra_metadata["placeholder"] = placeholder
        if items and tag not in {"comboBox", "listBox"}:
            extra_metadata["items"] = list(items)
        if extra_metadata:
            form.set(
                "command",
                _FORM_METADATA_COMMAND_PREFIX + json.dumps(extra_metadata, ensure_ascii=False, separators=(",", ":"), sort_keys=True),
            )

        section = self._append_run_child(section_index, paragraph_index, form, char_pr_id=char_pr_id)
        return FormXml(self, section, form)

    def append_memo(
        self,
        text: str = "",
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> MemoXml:
        self._ensure_editable_sections()
        memo = etree.Element(qname("hp", "hiddenComment"))
        memo.append(self._build_sublist_with_text(text, char_pr_id=char_pr_id, vertical_align="TOP"))
        section = self._append_control(section_index, paragraph_index, memo, char_pr_id=char_pr_id)
        return MemoXml(self, section, memo)

    def append_memo_shape(
        self,
        *,
        memo_shape_id: int | str | None = None,
        width: int | str = 12000,
        line_width: int | str = 12,
        line_type: str = "SOLID",
        line_color: str = "#808080",
        fill_color: str = "#FFF2CC",
        active_color: str = "#FFC000",
        memo_type: str = "NOMAL",
    ) -> MemoShapeXml:
        resolved_id = str(memo_shape_id) if memo_shape_id is not None else self._next_header_numeric_id(".//hh:memoPr/@id")
        memo_shape = etree.Element(qname("hh", "memoPr"))
        memo_shape.set("id", resolved_id)
        memo_shape.set("width", str(width))
        memo_shape.set("lineWidth", str(line_width))
        memo_shape.set("lineType", line_type)
        memo_shape.set("lineColor", line_color)
        memo_shape.set("fillColor", fill_color)
        memo_shape.set("activeColor", active_color)
        memo_shape.set("memoType", memo_type)

        parent = self._ensure_header_collection(".//hh:memoProperties", "memoProperties")
        parent.append(memo_shape)
        parent.set("itemCnt", str(len(parent.xpath("./hh:memoPr", namespaces=NS))))
        self.header.mark_modified()
        return MemoShapeXml(self.header, memo_shape)

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
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
        text_wrap: str = "TOP_AND_BOTTOM",
        text_flow: str = "BOTH_SIDES",
        treat_as_char: bool = True,
        affect_line_spacing: bool = False,
        flow_with_text: bool = True,
        allow_overlap: bool = False,
        hold_anchor_and_so: bool = False,
        vert_rel_to: str = "PARA",
        horz_rel_to: str = "COLUMN",
        vert_align: str = "TOP",
        horz_align: str = "LEFT",
        vert_offset: int = 0,
        horz_offset: int = 0,
        out_margin_left: int = 0,
        out_margin_right: int = 0,
        out_margin_top: int = 0,
        out_margin_bottom: int = 0,
    ) -> ChartXml:
        self._ensure_editable_sections()

        chart_path = self._next_chart_part_path()
        chart_control_id = self._next_control_number(".//hp:chart/@id | .//hp:ole/@id")
        chart_z_order = self._next_control_number(".//hp:chart/@zOrder | .//hp:ole/@zOrder")
        metadata: dict[str, object] = {"chartType": str(chart_type or "BAR").upper()}
        chart_part = self.add_part(
            chart_path,
            etree.tostring(
                _build_chart_part_root(
                    title=title,
                    chart_type=str(chart_type or "BAR").upper(),
                    categories=[str(value) for value in (categories or [])],
                    series=[dict(value) for value in (series or [])],
                    data_ref=None if data_ref in (None, "") else str(data_ref),
                    legend_visible=legend_visible,
                    metadata=metadata,
                ),
                encoding="UTF-8",
                xml_declaration=True,
                standalone=True,
                pretty_print=False,
            ),
        )

        chart = etree.Element(qname("hp", "chart"))
        chart.set("id", chart_control_id)
        chart.set("zOrder", chart_z_order)
        chart.set("numberingType", "PICTURE")
        chart.set("textWrap", text_wrap)
        chart.set("textFlow", text_flow)
        chart.set("lock", "0")
        chart.set("dropcapstyle", "None")
        chart.set("chartIDRef", chart_path)
        if shape_comment:
            comment = etree.SubElement(chart, qname("hp", "shapeComment"))
            comment.text = shape_comment

        size = etree.SubElement(chart, qname("hp", "sz"))
        size.set("width", str(width))
        size.set("widthRelTo", "ABSOLUTE")
        size.set("height", str(height))
        size.set("heightRelTo", "ABSOLUTE")
        size.set("protect", "0")

        position = etree.SubElement(chart, qname("hp", "pos"))
        position.set("treatAsChar", "1" if treat_as_char else "0")
        position.set("affectLSpacing", "0")
        position.set("flowWithText", "1")
        position.set("allowOverlap", "0")
        position.set("holdAnchorAndSO", "0")
        position.set("vertRelTo", "PARA")
        position.set("horzRelTo", "COLUMN")
        position.set("vertAlign", "TOP")
        position.set("horzAlign", "LEFT")
        position.set("vertOffset", "0")
        position.set("horzOffset", "0")
        _set_graphic_layout(
            chart,
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

        out_margin = etree.SubElement(chart, qname("hp", "outMargin"))
        out_margin.set("left", "0")
        out_margin.set("right", "0")
        out_margin.set("top", "0")
        out_margin.set("bottom", "0")
        _set_margin_values(
            chart,
            "./hp:outMargin",
            left=out_margin_left,
            right=out_margin_right,
            top=out_margin_top,
            bottom=out_margin_bottom,
        )

        switch = etree.Element(qname("hp", "switch"))
        case = etree.SubElement(switch, qname("hp", "case"))
        case.set(qname("hp", "required-namespace"), NS["ooxmlchart"])
        case.append(chart)
        default = etree.SubElement(switch, qname("hp", "default"))
        default.append(
            self._build_chart_fallback_ole(
                control_id=chart_control_id,
                z_order=chart_z_order,
                width=width,
                height=height,
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
                out_margin_left=out_margin_left,
                out_margin_right=out_margin_right,
                out_margin_top=out_margin_top,
                out_margin_bottom=out_margin_bottom,
            )
        )

        paragraph = self._resolve_paragraph_for_insert(section_index=section_index, paragraph_index=paragraph_index)
        run = self._append_run(paragraph, char_pr_id=char_pr_id)
        run.append(switch)
        etree.SubElement(run, qname("hp", "t"))

        section = self.sections[section_index]
        section.mark_modified()
        chart_part.mark_modified()
        return ChartXml(self, section, chart)

    def append_auto_number(
        self,
        *,
        number: int | str = 1,
        number_type: str = "PAGE",
        kind: str = "newNum",
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> AutoNumberXml:
        self._ensure_editable_sections()
        if kind not in {"newNum", "autoNum"}:
            raise ValueError("append_auto_number() kind must be 'newNum' or 'autoNum'.")
        node = etree.Element(qname("hp", kind))
        node.set("num", str(number))
        node.set("numType", number_type)
        section = self._append_control(section_index, paragraph_index, node, char_pr_id=char_pr_id)
        return AutoNumberXml(self, section, node)

    def append_new_number(
        self,
        *,
        number: int | str = 1,
        number_type: str = "PAGE",
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> AutoNumberXml:
        return self.append_auto_number(
            number=number,
            number_type=number_type,
            kind="newNum",
            section_index=section_index,
            paragraph_index=paragraph_index,
            char_pr_id=char_pr_id,
        )

    def append_equation(
        self,
        script: str,
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
        width: int = 4800,
        height: int = 2300,
        treat_as_char: bool = True,
        shape_comment: str | None = None,
        text_color: str = "#000000",
        base_unit: int = 1100,
        font: str = "HYhwpEQ",
        text_wrap: str = "TOP_AND_BOTTOM",
        text_flow: str = "BOTH_SIDES",
        affect_line_spacing: bool = False,
        flow_with_text: bool = True,
        allow_overlap: bool = False,
        hold_anchor_and_so: bool = False,
        vert_rel_to: str = "PARA",
        horz_rel_to: str = "COLUMN",
        vert_align: str = "TOP",
        horz_align: str = "LEFT",
        vert_offset: int = 0,
        horz_offset: int = 0,
        out_margin_left: int = 0,
        out_margin_right: int = 0,
        out_margin_top: int = 0,
        out_margin_bottom: int = 0,
    ) -> EquationXml:
        self._ensure_editable_sections()
        if width < 1 or height < 1:
            raise ValueError("width and height must be positive.")

        equation = etree.Element(qname("hp", "equation"))
        equation.set("id", self._next_control_number(_graphic_attribute_xpath("id")))
        equation.set("zOrder", self._next_control_number(_graphic_attribute_xpath("zOrder")))
        equation.set("numberingType", "EQUATION")
        equation.set("textWrap", "TOP_AND_BOTTOM")
        equation.set("textFlow", "BOTH_SIDES")
        equation.set("lock", "0")
        equation.set("dropcapstyle", "None")
        equation.set("version", "Equation Version 60")
        equation.set("baseLine", "93")
        equation.set("textColor", text_color)
        equation.set("baseUnit", str(base_unit))
        equation.set("lineMode", "CHAR")
        equation.set("font", font)

        size = etree.SubElement(equation, qname("hp", "sz"))
        size.set("width", str(width))
        size.set("widthRelTo", "ABSOLUTE")
        size.set("height", str(height))
        size.set("heightRelTo", "ABSOLUTE")
        size.set("protect", "0")

        position = etree.SubElement(equation, qname("hp", "pos"))
        position.set("treatAsChar", "1" if treat_as_char else "0")
        position.set("affectLSpacing", "0")
        position.set("flowWithText", "1")
        position.set("allowOverlap", "0")
        position.set("holdAnchorAndSO", "0")
        position.set("vertRelTo", "PARA")
        position.set("horzRelTo", "COLUMN")
        position.set("vertAlign", "TOP")
        position.set("horzAlign", "LEFT")
        position.set("vertOffset", "0")
        position.set("horzOffset", "0")

        out_margin = etree.SubElement(equation, qname("hp", "outMargin"))
        for key in ("left", "right", "top", "bottom"):
            out_margin.set(key, "0")
        _set_graphic_layout(
            equation,
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
        _set_margin_values(
            equation,
            "./hp:outMargin",
            left=out_margin_left,
            right=out_margin_right,
            top=out_margin_top,
            bottom=out_margin_bottom,
        )

        if shape_comment:
            comment = etree.SubElement(equation, qname("hp", "shapeComment"))
            comment.text = shape_comment

        script_node = etree.SubElement(equation, qname("hp", "script"))
        script_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        script_node.text = script

        section = self._append_run_child(section_index, paragraph_index, equation, char_pr_id=char_pr_id)
        return EquationXml(self, section, equation)

    def append_paragraph_style(
        self,
        *,
        style_id: str | None = None,
        template_id: str | None = None,
        alignment_horizontal: str | None = None,
        alignment_vertical: str | None = None,
        line_spacing: int | str | None = None,
    ) -> ParagraphStyleXml:
        template = self._clone_header_template(".//hh:paraPr", template_id=template_id)
        resolved_id = style_id or self._next_header_numeric_id(".//hh:paraPr/@id")
        template.set("id", resolved_id)
        parent = self._ensure_header_collection(".//hh:paraProperties", "paraProperties")
        parent.append(template)
        parent.set("itemCnt", str(len(parent.xpath("./hh:paraPr", namespaces=NS))))
        self.header.mark_modified()

        style = ParagraphStyleXml(self.header, template)
        if alignment_horizontal is not None or alignment_vertical is not None:
            style.set_alignment(horizontal=alignment_horizontal, vertical=alignment_vertical)
        if line_spacing is not None:
            style.set_line_spacing(line_spacing)
        return style

    def append_character_style(
        self,
        *,
        style_id: str | None = None,
        template_id: str | None = None,
        text_color: str | None = None,
        height: int | str | None = None,
    ) -> CharacterStyleXml:
        template = self._clone_header_template(".//hh:charPr", template_id=template_id)
        resolved_id = style_id or self._next_header_numeric_id(".//hh:charPr/@id")
        template.set("id", resolved_id)
        parent = self._ensure_header_collection(".//hh:charProperties", "charProperties")
        parent.append(template)
        parent.set("itemCnt", str(len(parent.xpath("./hh:charPr", namespaces=NS))))
        self.header.mark_modified()

        style = CharacterStyleXml(self.header, template)
        if text_color is not None:
            style.set_text_color(text_color)
        if height is not None:
            style.set_height(height)
        return style

    def append_style(
        self,
        name: str,
        *,
        english_name: str | None = None,
        style_id: str | None = None,
        style_type: str = "PARA",
        para_pr_id: str | None = None,
        char_pr_id: str | None = None,
        next_style_id: str | None = None,
        lang_id: str = "1042",
        lock_form: bool = False,
        template_id: str | None = None,
    ) -> StyleDefinitionXml:
        template = self._clone_header_template(".//hh:style", template_id=template_id)
        resolved_id = style_id or self._next_header_numeric_id(".//hh:style/@id")
        template.set("id", resolved_id)
        template.set("type", style_type)
        template.set("name", name)
        template.set("engName", english_name or name)
        template.set("paraPrIDRef", para_pr_id or template.get("paraPrIDRef", "0"))
        template.set("charPrIDRef", char_pr_id or template.get("charPrIDRef", "0"))
        template.set("nextStyleIDRef", next_style_id or template.get("nextStyleIDRef", resolved_id))
        template.set("langID", lang_id)
        template.set("lockForm", "1" if lock_form else "0")

        parent = self._ensure_header_collection(".//hh:styles", "styles")
        parent.append(template)
        parent.set("itemCnt", str(len(parent.xpath("./hh:style", namespaces=NS))))
        self.header.mark_modified()
        return StyleDefinitionXml(self.header, template)

    def append_table(
        self,
        rows: int,
        columns: int,
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
        cell_texts: list[list[str]] | None = None,
        cell_width: int = 16000,
        row_height: int = 1800,
        table_width: int | None = None,
        treat_as_char: bool = True,
        text_wrap: str = "TOP_AND_BOTTOM",
        text_flow: str = "BOTH_SIDES",
        affect_line_spacing: bool = False,
        flow_with_text: bool = True,
        allow_overlap: bool = False,
        hold_anchor_and_so: bool = False,
        vert_rel_to: str = "PARA",
        horz_rel_to: str = "COLUMN",
        vert_align: str = "TOP",
        horz_align: str = "LEFT",
        vert_offset: int = 0,
        horz_offset: int = 0,
        out_margin_left: int = 0,
        out_margin_right: int = 0,
        out_margin_top: int = 0,
        out_margin_bottom: int = 0,
    ) -> TableXml:
        self._ensure_editable_sections()
        if rows < 1 or columns < 1:
            raise ValueError("rows and columns must be at least 1.")
        if cell_width < 1 or row_height < 1:
            raise ValueError("cell_width and row_height must be positive.")

        width = table_width if table_width is not None else cell_width * columns
        if width < 1:
            raise ValueError("table_width must be positive.")

        table = etree.Element(qname("hp", "tbl"))
        table.set("id", self._next_control_number(_graphic_attribute_xpath("id")))
        table.set("zOrder", self._next_control_number(_graphic_attribute_xpath("zOrder")))
        table.set("numberingType", "TABLE")
        table.set("textWrap", "TOP_AND_BOTTOM")
        table.set("textFlow", "BOTH_SIDES")
        table.set("lock", "0")
        table.set("dropcapstyle", "None")
        table.set("pageBreak", "CELL")
        table.set("repeatHeader", "1")
        table.set("rowCnt", str(rows))
        table.set("colCnt", str(columns))
        table.set("cellSpacing", "0")
        table.set("borderFillIDRef", "0")
        table.set("noAdjust", "0")

        size = etree.SubElement(table, qname("hp", "sz"))
        size.set("width", str(width))
        size.set("widthRelTo", "ABSOLUTE")
        size.set("height", str(row_height * rows))
        size.set("heightRelTo", "ABSOLUTE")
        size.set("protect", "0")

        position = etree.SubElement(table, qname("hp", "pos"))
        position.set("treatAsChar", "1" if treat_as_char else "0")
        position.set("affectLSpacing", "0")
        position.set("flowWithText", "1")
        position.set("allowOverlap", "0")
        position.set("holdAnchorAndSO", "0")
        position.set("vertRelTo", "PARA")
        position.set("horzRelTo", "COLUMN")
        position.set("vertAlign", "TOP")
        position.set("horzAlign", "LEFT")
        position.set("vertOffset", "0")
        position.set("horzOffset", "0")

        out_margin = etree.SubElement(table, qname("hp", "outMargin"))
        for key in ("left", "right", "top", "bottom"):
            out_margin.set(key, "0")
        _set_graphic_layout(
            table,
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
        _set_margin_values(
            table,
            "./hp:outMargin",
            left=out_margin_left,
            right=out_margin_right,
            top=out_margin_top,
            bottom=out_margin_bottom,
        )

        in_margin = etree.SubElement(table, qname("hp", "inMargin"))
        in_margin.set("left", "141")
        in_margin.set("right", "141")
        in_margin.set("top", "141")
        in_margin.set("bottom", "141")

        nested_paragraph_id = self._next_control_number(".//hp:p/@id")
        for row_index in range(rows):
            row = etree.SubElement(table, qname("hp", "tr"))
            for column_index in range(columns):
                cell = etree.SubElement(row, qname("hp", "tc"))
                cell.set("name", "")
                cell.set("header", "0")
                cell.set("hasMargin", "0")
                cell.set("protect", "0")
                cell.set("editable", "0")
                cell.set("dirty", "0")
                cell.set("borderFillIDRef", "0")

                sub_list = etree.SubElement(cell, qname("hp", "subList"))
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
                paragraph.set("id", nested_paragraph_id)
                nested_paragraph_id = str(int(nested_paragraph_id) + 1)
                paragraph.set("paraPrIDRef", "0")
                paragraph.set("styleIDRef", "0")
                paragraph.set("pageBreak", "0")
                paragraph.set("columnBreak", "0")
                paragraph.set("merged", "0")

                run = etree.SubElement(paragraph, qname("hp", "run"))
                run.set("charPrIDRef", char_pr_id or "0")
                text_node = etree.SubElement(run, qname("hp", "t"))
                value = ""
                if cell_texts is not None and row_index < len(cell_texts) and column_index < len(cell_texts[row_index]):
                    value = cell_texts[row_index][column_index]
                text_node.text = value

                cell_addr = etree.SubElement(cell, qname("hp", "cellAddr"))
                cell_addr.set("colAddr", str(column_index))
                cell_addr.set("rowAddr", str(row_index))

                cell_span = etree.SubElement(cell, qname("hp", "cellSpan"))
                cell_span.set("colSpan", "1")
                cell_span.set("rowSpan", "1")

                cell_size = etree.SubElement(cell, qname("hp", "cellSz"))
                cell_size.set("width", str(cell_width))
                cell_size.set("height", str(row_height))

                cell_margin = etree.SubElement(cell, qname("hp", "cellMargin"))
                cell_margin.set("left", "141")
                cell_margin.set("right", "141")
                cell_margin.set("top", "141")
                cell_margin.set("bottom", "141")

        section = self._append_run_child(section_index, paragraph_index, table, char_pr_id=char_pr_id)
        return TableXml(self, section, table)

    def append_picture(
        self,
        name: str,
        data: bytes,
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
        media_type: str | None = None,
        manifest_id: str | None = None,
        width: int = 7200,
        height: int = 7200,
        original_width: int | None = None,
        original_height: int | None = None,
        treat_as_char: bool = True,
        shape_comment: str | None = None,
        text_wrap: str = "TOP_AND_BOTTOM",
        text_flow: str = "BOTH_SIDES",
        affect_line_spacing: bool = False,
        flow_with_text: bool = True,
        allow_overlap: bool = False,
        hold_anchor_and_so: bool = False,
        vert_rel_to: str = "PARA",
        horz_rel_to: str = "COLUMN",
        vert_align: str = "TOP",
        horz_align: str = "LEFT",
        vert_offset: int = 0,
        horz_offset: int = 0,
        out_margin_left: int = 0,
        out_margin_right: int = 0,
        out_margin_top: int = 0,
        out_margin_bottom: int = 0,
    ) -> PictureXml:
        self._ensure_editable_sections()
        if width < 1 or height < 1:
            raise ValueError("width and height must be positive.")

        resolved_manifest_id = manifest_id or self.content_hpf.next_manifest_id(PurePosixPath(name).stem or "image")
        self.add_or_replace_binary(name, data, media_type=media_type, manifest_id=resolved_manifest_id)

        image_width = original_width or width
        image_height = original_height or height
        picture = etree.Element(qname("hp", "pic"))
        picture.set("id", self._next_control_number(_graphic_attribute_xpath("id")))
        picture.set("zOrder", self._next_control_number(_graphic_attribute_xpath("zOrder")))
        picture.set("numberingType", "PICTURE")
        picture.set("textWrap", "TOP_AND_BOTTOM")
        picture.set("textFlow", "BOTH_SIDES")
        picture.set("lock", "0")
        picture.set("dropcapstyle", "None")
        picture.set("href", "")
        picture.set("groupLevel", "0")
        picture.set("instid", self._next_control_number(_graphic_attribute_xpath("instid")))
        picture.set("reverse", "0")

        offset = etree.SubElement(picture, qname("hp", "offset"))
        offset.set("x", "0")
        offset.set("y", "0")

        original_size = etree.SubElement(picture, qname("hp", "orgSz"))
        original_size.set("width", str(image_width))
        original_size.set("height", str(image_height))

        current_size = etree.SubElement(picture, qname("hp", "curSz"))
        current_size.set("width", str(width))
        current_size.set("height", str(height))

        flip = etree.SubElement(picture, qname("hp", "flip"))
        flip.set("horizontal", "0")
        flip.set("vertical", "0")

        rotation = etree.SubElement(picture, qname("hp", "rotationInfo"))
        rotation.set("angle", "0")
        rotation.set("centerX", str(width // 2))
        rotation.set("centerY", str(height // 2))
        rotation.set("rotateimage", "1")

        rendering = etree.SubElement(picture, qname("hp", "renderingInfo"))
        self._append_identity_matrix(rendering, "transMatrix")
        scale_x = width / image_width if image_width else 1
        scale_y = height / image_height if image_height else 1
        self._append_matrix(rendering, "scaMatrix", e1=f"{scale_x:.6f}", e5=f"{scale_y:.6f}")
        self._append_identity_matrix(rendering, "rotMatrix")

        image = etree.SubElement(picture, qname("hc", "img"))
        image.set("binaryItemIDRef", resolved_manifest_id)
        image.set("bright", "0")
        image.set("contrast", "0")
        image.set("effect", "REAL_PIC")
        image.set("alpha", "0")

        image_rect = etree.SubElement(picture, qname("hp", "imgRect"))
        for index, (x, y) in enumerate(((0, 0), (image_width, 0), (image_width, image_height), (0, image_height))):
            point = etree.SubElement(image_rect, qname("hc", f"pt{index}"))
            point.set("x", str(x))
            point.set("y", str(y))

        image_clip = etree.SubElement(picture, qname("hp", "imgClip"))
        image_clip.set("left", "0")
        image_clip.set("right", str(image_width))
        image_clip.set("top", "0")
        image_clip.set("bottom", str(image_height))

        in_margin = etree.SubElement(picture, qname("hp", "inMargin"))
        for key in ("left", "right", "top", "bottom"):
            in_margin.set(key, "0")

        image_dim = etree.SubElement(picture, qname("hp", "imgDim"))
        image_dim.set("dimwidth", str(image_width))
        image_dim.set("dimheight", str(image_height))
        etree.SubElement(picture, qname("hp", "effects"))

        size = etree.SubElement(picture, qname("hp", "sz"))
        size.set("width", str(width))
        size.set("widthRelTo", "ABSOLUTE")
        size.set("height", str(height))
        size.set("heightRelTo", "ABSOLUTE")
        size.set("protect", "0")

        position = etree.SubElement(picture, qname("hp", "pos"))
        position.set("treatAsChar", "1" if treat_as_char else "0")
        position.set("affectLSpacing", "0")
        position.set("flowWithText", "1")
        position.set("allowOverlap", "0")
        position.set("holdAnchorAndSO", "0")
        position.set("vertRelTo", "PARA")
        position.set("horzRelTo", "COLUMN")
        position.set("vertAlign", "TOP")
        position.set("horzAlign", "LEFT")
        position.set("vertOffset", "0")
        position.set("horzOffset", "0")

        out_margin = etree.SubElement(picture, qname("hp", "outMargin"))
        for key in ("left", "right", "top", "bottom"):
            out_margin.set(key, "0")
        _set_graphic_layout(
            picture,
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
        _set_margin_values(
            picture,
            "./hp:outMargin",
            left=out_margin_left,
            right=out_margin_right,
            top=out_margin_top,
            bottom=out_margin_bottom,
        )

        if shape_comment:
            comment = etree.SubElement(picture, qname("hp", "shapeComment"))
            comment.text = shape_comment

        section = self._append_run_child(section_index, paragraph_index, picture, char_pr_id=char_pr_id)
        return PictureXml(self, section, picture)

    def append_shape(
        self,
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
        kind: str = "rect",
        text: str = "",
        width: int = 12000,
        height: int = 3200,
        treat_as_char: bool = True,
        shape_comment: str | None = None,
        fill_color: str = "#FFFFFF",
        line_color: str = "#000000",
        text_wrap: str = "TOP_AND_BOTTOM",
        text_flow: str = "BOTH_SIDES",
        affect_line_spacing: bool = False,
        flow_with_text: bool = True,
        allow_overlap: bool = False,
        hold_anchor_and_so: bool = False,
        vert_rel_to: str = "PARA",
        horz_rel_to: str = "COLUMN",
        vert_align: str = "TOP",
        horz_align: str = "LEFT",
        vert_offset: int = 0,
        horz_offset: int = 0,
        out_margin_left: int = 0,
        out_margin_right: int = 0,
        out_margin_top: int = 0,
        out_margin_bottom: int = 0,
    ) -> ShapeXml:
        self._ensure_editable_sections()
        supported_kinds = {"rect", "ellipse", "arc", "polygon", "curve", "connectLine", "line", "textart", "container"}
        if kind not in supported_kinds:
            supported = ", ".join(sorted(supported_kinds))
            raise ValueError(f"append_shape() supports kind values: {supported}")
        if width < 1 or height < 1:
            raise ValueError("width and height must be positive.")

        shape = etree.Element(qname("hp", kind))
        shape.set("id", self._next_control_number(_graphic_attribute_xpath("id")))
        shape.set("zOrder", self._next_control_number(_graphic_attribute_xpath("zOrder")))
        shape.set("numberingType", "PICTURE")
        shape.set("textWrap", "TOP_AND_BOTTOM")
        shape.set("textFlow", "BOTH_SIDES")
        shape.set("lock", "0")
        shape.set("dropcapstyle", "None")
        shape.set("href", "")
        shape.set("groupLevel", "0")
        shape.set("instid", self._next_control_number(_graphic_attribute_xpath("instid")))
        if kind not in {"line", "connectLine"}:
            shape.set("ratio", "0")

        offset = etree.SubElement(shape, qname("hp", "offset"))
        offset.set("x", "0")
        offset.set("y", "0")

        original_size = etree.SubElement(shape, qname("hp", "orgSz"))
        original_size.set("width", str(width))
        original_size.set("height", str(height))

        current_size = etree.SubElement(shape, qname("hp", "curSz"))
        current_size.set("width", str(width))
        current_size.set("height", str(height))

        flip = etree.SubElement(shape, qname("hp", "flip"))
        flip.set("horizontal", "0")
        flip.set("vertical", "0")

        rotation = etree.SubElement(shape, qname("hp", "rotationInfo"))
        rotation.set("angle", "0")
        rotation.set("centerX", str(width // 2))
        rotation.set("centerY", str(height // 2))
        rotation.set("rotateimage", "1")

        rendering = etree.SubElement(shape, qname("hp", "renderingInfo"))
        self._append_identity_matrix(rendering, "transMatrix")
        self._append_identity_matrix(rendering, "scaMatrix")
        self._append_identity_matrix(rendering, "rotMatrix")

        line_shape = etree.SubElement(shape, qname("hp", "lineShape"))
        line_shape.set("color", line_color)
        line_shape.set("width", "33")
        line_shape.set("style", "SOLID")
        line_shape.set("endCap", "FLAT")
        line_shape.set("headStyle", "NORMAL")
        line_shape.set("tailStyle", "NORMAL")
        line_shape.set("headfill", "1")
        line_shape.set("tailfill", "1")
        line_shape.set("headSz", "MEDIUM_MEDIUM")
        line_shape.set("tailSz", "MEDIUM_MEDIUM")
        line_shape.set("outlineStyle", "NORMAL")
        line_shape.set("alpha", "0")

        if kind not in {"line", "connectLine"}:
            fill_brush = etree.SubElement(shape, qname("hc", "fillBrush"))
            win_brush = etree.SubElement(fill_brush, qname("hc", "winBrush"))
            win_brush.set("faceColor", fill_color)
            win_brush.set("hatchColor", line_color)
            win_brush.set("alpha", "0")

        shadow = etree.SubElement(shape, qname("hp", "shadow"))
        shadow.set("type", "NONE")
        shadow.set("color", "#B2B2B2")
        shadow.set("offsetX", "0")
        shadow.set("offsetY", "0")
        shadow.set("alpha", "0")

        if kind in {"rect", "ellipse", "arc", "polygon", "curve", "textart", "container"}:
            draw_text = etree.SubElement(shape, qname("hp", "drawText"))
            draw_text.set("lastWidth", str(width))
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
            paragraph.set("id", self._next_control_number(".//hp:p/@id"))
            paragraph.set("paraPrIDRef", "0")
            paragraph.set("styleIDRef", "0")
            paragraph.set("pageBreak", "0")
            paragraph.set("columnBreak", "0")
            paragraph.set("merged", "0")

            run = etree.SubElement(paragraph, qname("hp", "run"))
            run.set("charPrIDRef", char_pr_id or "0")
            text_node = etree.SubElement(run, qname("hp", "t"))
            text_node.text = text

            text_margin = etree.SubElement(draw_text, qname("hp", "textMargin"))
            text_margin.set("left", "283")
            text_margin.set("right", "283")
            text_margin.set("top", "283")
            text_margin.set("bottom", "283")

        if kind in {"line", "connectLine"}:
            points = ((0, 0), (width, height))
        else:
            points = ((0, 0), (width, 0), (width, height), (0, height))

        for index, (x, y) in enumerate(points):
            point = etree.SubElement(shape, qname("hc", f"pt{index}"))
            point.set("x", str(x))
            point.set("y", str(y))

        size = etree.SubElement(shape, qname("hp", "sz"))
        size.set("width", str(width))
        size.set("widthRelTo", "ABSOLUTE")
        size.set("height", str(height))
        size.set("heightRelTo", "ABSOLUTE")
        size.set("protect", "0")

        position = etree.SubElement(shape, qname("hp", "pos"))
        position.set("treatAsChar", "1" if treat_as_char else "0")
        position.set("affectLSpacing", "0")
        position.set("flowWithText", "1")
        position.set("allowOverlap", "0")
        position.set("holdAnchorAndSO", "0")
        position.set("vertRelTo", "PARA")
        position.set("horzRelTo", "COLUMN")
        position.set("vertAlign", "TOP")
        position.set("horzAlign", "LEFT")
        position.set("vertOffset", "0")
        position.set("horzOffset", "0")

        out_margin = etree.SubElement(shape, qname("hp", "outMargin"))
        for key in ("left", "right", "top", "bottom"):
            out_margin.set(key, "0")
        _set_graphic_layout(
            shape,
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
        _set_margin_values(
            shape,
            "./hp:outMargin",
            left=out_margin_left,
            right=out_margin_right,
            top=out_margin_top,
            bottom=out_margin_bottom,
        )

        if shape_comment:
            comment = etree.SubElement(shape, qname("hp", "shapeComment"))
            comment.text = shape_comment

        section = self._append_run_child(section_index, paragraph_index, shape, char_pr_id=char_pr_id)
        return ShapeXml(self, section, shape)

    def append_ole(
        self,
        name: str,
        data: bytes,
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
        media_type: str = "application/ole",
        manifest_id: str | None = None,
        width: int = 42001,
        height: int = 13501,
        original_width: int | None = None,
        original_height: int | None = None,
        current_width: int = 0,
        current_height: int = 0,
        treat_as_char: bool = False,
        shape_comment: str | None = None,
        object_type: str = "EMBEDDED",
        has_moniker: bool = False,
        draw_aspect: str = "CONTENT",
        eq_baseline: int = 0,
        text_wrap: str = "SQUARE",
        text_flow: str = "BOTH_SIDES",
        affect_line_spacing: bool = False,
        flow_with_text: bool = True,
        allow_overlap: bool = False,
        hold_anchor_and_so: bool = False,
        vert_rel_to: str = "PARA",
        horz_rel_to: str = "COLUMN",
        vert_align: str = "TOP",
        horz_align: str = "LEFT",
        vert_offset: int = 0,
        horz_offset: int = 0,
        out_margin_left: int = 0,
        out_margin_right: int = 0,
        out_margin_top: int = 0,
        out_margin_bottom: int = 0,
    ) -> OleXml:
        self._ensure_editable_sections()
        if width < 1 or height < 1:
            raise ValueError("width and height must be positive.")

        resolved_manifest_id = manifest_id or self.content_hpf.next_manifest_id(PurePosixPath(name).stem or "ole")
        self.add_or_replace_binary(
            name,
            data,
            media_type=media_type,
            manifest_id=resolved_manifest_id,
            is_embedded=False,
        )

        resolved_original_width = original_width or width
        resolved_original_height = original_height or height

        ole = etree.Element(qname("hp", "ole"))
        ole.set("id", self._next_control_number(_graphic_attribute_xpath("id")))
        ole.set("zOrder", self._next_control_number(_graphic_attribute_xpath("zOrder")))
        ole.set("numberingType", "PICTURE")
        ole.set("textWrap", text_wrap)
        ole.set("textFlow", text_flow)
        ole.set("lock", "0")
        ole.set("dropcapstyle", "None")
        ole.set("href", "")
        ole.set("groupLevel", "0")
        ole.set("instid", self._next_control_number(_graphic_attribute_xpath("instid")))
        ole.set("objectType", object_type)
        ole.set("binaryItemIDRef", resolved_manifest_id)
        ole.set("hasMoniker", "1" if has_moniker else "0")
        ole.set("drawAspect", draw_aspect)
        ole.set("eqBaseLine", str(eq_baseline))

        offset = etree.SubElement(ole, qname("hp", "offset"))
        offset.set("x", "0")
        offset.set("y", "0")

        original_size = etree.SubElement(ole, qname("hp", "orgSz"))
        original_size.set("width", str(resolved_original_width))
        original_size.set("height", str(resolved_original_height))

        current_size = etree.SubElement(ole, qname("hp", "curSz"))
        current_size.set("width", str(current_width))
        current_size.set("height", str(current_height))

        flip = etree.SubElement(ole, qname("hp", "flip"))
        flip.set("horizontal", "0")
        flip.set("vertical", "0")

        rotation = etree.SubElement(ole, qname("hp", "rotationInfo"))
        rotation.set("angle", "0")
        rotation.set("centerX", str(width // 2))
        rotation.set("centerY", str(height // 2))
        rotation.set("rotateimage", "1")

        rendering = etree.SubElement(ole, qname("hp", "renderingInfo"))
        self._append_identity_matrix(rendering, "transMatrix")
        self._append_identity_matrix(rendering, "scaMatrix")
        self._append_identity_matrix(rendering, "rotMatrix")

        extent = etree.SubElement(ole, qname("hc", "extent"))
        extent.set("x", str(width))
        extent.set("y", str(height))

        line_shape = etree.SubElement(ole, qname("hp", "lineShape"))
        line_shape.set("color", "#000000")
        line_shape.set("width", "0")
        line_shape.set("style", "NONE")
        line_shape.set("endCap", "ROUND")
        line_shape.set("headStyle", "NORMAL")
        line_shape.set("tailStyle", "NORMAL")
        line_shape.set("headfill", "0")
        line_shape.set("tailfill", "0")
        line_shape.set("headSz", "SMALL_SMALL")
        line_shape.set("tailSz", "SMALL_SMALL")
        line_shape.set("outlineStyle", "OUTER")
        line_shape.set("alpha", "0")

        size = etree.SubElement(ole, qname("hp", "sz"))
        size.set("width", str(width))
        size.set("widthRelTo", "ABSOLUTE")
        size.set("height", str(height))
        size.set("heightRelTo", "ABSOLUTE")
        size.set("protect", "0")

        position = etree.SubElement(ole, qname("hp", "pos"))
        position.set("treatAsChar", "1" if treat_as_char else "0")
        position.set("affectLSpacing", "0")
        position.set("flowWithText", "1")
        position.set("allowOverlap", "0")
        position.set("holdAnchorAndSO", "0")
        position.set("vertRelTo", "PARA")
        position.set("horzRelTo", "COLUMN")
        position.set("vertAlign", "TOP")
        position.set("horzAlign", "LEFT")
        position.set("vertOffset", "0")
        position.set("horzOffset", "0")

        out_margin = etree.SubElement(ole, qname("hp", "outMargin"))
        for key in ("left", "right", "top", "bottom"):
            out_margin.set(key, "0")

        _set_graphic_layout(
            ole,
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
        _set_margin_values(
            ole,
            "./hp:outMargin",
            left=out_margin_left,
            right=out_margin_right,
            top=out_margin_top,
            bottom=out_margin_bottom,
        )

        if shape_comment:
            comment = etree.SubElement(ole, qname("hp", "shapeComment"))
            comment.text = shape_comment

        section = self._append_run_child(section_index, paragraph_index, ole, char_pr_id=char_pr_id)
        return OleXml(self, section, ole)

    def append_bookmark(
        self,
        name: str,
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> BookmarkXml:
        self._ensure_editable_sections()
        paragraph = self._resolve_paragraph_for_insert(section_index=section_index, paragraph_index=paragraph_index)
        run = self._append_run(paragraph, char_pr_id=char_pr_id)
        ctrl = etree.SubElement(run, qname("hp", "ctrl"))
        bookmark = etree.SubElement(ctrl, qname("hp", "bookmark"))
        bookmark.set("name", name)
        etree.SubElement(run, qname("hp", "t")).text = ""
        section = self.sections[section_index]
        section.mark_modified()
        return BookmarkXml(self, section, bookmark)

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
    ) -> FieldXml:
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
        return FieldXml(self, section, field_begin)

    def append_hyperlink(
        self,
        target: str,
        *,
        display_text: str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> FieldXml:
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
    ) -> FieldXml:
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
    ) -> FieldXml:
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
    ) -> FieldXml:
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

    def to_hwp_document(self, *, converter=None):
        from .hancom_document import HancomDocument

        return HancomDocument.from_hwpx_document(self, converter=converter).to_hwp_document(converter=converter)

    def to_hancom_document(self, *, converter=None):
        from .hancom_document import HancomDocument

        return HancomDocument.from_hwpx_document(self, converter=converter)

    def save_as_hwp(self, path: str | os.PathLike[str], *, converter=None, force_refresh: bool = True) -> Path:
        return self.to_hwp_document(converter=converter).save(path)

    def bridge(self, *, converter=None):
        from .bridge import HwpHwpxBridge

        return HwpHwpxBridge.from_hwpx(self, converter=converter)

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
        temp_dir = _new_temp_dir("jakal_hwpx_")
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
            if (field.field_type or "").upper() == _MEMO_CARRIER_FIELD_TYPE:
                parameters = field.parameter_map()
                if not _bool_like_literal_is_valid(parameters.get("Visible")):
                    errors.append(
                        _issue(
                            "metadata_schema",
                            "JAKAL_MEMO carrier Visible parameter must be a bool-like literal.",
                            part_path=getattr(field.section, "path", None),
                            xpath=".//hp:fieldBegin[@type='JAKAL_MEMO']",
                        )
                    )
                order_value = parameters.get("Order")
                if order_value not in (None, ""):
                    try:
                        int(order_value)
                    except (TypeError, ValueError):
                        errors.append(
                            _issue(
                                "metadata_schema",
                                "JAKAL_MEMO carrier Order parameter must be an integer.",
                                part_path=getattr(field.section, "path", None),
                                xpath=".//hp:fieldBegin[@type='JAKAL_MEMO']",
                            )
                        )

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
            for binary_ref in section.root_element.xpath(".//hp:ole/@binaryItemIDRef", namespaces=NS):
                if binary_ref not in manifest_ids:
                    errors.append(
                        _issue(
                            "binary_reference",
                            f"ole references unknown manifest id {binary_ref}",
                            part_path=section.path,
                            section_index=section_index,
                        )
                    )
            for chart in self.charts(section_index=section_index):
                chart_ref = chart.chart_part_path
                if not chart_ref:
                    errors.append(
                        _issue(
                            "chart_reference",
                            "chart is missing chartIDRef",
                            part_path=section.path,
                            section_index=section_index,
                        )
                    )
                    continue
                if _normalize_path(chart_ref) not in self._parts:
                    errors.append(
                        _issue(
                            "chart_reference",
                            f"chart references missing chart part: {chart_ref}",
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
        temp_dir = _new_temp_dir("jakal_hwpx_reopen_")
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

    def strict_lint_errors(self) -> list[ValidationIssue]:
        errors = list(self.reference_validation_errors())

        ordered_paths = self.list_part_paths()
        if ordered_paths and ordered_paths[0] != "mimetype":
            errors.append(
                _issue(
                    "mimetype_packaging",
                    "mimetype must be the first package entry",
                    part_path="mimetype",
                )
            )

        mimetype_info = self._zip_infos.get("mimetype")
        if mimetype_info is not None and mimetype_info.compress_type != zipfile.ZIP_STORED:
            errors.append(
                _issue(
                    "mimetype_packaging",
                    "mimetype must be stored without compression",
                    part_path="mimetype",
                )
            )

        if not self.is_distribution_protected:
            style_ids = {style.style_id for style in self.styles() if style.style_id is not None}
            para_pr_ids = {style.style_id for style in self.paragraph_styles() if style.style_id is not None}
            char_pr_ids = {style.style_id for style in self.character_styles() if style.style_id is not None}
            memo_shape_ids = {memo_shape.memo_shape_id for memo_shape in self.memo_shapes() if memo_shape.memo_shape_id is not None}
            memo_properties_nodes = self.header.root_element.xpath(".//hh:memoProperties", namespaces=NS)
            for memo_properties in memo_properties_nodes:
                memo_pr_nodes = memo_properties.xpath("./hh:memoPr", namespaces=NS)
                expected_item_count = memo_properties.get("itemCnt")
                if expected_item_count not in (None, "", str(len(memo_pr_nodes))):
                    errors.append(
                        _issue(
                            "memo_reference",
                            f"memoProperties itemCnt={expected_item_count} does not match memoPr count={len(memo_pr_nodes)}",
                            part_path="Contents/header.xml",
                        )
                    )
                seen_memo_shape_ids: set[str] = set()
                for memo_pr in memo_pr_nodes:
                    memo_shape_id = memo_pr.get("id")
                    if not memo_shape_id:
                        continue
                    if memo_shape_id in seen_memo_shape_ids:
                        errors.append(
                            _issue(
                                "memo_reference",
                                f"Duplicate hh:memoPr id: {memo_shape_id}",
                                part_path="Contents/header.xml",
                            )
                        )
                    seen_memo_shape_ids.add(memo_shape_id)
            for section_index, section in enumerate(self.sections):
                sec_pr_nodes = section.root_element.xpath("./hp:p[1]//hp:secPr[1]", namespaces=NS)
                if sec_pr_nodes:
                    memo_shape_id = sec_pr_nodes[0].get("memoShapeIDRef")
                    if memo_shape_id not in (None, "", "0") and memo_shape_id not in memo_shape_ids:
                        errors.append(
                            _issue(
                                "memo_reference",
                                f"memoShapeIDRef points to unknown hh:memoPr id {memo_shape_id}",
                                part_path=section.path,
                                section_index=section_index,
                            )
                        )
                paragraphs = section.root_element.xpath(".//hp:p", namespaces=NS)
                for paragraph_index, paragraph in enumerate(paragraphs):
                    if style_ids and not paragraph.get("styleIDRef"):
                        errors.append(
                            _issue(
                                "style_binding",
                                "paragraph is missing styleIDRef",
                                part_path=section.path,
                                section_index=section_index,
                                paragraph_index=paragraph_index,
                            )
                        )
                    if para_pr_ids and not paragraph.get("paraPrIDRef"):
                        errors.append(
                            _issue(
                                "style_binding",
                                "paragraph is missing paraPrIDRef",
                                part_path=section.path,
                                section_index=section_index,
                                paragraph_index=paragraph_index,
                            )
                        )
                    for run_index, run in enumerate(paragraph.xpath("./hp:run", namespaces=NS)):
                        if char_pr_ids and not run.get("charPrIDRef"):
                            errors.append(
                                _issue(
                                    "style_binding",
                                    "run is missing charPrIDRef",
                                    part_path=section.path,
                                    section_index=section_index,
                                    paragraph_index=paragraph_index,
                                    context=f"run[{run_index}]",
                                )
                            )
                    form_nodes = paragraph.xpath(" | ".join(f".//hp:{tag}" for tag in _FORM_TAGS), namespaces=NS)
                    for form_index, form in enumerate(form_nodes):
                        kind = etree.QName(form).localname
                        command = form.get("command", "")
                        if command.startswith(_FORM_METADATA_COMMAND_PREFIX):
                            payload_text = command[len(_FORM_METADATA_COMMAND_PREFIX) :]
                            try:
                                payload = json.loads(payload_text)
                            except json.JSONDecodeError as exc:
                                errors.append(
                                    _issue(
                                        "metadata_schema",
                                        f"{kind} command metadata must be valid JSON: {exc}",
                                        part_path=section.path,
                                        section_index=section_index,
                                        paragraph_index=paragraph_index,
                                        context=f"form[{form_index}]",
                                    )
                                )
                                payload = None
                            if payload is not None and not isinstance(payload, dict):
                                errors.append(
                                    _issue(
                                        "metadata_schema",
                                        f"{kind} command metadata must decode to a JSON object.",
                                        part_path=section.path,
                                        section_index=section_index,
                                        paragraph_index=paragraph_index,
                                        context=f"form[{form_index}]",
                                    )
                                )
                            elif isinstance(payload, dict):
                                if "placeholder" in payload and not isinstance(payload["placeholder"], str):
                                    errors.append(
                                        _issue(
                                            "metadata_schema",
                                            f"{kind} placeholder metadata must be a string.",
                                            part_path=section.path,
                                            section_index=section_index,
                                            paragraph_index=paragraph_index,
                                            context=f"form[{form_index}]",
                                        )
                                    )
                                if "items" in payload and not isinstance(payload["items"], list):
                                    errors.append(
                                        _issue(
                                            "metadata_schema",
                                            f"{kind} items metadata must be a list.",
                                            part_path=section.path,
                                            section_index=section_index,
                                            paragraph_index=paragraph_index,
                                            context=f"form[{form_index}]",
                                        )
                                    )
                                elif isinstance(payload.get("items"), list) and any(
                                    not isinstance(item, str) for item in payload["items"]
                                ):
                                    errors.append(
                                        _issue(
                                            "metadata_schema",
                                            f"{kind} items metadata entries must be strings.",
                                            part_path=section.path,
                                            section_index=section_index,
                                            paragraph_index=paragraph_index,
                                            context=f"form[{form_index}]",
                                        )
                                    )
                                if "label" in payload and not isinstance(payload["label"], str):
                                    errors.append(
                                        _issue(
                                            "metadata_schema",
                                            f"{kind} label metadata must be a string.",
                                            part_path=section.path,
                                            section_index=section_index,
                                            paragraph_index=paragraph_index,
                                            context=f"form[{form_index}]",
                                        )
                                    )
                        for required_xpath, label in (
                            ("./hp:sz", "hp:sz"),
                            ("./hp:pos", "hp:pos"),
                            ("./hp:outMargin", "hp:outMargin"),
                        ):
                            if not form.xpath(required_xpath, namespaces=NS):
                                errors.append(
                                    _issue(
                                        "control_subtree",
                                        f"{kind} is missing {label}",
                                        part_path=section.path,
                                        section_index=section_index,
                                        paragraph_index=paragraph_index,
                                        context=f"form[{form_index}]",
                                    )
                                )
                        if form.xpath("./hp:caption", namespaces=NS) and not form.xpath("./hp:caption/hp:subList", namespaces=NS):
                            errors.append(
                                _issue(
                                    "control_subtree",
                                    f"{kind} caption is missing hp:subList",
                                    part_path=section.path,
                                    section_index=section_index,
                                    paragraph_index=paragraph_index,
                                    context=f"form[{form_index}]",
                                )
                            )
                        if kind != "scrollBar" and not form.xpath("./hp:formCharPr", namespaces=NS):
                            errors.append(
                                _issue(
                                    "control_subtree",
                                    f"{kind} is missing hp:formCharPr",
                                    part_path=section.path,
                                    section_index=section_index,
                                    paragraph_index=paragraph_index,
                                    context=f"form[{form_index}]",
                                )
                            )
                        if kind == "edit" and not form.xpath("./hp:text", namespaces=NS):
                            errors.append(
                                _issue(
                                    "control_subtree",
                                    "edit is missing hp:text",
                                    part_path=section.path,
                                    section_index=section_index,
                                    paragraph_index=paragraph_index,
                                    context=f"form[{form_index}]",
                                )
                            )
                    memo_nodes = paragraph.xpath(".//hp:hiddenComment", namespaces=NS)
                    for memo_index, memo in enumerate(memo_nodes):
                        if not memo.xpath("./hp:subList", namespaces=NS):
                            errors.append(
                                _issue(
                                    "control_subtree",
                                    "hiddenComment is missing hp:subList",
                                    part_path=section.path,
                                    section_index=section_index,
                                    paragraph_index=paragraph_index,
                                    context=f"memo[{memo_index}]",
                                )
                            )
                    chart_nodes = paragraph.xpath(".//hp:chart", namespaces=NS)
                    for chart_index, chart in enumerate(chart_nodes):
                        for required_xpath, label in (
                            ("./hp:sz", "hp:sz"),
                            ("./hp:pos", "hp:pos"),
                            ("./hp:outMargin", "hp:outMargin"),
                        ):
                            if not chart.xpath(required_xpath, namespaces=NS):
                                errors.append(
                                    _issue(
                                        "control_subtree",
                                        f"chart is missing {label}",
                                        part_path=section.path,
                                        section_index=section_index,
                                        paragraph_index=paragraph_index,
                                        context=f"chart[{chart_index}]",
                                    )
                                )
                        case = chart.getparent()
                        switch = case.getparent() if case is not None else None
                        if (
                            case is None
                            or etree.QName(case).localname != "case"
                            or switch is None
                            or etree.QName(switch).localname != "switch"
                        ):
                            errors.append(
                                _issue(
                                    "chart_closure",
                                    "native hp:chart control must be wrapped by hp:switch/hp:case",
                                    part_path=section.path,
                                    section_index=section_index,
                                    paragraph_index=paragraph_index,
                                    context=f"chart[{chart_index}]",
                                )
                            )
                        elif not switch.xpath("./hp:default/hp:ole", namespaces=NS):
                            errors.append(
                                _issue(
                                    "chart_closure",
                                    "native hp:chart control is missing hp:default/hp:ole fallback",
                                    part_path=section.path,
                                    section_index=section_index,
                                    paragraph_index=paragraph_index,
                                    context=f"chart[{chart_index}]",
                                )
                            )
                        elif not switch.xpath("./hp:default/hp:ole/hp:rotationInfo", namespaces=NS):
                            errors.append(
                                _issue(
                                    "chart_closure",
                                    "native hp:chart fallback hp:ole is missing hp:rotationInfo",
                                    part_path=section.path,
                                    section_index=section_index,
                                    paragraph_index=paragraph_index,
                                    context=f"chart[{chart_index}]",
                                )
                            )

        binary_part_paths = {
            path
            for path, part in self._parts.items()
            if path.startswith("BinData/") and isinstance(part, BinaryDataPart)
        }
        manifest_binary_items: dict[str, etree._Element] = {}
        manifest_binary_hrefs: set[str] = set()
        for item in self.content_hpf.manifest_items():
            item_id = item.get("id")
            href = item.get("href")
            if not item_id or not href:
                continue
            normalized_href = _normalize_path(href)
            if not normalized_href.startswith("BinData/"):
                continue
            manifest_binary_items[item_id] = item
            manifest_binary_hrefs.add(normalized_href)
            expected_media_type = _expected_binary_media_type(normalized_href)
            actual_media_type = item.get("media-type")
            if expected_media_type and actual_media_type and actual_media_type.lower() != expected_media_type.lower():
                errors.append(
                    _issue(
                        "binary_media_type",
                        f"manifest media-type {actual_media_type} does not match expected {expected_media_type} for {normalized_href}",
                        part_path="Contents/content.hpf",
                        context=f"manifest id={item_id}",
                    )
                )

        for part_path in sorted(binary_part_paths - manifest_binary_hrefs):
            errors.append(
                _issue(
                    "binary_closure",
                    f"BinData part is not referenced by content.hpf manifest: {part_path}",
                    part_path=part_path,
                )
            )

        referenced_binary_item_ids: set[str] = set()
        for section in self.sections:
            values = section.root_element.xpath(
                ".//hc:img/@binaryItemIDRef | " + _graphic_attribute_xpath("binaryItemIDRef"),
                namespaces=NS,
            )
            referenced_binary_item_ids.update(str(value) for value in values if str(value).strip())

        for item_id, item in manifest_binary_items.items():
            href = _normalize_path(item.get("href") or "")
            if item_id not in referenced_binary_item_ids:
                errors.append(
                    _issue(
                        "binary_closure",
                        f"manifest binary item is not referenced by any control: {item_id}",
                        part_path="Contents/content.hpf",
                        context=f"href={href}",
                    )
                )

        for item_id in sorted(referenced_binary_item_ids):
            item = manifest_binary_items.get(item_id)
            if item is None:
                continue
            href = _normalize_path(item.get("href") or "")
            if not href.startswith("BinData/"):
                errors.append(
                    _issue(
                        "binary_closure",
                        f"binaryItemIDRef {item_id} must resolve to a BinData part, got {href}",
                        part_path="Contents/content.hpf",
                        context=f"manifest id={item_id}",
                    )
                )

        referenced_chart_parts = {
            _normalize_path(chart.chart_part_path)
            for chart in self.charts()
            if chart.chart_part_path
        }
        chart_part_paths = {
            path
            for path, part in self._parts.items()
            if path.startswith("Chart/") and path.endswith(".xml")
        }
        for chart_path in sorted(chart_part_paths - referenced_chart_parts):
            errors.append(
                _issue(
                    "chart_closure",
                    f"Chart part is not referenced by any native hp:chart control: {chart_path}",
                    part_path=chart_path,
                )
            )
        for chart_path in sorted(referenced_chart_parts):
            part = self._parts.get(chart_path)
            if part is None:
                continue
            if not isinstance(part, XmlPart):
                errors.append(
                    _issue(
                        "chart_closure",
                        f"Chart part must be XML: {chart_path}",
                        part_path=chart_path,
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
        _invalidate_paragraph_layout(paragraph)
        return run

    def _default_char_pr_id(self, paragraph: etree._Element) -> str:
        runs = paragraph.xpath("./hp:run", namespaces=NS)
        for run in runs:
            value = run.get("charPrIDRef")
            if value is not None:
                return value
        return "0"

    def _append_control(
        self,
        section_index: int,
        paragraph_index: int | None,
        control_element: etree._Element,
        *,
        char_pr_id: str | None = None,
    ):
        paragraph = self._resolve_paragraph_for_insert(section_index=section_index, paragraph_index=paragraph_index)
        run = self._append_run(paragraph, char_pr_id=char_pr_id)
        ctrl = etree.SubElement(run, qname("hp", "ctrl"))
        ctrl.append(control_element)
        section = self.sections[section_index]
        section.mark_modified()
        return section

    def _append_run_child(
        self,
        section_index: int,
        paragraph_index: int | None,
        element: etree._Element,
        *,
        char_pr_id: str | None = None,
    ):
        paragraph = self._resolve_paragraph_for_insert(section_index=section_index, paragraph_index=paragraph_index)
        run = self._append_run(paragraph, char_pr_id=char_pr_id)
        run.append(element)
        section = self.sections[section_index]
        section.mark_modified()
        return section

    def _build_header_footer_block(
        self,
        kind: str,
        *,
        text: str,
        char_pr_id: str | None = None,
        apply_page_type: str = "BOTH",
    ) -> etree._Element:
        block = etree.Element(qname("hp", kind))
        block.set("id", self._next_control_number(".//hp:header/@id | .//hp:footer/@id"))
        block.set("applyPageType", apply_page_type)
        block.append(self._build_sublist_with_text(text, char_pr_id=char_pr_id, vertical_align="TOP"))
        return block

    def _build_sublist_with_text(
        self,
        text: str,
        *,
        char_pr_id: str | None = None,
        vertical_align: str = "CENTER",
    ) -> etree._Element:
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

        paragraph = etree.SubElement(sublist, qname("hp", "p"))
        paragraph.set("id", self._next_control_number(".//hp:p/@id"))
        paragraph.set("paraPrIDRef", "0")
        paragraph.set("styleIDRef", "0")
        paragraph.set("pageBreak", "0")
        paragraph.set("columnBreak", "0")
        paragraph.set("merged", "0")

        run = etree.SubElement(paragraph, qname("hp", "run"))
        run.set("charPrIDRef", char_pr_id or "0")
        text_node = etree.SubElement(run, qname("hp", "t"))
        text_node.text = text
        return sublist

    def _ensure_header_collection(self, expression: str, tag_name: str) -> etree._Element:
        nodes = self.header.root_element.xpath(expression, namespaces=NS)
        if nodes:
            return nodes[0]
        ref_lists = self.header.root_element.xpath("./hh:refList", namespaces=NS)
        if ref_lists:
            ref_list = ref_lists[0]
        else:
            ref_list = etree.SubElement(self.header.root_element, qname("hh", "refList"))
        collection = etree.SubElement(ref_list, qname("hh", tag_name))
        collection.set("itemCnt", "0")
        self.header.mark_modified()
        return collection

    def _clone_header_template(self, expression: str, *, template_id: str | None = None) -> etree._Element:
        nodes = self.header.root_element.xpath(expression, namespaces=NS)
        if not nodes:
            raise ValueError(f"Header part does not contain a template for {expression}.")
        if template_id is not None:
            for node in nodes:
                if node.get("id") == template_id:
                    return etree.fromstring(etree.tostring(node))
            raise KeyError(template_id)
        return etree.fromstring(etree.tostring(nodes[0]))

    def _next_header_numeric_id(self, expression: str) -> str:
        highest = -1
        for value in self.header.root_element.xpath(expression, namespaces=NS):
            try:
                highest = max(highest, int(value))
            except (TypeError, ValueError):
                continue
        return str(highest + 1)

    def _append_identity_matrix(self, parent: etree._Element, name: str) -> etree._Element:
        return self._append_matrix(parent, name, e1="1", e2="0", e3="0", e4="0", e5="1", e6="0")

    def _append_matrix(self, parent: etree._Element, name: str, **entries: str) -> etree._Element:
        matrix = etree.SubElement(parent, qname("hc", name))
        for key, value in entries.items():
            matrix.set(key, value)
        return matrix

    def _build_chart_fallback_ole(
        self,
        *,
        control_id: str,
        z_order: str,
        width: int,
        height: int,
        text_wrap: str,
        text_flow: str,
        treat_as_char: bool,
        affect_line_spacing: bool,
        flow_with_text: bool,
        allow_overlap: bool,
        hold_anchor_and_so: bool,
        vert_rel_to: str,
        horz_rel_to: str,
        vert_align: str,
        horz_align: str,
        vert_offset: int,
        horz_offset: int,
        out_margin_left: int,
        out_margin_right: int,
        out_margin_top: int,
        out_margin_bottom: int,
    ) -> etree._Element:
        manifest_id = self._next_chart_ole_manifest_id()
        self.add_or_replace_binary(
            f"{manifest_id}.ole",
            _default_chart_fallback_ole_bytes(),
            media_type="application/ole",
            manifest_id=manifest_id,
            is_embedded=False,
        )

        ole = etree.Element(qname("hp", "ole"))
        ole.set("id", control_id)
        ole.set("zOrder", z_order)
        ole.set("numberingType", "PICTURE")
        ole.set("textWrap", text_wrap)
        ole.set("textFlow", text_flow)
        ole.set("lock", "0")
        ole.set("dropcapstyle", "None")
        ole.set("href", "")
        ole.set("groupLevel", "0")
        ole.set("instid", "0")
        ole.set("objectType", "UNKNOWN")
        ole.set("binaryItemIDRef", manifest_id)
        ole.set("hasMoniker", "0")
        ole.set("drawAspect", "CONTENT")
        ole.set("eqBaseLine", "0")

        offset = etree.SubElement(ole, qname("hp", "offset"))
        offset.set("x", "0")
        offset.set("y", "0")

        original_size = etree.SubElement(ole, qname("hp", "orgSz"))
        original_size.set("width", "7200")
        original_size.set("height", "7200")

        current_size = etree.SubElement(ole, qname("hp", "curSz"))
        current_size.set("width", "0")
        current_size.set("height", "0")

        flip = etree.SubElement(ole, qname("hp", "flip"))
        flip.set("horizontal", "0")
        flip.set("vertical", "0")

        rotation = etree.SubElement(ole, qname("hp", "rotationInfo"))
        rotation.set("angle", "0")
        rotation.set("centerX", "0")
        rotation.set("centerY", "0")
        rotation.set("rotateimage", "1")

        rendering = etree.SubElement(ole, qname("hp", "renderingInfo"))
        self._append_identity_matrix(rendering, "transMatrix")
        self._append_identity_matrix(rendering, "scaMatrix")
        self._append_identity_matrix(rendering, "rotMatrix")

        extent = etree.SubElement(ole, qname("hc", "extent"))
        extent.set("x", "7200")
        extent.set("y", "7200")

        line_shape = etree.SubElement(ole, qname("hp", "lineShape"))
        line_shape.set("color", "#000000")
        line_shape.set("width", "0")
        line_shape.set("style", "NONE")
        line_shape.set("endCap", "ROUND")
        line_shape.set("headStyle", "NORMAL")
        line_shape.set("tailStyle", "NORMAL")
        line_shape.set("headfill", "0")
        line_shape.set("tailfill", "0")
        line_shape.set("headSz", "SMALL_SMALL")
        line_shape.set("tailSz", "SMALL_SMALL")
        line_shape.set("outlineStyle", "NORMAL")
        line_shape.set("alpha", "0")

        size = etree.SubElement(ole, qname("hp", "sz"))
        size.set("width", str(width))
        size.set("widthRelTo", "ABSOLUTE")
        size.set("height", str(height))
        size.set("heightRelTo", "ABSOLUTE")
        size.set("protect", "0")

        position = etree.SubElement(ole, qname("hp", "pos"))
        position.set("treatAsChar", "1" if treat_as_char else "0")
        position.set("affectLSpacing", "0")
        position.set("flowWithText", "1")
        position.set("allowOverlap", "0")
        position.set("holdAnchorAndSO", "0")
        position.set("vertRelTo", "PARA")
        position.set("horzRelTo", "COLUMN")
        position.set("vertAlign", "TOP")
        position.set("horzAlign", "LEFT")
        position.set("vertOffset", "0")
        position.set("horzOffset", "0")
        _set_graphic_layout(
            ole,
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

        out_margin = etree.SubElement(ole, qname("hp", "outMargin"))
        out_margin.set("left", "0")
        out_margin.set("right", "0")
        out_margin.set("top", "0")
        out_margin.set("bottom", "0")
        _set_margin_values(
            ole,
            "./hp:outMargin",
            left=out_margin_left,
            right=out_margin_right,
            top=out_margin_top,
            bottom=out_margin_bottom,
        )
        return ole

    def _next_chart_ole_manifest_id(self) -> str:
        existing = {
            item.get("id", "")
            for item in self.content_hpf.manifest_items()
            if (item.get("href") or "").startswith("BinData/") and item.get("id")
        }
        index = 1
        while f"ole{index}" in existing:
            index += 1
        return f"ole{index}"

    def _next_chart_part_path(self) -> str:
        highest = 0
        pattern = re.compile(r"^Chart/chart(\d+)\.xml$")
        for path in self.list_part_paths():
            match = pattern.match(path)
            if match is None:
                continue
            highest = max(highest, int(match.group(1)))
        return f"Chart/chart{highest + 1}.xml"

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
        if temp_dir.exists() and isinstance(last_error, PermissionError):
            return
        if temp_dir.exists() and last_error is not None:
            raise last_error
