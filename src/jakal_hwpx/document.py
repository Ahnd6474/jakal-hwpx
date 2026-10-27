from __future__ import annotations

import io
import mimetypes
import os
import re
import shutil
import tempfile
import time
import zipfile
from collections import Counter
from copy import deepcopy
from dataclasses import replace
from pathlib import Path, PurePosixPath
from lxml import etree

from ._hwpx_archive import read_archive_state, write_archive_state
from ._hwpx_common import (  # noqa: F401 - compatibility re-exports
    _BUTTON_FORM_TAGS,
    _FORM_METADATA_COMMAND_PREFIX,
    _FORM_TAG_BY_TYPE,
    _FORM_TAGS,
    _MEMO_CARRIER_FIELD_TYPE,
    _bool_like_literal_is_valid,
    _expected_binary_media_type,
    _graphic_attribute_xpath,
    _issue,
    _new_temp_dir,
    _normalize_path,
)
from ._hwpx_factory import (  # noqa: F401 - compatibility re-exports
    _build_blank_container_part,
    _build_blank_content_hpf_part,
    _build_blank_document_state,
    _build_blank_header_part,
    _build_blank_section_part,
    _build_blank_settings_part,
    _build_blank_version_part,
    _full_nsmap,
)
from ._hwpx_document_controls import (
    _APPEND_BLOCK_TYPE_ALIASES,  # noqa: F401 - compatibility re-export
    HwpxDocumentControlMixin,
    _canonical_append_block_type,  # noqa: F401 - compatibility re-export
)
from ._hwpx_validation import HwpxValidationMixin
from ._ole_payloads import default_ole_cfb_bytes, normalize_hwpx_ole_payload
from .elements import (
    AutoNumberXml,
    BookmarkXml,
    CharacterStyleXml,
    ChartXml,
    EquationXml,
    FieldXml,
    FormXml,
    HeaderFooterXml,
    MemoShapeXml,
    MemoXml,
    NoteXml,
    OleXml,
    ParagraphStyleXml,
    PictureXml,
    SectionSettingsXml,
    ShapeXml,
    StyleDefinitionXml,
    TableXml,
    _invalidate_paragraph_layout,
    _missing_preserved_tokens,
    _preserved_structure_signature,
    _set_graphic_layout,
    _set_margin_values,
)
from .exceptions import HwpxValidationError, InvalidHwpxFileError, ValidationIssue  # noqa: F401
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
    XmlPart,
    infer_part_class,
)
from .xmlnode import HwpxXmlNode

_DEFAULT_CHART_FALLBACK_OLE_FIXTURE = "tests/fixtures/charts/chart_01_single_column.hwpx"
_DEFAULT_CHART_FALLBACK_OLE_BYTES: bytes | None = None

def _normalize_existing_index(index: int, count: int, label: str) -> int:
    if count <= 0:
        raise IndexError(f"No {label} exists.")
    normalized = index + count if index < 0 else index
    if normalized < 0 or normalized >= count:
        raise IndexError(f"{label} index out of range: {index}")
    return normalized


def _normalize_insert_index(index: int | None, count: int, label: str) -> int:
    if index is None:
        return count
    normalized = index + count if index < 0 else index
    if normalized < 0 or normalized > count:
        raise IndexError(f"{label} insert index out of range: {index}")
    return normalized


def _parse_xml_element(xml: str | bytes) -> etree._Element:
    raw = xml if isinstance(xml, bytes) else xml.encode("utf-8")
    element = etree.fromstring(raw)
    if not isinstance(element.tag, str):
        raise ValueError("XML payload must contain a concrete element.")
    return element


def _local_name(element: etree._Element) -> str:
    return etree.QName(element).localname


_DIRECT_RUN_CONTROL_TAGS = {
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
    "equation",
    "chart",
    "switch",
    "checkBtn",
    "radioBtn",
    "btn",
    "edit",
    "comboBox",
    "listBox",
    "scrollBar",
    "autoNum",
    "newNum",
}


def _node_for_control_return(element: etree._Element) -> etree._Element:
    if _local_name(element) == "ctrl" and len(element):
        return element[0]
    return element


def _paragraph_has_visible_text(paragraph: etree._Element) -> bool:
    return any((text or "").strip() for text in paragraph.xpath(".//hp:t//text()", namespaces=NS))


def _paragraph_has_nontext_run_content(paragraph: etree._Element) -> bool:
    return bool(paragraph.xpath("./hp:run/*[not(self::hp:t)]", namespaces=NS))


def _paragraph_plain_text(paragraph: etree._Element) -> str:
    return "".join(paragraph.xpath(".//hp:t/text()", namespaces=NS))


def _paragraph_text_node_count(paragraph: etree._Element) -> int:
    return len(paragraph.xpath("./hp:run/hp:t", namespaces=NS))


def _paragraph_layout_text_length(paragraph: etree._Element) -> int | None:
    """Count LineSeg positions, including explicit breaks and UTF-16 pairs.

    Unknown inline controls can occupy multiple HWP character positions. Leave
    their cache alone rather than guessing a length and deleting valid layout.
    """
    breaks = 0
    for text_node in paragraph.xpath(".//hp:t", namespaces=NS):
        for child in text_node:
            if child.tag != qname("hp", "lineBreak") or len(child) or child.text:
                return None
            breaks += 1
    return len(_paragraph_plain_text(paragraph).encode("utf-16-le")) // 2 + breaks


def _containing_top_level_paragraph_index(
    section_root: etree._Element,
    paragraph: etree._Element,
    top_level_indexes: dict[int, int],
) -> int | None:
    current: etree._Element | None = paragraph
    while current is not None:
        if current.getparent() is section_root:
            return top_level_indexes.get(id(current))
        parent = current.getparent()
        current = parent if isinstance(parent, etree._Element) else None
    return None
def _resolve_append_kind_argument(
    kind: str | None,
    extra_kwargs: dict[str, object],
) -> str:
    legacy_type = extra_kwargs.pop("type", None)
    if kind is not None and legacy_type is not None:
        raise TypeError("Pass either kind=... or type=..., not both.")
    resolved = kind if kind is not None else legacy_type
    if not isinstance(resolved, str) or not resolved.strip():
        raise TypeError("kind must be a non-empty string.")
    return resolved


def _is_chart_fallback_ole(node: etree._Element) -> bool:
    return bool(node.xpath("ancestor::hp:default[parent::hp:switch[hp:case/hp:chart]]", namespaces=NS))


def _default_chart_fallback_ole_bytes() -> bytes:
    global _DEFAULT_CHART_FALLBACK_OLE_BYTES
    if _DEFAULT_CHART_FALLBACK_OLE_BYTES is not None:
        return _DEFAULT_CHART_FALLBACK_OLE_BYTES

    candidate = Path(__file__).resolve().parents[2] / ".codex-temp" / "HwpForge" / _DEFAULT_CHART_FALLBACK_OLE_FIXTURE
    if candidate.exists():
        with zipfile.ZipFile(candidate) as archive:
            _DEFAULT_CHART_FALLBACK_OLE_BYTES = normalize_hwpx_ole_payload(archive.read("BinData/ole1.ole"))
            return _DEFAULT_CHART_FALLBACK_OLE_BYTES

    _DEFAULT_CHART_FALLBACK_OLE_BYTES = normalize_hwpx_ole_payload(default_ole_cfb_bytes())
    return _DEFAULT_CHART_FALLBACK_OLE_BYTES


class HwpxDocument(HwpxDocumentControlMixin, HwpxValidationMixin):
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
    def open(cls, path: str | os.PathLike[str], *, validate: bool = True) -> "HwpxDocument":
        input_path = Path(path)
        if not input_path.exists():
            raise FileNotFoundError(input_path)
        if not zipfile.is_zipfile(input_path):
            raise InvalidHwpxFileError(f"{input_path} is not a valid zip-based HWPX package.")
        with zipfile.ZipFile(input_path) as archive:
            return cls._from_zipfile(archive=archive, source_path=input_path, validate=validate)

    @classmethod
    def from_bytes(cls, raw_bytes: bytes, *, validate: bool = True) -> "HwpxDocument":
        try:
            with zipfile.ZipFile(io.BytesIO(raw_bytes)) as archive:
                return cls._from_zipfile(archive=archive, source_path=None, validate=validate)
        except zipfile.BadZipFile as exc:
            raise InvalidHwpxFileError("The supplied bytes are not a valid zip-based HWPX package.") from exc

    @classmethod
    def _from_zipfile(
        cls,
        archive: zipfile.ZipFile,
        source_path: Path | None,
        *,
        validate: bool,
    ) -> "HwpxDocument":
        parts, part_order, zip_infos, duplicates = cls._read_zipfile_state(archive)
        document = cls(
            parts=parts,
            part_order=part_order,
            zip_infos=zip_infos,
            source_path=source_path,
            duplicate_entries=duplicates,
        )
        if validate:
            document.validate()
        return document

    @classmethod
    def repair(
        cls,
        path: str | os.PathLike[str],
        output_path: str | os.PathLike[str] | None = None,
    ) -> Path:
        source_path = Path(path)
        if output_path is None:
            target_path = source_path.with_name(f"{source_path.stem}_repaired{source_path.suffix}")
        else:
            target_path = Path(output_path)
        document = cls.open(source_path, validate=False)
        document.repair_stale_paragraph_layout()
        return document.save(target_path, validate=True)

    @staticmethod
    def _read_zipfile_state(
        archive: zipfile.ZipFile,
    ) -> tuple[dict[str, HwpxPart], list[str], dict[str, zipfile.ZipInfo], list[str]]:
        return read_archive_state(archive)

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
            for node in section.readonly_root_element.xpath(".//hp:header", namespaces=NS):
                blocks.append(HeaderFooterXml(self, section, node))
        return blocks

    def footers(self, section_index: int | None = None) -> list[HeaderFooterXml]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        blocks: list[HeaderFooterXml] = []
        for section in sections:
            for node in section.readonly_root_element.xpath(".//hp:footer", namespaces=NS):
                blocks.append(HeaderFooterXml(self, section, node))
        return blocks

    def tables(self, section_index: int | None = None) -> list[TableXml]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        tables: list[TableXml] = []
        for section in sections:
            for node in section.readonly_root_element.xpath(".//hp:tbl", namespaces=NS):
                tables.append(TableXml(self, section, node))
        return tables

    def pictures(self, section_index: int | None = None) -> list[PictureXml]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        pictures: list[PictureXml] = []
        for section in sections:
            for node in section.readonly_root_element.xpath(".//hp:pic", namespaces=NS):
                pictures.append(PictureXml(self, section, node))
        return pictures

    def oles(self, section_index: int | None = None) -> list[OleXml]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        values: list[OleXml] = []
        for section in sections:
            for node in section.readonly_root_element.xpath(".//hp:ole", namespaces=NS):
                if _is_chart_fallback_ole(node):
                    continue
                values.append(OleXml(self, section, node))
        return values

    def section_settings(self, section_index: int = 0) -> SectionSettingsXml:
        self._ensure_editable_sections()
        nodes = self.sections[section_index].readonly_root_element.xpath("./hp:p[1]//hp:secPr[1]", namespaces=NS)
        if not nodes:
            raise ValueError("Section does not contain hp:secPr.")
        return SectionSettingsXml(self.sections[section_index], nodes[0])

    def notes(self, section_index: int | None = None) -> list[NoteXml]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        notes: list[NoteXml] = []
        for section in sections:
            for tag in (".//hp:footNote", ".//hp:endNote"):
                for node in section.readonly_root_element.xpath(tag, namespaces=NS):
                    notes.append(NoteXml(self, section, node))
        return notes

    def memos(self, section_index: int | None = None) -> list[MemoXml]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        values: list[MemoXml] = []
        for section in sections:
            for node in section.readonly_root_element.xpath(".//hp:hiddenComment", namespaces=NS):
                values.append(MemoXml(self, section, node))
        return values

    def forms(self, section_index: int | None = None) -> list[FormXml]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        values: list[FormXml] = []
        xpath = " | ".join(f".//hp:{tag}" for tag in _FORM_TAGS)
        for section in sections:
            for node in section.readonly_root_element.xpath(xpath, namespaces=NS):
                values.append(FormXml(self, section, node))
        return values

    def charts(self, section_index: int | None = None) -> list[ChartXml]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        values: list[ChartXml] = []
        for section in sections:
            for node in section.readonly_root_element.xpath(".//hp:chart", namespaces=NS):
                values.append(ChartXml(self, section, node))
        return values

    def bookmarks(self, section_index: int | None = None) -> list[BookmarkXml]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        bookmarks: list[BookmarkXml] = []
        for section in sections:
            for node in section.readonly_root_element.xpath(".//hp:bookmark", namespaces=NS):
                bookmarks.append(BookmarkXml(self, section, node))
        return bookmarks

    def fields(self, section_index: int | None = None) -> list[FieldXml]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        fields: list[FieldXml] = []
        for section in sections:
            for node in section.readonly_root_element.xpath(".//hp:fieldBegin", namespaces=NS):
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
                for node in section.readonly_root_element.xpath(tag, namespaces=NS):
                    values.append(AutoNumberXml(self, section, node))
        return values

    def equations(self, section_index: int | None = None) -> list[EquationXml]:
        sections = self.sections if section_index is None else [self.sections[section_index]]
        values: list[EquationXml] = []
        for section in sections:
            for node in section.readonly_root_element.xpath(".//hp:equation", namespaces=NS):
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
                for node in section.readonly_root_element.xpath(f".//{tag}", namespaces=NS):
                    if etree.QName(node).localname == "ole":
                        if _is_chart_fallback_ole(node):
                            continue
                        values.append(OleXml(self, section, node))
                    else:
                        values.append(ShapeXml(self, section, node))
        return values

    def styles(self) -> list[StyleDefinitionXml]:
        return [
            StyleDefinitionXml(self.header, node)
            for node in self.header.readonly_root_element.xpath(".//hh:style", namespaces=NS)
        ]

    def paragraph_styles(self) -> list[ParagraphStyleXml]:
        return [
            ParagraphStyleXml(self.header, node)
            for node in self.header.readonly_root_element.xpath(".//hh:paraPr", namespaces=NS)
        ]

    def character_styles(self) -> list[CharacterStyleXml]:
        return [
            CharacterStyleXml(self.header, node)
            for node in self.header.readonly_root_element.xpath(".//hh:charPr", namespaces=NS)
        ]

    def memo_shapes(self) -> list[MemoShapeXml]:
        return [
            MemoShapeXml(self.header, node)
            for node in self.header.readonly_root_element.xpath(".//hh:memoPr", namespaces=NS)
        ]

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

    def _stale_paragraph_layout_targets(
        self,
    ) -> list[tuple[SectionPart, int, int | None, etree._Element, str, str]]:
        targets: list[tuple[SectionPart, int, int | None, etree._Element, str, str]] = []
        for section_index, section in enumerate(self.sections):
            section_root = section.readonly_root_element
            top_level_paragraphs = section_root.xpath("./hp:p", namespaces=NS)
            top_level_indexes = {id(paragraph): index for index, paragraph in enumerate(top_level_paragraphs)}

            for paragraph in section.root_element.xpath(".//hp:p[hp:linesegarray]", namespaces=NS):
                has_nontext_content = _paragraph_has_nontext_run_content(paragraph)
                text_length = _paragraph_layout_text_length(paragraph)
                if text_length is None:
                    continue
                numeric_positions: list[int] = []
                raw_positions: list[str] = []
                for line_seg in paragraph.xpath("./hp:linesegarray/hp:lineseg[@textpos]", namespaces=NS):
                    value = line_seg.get("textpos")
                    if value is None:
                        continue
                    raw_positions.append(value)
                    try:
                        numeric_positions.append(int(value))
                    except ValueError:
                        continue
                if not numeric_positions:
                    continue


                paragraph_index = _containing_top_level_paragraph_index(section.root_element, paragraph, top_level_indexes)
                if text_length == 0 and not has_nontext_content:
                    stale_positions = [str(value) for value in numeric_positions if value > 0]
                    if stale_positions:
                        targets.append(
                            (
                                section,
                                section_index,
                                paragraph_index,
                                paragraph,
                                "empty_paragraph_textpos",
                                f"textpos={','.join(stale_positions)}",
                            )
                        )
                    continue

                if has_nontext_content or _paragraph_text_node_count(paragraph) != 1:
                    continue

                max_textpos = max(numeric_positions)
                if max_textpos > text_length:
                    targets.append(
                        (
                            section,
                            section_index,
                            paragraph_index,
                            paragraph,
                            "plain_text_textpos_overflow",
                            f"max_textpos={max_textpos};text_len={text_length};textpos={','.join(raw_positions)}",
                        )
                    )

            for paragraph in section_root.xpath(".//hp:caption/hp:subList/hp:p[hp:linesegarray]", namespaces=NS):
                text = _paragraph_plain_text(paragraph)
                text_length = _paragraph_layout_text_length(paragraph)
                if text_length is None:
                    continue
                if not text.strip():
                    continue
                raw_positions = [
                    value
                    for value in paragraph.xpath("./hp:linesegarray/hp:lineseg/@textpos", namespaces=NS)
                    if str(value).lstrip("-").isdigit()
                ]
                if not raw_positions:
                    continue
                numeric_positions = [int(value) for value in raw_positions]
                max_textpos = max(numeric_positions)
                if max_textpos <= text_length:
                    continue
                paragraph_index = _containing_top_level_paragraph_index(section_root, paragraph, top_level_indexes)
                targets.append(
                    (
                        section,
                        section_index,
                        paragraph_index,
                        paragraph,
                        "caption_textpos_overflow",
                        f"max_textpos={max_textpos};text_len={text_length};textpos={','.join(raw_positions)}",
                    )
                )
        return targets

    def stale_paragraph_layout_validation_errors(self) -> list[ValidationIssue]:
        errors: list[ValidationIssue] = []
        for section, section_index, paragraph_index, _paragraph, reason, context in self._stale_paragraph_layout_targets():
            if reason == "plain_text_textpos_overflow":
                message = (
                    "Plain-text paragraph carries linesegarray text positions beyond the text length; "
                    "Hancom may drop following content."
                )
            elif reason == "caption_textpos_overflow":
                message = (
                    "Caption paragraph carries linesegarray text positions beyond the caption text length; "
                    "Hancom may drop the anchored object."
                )
            else:
                message = "Empty paragraph carries stale linesegarray text positions; Hancom may drop following content."
            errors.append(
                _issue(
                    "paragraph_layout_cache",
                    message,
                    part_path=section.path,
                    section_index=section_index,
                    paragraph_index=paragraph_index,
                    xpath="./hp:linesegarray",
                    context=context,
                )
            )
        return errors

    def repair_stale_paragraph_layout(self) -> list[ValidationIssue]:
        repairs: list[ValidationIssue] = []
        for section, section_index, paragraph_index, paragraph, reason, context in self._stale_paragraph_layout_targets():
            _invalidate_paragraph_layout(paragraph)
            section.mark_modified()
            if reason == "plain_text_textpos_overflow":
                message = "Removed stale linesegarray from plain-text paragraph."
            elif reason == "caption_textpos_overflow":
                message = "Removed stale linesegarray from caption paragraph."
            else:
                message = "Removed stale linesegarray from empty paragraph."
            repairs.append(
                _issue(
                    "paragraph_layout_cache_repair",
                    message,
                    part_path=section.path,
                    section_index=section_index,
                    paragraph_index=paragraph_index,
                    context=context,
                )
            )
        return repairs

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

    def append(
        self,
        kind: str | None = None,
        content: object | None = None,
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
        **kwargs,
    ):
        return self.append_block(
            kind=kind,
            content=content,
            section_index=section_index,
            paragraph_index=paragraph_index,
            char_pr_id=char_pr_id,
            **kwargs,
        )

    def append_block(
        self,
        kind: str | None = None,
        content: object | None = None,
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
        **kwargs,
    ):
        self._ensure_editable_sections()
        resolved_kind = _resolve_append_kind_argument(kind, kwargs)
        return self._dispatch_append_block(
            resolved_kind,
            content,
            section_index=section_index,
            paragraph_index=paragraph_index,
            char_pr_id=char_pr_id,
            inline=False,
            extra_kwargs=kwargs,
        )

    def append_inline(
        self,
        kind: str | None = None,
        content: object | None = None,
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
        **kwargs,
    ):
        self._ensure_editable_sections()
        resolved_kind = _resolve_append_kind_argument(kind, kwargs)
        return self._dispatch_append_block(
            resolved_kind,
            content,
            section_index=section_index,
            paragraph_index=paragraph_index,
            char_pr_id=char_pr_id,
            inline=True,
            extra_kwargs=kwargs,
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

    def paragraph_count(self, section_index: int = 0) -> int:
        self._ensure_editable_sections()
        return len(self._paragraph_elements(section_index))

    def paragraph_xml(self, section_index: int, paragraph_index: int, *, pretty_print: bool = False) -> str:
        self._ensure_editable_sections()
        paragraph = self._paragraph_element(section_index, paragraph_index)
        return etree.tostring(paragraph, encoding="unicode", pretty_print=pretty_print)

    def insert_paragraph_xml(
        self,
        section_index: int,
        paragraph_index: int | None,
        xml: str | bytes,
    ) -> HwpxXmlNode:
        self._ensure_editable_sections()
        section = self.sections[section_index]
        paragraph = self._parse_paragraph_xml(xml)
        paragraphs = self._paragraph_elements(section_index)
        insert_index = _normalize_insert_index(paragraph_index, len(paragraphs), "paragraph")
        if insert_index == len(paragraphs):
            section.root_element.append(paragraph)
        else:
            paragraphs[insert_index].addprevious(paragraph)
        _invalidate_paragraph_layout(paragraph)
        section.mark_modified()
        return HwpxXmlNode(paragraph, section)

    def replace_paragraph_xml(
        self,
        section_index: int,
        paragraph_index: int,
        xml: str | bytes,
    ) -> HwpxXmlNode:
        self._ensure_editable_sections()
        section = self.sections[section_index]
        paragraphs = self._paragraph_elements(section_index)
        index = _normalize_existing_index(paragraph_index, len(paragraphs), "paragraph")
        paragraph = self._parse_paragraph_xml(xml)
        paragraphs[index].getparent().replace(paragraphs[index], paragraph)
        _invalidate_paragraph_layout(paragraph)
        section.mark_modified()
        return HwpxXmlNode(paragraph, section)

    def move_paragraph(
        self,
        source_section_index: int,
        source_paragraph_index: int,
        *,
        target_section_index: int | None = None,
        target_paragraph_index: int | None = None,
    ) -> HwpxXmlNode:
        self._ensure_editable_sections()
        target_section_index = source_section_index if target_section_index is None else target_section_index
        source_section = self.sections[source_section_index]
        target_section = self.sections[target_section_index]
        source_paragraphs = self._paragraph_elements(source_section_index)
        source_index = _normalize_existing_index(source_paragraph_index, len(source_paragraphs), "paragraph")
        if source_section_index != target_section_index and len(source_paragraphs) <= 1:
            raise ValueError("A source section must keep at least one paragraph.")
        paragraph = source_paragraphs[source_index]
        source_section.root_element.remove(paragraph)

        target_paragraphs = self._paragraph_elements(target_section_index)
        target_index = _normalize_insert_index(target_paragraph_index, len(target_paragraphs), "paragraph")
        if target_index == len(target_paragraphs):
            target_section.root_element.append(paragraph)
        else:
            target_paragraphs[target_index].addprevious(paragraph)
        _invalidate_paragraph_layout(paragraph)
        source_section.mark_modified()
        target_section.mark_modified()
        return HwpxXmlNode(paragraph, target_section)

    def copy_paragraph(
        self,
        source_section_index: int,
        source_paragraph_index: int,
        *,
        target_section_index: int | None = None,
        target_paragraph_index: int | None = None,
    ) -> HwpxXmlNode:
        self._ensure_editable_sections()
        target_section_index = source_section_index if target_section_index is None else target_section_index
        target_section = self.sections[target_section_index]
        paragraph = deepcopy(self._paragraph_element(source_section_index, source_paragraph_index))
        paragraph.set("id", target_section._next_paragraph_id())

        target_paragraphs = self._paragraph_elements(target_section_index)
        target_index = _normalize_insert_index(target_paragraph_index, len(target_paragraphs), "paragraph")
        if target_index == len(target_paragraphs):
            target_section.root_element.append(paragraph)
        else:
            target_paragraphs[target_index].addprevious(paragraph)
        _invalidate_paragraph_layout(paragraph)
        target_section.mark_modified()
        return HwpxXmlNode(paragraph, target_section)

    def control_count(self, section_index: int, paragraph_index: int) -> int:
        self._ensure_editable_sections()
        paragraph = self._paragraph_element(section_index, paragraph_index)
        return len(self._paragraph_control_wrappers(paragraph))

    def insert_control_xml_at(
        self,
        xml: str | bytes,
        *,
        section_index: int = 0,
        paragraph_index: int = 0,
        control_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> HwpxXmlNode:
        self._ensure_editable_sections()
        control = _parse_xml_element(xml)
        if _local_name(control) == "ctrl":
            control_element = control
        elif _local_name(control) in _DIRECT_RUN_CONTROL_TAGS:
            control_element = control
        else:
            control_element = etree.Element(qname("hp", "ctrl"))
            control_element.append(control)
        paragraph = self._paragraph_element(section_index, paragraph_index)
        section = self.sections[section_index]
        self._insert_control_wrapper(paragraph, control_element, control_index=control_index, section=section, char_pr_id=char_pr_id)
        return HwpxXmlNode(_node_for_control_return(control_element), section)

    def delete_control(self, section_index: int, paragraph_index: int, control_index: int) -> None:
        self._ensure_editable_sections()
        paragraph = self._paragraph_element(section_index, paragraph_index)
        wrappers = self._paragraph_control_wrappers(paragraph)
        index = _normalize_existing_index(control_index, len(wrappers), "control")
        self._detach_control_wrapper(wrappers[index])
        _invalidate_paragraph_layout(paragraph)
        self.sections[section_index].mark_modified()

    def move_control(
        self,
        source_section_index: int,
        source_paragraph_index: int,
        control_index: int,
        *,
        target_section_index: int | None = None,
        target_paragraph_index: int,
        target_control_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> HwpxXmlNode:
        self._ensure_editable_sections()
        target_section_index = source_section_index if target_section_index is None else target_section_index
        source_section = self.sections[source_section_index]
        target_section = self.sections[target_section_index]
        source_paragraph = self._paragraph_element(source_section_index, source_paragraph_index)
        wrappers = self._paragraph_control_wrappers(source_paragraph)
        source_index = _normalize_existing_index(control_index, len(wrappers), "control")
        wrapper = wrappers[source_index]
        source_run = wrapper.getparent()
        resolved_char_pr_id = char_pr_id or (source_run.get("charPrIDRef") if source_run is not None else None)
        control = _node_for_control_return(wrapper)
        self._detach_control_wrapper(wrapper)

        target_paragraph = self._paragraph_element(target_section_index, target_paragraph_index)
        self._insert_control_wrapper(
            target_paragraph,
            wrapper,
            control_index=target_control_index,
            section=target_section,
            char_pr_id=resolved_char_pr_id,
        )
        _invalidate_paragraph_layout(source_paragraph)
        source_section.mark_modified()
        target_section.mark_modified()
        return HwpxXmlNode(control, target_section)

    def copy_control(
        self,
        source_section_index: int,
        source_paragraph_index: int,
        control_index: int,
        *,
        target_section_index: int | None = None,
        target_paragraph_index: int,
        target_control_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> HwpxXmlNode:
        self._ensure_editable_sections()
        target_section_index = source_section_index if target_section_index is None else target_section_index
        source_paragraph = self._paragraph_element(source_section_index, source_paragraph_index)
        wrappers = self._paragraph_control_wrappers(source_paragraph)
        source_index = _normalize_existing_index(control_index, len(wrappers), "control")
        source_run = wrappers[source_index].getparent()
        resolved_char_pr_id = char_pr_id or (source_run.get("charPrIDRef") if source_run is not None else None)
        wrapper = deepcopy(wrappers[source_index])
        control = _node_for_control_return(wrapper)
        target_paragraph = self._paragraph_element(target_section_index, target_paragraph_index)
        target_section = self.sections[target_section_index]
        self._insert_control_wrapper(
            target_paragraph,
            wrapper,
            control_index=target_control_index,
            section=target_section,
            char_pr_id=resolved_char_pr_id,
        )
        return HwpxXmlNode(control, target_section)

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

    def _unused_binary_name(self, name: str) -> str:
        original = PurePosixPath(name).name
        candidate = original
        suffix = PurePosixPath(original).suffix
        stem = PurePosixPath(original).stem or "bindata"
        index = 2
        used_hrefs = {item.get("href", "") for item in self.content_hpf.manifest_items()}
        while f"BinData/{candidate}" in self._parts or f"BinData/{candidate}" in used_hrefs:
            candidate = f"{stem}_{index}{suffix}"
            index += 1
        return candidate

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


    def compile(self, validate: bool = True, *, auto_repair: bool = True) -> bytes:
        self._prepare_for_write(validate=validate, auto_repair=auto_repair)

        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, mode="w") as archive:
            self._write_archive(archive)
        return buffer.getvalue()

    def save(self, path: str | os.PathLike[str], validate: bool = True, *, auto_repair: bool = True) -> Path:
        output_path = Path(path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        self._prepare_for_write(validate=validate, auto_repair=auto_repair)

        descriptor, temporary_name = tempfile.mkstemp(
            prefix=f".{output_path.name}.",
            suffix=".tmp",
            dir=output_path.parent,
        )
        os.close(descriptor)
        temporary_path = Path(temporary_name)
        try:
            with zipfile.ZipFile(temporary_path, mode="w") as archive:
                self._write_archive(archive)
            os.replace(temporary_path, output_path)
        finally:
            temporary_path.unlink(missing_ok=True)
        return output_path

    def _prepare_for_write(self, *, validate: bool, auto_repair: bool) -> None:
        if auto_repair:
            self.repair_stale_paragraph_layout()
        if validate:
            self.validate()

    def _write_archive(self, archive: zipfile.ZipFile) -> None:
        write_archive_state(
            archive,
            part_paths=self.list_part_paths(),
            parts=self._parts,
            zip_infos=self._zip_infos,
        )

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



    def _ensure_editable_sections(self) -> None:
        if self.sections:
            return
        if self.is_distribution_protected:
            raise RuntimeError(
                "This HWPX package is distribution protected. The encrypted section/header parts can be preserved, "
                "but high-level editing is not available."
            )
        raise RuntimeError("This HWPX package does not contain editable section XML parts.")

    def _paragraph_elements(self, section_index: int) -> list[etree._Element]:
        return self.sections[section_index].readonly_root_element.xpath("./hp:p", namespaces=NS)

    def _paragraph_element(self, section_index: int, paragraph_index: int) -> etree._Element:
        paragraphs = self._paragraph_elements(section_index)
        index = _normalize_existing_index(paragraph_index, len(paragraphs), "paragraph")
        return paragraphs[index]

    def _parse_paragraph_xml(self, xml: str | bytes) -> etree._Element:
        paragraph = _parse_xml_element(xml)
        if _local_name(paragraph) != "p":
            raise ValueError("Paragraph XML must use an hp:p root element.")
        return paragraph

    def _paragraph_control_wrappers(self, paragraph: etree._Element) -> list[etree._Element]:
        values: list[etree._Element] = []
        for node in paragraph.xpath("./hp:run/*", namespaces=NS):
            name = _local_name(node)
            if name == "t":
                continue
            if name == "ctrl" or name in _DIRECT_RUN_CONTROL_TAGS:
                values.append(node)
        return values

    def _insert_control_wrapper(
        self,
        paragraph: etree._Element,
        wrapper: etree._Element,
        *,
        control_index: int | None,
        section: SectionPart,
        char_pr_id: str | None = None,
    ) -> None:
        controls = self._paragraph_control_wrappers(paragraph)
        insert_index = _normalize_insert_index(control_index, len(controls), "control")
        run = etree.Element(qname("hp", "run"))
        run.set("charPrIDRef", char_pr_id or self._default_char_pr_id(paragraph))
        run.append(wrapper)
        if insert_index == len(controls):
            line_seg = paragraph.xpath("./hp:linesegarray[1]", namespaces=NS)
            if line_seg:
                line_seg[0].addprevious(run)
            else:
                paragraph.append(run)
        else:
            target_run = controls[insert_index].getparent()
            target_run.addprevious(run)
        _invalidate_paragraph_layout(paragraph)
        section.mark_modified()

    def _detach_control_wrapper(self, wrapper: etree._Element) -> None:
        parent = wrapper.getparent()
        if parent is None:
            return
        parent.remove(wrapper)
        if _local_name(parent) != "run" or len(parent) > 0:
            return
        if (parent.text or "").strip():
            return
        grandparent = parent.getparent()
        if grandparent is not None:
            grandparent.remove(parent)

    def _resolve_paragraph_for_insert(self, *, section_index: int, paragraph_index: int | None) -> etree._Element:
        if paragraph_index is None:
            return self.append_paragraph("", section_index=section_index).element
        paragraphs = self.sections[section_index].readonly_root_element.xpath("./hp:p", namespaces=NS)
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
        nodes = self.header.readonly_root_element.xpath(expression, namespaces=NS)
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
        nodes = self.header.readonly_root_element.xpath(expression, namespaces=NS)
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
        for value in self.header.readonly_root_element.xpath(expression, namespaces=NS):
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
            for value in section.readonly_root_element.xpath(expression, namespaces=NS):
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
