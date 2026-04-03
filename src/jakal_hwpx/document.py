from __future__ import annotations

import io
import os
import re
import shutil
import subprocess
import tempfile
import time
import zipfile
from pathlib import Path, PurePosixPath

from lxml import etree

from .elements import (
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
from .exceptions import HwpxValidationError, InvalidHwpxFileError
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
    XmlPart,
    infer_part_class,
)


def _normalize_path(path: str) -> str:
    return PurePosixPath(path).as_posix()


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


class HwpxDocument:
    def __init__(
        self,
        *,
        parts: dict[str, HwpxPart],
        part_order: list[str],
        zip_infos: dict[str, zipfile.ZipInfo],
        source_path: Path | None = None,
        duplicate_entries: list[str] | None = None,
    ):
        self._parts = parts
        self._part_order = part_order
        self._zip_infos = zip_infos
        self.source_path = source_path
        self.duplicate_entries = duplicate_entries or []

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

        document = cls(
            parts=parts,
            part_order=part_order,
            zip_infos=zip_infos,
            source_path=source_path,
            duplicate_entries=duplicates,
        )
        document.validate()
        return document

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
            self.container.ensure_rootfile("Preview/PrvText.txt", "text/plain")
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

    def validation_errors(self) -> list[str]:
        errors: list[str] = []
        required = [
            "mimetype",
            "version.xml",
            "Contents/content.hpf",
            "META-INF/container.xml",
        ]
        if not self.is_distribution_protected:
            required.append("Contents/header.xml")
        for path in required:
            if path not in self._parts:
                errors.append(f"Missing required part: {path}")

        if self.duplicate_entries:
            errors.append(f"Duplicate zip entries found: {', '.join(sorted(set(self.duplicate_entries)))}")

        mimetype = self._parts.get("mimetype")
        if isinstance(mimetype, MimetypePart) and mimetype.mime.strip() != "application/hwp+zip":
            errors.append("mimetype must be application/hwp+zip")

        xml_parts = [part for part in self._parts.values() if isinstance(part, XmlPart)]
        for part in xml_parts:
            try:
                part.to_bytes()
            except Exception as exc:  # noqa: BLE001
                errors.append(f"Invalid XML part {part.path}: {exc}")

        container = self._parts.get("META-INF/container.xml")
        if isinstance(container, ContainerPart):
            for rootfile_path in container.rootfile_paths():
                normalized = _normalize_path(rootfile_path)
                if normalized not in self._parts:
                    errors.append(f"container.xml references missing rootfile: {normalized}")

        content_hpf = self._parts.get("Contents/content.hpf")
        if isinstance(content_hpf, ContentHpfPart):
            manifest_ids: set[str] = set()
            for item in content_hpf.manifest_items():
                item_id = item.get("id")
                href = item.get("href")
                if not item_id:
                    errors.append("content.hpf manifest item is missing id")
                elif item_id in manifest_ids:
                    errors.append(f"Duplicate content.hpf manifest id: {item_id}")
                else:
                    manifest_ids.add(item_id)
                if not href:
                    errors.append("content.hpf manifest item is missing href")
                    continue
                if _normalize_path(href) not in self._parts:
                    errors.append(f"content.hpf manifest references missing part: {href}")

            for itemref in content_hpf.spine_itemrefs():
                item_id = itemref.get("idref")
                if not item_id:
                    errors.append("content.hpf spine itemref is missing idref")
                elif item_id not in manifest_ids:
                    errors.append(f"content.hpf spine references unknown manifest id: {item_id}")

        header = self._parts.get("Contents/header.xml")
        if isinstance(header, HeaderPart):
            if header.section_count and header.section_count != len(self.sections):
                errors.append(
                    f"header.xml secCnt={header.section_count} does not match section count={len(self.sections)}"
                )

        if not self.is_distribution_protected and not self.sections:
            errors.append("Document must contain at least one section part")

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

    def xml_validation_errors(self) -> list[str]:
        errors: list[str] = []
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
                errors.append(f"{path} is expected to be XML but is loaded as {type(part).__name__}")
                continue
            actual = part.root.local_name
            if actual != expected_local_name:
                errors.append(f"{path} root element must be {expected_local_name}, got {actual}")
        for section in self.sections:
            if section.root.local_name != "sec":
                errors.append(f"{section.path} root element must be sec, got {section.root.local_name}")
            if not section.root_element.xpath("./hp:p", namespaces=NS):
                errors.append(f"{section.path} must contain at least one paragraph")
            for table in self.tables(section_index=self.sections.index(section)):
                if table.row_count < 1 or table.column_count < 1:
                    errors.append(f"{section.path} contains a table with invalid row/column count")
                for cell in table.cells():
                    if cell.row_span < 1 or cell.col_span < 1:
                        errors.append(f"{section.path} table cell has invalid span at ({cell.row}, {cell.column})")
        return errors

    def schema_validation_errors(self, schema_map: dict[str, str | os.PathLike[str]]) -> list[str]:
        validators: dict[Path, etree.XMLSchema] = {}
        errors: list[str] = []
        for part_path, schema_path in schema_map.items():
            part = self._parts.get(_normalize_path(part_path))
            if part is None:
                errors.append(f"Schema configured for missing part: {part_path}")
                continue
            if not isinstance(part, XmlPart):
                errors.append(f"Schema validation requires XML part: {part_path}")
                continue
            resolved_schema_path = Path(schema_path)
            if resolved_schema_path not in validators:
                schema_doc = etree.parse(str(resolved_schema_path))
                validators[resolved_schema_path] = etree.XMLSchema(schema_doc)
            validator = validators[resolved_schema_path]
            if not validator.validate(part._root):
                for entry in validator.error_log:
                    errors.append(f"{part_path}: {entry.message}")
        return errors

    def reference_validation_errors(self) -> list[str]:
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
                errors.append("bookmark is missing name")
                continue
            if bookmark.name in bookmark_names:
                errors.append(f"duplicate bookmark name: {bookmark.name}")
            bookmark_names.add(bookmark.name)

        for field in self.fields():
            if field.control_id:
                if field.control_id in field_begin_ids:
                    errors.append(f"duplicate fieldBegin id: {field.control_id}")
                field_begin_ids[field.control_id] = field.field_id or ""
            if field.field_id:
                if field.field_id in field_id_map:
                    errors.append(f"duplicate fieldid: {field.field_id}")
                field_id_map[field.field_id] = field.control_id or ""

        for section in self.sections:
            for paragraph in section.root_element.xpath("./hp:p", namespaces=NS):
                style_id = paragraph.get("styleIDRef")
                para_pr_id = paragraph.get("paraPrIDRef")
                if style_id and style_ids and style_id not in style_ids:
                    errors.append(f"{section.path} paragraph references unknown styleIDRef={style_id}")
                if para_pr_id and para_pr_ids and para_pr_id not in para_pr_ids:
                    errors.append(f"{section.path} paragraph references unknown paraPrIDRef={para_pr_id}")
                for run in paragraph.xpath("./hp:run", namespaces=NS):
                    char_pr_id = run.get("charPrIDRef")
                    if char_pr_id and char_pr_ids and char_pr_id not in char_pr_ids:
                        errors.append(f"{section.path} run references unknown charPrIDRef={char_pr_id}")
            for picture in self.pictures(section_index=self.sections.index(section)):
                if picture.binary_item_id and picture.binary_item_id not in manifest_ids:
                    errors.append(f"{section.path} picture references unknown manifest id {picture.binary_item_id}")
            for shape in self.shapes(section_index=self.sections.index(section)):
                binary_refs = shape.element.xpath("./@binaryItemIDRef", namespaces=NS)
                for binary_ref in binary_refs:
                    if binary_ref not in manifest_ids:
                        errors.append(f"{section.path} shape references unknown manifest id {binary_ref}")
            for field_end in section.root_element.xpath(".//hp:fieldEnd", namespaces=NS):
                begin_id = field_end.get("beginIDRef")
                field_id = field_end.get("fieldid")
                if not begin_id:
                    errors.append(f"{section.path} fieldEnd is missing beginIDRef")
                    continue
                if begin_id not in field_begin_ids:
                    errors.append(f"{section.path} fieldEnd references unknown fieldBegin id {begin_id}")
                    continue
                expected_field_id = field_begin_ids[begin_id]
                if field_id and expected_field_id and field_id != expected_field_id:
                    errors.append(
                        f"{section.path} fieldEnd fieldid={field_id} does not match fieldBegin fieldid={expected_field_id}"
                    )
        for field in self.cross_references():
            target = field.get_parameter("BookmarkName") or field.get_parameter("Path") or field.name
            if target and target not in bookmark_names:
                errors.append(f"cross reference points to unknown bookmark: {target}")
        return errors

    def save_reopen_validation_errors(self) -> list[str]:
        temp_dir = Path(tempfile.mkdtemp(prefix="jakal_hwpx_reopen_"))
        try:
            temp_path = temp_dir / "validation.hwpx"
            self.save(temp_path)
            reopened = HwpxDocument.open(temp_path)
            return reopened.validation_errors()
        except Exception as exc:  # noqa: BLE001
            return [str(exc)]
        finally:
            self._cleanup_temp_dir(temp_dir)

    @staticmethod
    def discover_hancom_executable() -> Path | None:
        env_path = os.environ.get("HWPX_HANCOM_EXE")
        if env_path:
            candidate = Path(env_path)
            if candidate.exists():
                return candidate

        for name in ("Hanword.exe", "Hword.exe", "Hwp.exe"):
            if located := shutil.which(name):
                return Path(located)

        for env_name in ("ProgramFiles", "ProgramFiles(x86)"):
            base = os.environ.get(env_name)
            if not base:
                continue
            base_path = Path(base)
            for stem in ("Hanword.exe", "Hword.exe", "Hwp.exe"):
                for candidate in base_path.glob(f"Hancom*/**/{stem}"):
                    if candidate.is_file():
                        return candidate
                for candidate in base_path.glob(f"HNC/**/{stem}"):
                    if candidate.is_file():
                        return candidate
        return None

    def hancom_open_validation_errors(
        self,
        executable_path: str | os.PathLike[str] | None = None,
        timeout_seconds: int = 15,
    ) -> list[str]:
        try:
            self.open_in_hancom(executable_path=executable_path, timeout_seconds=timeout_seconds)
        except Exception as exc:  # noqa: BLE001
            return [str(exc)]
        return []

    def open_in_hancom(
        self,
        executable_path: str | os.PathLike[str] | None = None,
        timeout_seconds: int = 15,
    ) -> None:
        if executable_path is None:
            discovered = self.discover_hancom_executable()
            if discovered is None:
                raise FileNotFoundError("Hancom executable was not found. Set HWPX_HANCOM_EXE or install Hanword.")
            executable_path = discovered
        temp_dir = Path(tempfile.mkdtemp(prefix="jakal_hwpx_hancom_"))
        try:
            temp_path = temp_dir / "open_test.hwpx"
            self.save(temp_path)
            process = subprocess.Popen([str(executable_path), str(temp_path)])
            timed_out = False
            try:
                process.wait(timeout=timeout_seconds)
            except subprocess.TimeoutExpired:
                timed_out = True
                self._terminate_process_tree(process)
            if not timed_out and process.returncode not in (None, 0):
                raise RuntimeError(f"Hancom process exited with code {process.returncode}")
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
    def _terminate_process_tree(process: subprocess.Popen[bytes] | subprocess.Popen[str]) -> None:
        if process.poll() is not None:
            return
        try:
            subprocess.run(
                ["taskkill", "/PID", str(process.pid), "/T", "/F"],
                check=False,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=10,
            )
        except Exception:  # noqa: BLE001
            process.kill()
        try:
            process.wait(timeout=10)
        except Exception:  # noqa: BLE001
            pass

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
