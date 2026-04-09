from __future__ import annotations

import mimetypes
import re
from copy import deepcopy
from dataclasses import dataclass, field
from pathlib import PurePosixPath
from typing import Any

from lxml import etree

from .elements import (
    _invalidate_paragraph_layout,
    _missing_preserved_tokens,
    _preserved_structure_signature,
    _replace_paragraph_text_preserving_controls,
)
from .exceptions import ValidationIssue
from .namespaces import NS, SECTION_PATTERN, qname
from .xmlnode import HwpxXmlNode

XML_DECLARATION_RE = re.compile(br"^\s*<\?xml[^>]*standalone=['\"](yes|no)['\"]", re.IGNORECASE)


def _build_xml_parser() -> etree.XMLParser:
    return etree.XMLParser(
        ns_clean=False,
        remove_blank_text=False,
        remove_comments=False,
        remove_pis=False,
        recover=False,
        resolve_entities=False,
        strip_cdata=False,
        huge_tree=True,
    )


def _detect_standalone(raw_bytes: bytes) -> bool | None:
    match = XML_DECLARATION_RE.match(raw_bytes[:200])
    if match is None:
        return None
    return match.group(1).lower() == b"yes"


def _looks_like_xml(raw_bytes: bytes) -> bool:
    return raw_bytes.lstrip().startswith(b"<")


def _replace_with_limit(source: str, old: str, new: str, count: int) -> tuple[str, int]:
    if count < 0:
        return source.replace(old, new), source.count(old)

    replaced = 0
    result = source
    while replaced < count and old in result:
        result = result.replace(old, new, 1)
        replaced += 1
    return result, replaced


@dataclass
class HwpxPart:
    path: str
    raw_bytes: bytes
    modified: bool = False
    document: Any | None = field(default=None, repr=False, compare=False)

    def __post_init__(self) -> None:
        self.path = PurePosixPath(self.path).as_posix()

    def mark_modified(self) -> None:
        self.modified = True

    def to_bytes(self) -> bytes:
        return self.raw_bytes


class GenericBinaryPart(HwpxPart):
    @property
    def data(self) -> bytes:
        return self.raw_bytes

    @data.setter
    def data(self, value: bytes) -> None:
        self.raw_bytes = value
        self.mark_modified()


class MimetypePart(HwpxPart):
    @property
    def mime(self) -> str:
        return self.raw_bytes.decode("utf-8")

    @mime.setter
    def mime(self, value: str) -> None:
        self.raw_bytes = value.encode("utf-8")
        self.mark_modified()


class GenericTextPart(HwpxPart):
    @property
    def text(self) -> str:
        return self.raw_bytes.decode("utf-8")

    @text.setter
    def text(self, value: str) -> None:
        self.raw_bytes = value.encode("utf-8")
        self.mark_modified()


class PreviewTextPart(GenericTextPart):
    pass


class PreviewImagePart(GenericBinaryPart):
    pass


class BinaryDataPart(GenericBinaryPart):
    pass


class ScriptPart(GenericTextPart):
    pass


class XmlPart(HwpxPart):
    def __init__(self, path: str, raw_bytes: bytes, modified: bool = False):
        super().__init__(path=path, raw_bytes=raw_bytes, modified=modified)
        self._standalone = _detect_standalone(raw_bytes)
        self._root = etree.fromstring(raw_bytes, parser=_build_xml_parser())

    @property
    def root(self) -> HwpxXmlNode:
        return HwpxXmlNode(self._root, self)

    @property
    def root_element(self) -> etree._Element:
        self.mark_modified()
        return self._root

    @classmethod
    def from_root(
        cls,
        path: str,
        root: etree._Element,
        standalone: bool | None = True,
    ) -> "XmlPart":
        instance = cls.__new__(cls)
        HwpxPart.__init__(instance, path=path, raw_bytes=b"", modified=True)
        instance._standalone = standalone
        instance._root = root
        return instance

    def clone_root(self) -> etree._Element:
        return deepcopy(self._root)

    def xpath(self, expression: str, namespaces: dict[str, str] | None = None) -> list[Any]:
        values = self._root.xpath(expression, namespaces=namespaces or NS)
        if not isinstance(values, list):
            values = [values]
        wrapped: list[Any] = []
        for value in values:
            if isinstance(value, etree._Element):
                wrapped.append(HwpxXmlNode(value, self))
            else:
                wrapped.append(value)
        return wrapped

    def find(self, expression: str, namespaces: dict[str, str] | None = None) -> HwpxXmlNode | None:
        return self.root.find(expression, namespaces=namespaces)

    def findall(self, expression: str, namespaces: dict[str, str] | None = None) -> list[HwpxXmlNode]:
        return self.root.findall(expression, namespaces=namespaces)

    def to_bytes(self) -> bytes:
        if not self.modified:
            return self.raw_bytes
        self.raw_bytes = etree.tostring(
            self._root,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=self._standalone,
            pretty_print=False,
        )
        self.modified = False
        return self.raw_bytes

    def extract_hp_text(self) -> str:
        fragments = []
        for node in self._root.xpath(".//hp:t", namespaces=NS):
            if node.text:
                fragments.append(node.text)
        return "".join(fragments)

    def replace_hp_text(self, old: str, new: str, count: int = -1) -> int:
        if not old:
            raise ValueError("old must be a non-empty string.")
        remaining = count
        replacements = 0
        layout_invalidated = False
        for node in self._root.xpath(".//hp:t", namespaces=NS):
            current = node.text or ""
            if old not in current:
                continue
            limit = remaining if remaining >= 0 else -1
            updated, changed = _replace_with_limit(current, old, new, limit)
            if changed:
                node.text = updated
                replacements += changed
                _invalidate_paragraph_layout(node)
                layout_invalidated = True
                self.mark_modified()
                if remaining >= 0:
                    remaining -= changed
                    if remaining <= 0:
                        break
        if replacements and not layout_invalidated:
            _invalidate_paragraph_layout(self._root)
        return replacements


class GenericXmlPart(XmlPart):
    pass


class VersionPart(XmlPart):
    pass


class ContainerPart(XmlPart):
    def rootfile_paths(self) -> list[str]:
        return [
            str(node)
            for node in self._root.xpath("./ocf:rootfiles/ocf:rootfile/@full-path", namespaces=NS)
        ]

    def ensure_rootfile(self, full_path: str, media_type: str) -> None:
        rootfiles = self._root.xpath("./ocf:rootfiles", namespaces=NS)
        if rootfiles:
            rootfiles_el = rootfiles[0]
        else:
            rootfiles_el = etree.SubElement(self._root, qname("ocf", "rootfiles"))
        for node in rootfiles_el.xpath("./ocf:rootfile", namespaces=NS):
            if node.get("full-path") == full_path:
                node.set("media-type", media_type)
                self.mark_modified()
                return
        new_node = etree.SubElement(rootfiles_el, qname("ocf", "rootfile"))
        new_node.set("full-path", full_path)
        new_node.set("media-type", media_type)
        self.mark_modified()


class ManifestPart(XmlPart):
    pass


class ContainerRdfPart(XmlPart):
    pass


@dataclass
class DocumentMetadata:
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


class ContentHpfPart(XmlPart):
    @property
    def distribution(self) -> str | None:
        return self._root.get(qname("hpf", "distribution"))

    @property
    def is_distribution_protected(self) -> bool:
        return self.distribution == "1"

    def metadata(self) -> DocumentMetadata:
        metadata = DocumentMetadata()
        metadata.title = self._first_text("./opf:metadata/opf:title")
        metadata.language = self._first_text("./opf:metadata/opf:language")
        meta_nodes = self._root.xpath("./opf:metadata/opf:meta", namespaces=NS)
        for node in meta_nodes:
            name = node.get("name")
            value = node.text or ""
            if name == "creator":
                metadata.creator = value
            elif name == "subject":
                metadata.subject = value
            elif name == "description":
                metadata.description = value
            elif name == "lastsaveby":
                metadata.lastsaveby = value
            elif name == "CreatedDate":
                metadata.created = value
            elif name == "ModifiedDate":
                metadata.modified = value
            elif name == "date":
                metadata.date = value
            elif name == "keyword":
                metadata.keyword = value
            elif name:
                metadata.extra[name] = value
        return metadata

    def set_metadata(self, **values: str | None) -> None:
        metadata = self._ensure_metadata_element()
        if "title" in values:
            self._ensure_text_child(metadata, "opf:title", values["title"])
        if "language" in values:
            self._ensure_text_child(metadata, "opf:language", values["language"])

        meta_mapping = {
            "creator": "creator",
            "subject": "subject",
            "description": "description",
            "lastsaveby": "lastsaveby",
            "created": "CreatedDate",
            "modified": "ModifiedDate",
            "date": "date",
            "keyword": "keyword",
        }
        for key, meta_name in meta_mapping.items():
            if key in values:
                self._ensure_meta(metadata, meta_name, values[key])
        self.mark_modified()

    def manifest_items(self) -> list[etree._Element]:
        manifest = self._ensure_manifest_element()
        return list(manifest.xpath("./opf:item", namespaces=NS))

    def spine_itemrefs(self) -> list[etree._Element]:
        spine = self._ensure_spine_element()
        return list(spine.xpath("./opf:itemref", namespaces=NS))

    def section_items(self) -> list[etree._Element]:
        return [
            node
            for node in self.manifest_items()
            if re.match(SECTION_PATTERN, node.get("href", ""))
        ]

    def ensure_manifest_item(
        self,
        item_id: str,
        href: str,
        media_type: str,
        **extra_attributes: str,
    ) -> etree._Element:
        manifest = self._ensure_manifest_element()
        for item in self.manifest_items():
            if item.get("id") == item_id or item.get("href") == href:
                item.set("id", item_id)
                item.set("href", href)
                item.set("media-type", media_type)
                for key, value in extra_attributes.items():
                    item.set(key, str(value))
                self.mark_modified()
                return item

        item = etree.SubElement(manifest, qname("opf", "item"))
        item.set("id", item_id)
        item.set("href", href)
        item.set("media-type", media_type)
        for key, value in extra_attributes.items():
            item.set(key, str(value))
        self.mark_modified()
        return item

    def remove_manifest_item(self, href: str | None = None, item_id: str | None = None) -> None:
        manifest = self._ensure_manifest_element()
        for item in list(self.manifest_items()):
            if href and item.get("href") == href:
                manifest.remove(item)
                self.mark_modified()
            elif item_id and item.get("id") == item_id:
                manifest.remove(item)
                self.mark_modified()

    def ensure_spine_itemref(self, item_id: str, linear: str | None = None) -> etree._Element:
        spine = self._ensure_spine_element()
        for itemref in self.spine_itemrefs():
            if itemref.get("idref") == item_id:
                if linear is not None:
                    itemref.set("linear", linear)
                    self.mark_modified()
                return itemref

        itemref = etree.SubElement(spine, qname("opf", "itemref"))
        itemref.set("idref", item_id)
        if linear is not None:
            itemref.set("linear", linear)
        self.mark_modified()
        return itemref

    def remove_spine_itemref(self, item_id: str) -> None:
        spine = self._ensure_spine_element()
        for itemref in list(self.spine_itemrefs()):
            if itemref.get("idref") == item_id:
                spine.remove(itemref)
                self.mark_modified()

    def next_section_path(self) -> str:
        indexes = [
            int(match.group(1))
            for item in self.section_items()
            if (match := re.match(SECTION_PATTERN, item.get("href", "")))
        ]
        return f"Contents/section{max(indexes, default=-1) + 1}.xml"

    def next_manifest_id(self, prefix: str) -> str:
        existing = {item.get("id", "") for item in self.manifest_items()}
        if prefix not in existing:
            return prefix
        index = 1
        while f"{prefix}_{index}" in existing:
            index += 1
        return f"{prefix}_{index}"

    def _ensure_metadata_element(self) -> etree._Element:
        nodes = self._root.xpath("./opf:metadata", namespaces=NS)
        if nodes:
            return nodes[0]
        metadata = etree.Element(qname("opf", "metadata"))
        self._root.insert(0, metadata)
        self.mark_modified()
        return metadata

    def _ensure_manifest_element(self) -> etree._Element:
        nodes = self._root.xpath("./opf:manifest", namespaces=NS)
        if nodes:
            return nodes[0]
        manifest = etree.SubElement(self._root, qname("opf", "manifest"))
        self.mark_modified()
        return manifest

    def _ensure_spine_element(self) -> etree._Element:
        nodes = self._root.xpath("./opf:spine", namespaces=NS)
        if nodes:
            return nodes[0]
        spine = etree.SubElement(self._root, qname("opf", "spine"))
        self.mark_modified()
        return spine

    def _ensure_text_child(self, parent: etree._Element, tag: str, value: str | None) -> None:
        resolved = qname(*tag.split(":", 1))
        nodes = parent.xpath(f"./{tag}", namespaces=NS)
        if nodes:
            nodes[0].text = value
        else:
            child = etree.SubElement(parent, resolved)
            child.text = value
        self.mark_modified()

    def _ensure_meta(self, parent: etree._Element, name: str, value: str | None) -> None:
        for meta in parent.xpath("./opf:meta", namespaces=NS):
            if meta.get("name") == name:
                meta.text = value
                if meta.get("content") is None:
                    meta.set("content", "text")
                self.mark_modified()
                return
        meta = etree.SubElement(parent, qname("opf", "meta"))
        meta.set("name", name)
        meta.set("content", "text")
        meta.text = value
        self.mark_modified()

    def _first_text(self, expression: str) -> str | None:
        values = self._root.xpath(expression, namespaces=NS)
        if not values:
            return None
        node = values[0]
        if isinstance(node, etree._Element):
            return node.text
        return str(node)


class HeaderPart(XmlPart):
    @property
    def section_count(self) -> int:
        value = self._root.get("secCnt")
        return int(value) if value else 0

    def set_section_count(self, count: int) -> None:
        self._root.set("secCnt", str(count))
        self.mark_modified()


class SettingsPart(XmlPart):
    pass


class SectionPart(XmlPart):
    def section_index(self) -> int:
        match = re.match(SECTION_PATTERN, self.path)
        if match is None:
            raise ValueError(f"Not a section path: {self.path}")
        return int(match.group(1))

    def paragraphs(self) -> list[HwpxXmlNode]:
        return [HwpxXmlNode(node, self) for node in self._root.xpath("./hp:p", namespaces=NS)]

    def text_fragments(self) -> list[str]:
        values = []
        for paragraph in self._root.xpath("./hp:p", namespaces=NS):
            chunks = [node.text or "" for node in paragraph.xpath(".//hp:t", namespaces=NS)]
            values.append("".join(chunks))
        return values

    def extract_text(self, paragraph_separator: str = "\n") -> str:
        return paragraph_separator.join(self.text_fragments())

    def append_paragraph(
        self,
        text: str,
        *,
        template_index: int | None = None,
        para_pr_id: str | None = None,
        style_id: str | None = None,
        char_pr_id: str | None = None,
    ) -> HwpxXmlNode:
        paragraph = self._build_text_paragraph(
            text,
            template_index=template_index,
            para_pr_id=para_pr_id,
            style_id=style_id,
            char_pr_id=char_pr_id,
        )
        self._root.append(paragraph)
        self.mark_modified()
        return HwpxXmlNode(paragraph, self)

    def insert_paragraph(
        self,
        index: int,
        text: str,
        *,
        template_index: int | None = None,
        para_pr_id: str | None = None,
        style_id: str | None = None,
        char_pr_id: str | None = None,
    ) -> HwpxXmlNode:
        paragraphs = self._root.xpath("./hp:p", namespaces=NS)
        paragraph = self._build_text_paragraph(
            text,
            template_index=template_index,
            para_pr_id=para_pr_id,
            style_id=style_id,
            char_pr_id=char_pr_id,
        )
        if index < 0:
            index = max(len(paragraphs) + index, 0)
        if index >= len(paragraphs):
            self._root.append(paragraph)
        else:
            paragraphs[index].addprevious(paragraph)
        self.mark_modified()
        return HwpxXmlNode(paragraph, self)

    def set_paragraph_text(
        self,
        index: int,
        text: str,
        *,
        char_pr_id: str | None = None,
    ) -> HwpxXmlNode:
        paragraphs = self._root.xpath("./hp:p", namespaces=NS)
        paragraph = paragraphs[index]
        expected_signature = _preserved_structure_signature(paragraph)
        _replace_paragraph_text_preserving_controls(paragraph, text, char_pr_id=char_pr_id)
        if self.document is not None:
            missing = _missing_preserved_tokens(expected_signature, _preserved_structure_signature(paragraph))
            if expected_signature:
                self.document._track_control_preservation_targets(
                    [paragraph],
                    [expected_signature],
                    issues=[
                        ValidationIssue(
                            kind="control_preservation",
                            message="",
                            part_path=self.path,
                            section_index=self.section_index(),
                            paragraph_index=index,
                        )
                    ],
                )
            if missing:
                self.document._record_control_preservation_errors(
                    [
                        ValidationIssue(
                            kind="control_preservation",
                            message=f"lost preserved nodes during set_paragraph_text: {', '.join(missing)}",
                            part_path=self.path,
                            section_index=self.section_index(),
                            paragraph_index=index,
                        )
                    ]
                )
        _invalidate_paragraph_layout(paragraph)
        self.mark_modified()
        return HwpxXmlNode(paragraph, self)

    def delete_paragraph(self, index: int) -> None:
        paragraphs = self._root.xpath("./hp:p", namespaces=NS)
        if len(paragraphs) <= 1:
            raise ValueError("A section must keep at least one paragraph.")
        self._root.remove(paragraphs[index])
        self.mark_modified()

    def _build_text_paragraph(
        self,
        text: str,
        *,
        template_index: int | None = None,
        para_pr_id: str | None = None,
        style_id: str | None = None,
        char_pr_id: str | None = None,
    ) -> etree._Element:
        template = self._select_template_paragraph(template_index=template_index)
        protected = _preserved_structure_signature(template)
        if template_index is not None and protected:
            protected_tokens = ", ".join(protected.elements())
            raise ValueError(
                "template_index refers to a paragraph containing preserved controls and cannot be used "
                f"as a text paragraph template: {protected_tokens}"
            )
        paragraph = deepcopy(template)
        paragraph.set("id", self._next_paragraph_id())
        if para_pr_id is not None:
            paragraph.set("paraPrIDRef", para_pr_id)
        if style_id is not None:
            paragraph.set("styleIDRef", style_id)
        for child in list(paragraph):
            paragraph.remove(child)
        run = etree.SubElement(paragraph, qname("hp", "run"))
        run.set("charPrIDRef", char_pr_id or self._default_char_pr_id(template))
        text_node = etree.SubElement(run, qname("hp", "t"))
        text_node.text = text
        return paragraph

    def _select_template_paragraph(self, template_index: int | None = None) -> etree._Element:
        paragraphs = self._root.xpath("./hp:p", namespaces=NS)
        if not paragraphs:
            paragraph = etree.Element(qname("hp", "p"))
            paragraph.set("id", "0")
            paragraph.set("paraPrIDRef", "0")
            paragraph.set("styleIDRef", "0")
            paragraph.set("pageBreak", "0")
            paragraph.set("columnBreak", "0")
            paragraph.set("merged", "0")
            return paragraph
        if template_index is not None:
            return paragraphs[template_index]
        safe_text_paragraphs = [
            paragraph
            for paragraph in reversed(paragraphs)
            if paragraph.xpath(".//hp:t", namespaces=NS) and not _preserved_structure_signature(paragraph)
        ]
        if safe_text_paragraphs:
            return safe_text_paragraphs[0]
        for paragraph in reversed(paragraphs):
            if paragraph.xpath(".//hp:t", namespaces=NS):
                return paragraph
        for paragraph in reversed(paragraphs):
            if not _preserved_structure_signature(paragraph):
                return paragraph
        return paragraphs[-1]

    def _next_paragraph_id(self) -> str:
        max_id = -1
        for paragraph in self._root.xpath("./hp:p", namespaces=NS):
            value = paragraph.get("id")
            if value is None:
                continue
            try:
                numeric = int(value)
            except ValueError:
                continue
            max_id = max(max_id, numeric)
        return str(max_id + 1)

    def _default_char_pr_id(self, paragraph: etree._Element) -> str:
        runs = paragraph.xpath("./hp:run", namespaces=NS)
        for run in runs:
            value = run.get("charPrIDRef")
            if value is not None:
                return value
        return "0"


def infer_part_class(path: str, raw_bytes: bytes) -> type[HwpxPart]:
    normalized = PurePosixPath(path).as_posix()
    if normalized == "mimetype":
        return MimetypePart
    if normalized == "version.xml":
        return VersionPart
    if normalized == "Contents/content.hpf":
        return ContentHpfPart
    if normalized == "Contents/header.xml":
        return HeaderPart if _looks_like_xml(raw_bytes) else GenericBinaryPart
    if re.match(SECTION_PATTERN, normalized):
        return SectionPart if _looks_like_xml(raw_bytes) else GenericBinaryPart
    if normalized == "settings.xml":
        return SettingsPart if _looks_like_xml(raw_bytes) else GenericBinaryPart
    if normalized == "META-INF/container.xml":
        return ContainerPart
    if normalized == "META-INF/manifest.xml":
        return ManifestPart
    if normalized == "META-INF/container.rdf":
        return ContainerRdfPart
    if normalized == "Preview/PrvText.txt":
        return PreviewTextPart
    if normalized.startswith("Preview/") and normalized.endswith((".png", ".jpg", ".jpeg", ".bmp", ".gif")):
        return PreviewImagePart
    if normalized.startswith("BinData/"):
        return BinaryDataPart
    if normalized.startswith("Scripts/"):
        return ScriptPart
    if normalized.endswith((".xml", ".rdf", ".hpf")):
        return GenericXmlPart
    guessed, _ = mimetypes.guess_type(normalized)
    if guessed and guessed.startswith("text/"):
        return GenericTextPart
    return GenericBinaryPart
