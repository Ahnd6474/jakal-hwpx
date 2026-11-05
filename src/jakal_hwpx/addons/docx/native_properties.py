"""Property-only HWPX records attached to the real, editable Word objects.

No source package, body XML, text, image data or old equation operands are saved.
Formatting is restored only while the projected Word properties still match.
"""
from __future__ import annotations

from dataclasses import fields
from hashlib import sha256
import base64
import json

from lxml import etree

from ..._hancom_models import CharacterStyle, ParagraphStyle, Paragraph, StyleDefinition

NS = "urn:jakal-hwpx:native-properties:1"
PART = "/customXml/jakal-hwpx-properties.xml"


def encode(value):
    if isinstance(value, bytes):
        return {"bytes": base64.b64encode(value).decode("ascii")}
    if isinstance(value, dict):
        return {"map": [[encode(k), encode(v)] for k, v in value.items()]}
    if isinstance(value, tuple):
        return {"tuple": [encode(item) for item in value]}
    if isinstance(value, list):
        return [encode(item) for item in value]
    return value


def decode(value):
    if isinstance(value, dict):
        if "tuple" in value:
            return tuple(decode(item) for item in value["tuple"])
        if "bytes" in value:
            return base64.b64decode(value["bytes"], validate=True)
        return {tuple(key) if isinstance(key := decode(k), list) else key: decode(v) for k, v in value["map"]}
    if isinstance(value, list):
        return [decode(item) for item in value]
    return value


def properties(value, *, exclude=()):
    return {item.name: getattr(value, item.name) for item in fields(value) if item.name not in exclude}


def digest(nodes):
    return sha256(b"".join(etree.tostring(node, method="c14n", exclusive=True) for node in nodes if node is not None)).hexdigest()


def formatting(node):
    """Hash formatting/geometry, never the current visible content."""
    from docx.oxml.ns import qn
    name = etree.QName(node).localname
    if name == "tbl":
        return digest(node.xpath("./w:tblPr | ./w:tblGrid | ./w:tr/w:trPr | ./w:tr/w:tc/w:tcPr"))
    if name == "sectPr":
        return digest([child for child in node if etree.QName(child).localname not in {"headerReference", "footerReference"}])
    if name == "oMath":
        return digest([child for child in node.iter() if etree.QName(child).localname.endswith("Pr")])
    return digest([node.find(qn("w:" + {"p": "pPr", "r": "rPr"}.get(name, "not-a-property")))])


def paragraph_signature(node):
    """Guard native empty glyph positions against text or inline format edits."""
    from .identity import NS as identity_ns, W, private_ids
    anchors = private_ids(node)

    def value(element):
        if element.tag in {W + "bookmarkStart", W + "bookmarkEnd"} and element.get(W + "id") in anchors:
            return None
        attrs = sorted((key, item) for key, item in element.attrib.items()
                       if etree.QName(key).namespace != identity_ns
                       and etree.QName(key).localname not in {"paraId", "textId"}
                       and not etree.QName(key).localname.startswith("rsid"))
        return [element.tag, attrs, element.text or "", [item for child in element if (item := value(child)) is not None]]

    return sha256(json.dumps(value(node), ensure_ascii=True).encode("utf-8")).hexdigest()


class Capture:
    def __init__(self, source):
        self.source = source
        self.records: dict[str, dict] = {}
        resources = getattr(source, "_word_resources", {})
        self.enabled = source.source_format == "hwpx" and (not resources or resources.get("native_hwpx_origin", False))

    def named_properties(self, named):
        character = next((item for item in self.source.character_styles if item.style_id == named.char_pr_id), None)
        paragraph = next((item for item in self.source.paragraph_styles if item.style_id == named.para_pr_id), None)
        return {"style": properties(named), "character": properties(character) if character else None,
                "paragraph": properties(paragraph) if paragraph else None}

    def add(self, node, kind, values):
        if self.enabled:
            from .identity import identify
            self.records[identify(node)] = {"kind": kind, "format": formatting(node), "values": values}

    def paragraph(self, node, block, style):
        if self.enabled and style is not None:
            character = next((item for item in self.source.character_styles if item.style_id == block.char_pr_id), None)
            named = next((item for item in self.source.style_definitions if item.style_id == block.style_id), None)
            self.add(node, "paragraph", {"style": properties(style),
                "character": properties(character) if character else None,
                "named": self.named_properties(named) if named else None,
                "attributes": getattr(block, "_hwpx_paragraph_attributes", {}),
                "positions": {key: getattr(block, key) for key in ("source_paragraph_index", "nested_paragraph_index") if hasattr(block, key)}})
            empty = getattr(block, "_hwpx_empty_runs", getattr(block, "_word_extra", {}).get("native_empty_runs"))
            if empty is not None:
                from .identity import ID
                characters = {item.style_id: item for item in self.source.character_styles}
                self.records[node.get(ID)]["values"]["empty_runs"] = {
                    "text_hash": empty["text_hash"],
                    "runs": [dict(item, character=properties(characters[item["character"]])) for item in empty["runs"]],
                }

    def character(self, node, style):
        if self.enabled:
            self.add(node, "character", properties(style))

    def equation(self, node, block):
        if self.enabled:
            character = next((item for item in self.source.character_styles if item.style_id == getattr(block, "_hwpx_char_pr_id", None)), None)
            self.add(node, "equation", {"properties": properties(block, exclude=("script",)),
                "character": properties(character) if character else None,
                "baseline": getattr(block, "_hwpx_equation_baseline", None),
                "content_signature": paragraph_signature(node),
                "content_hash": sha256("".join(node.itertext("{http://schemas.openxmlformats.org/officeDocument/2006/math}t")).encode("utf-8")).hexdigest()})

    def table(self, node, block):
        if not self.enabled:
            return
        anchor = getattr(block, "_hwpx_anchor_paragraph", None)
        para = next((item for item in self.source.paragraph_styles if anchor is not None and item.style_id == anchor.para_pr_id), None)
        char = next((item for item in self.source.character_styles if anchor is not None and item.style_id == anchor.char_pr_id), None)
        named = next((item for item in self.source.style_definitions if anchor is not None and item.style_id == anchor.style_id), None)
        self.add(node, "table", {"properties": properties(block, exclude=("cell_texts",)),
            "anchor": {"paragraph": properties(para) if para else None, "character": properties(char) if char else None,
                       "named": self.named_properties(named) if named else None},
            "geometry": getattr(block, "_hwpx_table_properties", {}),
            "positions": {key: getattr(block, key) for key in ("source_paragraph_index", "nested_paragraph_index") if hasattr(block, key)}})

    def section(self, node, settings, *, section=None, source=None):
        if self.enabled:
            from .identity import ID
            node.attrib.pop(ID, None)  # add_section clones the preceding sectPr.
            values = properties(settings)
            absent = []
            for kind in ("header", "footer"):
                for variant, prefix, types in (("default", "", {"BOTH", "ODD"}), ("even", "even_page_", {"BOTH", "EVEN"}), ("first", "first_page_", {"FIRST"})):
                    if not any(item.kind == kind and item.apply_page_type in types for item in source.header_footer_blocks):
                        story = getattr(section, prefix + kind)
                        absent.append({"kind": kind, "variant": variant, "hash": digest([story._element])})
            values["generated_stories"] = absent
            self.add(node, "section", values)

    def save(self, word):
        if not self.enabled or not self.records:
            return
        from docx.opc.part import Part
        from docx.opc.packuri import PackURI
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        from .identity import ID, W
        for part in word.part.package.parts:
            element = getattr(part, "element", None)
            if element is not None:
                for paragraph in element.iter(W + "p"):
                    record = self.records.get(paragraph.get(ID), {})
                    if "empty_runs" in record.get("values", {}):
                        record["values"]["empty_signature"] = paragraph_signature(paragraph)
        root = etree.Element("{" + NS + "}properties")
        root.text = json.dumps(encode({"records": self.records, "styles_hash": digest([word.styles.element]),
            "named_catalog": [self.named_properties(item) for item in self.source.style_definitions],
            "defaults": {kind: properties(next(item for item in collection if item.style_id == "0"))
                         for kind, collection in (("character", self.source.character_styles), ("paragraph", self.source.paragraph_styles), ("named", self.source.style_definitions))
                         if any(item.style_id == "0" for item in collection)},
            "version": getattr(self.source, "_hwpx_version_attributes", {}),
            "application_settings": getattr(self.source, "_hwpx_application_settings", []),
            "resources": getattr(self.source, "_hwpx_ref_resources", [])}), ensure_ascii=True)
        part = Part(PackURI(PART), "application/xml", etree.tostring(root), word.part.package)
        word.part.relate_to(part, RT.CUSTOM_XML)


class Restore:
    def __init__(self, word, target):
        self.target = target
        self.records: dict[str, dict] = {}
        self.styles: dict[tuple[str, str], str] = {}
        self.named_map: dict[str, str] = {}
        self.named_catalog: list[dict] = []
        for part in word.part.package.parts:
            if str(part.partname) != PART:
                continue
            root = etree.fromstring(part.blob, parser=etree.XMLParser(resolve_entities=False, no_network=True))
            if root.tag != "{" + NS + "}properties":
                raise ValueError("Invalid native HWPX property record.")
            data = decode(json.loads(root.text or "{}"))
            target._hwpx_version_attributes = data.get("version", {})
            target._hwpx_application_settings = data.get("application_settings", [])
            if data.get("styles_hash") == digest([word.styles.element]):
                self.records = data["records"]
                self.named_catalog = data.get("named_catalog", [])
                target._hwpx_ref_resources = data.get("resources", [])
                for kind, value in data.get("defaults", {}).items():
                    cls = {"character": CharacterStyle, "paragraph": ParagraphStyle, "named": StyleDefinition}[kind]
                    collection = {"character": target.character_styles, "paragraph": target.paragraph_styles, "named": target.style_definitions}[kind]
                    collection[:] = [item for item in collection if item.style_id != "0"]
                    collection.append(cls(**value))
                    if kind != "named":
                        self.styles[kind, json.dumps(encode(value), sort_keys=True)] = "0"
                    else:
                        self.named_map["0"] = "0"
            else:
                target.import_warnings.append("Word styles changed; native HWPX property fallbacks were discarded to preserve the current Word formatting.")

    def get(self, node, kind):
        from .identity import ID
        record = self.records.get(node.get(ID), {})
        if record.get("kind") == kind and record.get("format") == formatting(node):
            return record["values"]
        return None

    def style(self, kind, value):
        key = kind, json.dumps(encode(value), sort_keys=True)
        if key not in self.styles:
            collection = self.target.character_styles if kind == "character" else self.target.paragraph_styles
            identifier = str(max([int(item.style_id) for item in collection if str(item.style_id).isdigit()] or [0]) + 1)
            value = dict(value, style_id=identifier)
            if value.get("native_hwpx_xml"):
                node = etree.fromstring(value["native_hwpx_xml"])
                node.set("id", identifier)
                value["native_hwpx_xml"] = etree.tostring(node)
            collection.append((CharacterStyle if kind == "character" else ParagraphStyle)(**value))
            self.styles[key] = identifier
        return self.styles[key]

    def character(self, node):
        value = self.get(node, "character")
        return self.style("character", value) if value is not None else None

    def named(self, value):
        original_id = value["style"]["style_id"]
        if original_id in self.named_map:
            return self.named_map[original_id]
        value = next((item for item in self.named_catalog if item["style"]["style_id"] == original_id), value)
        key = "named", json.dumps(encode(value), sort_keys=True)
        if key not in self.styles:
            collection = self.target.style_definitions
            identifier = str(max([int(item.style_id) for item in collection if str(item.style_id).isdigit()] or [0]) + 1)
            record = dict(value["style"], style_id=identifier, next_style_id=identifier,
                para_pr_id=self.style("paragraph", value["paragraph"]) if value.get("paragraph") else None,
                char_pr_id=self.style("character", value["character"]) if value.get("character") else None)
            collection.append(StyleDefinition(**record))
            self.styles[key] = identifier
            self.named_map[original_id] = identifier
        return self.styles[key]

    def finish(self):
        for value in self.named_catalog:
            self.named(value)
        for value in self.named_catalog:
            style = next(item for item in self.target.style_definitions if item.style_id == self.named_map[value["style"]["style_id"]])
            style.para_pr_id = self.style("paragraph", value["paragraph"]) if value.get("paragraph") else None
            style.char_pr_id = self.style("character", value["character"]) if value.get("character") else None
            style.next_style_id = self.named_map.get(value["style"].get("next_style_id"), style.style_id)

    def equation(self, node, block):
        value = self.get(node, "equation")
        if value is not None:
            current = sha256("".join(node.itertext("{http://schemas.openxmlformats.org/officeDocument/2006/math}t")).encode("utf-8")).hexdigest()
            unchanged = (paragraph_signature(node) == value["content_signature"] if "content_signature" in value
                         else current == value["content_hash"])
            if unchanged and "baseline" in value:
                setattr(block, "_hwpx_equation_baseline", value["baseline"])
            if value.get("character"):
                setattr(block, "_hwpx_char_pr_id", self.style("character", value["character"]))
            for key, item in value["properties"].items():
                if unchanged or key not in {"width", "height", "shape_comment"}:
                    setattr(block, key, item)

    def paragraph(self, node, blocks):
        value = self.get(node, "paragraph")
        if value is None:
            return
        identifier = self.style("paragraph", value["style"])
        empty = value.get("empty_runs")
        if empty is not None and value.get("empty_signature") == paragraph_signature(node):
            empty = {"text_hash": empty["text_hash"],
                     "runs": [dict(item, character=self.style("character", item["character"])) for item in empty["runs"]]}
        else:
            empty = None
        for block in blocks:
            if hasattr(block, "para_pr_id"):
                block.para_pr_id = identifier
                if value.get("named"):
                    block.style_id = self.named(value["named"])
            if isinstance(block, Paragraph) and not block.text and block.char_pr_id is None and value.get("character"):
                block.char_pr_id = self.style("character", value["character"])
            for key, item in value["positions"].items():
                setattr(block, key, item)
            setattr(block, "_hwpx_paragraph_attributes", value["attributes"])
            if isinstance(block, Paragraph) and empty is not None:
                extra = dict(getattr(block, "_word_extra", {}), native_empty_runs=empty)
                setattr(block, "_word_extra", extra)

    def table(self, node, block):
        value = self.get(node, "table")
        if value is None:
            return
        for key, item in value["properties"].items():
            setattr(block, key, item)
        for key, item in value["positions"].items():
            setattr(block, key, item)
        setattr(block, "_hwpx_table_properties", value["geometry"])
        anchor = value.get("anchor", {})
        if anchor.get("paragraph"):
            paragraph = Paragraph("", para_pr_id=self.style("paragraph", anchor["paragraph"]),
                char_pr_id=self.style("character", anchor["character"]) if anchor.get("character") else None)
            if anchor.get("named"):
                named = anchor["named"]
                if "style" not in named:
                    named = {"style": named, "paragraph": anchor["paragraph"], "character": anchor.get("character")}
                paragraph.style_id = self.named(named)
            setattr(block, "_hwpx_anchor_paragraph", paragraph)

    def section(self, node, settings, *, section=None, source=None):
        value = self.get(node, "section")
        if value is not None:
            for key, item in value.items():
                if key != "generated_stories":
                    setattr(settings, key, item)
            for record in value.get("generated_stories", []):
                variant = record["variant"]
                prefix = {"default": "", "even": "even_page_", "first": "first_page_"}[variant]
                if digest([getattr(source, prefix + record["kind"])._element]) == record["hash"]:
                    types = {"BOTH", "ODD"} if variant == "default" else {"EVEN"} if variant == "even" else {"FIRST"}
                    section.header_footer_blocks[:] = [story for story in section.header_footer_blocks
                        if not (story.kind == record["kind"] and story.apply_page_type in types)]
