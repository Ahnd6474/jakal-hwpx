"""Small reversible Word property records, never an embedded source document.

Visible content remains in native HWPX objects. Empty, ordinary CLICK_HERE fields
carry properties which have no equivalent in the shared model. Their names also
survive Hancom's field-parameter normalization.
"""
from __future__ import annotations

from copy import copy, deepcopy
import json
from typing import Any

from lxml import etree

from ..._hancom_models import Field, Paragraph, Table

FEATURE = "JAKAL_WORD_FEATURE"


def xml(element: Any) -> str | None:
    return etree.tostring(element, encoding="unicode") if element is not None else None


def parse(value: str | bytes) -> Any:
    from docx.oxml import parse_xml

    return parse_xml(value.encode("utf-8") if isinstance(value, str) else value)


def set_property(parent: Any, node: Any) -> None:
    for old in list(parent):
        if old.tag == node.tag:
            parent.remove(old)
    parent.append(deepcopy(node))


def children(block: Any):
    if isinstance(block, Table):
        for baseline, values in getattr(block, "_docx_cell_content", {}).values():
            yield values
    if getattr(block, "nested_blocks", None):
        yield block.nested_blocks


def carrier_model(document: Any) -> Any:
    result = copy(document)
    result.sections = deepcopy(document.sections)
    result._word_carriers_prepared = True

    def marker(properties, block=None):
        field = Field(FEATURE, display_text="", parameters={"Record": json.dumps(properties, ensure_ascii=True)})
        if block is not None:
            for key in ("source_paragraph_index", "nested_paragraph_index"):
                if hasattr(block, key):
                    setattr(field, key, getattr(block, key))
        return field

    def prepare(values):
        output = []
        for block in values:
            for nested in children(block):
                nested[:] = prepare(nested)
            properties = getattr(block, "_word_extra", None)
            if properties:
                record = {"kind": "block", "type": type(block).__name__, "properties": properties}
                if isinstance(block, Paragraph) and not block.text:
                    record["empty_paragraph"] = {"para_pr_id": block.para_pr_id, "style_id": block.style_id,
                                                  "char_pr_id": block.char_pr_id}
                output.append(marker(record, block))
            output.append(block)
        for index, block in enumerate(output):
            setattr(block, "source_order", index)
            setattr(block, "nested_source_order", index)
        return output

    for section in result.sections:
        section.blocks = prepare(section.blocks)
        for story in section.header_footer_blocks:
            if getattr(story, "nested_blocks", None):
                story.nested_blocks = prepare(story.nested_blocks)
    if result.sections:
        values = result.sections[0].blocks
        first = values[0] if values else None
        prefixes = []
        for style in document.character_styles:
            if getattr(style, "_word_extra", None):
                field = marker({"kind": "character", "id": style.style_id, "properties": style._word_extra}, first)
                setattr(field, "_docx_char_pr_id", style.style_id)
                prefixes.append(field)
        resources = getattr(document, "_word_resources", None)
        if resources:
            prefixes.append(marker({"kind": "resources", "properties": resources}, first))
        for index, field in enumerate(prefixes):
            setattr(field, "source_order", -len(prefixes) + index)
        values[:0] = prefixes
    return result


def restore_carriers(document: Any) -> None:
    characters = {style.style_id: style for style in document.character_styles}

    def restore(values):
        output = []
        pending = []
        for block in values:
            if isinstance(block, Field) and block.field_type == FEATURE:
                try:
                    record = json.loads(str(block.parameters["Record"]))
                except (KeyError, ValueError, TypeError) as error:
                    raise ValueError("Invalid Word feature property record.") from error
                kind = record.get("kind")
                if kind == "character":
                    style_id = getattr(block, "_docx_char_pr_id", record.get("id"))
                    if style_id in characters:
                        setattr(characters[style_id], "_word_extra", record["properties"])
                elif kind == "resources":
                    document._word_resources = record["properties"]
                elif kind == "block":
                    if "empty_paragraph" in record:
                        paragraph = Paragraph("", **record["empty_paragraph"])
                        for attribute, key in (("para_pr_id", "_word_native_para_pr_id"), ("style_id", "_word_native_style_id"),
                                               ("char_pr_id", "_word_native_text_char_id")):
                            if getattr(block, key, None) is not None:
                                setattr(paragraph, attribute, getattr(block, key))
                        setattr(paragraph, "_word_extra", record["properties"])
                        for key in ("source_paragraph_index", "source_order", "nested_paragraph_index", "nested_source_order"):
                            if hasattr(block, key):
                                setattr(paragraph, key, getattr(block, key))
                        output.append(paragraph)
                    else:
                        pending.append(record)
                else:
                    raise ValueError("Unknown Word feature property record.")
                continue
            for record in list(pending):
                if type(block).__name__ == record.get("type"):
                    setattr(block, "_word_extra", record["properties"])
                    pending.remove(record)
                    break
            for nested in children(block):
                nested[:] = restore(nested)
            output.append(block)
        if pending:
            # A deleted native object must not cause its old properties to be
            # silently attached to an unrelated object of a different kind.
            document.import_warnings.append("A Word feature carrier has no corresponding native object.")
        return output

    for section in document.sections:
        section.blocks = restore(section.blocks)
        for story in section.header_footer_blocks:
            if getattr(story, "nested_blocks", None):
                story.nested_blocks = restore(story.nested_blocks)
    document._word_feature_model = True
    from .native_effects import restore_list_styles

    restore_list_styles(document)


def capture_controls(root: Any) -> dict[Any, list[dict[str, Any]]]:
    from docx.oxml.ns import qn

    result: dict[Any, list[dict[str, Any]]] = {}
    for index, node in enumerate(root.iter(qn("w:sdt"))):
        content = node.find(qn("w:sdtContent"))
        if content is None:
            continue
        parent = node.getparent()
        while parent is not None and parent.tag in {qn("w:sdtContent"), qn("w:sdt")}:
            parent = parent.getparent()
        level = "inline" if parent.tag == qn("w:p") else "row" if parent.tag == qn("w:tbl") else "cell" if parent.tag == qn("w:tr") else "block"
        info = {"id": str(index), "level": level, "properties": xml(node.find(qn("w:sdtPr"))),
                "end": xml(node.find(qn("w:sdtEndPr")))}
        for element in content.iter():
            if level in {"row", "cell"}:
                if element.tag != qn("w:tr" if level == "row" else "w:tc") or element.getparent() is not content:
                    continue
            if level == "block":
                if element.tag not in {qn("w:p"), qn("w:tbl")}:
                    continue
                ancestor = element.getparent()
                nested = False
                while ancestor is not None and ancestor is not content:
                    if ancestor.tag in {qn("w:p"), qn("w:tbl"), qn("w:tc")}:
                        nested = True
                        break
                    ancestor = ancestor.getparent()
                if nested:
                    continue
            result.setdefault(element, []).append(info)
    return result


def wrap_controls(node: Any, controls: list[dict[str, Any]], *, level: str) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    parent = node.getparent()
    if parent is None:
        return
    # Move the real, edited output into wrappers; never restore stale text.
    for control in (item for item in controls if item["level"] == level):
        previous = node.getprevious()
        if previous is not None and previous.tag == qn("w:sdt") and previous.get("data-jakal-control") == control["id"]:
            wrapper = previous
        else:
            wrapper = OxmlElement("w:sdt")
            wrapper.set("data-jakal-control", control["id"])
            for key in ("properties", "end"):
                if control.get(key):
                    wrapper.append(parse(control[key]))
            wrapper.append(OxmlElement("w:sdtContent"))
            node.addprevious(wrapper)
        wrapper.find(qn("w:sdtContent")).append(node)
        parent = wrapper.find(qn("w:sdtContent"))
        # A following nested control is inserted into this content container.
        if node.getparent() is not parent:
            parent.append(node)


def clean_control_markers(word: Any) -> None:
    for part in word.part.package.parts:
        element = getattr(part, "element", None)
        if element is not None:
            for node in element.iter():
                node.attrib.pop("data-jakal-control", None)


def synchronize_bindings(word: Any, *, controls: list[Any] | None = None) -> None:
    """Update bound XML from the edited control, avoiding stale values on open."""
    import re
    from docx.oxml.ns import qn

    stores = {}
    for part in word.part.package.parts:
        if not str(part.partname).startswith("/customXml/") or "customXmlProperties" in part.content_type:
            continue
        for rel in part.rels.values():
            if rel.reltype.endswith("/customXmlProps") and not rel.is_external:
                props = parse(rel.target_part.blob)
                identity = props.get("{http://schemas.openxmlformats.org/officeDocument/2006/customXml}itemID", "").upper()
                stores[identity] = part
    if controls is None:
        controls = [node for part in word.part.package.parts if getattr(part, "element", None) is not None
                    for node in part.element.iter(qn("w:sdt"))]
    for node in controls:
        binding = node.find(f"./{qn('w:sdtPr')}/{qn('w:dataBinding')}")
        if binding is None:
            continue
        part = stores.get(binding.get(qn("w:storeItemID"), "").upper())
        if part is None:
            continue
        tree = parse(part.blob)
        namespaces = {name: value for name, quote, value in re.findall(r"xmlns:([\w.-]+)\s*=\s*(['\"])(.*?)\2", binding.get(qn("w:prefixMappings"), ""))}
        try:
            targets = tree.xpath(binding.get(qn("w:xpath"), ""), namespaces=namespaces)
        except etree.XPathError as error:
            raise ValueError("The content control has an invalid data-binding XPath.") from error
        if not isinstance(targets, list) or len(targets) != 1:
            raise ValueError("The content-control binding must select exactly one XML node.")
        content = node.find(qn("w:sdtContent"))
        visible = etree.XPath(".//w:t | .//w:tab | .//w:br", namespaces={"w": qn("w:p")[1:].split("}")[0]})(content)
        text = "".join("\t" if item.tag == qn("w:tab") else "\n" if item.tag == qn("w:br") else item.text or "" for item in visible)
        target = targets[0]
        if isinstance(target, etree._Element):
            if not len(target) and (target.text or "") == text:
                continue
            target.text = text
            for child in list(target):
                target.remove(child)
        elif getattr(target, "is_attribute", False):
            if str(target) == text:
                continue
            target.getparent().set(target.attrname, text)
        elif getattr(target, "is_text", False):
            if str(target) == text:
                continue
            target.getparent().text = text
        else:
            raise ValueError("The content-control binding must select an element, attribute, or text node.")
        part._blob = etree.tostring(tree, encoding="utf-8", xml_declaration=True, standalone=True)


def capture_resources(word: Any) -> dict:
    """Keep rendering compatibility, themes, table styles, and SDT bindings."""
    import base64
    from docx.oxml.ns import qn

    custom = {}
    for part in word.part.package.parts:
        if str(part.partname).startswith("/customXml/"):
            custom[str(part.partname)] = {
                "content_type": part.content_type, "data": base64.b64encode(part.blob).decode("ascii"),
                "rels": [{"id": rel.rId, "type": rel.reltype, "external": rel.is_external,
                          "target": rel.target_ref if rel.is_external else str(rel.target_part.partname)}
                         for rel in part.rels.values()],
            }
    theme = next((xml(parse(rel.target_part.blob)) for rel in word.part.rels.values()
                  if rel.reltype.endswith("/theme") and not rel.is_external), None)
    tables = {node.get(qn("w:styleId")): node for node in word.styles.element
              if node.tag == qn("w:style") and node.get(qn("w:type")) == "table"}
    # Default templates contain many unused table styles; carry referenced styles
    # and their base chains instead of embedding the entire style catalogue.
    used = set(word.element.xpath(".//w:tblStyle/@w:val"))
    selected: list[str] = []
    included: set[str] = set()
    while used:
        style_id = used.pop()
        node = tables.get(style_id)
        if node is None or style_id in included:
            continue
        selected.append(str(xml(node)))
        included.add(style_id)
        base = node.find(qn("w:basedOn"))
        if base is not None:
            used.add(base.get(qn("w:val")))
    # A fresh Word template can carry a border/spacing setting absent from the
    # source style. Retain the used non-list styles' direct paragraph properties,
    # including absence, rather than inheriting those fresh-template defaults.
    referenced = {value for part in word.part.package.parts if getattr(part, "element", None) is not None
                  for reference in part.element.iter(qn("w:pStyle")) if (value := reference.get(qn("w:val")))}
    paragraph_styles = []
    for style in word.styles:
        if style.element.get(qn("w:type")) != "paragraph":
            continue
        if style.style_id not in referenced and style.element.get(qn("w:default")) not in {"1", "true"}:
            continue
        cursor, numbered, seen = style, False, set()
        while cursor is not None and cursor.style_id not in seen:
            seen.add(cursor.style_id)
            numbered |= cursor.element.find(f"./{qn('w:pPr')}/{qn('w:numPr')}") is not None
            cursor = cursor.base_style
        if not numbered:
            paragraph_styles.append({"name": style.name, "pPr": xml(style.element.find(qn("w:pPr")))})
    return {"theme": theme, "table_styles": selected, "custom_xml": custom,
            "paragraph_styles": paragraph_styles,
            "compatibility": xml(word.settings.element.find(qn("w:compat"))),
            "custom_links": [{"type": rel.reltype, "target": str(rel.target_part.partname)}
                             for rel in word.part.rels.values() if not rel.is_external and rel.reltype.endswith("/customXml")]}


def restore_resources(word: Any, resources: dict) -> None:
    import base64
    from docx.opc.packuri import PackURI
    from docx.opc.part import Part
    from docx.oxml.ns import qn

    if "compatibility" in resources:
        for node in list(word.settings.element.findall(qn("w:compat"))):
            word.settings.element.remove(node)
        if resources["compatibility"]:
            word.settings.element.append(parse(resources["compatibility"]))
    if resources.get("theme"):
        for rel in word.part.rels.values():
            if rel.reltype.endswith("/theme") and not rel.is_external:
                rel.target_part._blob = resources["theme"].encode("utf-8")
    for value in resources.get("table_styles", []):
        node = parse(value)
        for existing in list(word.styles.element):
            if existing.tag == node.tag and existing.get(qn("w:styleId")) == node.get(qn("w:styleId")):
                word.styles.element.remove(existing)
        word.styles.element.append(node)
    for record in resources.get("paragraph_styles", []):
        from docx.enum.style import WD_STYLE_TYPE

        target = word.styles[record["name"]] if record["name"] in word.styles else word.styles.add_style(record["name"], WD_STYLE_TYPE.PARAGRAPH)
        if target.type != WD_STYLE_TYPE.PARAGRAPH:
            continue
        for previous in list(target.element.findall(qn("w:pPr"))):
            target.element.remove(previous)
        if record["pPr"]:
            # pPr must precede rPr in a style definition.
            properties = parse(record["pPr"])
            character = target.element.find(qn("w:rPr"))
            if character is not None:
                character.addprevious(properties)
            else:
                target.element.append(properties)
    records = resources.get("custom_xml", {})
    parts = {name: Part(PackURI(name), info["content_type"], base64.b64decode(info["data"]), word.part.package)
             for name, info in records.items()}
    for name, part in parts.items():
        for rel in records[name]["rels"]:
            target = rel["target"] if rel["external"] else parts.get(rel["target"])
            if target is not None:
                part.rels.add_relationship(rel["type"], target, rel["id"], is_external=rel["external"])
    # Remove template sample custom XML before connecting the actual bindings.
    for rid, rel in list(word.part.rels.items()):
        if rel.reltype.endswith("/customXml"):
            del word.part.rels[rid]
    for rel in resources.get("custom_links", []):
        if rel["target"] in parts:
            word.part.relate_to(parts[rel["target"]], rel["type"])
