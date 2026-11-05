"""Keep native-origin Word property carriers out of Hancom's line composer.

The extension stores only empty technical fields and their attachment points.
Body text and paragraph layout caches are never copied. Native content remains
editable; changed paragraphs discard obsolete attachment records on import.
"""
from __future__ import annotations

from copy import deepcopy
from hashlib import sha256
import json

from lxml import etree

from ...namespaces import qname

PART = "META-INF/jakal-word-properties.xml"
NS = "urn:jakal-hwpx:word-properties:1"


def signature(node):
    def value(element):
        return [element.tag, sorted((key, item) for key, item in element.attrib.items() if key != "id"),
                element.text or "", [value(child) for child in element if child.tag != qname("hp", "linesegarray")]]
    return sha256(json.dumps(value(node), ensure_ascii=True).encode("utf-8")).hexdigest()


def store_carriers(document):
    from ..._hancom_hwpx_write import _word_effect_from_field_name
    records = []
    for index, section in enumerate(document.sections):
        pending = []
        for paragraph in section.root_element.iter(qname("hp", "p")):
            groups = []
            group: list[str] = []
            identifier, position = None, 0
            for run in list(paragraph):
                begin = run.find(qname("hp", "ctrl") + "/" + qname("hp", "fieldBegin"))
                if begin is not None and (_word_effect_from_field_name(begin.get("name")) or {}).get("kind") == "JAKAL_WORD_FEATURE":
                    if identifier is not None:
                        raise ValueError("Nested Word property fields cannot be detached.")
                    identifier = begin.get("id")
                    group = []
                if identifier is None:
                    position += 1
                    continue
                group.append(etree.tostring(run, encoding="unicode"))
                paragraph.remove(run)
                end = run.find(qname("hp", "ctrl") + "/" + qname("hp", "fieldEnd"))
                if end is not None and end.get("beginIDRef") == identifier:
                    groups.append({"position": position, "runs": group})
                    identifier = None
            if identifier is not None:
                raise ValueError("Unclosed Word property field.")
            if groups:
                pending.append((paragraph, groups))
        # Resolve paths after all nested fields have been removed.
        for paragraph, groups in pending:
            path, node = [], paragraph
            while node is not section.root_element:
                parent = node.getparent()
                path.append(parent.index(node))
                node = parent
            records.append({"section": index, "path": path[::-1], "hash": signature(paragraph), "groups": groups})
        if pending:
            section.mark_modified()
    if records:
        root = etree.Element("{" + NS + "}properties", version="1", headerHash=signature(document.header.readonly_root_element))
        root.text = json.dumps(records, ensure_ascii=True)
        document.add_part(PART, etree.tostring(root, encoding="UTF-8", xml_declaration=True))


def inflate_carriers(document):
    """Return a reading copy; the caller's native document stays field-free."""
    try:
        part = document.get_part(PART)
    except KeyError:
        return document, []
    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    root = etree.fromstring(part.to_bytes(), parser=parser)
    if root.tag != "{" + NS + "}properties" or root.get("version") != "1":
        raise ValueError("Invalid HWPX Word property extension.")
    if root.get("headerHash") != signature(document.header.readonly_root_element):
        return document, ["Native HWPX styles changed; obsolete Word property attachments were discarded."]
    result = deepcopy(document)
    targets, skipped = [], 0
    for record in json.loads(root.text or "[]"):
        try:
            paragraph = result.sections[record["section"]].root_element
            for index in record["path"]:
                paragraph = paragraph[index]
        except IndexError:
            skipped += 1
            continue
        if paragraph.tag != qname("hp", "p") or signature(paragraph) != record["hash"]:
            skipped += 1
            continue
        targets.append((paragraph, record["groups"]))
    # Resolve every nested path before inserting fields into ancestor paragraphs.
    for paragraph, groups in targets:
        inserted = 0
        for group in groups:
            for xml in group["runs"]:
                run = etree.fromstring(xml.encode("utf-8"), parser=parser)
                if run.tag != qname("hp", "run") or any((node.text or "") or len(node) for node in run.iter(qname("hp", "t"))):
                    raise ValueError("A Word property attachment cannot contain visible content.")
                paragraph.insert(group["position"] + inserted, run)
                inserted += 1
    return result, ([f"Native HWPX content changed; {skipped} obsolete Word property attachments were discarded."] if skipped else [])
