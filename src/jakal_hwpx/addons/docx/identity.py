"""Persistent, markup-compatible identities for interactive Word objects."""
from __future__ import annotations

from uuid import uuid4
from hashlib import sha256
import json
import secrets
import re
from lxml import etree

NS = "urn:jakal-hwpx:word-editing:1"
ID = "{" + NS + "}id"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
TAGS = {"p", "tbl", "drawing", "hyperlink", "oMath", "bookmarkStart", "footnote", "endnote", "hdr", "ftr", "sectPr", "txbxContent", "num", "fldSimple", "fldChar", "sdt"}
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
PARA_ID = "{http://schemas.microsoft.com/office/word/2010/wordml}paraId"
LEDGER = "/customXml/jakal-word-identities.xml"
STORIES = {"document", "hdr", "ftr", "footnotes", "endnotes", "numbering"}


def internal_name(value):
    return bool(re.fullmatch(r"_JAKAL_[0-9a-f]{32}", value or ""))


def private_ids(paragraph):
    all_private = {node.get(W + "id") for node in paragraph.getroottree().getroot().iter(W + "bookmarkStart") if internal_name(node.get(W + "name"))}
    return {node.get(W + "id") for node in paragraph.iter() if node.tag in {W + "bookmarkStart", W + "bookmarkEnd"} and node.get(W + "id") in all_private}


def _ledger(backend):
    for part in backend.part.package.parts:
        if not str(part.partname).startswith("/customXml/"):
            continue
        try:
            root = etree.fromstring(part.blob, parser=etree.XMLParser(resolve_entities=False, no_network=True))
        except etree.XMLSyntaxError:
            continue
        if root.tag == "{" + NS + "}identities":
            return part
    return None


def enable(root):
    etree.register_namespace("jed", NS)
    root.set("{" + NS + "}schema", "1")
    prefix = next(key for key, value in root.nsmap.items() if value == NS and key is not None)
    values = (root.get("{" + MC + "}Ignorable") or "").split()
    if prefix not in values:
        root.set("{" + MC + "}Ignorable", " ".join(values + [prefix]))
    preserved = (root.get("{" + MC + "}PreserveAttributes") or "").split()
    for name in ("id", "schema"):
        if prefix + ":" + name not in preserved:
            preserved.append(prefix + ":" + name)
    root.set("{" + MC + "}PreserveAttributes", " ".join(preserved))


def identify(node):
    root = node.getroottree().getroot()
    enable(root)
    if not node.get(ID):
        node.set(ID, uuid4().hex)
    return node.get(ID)


def stamp(backend):
    from docx.opc.part import Part
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.opc.packuri import PackURI
    part = _ledger(backend)
    previous = []
    if part is not None:
        try:
            root = etree.fromstring(part.blob, parser=etree.XMLParser(resolve_entities=False, no_network=True))
            previous = json.loads(root.text or "{}").get("objects", [])
        except (ValueError, TypeError, etree.XMLSyntaxError):
            raise ValueError("Invalid Word object identity ledger.") from None
    nodes, used_paragraphs = [], set()
    for part in backend.part.package.parts:
        root = getattr(part, "element", None)
        if root is None or etree.QName(root).namespace != "http://schemas.openxmlformats.org/wordprocessingml/2006/main" or etree.QName(root).localname not in STORIES:
            continue
        for node in root.iter():
            if not isinstance(node.tag, str) or node.tag.endswith("}fldChar") and node.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldCharType") != "begin":
                continue
            notes = [item for item in [node, *node.iterancestors()] if item.tag in {W + "footnote", W + "endnote"}]
            if any(item.get(W + "type", "normal") != "normal" or int(item.get(W + "id", "0")) <= 0 for item in notes):
                continue
            if etree.QName(node).localname in TAGS and not (node.tag == W + "bookmarkStart" and internal_name(node.get(W + "name"))):
                nodes.append((node, str(part.partname)))
                if node.tag == W + "p":
                    if not node.get(PARA_ID) or node.get(PARA_ID) in used_paragraphs:
                        value = f"{secrets.randbelow(0x7FFFFFFE) + 1:08X}"
                        while value in used_paragraphs:
                            value = f"{secrets.randbelow(0x7FFFFFFE) + 1:08X}"
                        node.set(PARA_ID, value)
                    used_paragraphs.add(node.get(PARA_ID))

    # Word can regenerate paraId as well. Standard hidden bookmarks provide
    # a durable paragraph anchor even through a native Word SaveAs/edit.
    highest = max([0] + [int(item.get(W + "id", "0")) for node, _ in nodes if node.tag == W + "p"
                         for item in node.iter(W + "bookmarkStart") if item.get(W + "id", "").isdigit()])
    paragraph_uuids = set()
    anchored_roots = {}
    for node, _ in nodes:
        root = node.getroottree().getroot()
        if root not in anchored_roots:
            anchored_roots[root] = {item.get(W + "id"): item.get(W + "name")[7:] for item in root.iter(W + "bookmarkStart") if internal_name(item.get(W + "name"))}
    for node, _ in nodes:
        if node.tag != W + "p":
            continue
        starts = [item for item in node if item.tag == W + "bookmarkStart" and internal_name(item.get(W + "name"))]
        names = {item.get(W + "name")[7:] for item in starts}
        root = node.getroottree().getroot()
        outside = anchored_roots[root]
        endings = [item for item in node if item.tag == W + "bookmarkEnd" and item.get(W + "id") in outside]
        names.update(outside[item.get(W + "id")] for item in endings)
        if not node.get(ID) and len(names) == 1 and next(iter(names)) not in paragraph_uuids:
            node.set(ID, next(iter(names)))
        value = identify(node)
        if value in paragraph_uuids:
            value = uuid4().hex
            node.set(ID, value)
        paragraph_uuids.add(value)
        correct = [item for item in starts if item.get(W + "name") == "_JAKAL_" + value]
        ending = next((item for item in node if item.tag == W + "bookmarkEnd" and correct and item.get(W + "id") == correct[0].get(W + "id")), None)
        if len(starts) == 1 and len(correct) == 1 and ending is not None:
            position = 1 if len(node) and node[0].tag == W + "pPr" else 0
            if node.index(correct[0]) != position:
                node.remove(correct[0])
                node.insert(position, correct[0])
            if node[-1] is not ending:
                node.remove(ending)
                node.append(ending)
            continue
        identities = {item.get(W + "id") for item in starts + endings}
        for item in list(root.iter()):
            if item.tag in {W + "bookmarkStart", W + "bookmarkEnd"} and item.get(W + "id") in identities:
                item.getparent().remove(item)
        from docx.oxml import OxmlElement
        highest += 1
        start, end = OxmlElement("w:bookmarkStart"), OxmlElement("w:bookmarkEnd")
        start.set(W + "id", str(highest))
        start.set(W + "name", "_JAKAL_" + value)
        end.set(W + "id", str(highest))
        position = 1 if len(node) and node[0].tag == W + "pPr" else 0
        node.insert(position, start)
        node.append(end)

    from .editing_navigation import link_fields, R
    fields = link_fields(backend)
    parts = {str(part.partname): part for part in backend.part.package.parts}

    def descriptor(node, partname):
        kind = etree.QName(node).localname
        paragraphs = [node.get(ID)] if kind == "p" else [item.get(ID) for item in node.iter(W + "p") if item.get(ID)]
        if not paragraphs:
            paragraph = next((p for p in node.iterancestors(W + "p")), None)
            if paragraph is not None:
                paragraphs = [paragraph.get(ID)]
        native = node.get(W + "numId") if kind == "num" else node.get(W + "id") if kind in {"footnote", "endnote"} else ""
        if kind == "drawing":
            docpr = next(node.iter("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr"), None)
            native = docpr.get("id", "") if docpr is not None else ""
        text = "".join(node.itertext()) + "|" + str(node.get(W + "instr", "")) + "|" + str(node.get(W + "name", ""))
        record = {"kind": kind, "paragraphs": paragraphs, "native": native or "", "signature": sha256(text.encode("utf-8")).hexdigest(), "part": partname}
        link = fields.get(node)
        if kind == "hyperlink":
            rel = parts[partname].rels.get(node.get("{" + R + "}id"))
            address = rel.target_ref if rel is not None else ""
            anchor = node.get(W + "anchor")
            link = (address + ("#" + anchor if anchor else ""), "".join(node.itertext(W + "t")))
        if link is not None:
            record.update(family="link", link_key=sha256(json.dumps(link, ensure_ascii=False).encode("utf-8")).hexdigest())
        return record

    # Native Word removes unknown ignorable attributes on SaveAs. Its paraId
    # attributes and custom XML survive; use those anchors to recover UUIDs.
    # The ledger contains IDs and hashes only, never a copy of document text.
    seen = {node.get(ID) for node, _ in nodes if node.get(ID)}
    assigned, records = set(), []
    described = [(node, descriptor(node, partname)) for node, partname in nodes]
    matches = {}
    for node, record in described:
        if not node.get(ID):
            candidates = [old for old in previous if isinstance(old, dict) and old.get("id") not in seen and
                          (old.get("kind") == record["kind"] or record.get("family") == "link" and (old.get("family") == "link" or old.get("kind") == "hyperlink"))]
            anchors = set(record["paragraphs"])
            if anchors:
                candidates = [old for old in candidates if anchors.intersection(old.get("paragraphs", []))]
                if record.get("family") == "link":
                    candidates = [old for old in candidates if old.get("link_key") == record["link_key"]]
                exact = [old for old in candidates if old.get("signature") == record["signature"]]
                if record["kind"] == "drawing" and record["native"]:
                    exact = [old for old in candidates if old.get("native") == record["native"]]
                if len(exact) == 1:
                    candidates = exact
                elif record["kind"] == "tbl":
                    primary = [old for old in candidates if old.get("paragraphs", [None])[0] in anchors]
                    if len(primary) == 1:
                        candidates = primary
            else:
                candidates = [old for old in candidates if not old.get("paragraphs") and old.get("native", "") == record["native"] and old.get("part") == record["part"]]
                exact = [old for old in candidates if old.get("signature") == record["signature"]]
                if len(exact) == 1:
                    candidates = exact
            matches[node] = candidates
    # Both directions must be unique: duplicated normalized links cannot steal
    # an old ID merely because they happened to be visited first.
    claims = {}
    for candidates in matches.values():
        for old in candidates:
            claims[old.get("id")] = claims.get(old.get("id"), 0) + 1
    for node, record in described:
        candidates = matches.get(node, [])
        if len(candidates) == 1 and candidates[0].get("id") and claims[candidates[0]["id"]] == 1:
            node.set(ID, candidates[0]["id"])
        value = identify(node)
        if value in assigned:
            value = uuid4().hex
            node.set(ID, value)
        seen.add(value)
        assigned.add(value)
        record["id"] = value
        records.append(record)
    root = etree.Element("{" + NS + "}identities", nsmap={"jed": NS})
    root.text = json.dumps({"schema": 1, "objects": sorted(records, key=lambda item: item["id"])}, sort_keys=True, separators=(",", ":"))
    data = etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
    ledger = _ledger(backend)
    if ledger is None:
        ledger = Part(PackURI(LEDGER), "application/xml", data, backend.part.package)
        backend.part.relate_to(ledger, RT.CUSTOM_XML)
    else:
        ledger._blob = data


def assign(node, value):
    if value:
        enable(node.getroottree().getroot())
        node.set(ID, value)
