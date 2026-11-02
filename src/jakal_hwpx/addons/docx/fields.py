"""Preserve complex Word field boundaries while keeping results editable.

Only field codes and attributes travel in the empty native boundary records.
Results remain ordinary Hancom text, links, pictures, tables, and inner fields.
"""
from __future__ import annotations

from copy import deepcopy
import json
from typing import Any

from lxml import etree

from .features import parse, xml


def _xpath(node: Any, expression: str) -> list:
    return etree.XPath(expression, namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})(node)


def _marker(record: dict) -> Any:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    node = OxmlElement("w:fldSimple")
    node.set(qn("w:instr"), "JAKAL_FIELD_BOUNDARY " + json.dumps(record))
    return node


def normalize_ranges(root: Any) -> None:
    """Expand rich/nested or multi-paragraph fields into reversible boundaries."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    # Simple fields may contain inner fields, links, or pictures. A text-only
    # Field cannot own those results; let the ordinary paragraph reader own them.
    for node in reversed(list(root.iter(qn("w:fldSimple")))):
        instruction = node.get(qn("w:instr"), "")
        if instruction.startswith("JAKAL_FIELD_BOUNDARY "):
            continue
        rich = any(child.tag != qn("w:r") for child in node)
        rich |= bool(_xpath(node, ".//w:fldChar | .//w:drawing | .//w:pict"))
        if not rich:
            continue
        code = OxmlElement("w:p")
        run = OxmlElement("w:r")
        text = OxmlElement("w:instrText")
        text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        text.text = instruction
        run.append(text)
        code.append(run)
        attrs = {key: value for key, value in node.attrib.items() if key != qn("w:instr")}
        node.addprevious(_marker({"side": "begin", "code": xml(code), "instruction": instruction,
                                  "attributes": attrs, "separated": True}))
        for child in list(node):
            node.addprevious(child)
        node.addprevious(_marker({"side": "end", "attributes": {}}))
        node.getparent().remove(node)

    # Split control/instruction runs so a run containing both a field end and
    # trailing visible text cannot lose that text during normalization.
    for run in list(root.iter(qn("w:r"))):
        if not any(child.tag in {qn("w:fldChar"), qn("w:instrText")} for child in run):
            continue
        children = [child for child in run if child.tag != qn("w:rPr")]
        if len(children) <= 1:
            continue
        for child in children:
            single = OxmlElement("w:r")
            single.attrib.update(run.attrib)
            if run.rPr is not None:
                single.append(deepcopy(run.rPr))
            single.append(deepcopy(child))
            run.addprevious(single)
        run.getparent().remove(run)

    # A text box is a separate story even when its anchor is inside a body
    # paragraph whose field spans several following paragraphs.
    stories: dict[Any, list[Any]] = {}
    for paragraph in root.iter(qn("w:p")):
        story = next((node for node in paragraph.iterancestors()
                      if node.tag in {qn("w:txbxContent"), qn("w:footnote"), qn("w:endnote")}), root)
        stories.setdefault(story, []).append(paragraph)
    for paragraphs in stories.values():
        entries = [(paragraph, child) for paragraph in paragraphs for child in paragraph if child.tag != qn("w:pPr")]
        stack: list[dict] = []
        spans = []
        for index, (paragraph, node) in enumerate(entries):
            control = node.find(qn("w:fldChar")) if node.tag == qn("w:r") else None
            if control is None:
                continue
            kind = control.get(qn("w:fldCharType"))
            if kind == "begin":
                stack.append({"start": index, "begin": dict(control.attrib)})
            elif kind == "separate" and stack:
                stack[-1]["separator"] = index
            elif kind == "end" and stack:
                item = stack.pop()
                item.update(end=index, end_attributes=dict(control.attrib))
                spans.append(item)
        selected = []
        for item in spans:
            start, end = item["start"], item["end"]
            nested = any(start < other["start"] < other["end"] < end for other in spans)
            rich = any(node.tag != qn("w:r") or bool(_xpath(node, ".//w:drawing | .//w:pict"))
                       for _, node in entries[start + 1:end])
            if entries[start][0] is not entries[end][0] or nested or rich:
                selected.append(item)
        deleted = {index for item in selected
                   for index in range(item["start"] + 1, item.get("separator", item["end"] - 1) + 1)}
        replacements = {}
        for item in selected:
            start, end = item["start"], item["end"]
            if start in deleted:
                continue  # Nested code is already retained in its outer code.
            separator = item.get("separator", end)
            code = OxmlElement("w:p")
            code.extend(deepcopy(node) for _, node in entries[start + 1:separator])
            instruction = "".join(code.xpath(".//w:instrText/text()"))
            replacements[start] = _marker({"side": "begin", "code": xml(code), "instruction": instruction,
                                            "attributes": item["begin"], "separated": "separator" in item})
            replacements[end] = _marker({"side": "end", "attributes": item["end_attributes"]})
        for index, (_, node) in enumerate(entries):
            if index in replacements:
                node.addprevious(replacements[index])
                node.getparent().remove(node)
            elif index in deleted and node.getparent() is not None:
                node.getparent().remove(node)


def append_boundary(paragraph: Any, block: Any) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    record = json.loads(str(block.parameters["Record"]))
    if record.get("side") not in {"begin", "end"}:
        raise ValueError("Invalid Word field boundary.")

    def control(kind, attributes=None):
        run = OxmlElement("w:r")
        node = OxmlElement("w:fldChar")
        node.attrib.update(attributes or {})
        node.set(qn("w:fldCharType"), kind)
        run.append(node)
        paragraph._p.append(run)

    control(record["side"], record.get("attributes"))
    if record["side"] == "begin":
        for node in parse(record["code"]):
            paragraph._p.append(node)
        if record.get("separated", True):
            control("separate")
