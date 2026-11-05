"""Retain empty glyph runs which participate in Hancom line measurement.

Records contain character references and positions, never old paragraph text or
line-segment caches. A text hash prevents applying positions after a text edit.
"""
from __future__ import annotations

from copy import deepcopy
from hashlib import sha256

from lxml import etree

from .namespaces import NS, qname
from ._hwpx_element_helpers import _text_node_plain_text, _set_text_node_content


def text_hash(text):
    return sha256(text.encode("utf-8")).hexdigest()


def glyph_runs(paragraph):
    """Return glyph runs while keeping technical property fields out of layout."""
    from ._hancom_hwpx_write import _word_effect_from_field_name
    hidden = None
    output = []
    for run in paragraph.findall(qname("hp", "run")):
        begins = run.xpath("./hp:ctrl/hp:fieldBegin", namespaces=NS)
        if begins:
            record = _word_effect_from_field_name(begins[0].get("name"))
            if not record or record.get("kind") != "JAKAL_WORD_FEATURE":
                return None
            hidden = begins[0].get("id")
        if hidden is not None:
            if run.xpath("./hp:ctrl/hp:fieldEnd[@beginIDRef=$id]", namespaces=NS, id=hidden):
                hidden = None
            continue
        if any(child.tag != qname("hp", "t") for child in run):
            return None
        output.append(run)
    return output if hidden is None else None


def capture_empty_runs(paragraph):
    runs = glyph_runs(paragraph)
    if runs is None:
        return None
    text = ""
    empty = []
    for run in runs:
        value = "".join(_text_node_plain_text(node) for node in run.findall(qname("hp", "t"))).replace("\r", "")
        if not value:
            empty.append({"offset": len(text), "character": run.get("charPrIDRef", "0"),
                          "text_node": bool(run.findall(qname("hp", "t")))})
        text += value
    return {"text_hash": text_hash(text), "runs": empty}


def restore_empty_runs(paragraph, record):
    runs = glyph_runs(paragraph)
    if runs is None:
        return
    values = ["".join(_text_node_plain_text(t) for t in run.findall(qname("hp", "t"))) for run in runs]
    if text_hash("".join(values)) != record["text_hash"]:
        return
    for run, value in zip(runs, values):
        if not value:
            paragraph.remove(run)
    inserted: dict[int, etree._Element] = {}
    for item in record["runs"]:
        offset = item["offset"]
        run = etree.Element(qname("hp", "run"), charPrIDRef=item["character"])
        if item["text_node"]:
            etree.SubElement(run, qname("hp", "t"))
        if offset in inserted:
            inserted[offset].addnext(run)
        else:
            position = 0
            anchor = None
            for candidate in glyph_runs(paragraph) or []:
                value = "".join(_text_node_plain_text(t) for t in candidate.findall(qname("hp", "t")))
                if not value:
                    continue
                if position <= offset < position + len(value):
                    if offset == position:
                        candidate.addprevious(run)
                    else:
                        split = offset - position
                        after = deepcopy(candidate)
                        for target, text in ((candidate, value[:split]), (after, value[split:])):
                            for child in list(target):
                                target.remove(child)
                            node = etree.SubElement(target, qname("hp", "t"))
                            _set_text_node_content(node, text)
                        candidate.addnext(run)
                        run.addnext(after)
                    anchor = run
                    break
                position += len(value)
                anchor = candidate
            if anchor is not run:
                if anchor is not None:
                    anchor.addnext(run)
                else:
                    index = sum(child.tag == qname("hp", "run") for child in paragraph)
                    paragraph.insert(index, run)
        inserted[offset] = run
