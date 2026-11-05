"""Isolate HWPX body/math and check its editable DOCX roundtrip.

    python scripts/validate_hwpx_math_body.py input.hwpx --output-dir build/math-body

Creates an isolated reference, middle.docx, returned.hwpx, direct.hwpx and a
structure.json report. Export body-source.hwpx and returned.hwpx in the same
Hancom installation, then run compare_document_pdfs.py on every PDF page.
The reference excludes native shapes, pictures/diagrams, automatic numbers and
page decorations, while retaining body tables, text and editable equations.
Original files are never modified; layout caches are rebuilt, not reused.
"""
from __future__ import annotations

from collections import Counter
from pathlib import Path
from zipfile import ZipFile
import argparse
import json

from lxml import etree

from jakal_hwpx import Document
from jakal_hwpx._hwpx_element_helpers import _text_node_plain_text
from jakal_hwpx.addons.docx.equations import hwp_to_omml
from jakal_hwpx.namespaces import qname


EXCLUDED = {"rect", "line", "ellipse", "arc", "polygon", "curve", "connectLine", "textart", "container", "pic",
            "ole", "chart", "autoNum", "newNum", "pageNum", "header", "footer", "hiddenComment", "masterPage"}


def prepare_reference(source: Path, output: Path) -> dict[str, int]:
    if source.resolve() == output.resolve():
        raise ValueError("The reference output must differ from the original document.")
    output.parent.mkdir(parents=True, exist_ok=True)
    removed: Counter[str] = Counter()
    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    with ZipFile(source) as original, ZipFile(output, "w") as target:
        for item in original.infolist():
            data = original.read(item.filename)
            if item.filename.startswith("Contents/section") and item.filename.endswith(".xml"):
                root = etree.fromstring(data, parser=parser)
                # Walk only the retained tree: removed shape text is out of scope.
                def strip(parent):
                    for node in list(parent):
                        if etree.QName(node).localname in EXCLUDED:
                            removed[etree.QName(node).localname] += 1
                            parent.remove(node)
                        elif node.tag == qname("hp", "linesegarray"):
                            parent.remove(node)
                        else:
                            strip(node)
                            if node.tag == qname("hp", "ctrl") and not len(node):
                                parent.remove(node)
                strip(root)
                for node in root.iter(qname("hp", "secPr")):
                    node.set("masterPageCnt", "0")
                data = etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
            elif item.filename == "Contents/header.xml":
                root = etree.fromstring(data, parser=parser)
                for node in root.iter(qname("hh", "heading")):
                    if node.get("type", "NONE") != "NONE":
                        removed["paragraph_numbering_style"] += 1
                        node.attrib.update(type="NONE", idRef="0", level="0")
                data = etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
            target.writestr(item, data)
    return dict(removed)


def math_signature(script: str):
    """Compare expression hierarchy, operands and semantic properties."""
    math = hwp_to_omml(script)
    def value(node):
        name = etree.QName(node).localname
        # Generated nodes contain semantic properties plus deterministic defaults.
        return [name, sorted((etree.QName(k).localname, v) for k, v in node.attrib.items()),
                node.text or "", [value(child) for child in node]]
    return value(math)


def inspect(path: Path):
    flow, equations, paragraphs = [], [], 0
    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    with ZipFile(path) as package:
        for name in package.namelist():
            if not (name.startswith("Contents/section") and name.endswith(".xml")):
                continue
            root = etree.fromstring(package.read(name), parser=parser)
            def walk(node):
                nonlocal paragraphs
                if node.tag == qname("hp", "equation"):
                    script = node.findtext(qname("hp", "script")) or ""
                    equations.append(script)
                    flow.append(["equation", len(equations) - 1])
                elif node.tag == qname("hp", "t"):
                    text = _text_node_plain_text(node)
                    if text:
                        if flow and flow[-1][0] == "text":
                            flow[-1][1] += text
                        else:
                            flow.append(["text", text])
                elif etree.QName(node).localname not in EXCLUDED:
                    if node.tag == qname("hp", "p"):
                        paragraphs += 1
                        flow.append(["paragraph_start"])
                    for child in node:
                        walk(child)
                    if node.tag == qname("hp", "p"):
                        flow.append(["paragraph_end"])
            walk(root)
    return flow, equations, paragraphs


def compare(reference: Path, candidate: Path) -> dict:
    original, source_math, source_paragraphs = inspect(reference)
    returned, target_math, target_paragraphs = inspect(candidate)
    mismatches = []
    for index, (source, target) in enumerate(zip(source_math, target_math)):
        try:
            if math_signature(source) == math_signature(target):
                continue
            error = "Expression hierarchy, operands or semantic properties differ."
        except ValueError as failure:
            error = str(failure)
        mismatches.append({"index": index, "source": source, "returned": target, "error": error})
    order_equal = original == returned
    counts_equal = len(source_math) == len(target_math) and source_paragraphs == target_paragraphs
    return {"paragraph_counts": [source_paragraphs, target_paragraphs],
            "equation_counts": [len(source_math), len(target_math)], "body_order_equal": order_equal,
            "math_mismatches": mismatches, "structure_equal": order_equal and counts_equal and not mismatches}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    folder = args.output_dir
    outputs = (folder / name for name in ("body-source.hwpx", "direct.hwpx", "middle.docx", "returned.hwpx", "structure.json"))
    if args.source.resolve() in {path.resolve() for path in outputs}:
        parser.error("Choose an output directory that does not overwrite the original document.")
    reference = folder / "body-source.hwpx"
    excluded = prepare_reference(args.source, reference)
    source = Document.open(reference)
    direct = source.save(folder / "direct.hwpx")
    middle = source.save(folder / "middle.docx")
    returned = Document.open(middle).save(folder / "returned.hwpx")
    report = {"original": str(args.source), "reference": str(reference), "excluded": excluded,
              "roundtrip": compare(reference, returned), "direct_model": compare(reference, direct),
              "visual_validation": "Export both PDFs in the same native application; compare_document_pdfs.py must also pass."}
    (folder / "structure.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False))
    return 0 if report["roundtrip"]["structure_equal"] and report["direct_model"]["structure_equal"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
