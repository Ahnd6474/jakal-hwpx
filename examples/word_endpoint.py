"""Versioned Word endpoint demo; run with --serve to start the local HTTP adapter.

    python examples/word_endpoint.py
    python examples/word_endpoint.py --serve --input report.docx --port 8765
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import secrets

from jakal_hwpx import Document
from jakal_hwpx.addons.docx import WordDocument, WordEndpoint, make_http_server


def demo_document():
    doc = WordDocument.blank()
    doc.append_paragraphs(["Word interaction", "Reference target", "Review status: draft", "Find across", "paragraph boundaries."])
    with doc.edit_word_details() as edit:
        edit.format_style("Normal", font={"font_name": "Arial", "font_size_pt": 11, "color": "000000"}, paragraph={"space_after_pt": 8})
        edit.create_style("Interaction heading", base="Normal", next_style="Normal", font={"font_size_pt": 20, "bold": True})
        edit.body.paragraph(0).format(style="Interaction heading")
        edit.section().format_page(width_pt=595, height_pt=842, margins_pt={"left": 48, "right": 48, "top": 45, "bottom": 45})
        edit.body.paragraph(1).range().add_bookmark("reference")
        edit.body.paragraph(2).insert_hyperlink(0, "#reference", "Jump to reference. ")
        edit.body.find_text("across\nparagraph")[0].replace("over\nparagraph")
        table = edit.body.insert_table(len(edit.body.blocks), rows=3, cols=2, values=[["Operation", "Result"], ["Atomic edit", "Committed"], ["Old revision", "Conflict"]])
        table.format(widths_pt=[220, 240], repeat_header=True)
        for c in range(2):
            table.format_cell(0, c, fill_color="E8EDF2")
        other = table.split(2)
        table.join(other)
        for r in range(3):
            table.format_row(r, allow_split=False)
        for title, script in [
            ("Matrix", "left [ {matrix {a & b # c & d}} right ]"),
            ("Summation", "sum _ {i=1} ^ {n} {{x} _ {i}}"),
            ("Integral", "int _ {0} ^ {1} {{x} ^ {2} dx}"),
        ]:
            edit.body.append_paragraph(title).format_text(bold=True)
            edit.body.append_paragraph().add_equation(script)
    return doc


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path)
    parser.add_argument("--output-dir", type=Path, default=Path("build/word-interaction-example"))
    parser.add_argument("--serve", action="store_true")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--token")
    args = parser.parse_args()
    doc = WordDocument.open(args.input) if args.input else demo_document()
    endpoint = WordEndpoint(doc)
    if args.serve:
        token = args.token or secrets.token_urlsafe(24)
        server = make_http_server(endpoint, host=args.host, port=args.port, token=token)
        print(f"Listening on http://{args.host}:{server.server_address[1]}", flush=True)
        print("Authorization: Bearer " + token, flush=True)
        try:
            server.serve_forever()
        except KeyboardInterrupt:
            pass
        finally:
            server.server_close()
        return
    args.output_dir.mkdir(parents=True, exist_ok=True)
    before = endpoint.snapshot()
    operations = []
    if args.input is None:
        target = next(item["id"] for item in before["objects"] if item.get("text", "").endswith("Review status: draft"))
        operations = [
            {"target": target, "method": "replace_text", "args": ["draft", "reviewed"]},
            {"target": target, "method": "format_text", "kwargs": {"bold": True}},
        ]
    request = {"expected_version": before["version"], "operations": operations}
    result = endpoint.apply(**request)
    for name, value in (("snapshot", before), ("request", request), ("response", result)):
        (args.output_dir / (name + ".json")).write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")
    doc.save(args.output_dir / "edited.docx")
    native = doc.save(args.output_dir / "edited.hwpx")
    Document.open(native).save(args.output_dir / "restored.docx")
    print(args.output_dir.resolve())


if __name__ == "__main__":
    main()
