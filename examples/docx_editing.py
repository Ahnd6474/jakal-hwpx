"""Use the transactional Word detail API without a desktop Office dependency.

    python examples/docx_editing.py --output-dir build/docx-editing-example
    python examples/docx_editing.py --image path/to/image.png
"""
from __future__ import annotations

import argparse
from pathlib import Path

from jakal_hwpx import Document
from jakal_hwpx.addons.docx import WordDocument


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, default=Path("build/docx-editing-example"))
    parser.add_argument("--image", type=Path)
    args = parser.parse_args()
    args.output_dir.mkdir(parents=True, exist_ok=True)

    doc = WordDocument.blank()
    # Some installed Word templates give Title a decorative bottom border.
    # Keep the API example's title plain while retaining all other properties.
    with doc.edit_word() as backend:
        from docx.oxml.ns import qn
        properties = backend.styles["Title"].element.pPr
        if properties is not None:
            for border in list(properties.findall(qn("w:pBdr"))):
                properties.remove(border)
    doc.append_paragraph("Word detail editing")
    doc.append_paragraph("Status: draft")
    doc.append_paragraph("Remove this draft note")
    doc.append_table(2, 2, [["Period", "Value"], ["Q1", "120"]])
    doc.append_paragraph("Updated values")
    original = doc.save(args.output_dir / "original.docx")

    doc = Document.open(original)
    with doc.edit_word_details() as edit:
        edit.format_style("Normal", font={"font_name": "Arial", "font_size_pt": 11, "color": "000000"},
                          paragraph={"space_after_pt": 6})
        edit.format_style("Title", font={"font_name": "Arial", "font_size_pt": 20, "bold": True, "color": "000000"})
        body = edit.body
        body.paragraph(0).format(style="Title")
        body.insert_paragraph(1, "This example edits the Word package while retaining unrelated document details.")
        status = body.paragraph(2)
        status.replace_text("draft", "reviewed")
        status.format_text(0, 7, bold=True)
        body.paragraph(3).delete()
        body.paragraph(3).move(3)

        table = body.table(0)
        table.set_cell_text(1, 1, "135")
        table.add_row(["Q2", "180"])
        table.format(style="Table Grid", widths_pt=[240, 160], repeat_header=True)
        for row in range(3):
            table.format_row(row, height_pt=24)
            for col in range(2):
                table.format_cell(row, col, vertical_alignment="center",
                                  padding_pt={"left": 6, "right": 6, "top": 5, "bottom": 5})
                if row == 0:
                    table.format_cell(row, col, fill_color="E8EDF2")
                    table.cell(row, col).paragraph(0).format_text(bold=True)

        edit.header().paragraph(0).set_text("Jakal HWPX DOCX addon").format_text(font_size_pt=9)
        edit.footer().paragraph(0).set_text("Created with the Word detail editing API").format_text(font_size_pt=9)
        if args.image:
            paragraph = body.append_paragraph().format(space_before_pt=12)
            picture = paragraph.add_picture(args.image.read_bytes(), width_pt=72, alt_text="Editing example image")
            picture.resize(width_pt=96).crop(left=.1, right=.1).rotate(10)

    edited = doc.save(args.output_dir / "edited.docx")
    middle = doc.save(args.output_dir / "edited.hwpx")
    restored = Document.open(middle).save(args.output_dir / "restored.docx")
    for path in (original, edited, middle, restored):
        print(path.resolve())


if __name__ == "__main__":
    main()
