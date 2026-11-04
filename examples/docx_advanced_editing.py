"""Create a multi-page Word editing/roundtrip fixture without desktop Office.

Run from the repository root after installing jakal-hwpx[docx]:
    python examples/docx_advanced_editing.py --output-dir build/word-advanced-example
"""
from __future__ import annotations

import argparse
from pathlib import Path

from jakal_hwpx import Document
from jakal_hwpx.addons.docx import WordDocument


def create(output: Path, image: Path | None = None):
    output.mkdir(parents=True, exist_ok=True)
    doc = WordDocument.blank()
    with doc.edit_word_details() as edit:
        edit.format_style("Normal", font={"font_name": "Arial", "font_size_pt": 10, "color": "000000"},
                          paragraph={"space_after_pt": 5, "line_spacing": 1.1})
        edit.section().format_page(width_pt=595, height_pt=842, margins_pt={"left": 48, "right": 48, "top": 45, "bottom": 45}, header_distance_pt=20, footer_distance_pt=20)
        body = edit.body
        title = body.append_paragraph("Word advanced editing")
        title.format_text(font_size_pt=20, bold=True)
        subtitle = body.append_paragraph("한글과 English / merged tables, lists and notes")
        subtitle.format_text(font_names={"eastAsia": "맑은 고딕"}, languages={"eastAsia": "ko-KR"})
        listing = edit.create_list([{"format": "decimal", "text": "%1.", "left_indent_pt": 18, "hanging_pt": 12}])
        for text in ["Preserve existing XML", "Edit merged table geometry", "Roundtrip through HWPX"]:
            body.append_paragraph(text).set_list(listing)
        table = body.insert_table(len(body.blocks), rows=5, cols=3,
                                  values=[["Group", "Item", "Value"], ["A", "Alpha", "120"], ["", "Beta", "150"], ["B", "Gamma", "180"], ["C", "Delta", "210"]])
        table.merge(1, 0, 2, 0).add_row([None, "Inserted", "135"], index=2)
        table.add_column(["Status", "OK", "OK", "OK", "OK", "OK"], width_pt=60)
        table.format(widths_pt=[60, 220, 90, 60], repeat_header=True)
        for r in range(6):
            for c in range(4):
                table.format_cell(r, c, padding_pt={"top": 4, "bottom": 4, "left": 5, "right": 5},
                                  borders={edge: {"color": "B8C2CC", "width_pt": .5} for edge in ("top", "bottom", "left", "right")})
        for c in range(4):
            table.format_cell(0, c, fill_color="E8EDF2")
            table.cell(0, c).paragraph(0).format_text(bold=True)
        reference = body.append_paragraph("A genuine footnote is attached here")
        note = reference.add_note("Original footnote text")
        note.paragraph(0).replace_text("Original", "Edited")
        note.paragraph(0).format_text(font_size_pt=9)
        note.format_reference(font_size_pt=8, superscript=True)
        if image:
            body.append_paragraph().add_picture(image.read_bytes(), width_pt=130).crop(left=.05, right=.05)
        p = body.append_paragraph("Detailed text formatting")
        p.format(page_break_before=True, space_after_pt=12)
        p.format_text(font_size_pt=18, bold=True)
        p = body.append_paragraph("Highlighted text, followed by ordinary text.")
        p.format_text(0, 16, highlight="yellow", character_spacing_pt=.3, scale_percent=95)
        p = body.append_paragraph("Aligned tab\t120.50")
        p.format(tab_stops=[{"position_pt": 330, "alignment": "decimal", "leader": "dot"}], line_spacing_pt=18)
        p = body.append_paragraph("Report date: ")
        p.add_field('DATE \\@ "yyyy-MM-dd"', "2026-11-04", locked=True)
        for index in range(12):
            body.append_paragraph(f"Review {index + 1:02d}: Content remains editable after export. 한글 문단과 English text를 함께 확인합니다.")
        body.append_paragraph("Final section").format_text(font_size_pt=18, bold=True)
        final = edit.insert_section(len(body.blocks) - 1, break_type="next_page")
        final.format_page(orientation="landscape").format_columns(count=2, gap_pt=24)
        for index in range(20):
            final.append_paragraph(f"Column sample {index + 1}: Each column uses native section settings. The document keeps its layout properties across the roundtrip.")
        edit.header().paragraph(0).set_text("Word editing verification")
        edit.footer().paragraph(0).add_field("PAGE", "1")
    edited = doc.save(output / "edited.docx")
    native = doc.save(output / "edited.hwpx")
    restored = Document.open(native).save(output / "restored.docx")
    print(edited)
    print(native)
    print(restored)
    return edited, native, restored


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, default=Path("build/word-advanced-example"))
    parser.add_argument("--image", type=Path)
    args = parser.parse_args()
    create(args.output_dir, args.image)


if __name__ == "__main__":
    main()
