"""Exercise structured Word editing and native HWPX/DOCX equation mapping."""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from jakal_hwpx import Document
from jakal_hwpx.addons.docx import WordDocument


def create(output):
    output.mkdir(parents=True, exist_ok=True)
    doc = WordDocument.blank()
    with doc.edit_word() as word:
        from docx.oxml.ns import qn
        for border in list(word.styles["Title"].element.iter(qn("w:pBdr"))):
            border.getparent().remove(border)
    with doc.edit_word_details() as edit:
        edit.format_style("Normal", font={"font_name": "Arial", "font_size_pt": 10, "color": "000000"},
                          paragraph={"space_after_pt": 5})
        for name in ("Title", "Heading 1", "Heading 2"):
            edit.format_style(name, font={"color": "000000"})
        edit.section().format_page(width_pt=595, height_pt=842, margins_pt={"left": 48, "right": 48, "top": 42, "bottom": 42})
        body = edit.body
        body.append_paragraph("Word module extensions", style="Title")
        body.append_paragraph("Native equations, structured field results, content controls and explicit field refresh.")
        toc = body.append_paragraph().add_field('TOC \\o "1-2" \\h', "Contents", complex_field=True, locked=True)
        body.append_paragraph("Structured equations", style="Heading 1")
        for script in ["cases {x & x>0 # -x & x<0}", "lpile {a+b=c # d+e=f}", "{n} choose {k}",
                       "{a} atop {b}", "bmatrix {a & b # c & d}"]:
            body.append_paragraph().add_equation(script)
        body.append_paragraph("Editable content", style="Heading 1")
        control = body.insert_control(len(body.blocks), text="Original control text", tag="article", title="Article")
        control.paragraph(0).replace_text("Original", "Edited")
        control.append_paragraph("Another paragraph inside the same content control.")
        field = body.append_paragraph().add_field('IF 1 = 1 "Result" "Other"', "Result", complex_field=True, locked=True)
        with field.edit_result() as result:
            result.paragraph(0).format_text(bold=True)
            result.append_paragraph("A second result paragraph.")
            result.insert_table(1, rows=1, cols=2, values=[["Rich result", "Editable table"]])
        body.append_paragraph("Explicit refresh", style="Heading 1")
        p = body.append_paragraph("Condition: ")
        condition = p.add_field('IF 2 < 10 "numeric comparison passed" "failed"', "old")
        body.append_paragraph("한글English123 mixed text wrapping").format(
            word_wrap=False, auto_space_east_asian_latin=True, auto_space_east_asian_numbers=False,
        ).format_text(font_names={"eastAsia": "맑은 고딕"})
        p = body.append_paragraph("Moved between parts")
        p.move(0, destination=edit.header())
        p.move(len(body.blocks), destination=body)
        report = edit.update_fields(now="2026-11-04T13:14:15")
        condition.lock()  # Compare the refreshed cache without Word recalculating it on open.
        toc_report = toc.rebuild_toc()
    edited = doc.save(output / "edited.docx")
    middle = doc.save(output / "edited.hwpx")
    returned = Document.open(middle).save(output / "restored.docx")
    (output / "refresh.json").write_text(json.dumps({"fields": report, "toc": toc_report}, ensure_ascii=False, indent=2), encoding="utf-8")
    for path in (edited, middle, returned):
        print(path)
    return edited, middle, returned


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, default=Path("build/word-module-example"))
    create(parser.parse_args().output_dir)


if __name__ == "__main__":
    main()
