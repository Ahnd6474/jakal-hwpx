"""Create, edit, and round-trip a Word document without an Office installation.

From the repository root:
    python -m pip install -e ".[docx]"
    python examples/docx_addon.py --output-dir build/docx-example
"""
from __future__ import annotations

import argparse
from pathlib import Path

from docx.shared import Pt

from jakal_hwpx import Document
from jakal_hwpx.addons.docx import WordDocument


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, default=Path("build/docx-example"))
    args = parser.parse_args()
    output_dir = args.output_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    doc = WordDocument.blank()
    doc.append_paragraph("Quarterly report")
    doc.append_paragraph("Status: draft")
    doc.append_table(
        rows=2,
        cols=2,
        cell_texts=[["Item", "Value"], ["Q1", "120"]],
    )
    original = doc.save(output_dir / "original.docx")

    doc = Document.open(original)
    if not isinstance(doc, WordDocument):
        raise RuntimeError("DOCX input did not select the WordDocument addon.")
    untouched = doc.save(output_dir / "untouched-copy.docx")
    if untouched.read_bytes() != original.read_bytes():
        raise RuntimeError("The untouched DOCX copy differs from the original.")

    count = doc.replace_text("Status: draft", "Status: reviewed")
    doc.append_paragraph("Review complete")
    with doc.edit_word() as word:
        for run in word.paragraphs[0].runs:
            run.bold = True
            run.font.size = Pt(18)
    edited = doc.save(output_dir / "edited.docx")

    converted = Document.open(edited).save(output_dir / "converted.hwpx")
    restored = Document.open(converted).save(output_dir / "restored.docx")

    print(f"Untouched DOCX copy is byte-identical; replaced {count} occurrence(s).")
    for path in (original, untouched, edited, converted, restored):
        print(path.resolve())


if __name__ == "__main__":
    main()
