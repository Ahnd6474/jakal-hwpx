"""Create a four-page DOCX/HWPX roundtrip QA fixture (requires docx extra and Pillow).

    python examples/docx_complex_roundtrip.py --output-dir build/complex-word

Includes nested/merged tables, repeated headers, manual page breaks, a two-
paragraph footnote, an equation, pictures, alternating stories and landscape.
Rendering and foreground comparison are separate, native-application QA steps.
"""

import argparse
from pathlib import Path
from docx import Document as Word
from docx.shared import Pt, Inches
from docx.enum.section import WD_SECTION_START, WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image, ImageDraw
from jakal_hwpx import Document
from jakal_hwpx.addons.docx.equations import hwp_to_omml

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("--output-dir", type=Path, default=Path("build/complex-word"))
folder = parser.parse_args().output_dir
folder.mkdir(parents=True, exist_ok=True)
image = Image.new("RGB", (600, 180), "#f1f3f5")
draw = ImageDraw.Draw(image)
for i, color in enumerate(["#25476a", "#a15232", "#27806a", "#7b518d"]):
    draw.rectangle((25 + 140 * i, 20 + 15 * i, 125 + 140 * i, 160), fill=color)
image.save(folder / "diagram.png")
word = Word()
normal = word.styles["Normal"]
normal.font.name = "Arial"
normal.font.size = Pt(10)
normal.paragraph_format.space_after = Pt(5)
for name in ["Title", "Heading 1", "Heading 2"]:
    word.styles[name].font.color.rgb = __import__("docx").shared.RGBColor(0, 0, 0)
section = word.sections[0]
section.page_width, section.page_height = Pt(595), Pt(842)
section.top_margin = section.bottom_margin = Pt(42)
section.left_margin = section.right_margin = Pt(48)
section.header_distance = section.footer_distance = Pt(18)
section.different_first_page_header_footer = True
word.settings.odd_and_even_pages_header_footer = True
section.header.paragraphs[0].text = "Odd page / engineering review"
section.even_page_header.paragraphs[0].text = "Even page / engineering review"
section.first_page_header.paragraphs[0].text = "Confidential draft / first page"
for story in (section.footer, section.even_page_footer, section.first_page_footer):
    p = story.paragraphs[0]
    p.text = "QA / "
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "1"
    run.append(text)
    field.append(run)
    p._p.append(field)
word.add_heading("Complex document interoperability", 0)
word.add_paragraph("Nested tables, multi-page rows, notes, equations and section variants.")
word.add_picture(str(folder / "diagram.png"), width=Inches(4.8))
table = word.add_table(rows=3, cols=3, style="Table Grid")
table.autofit = False
for i, width in enumerate([80, 180, 220]):
    table.columns[i].width = Pt(width)
table.cell(0, 0).merge(table.cell(0, 2)).text = "Merged header"
table.cell(1, 0).merge(table.cell(2, 0)).text = "Vertical group"
table.cell(1, 1).text = "Rich paragraph"
table.cell(1, 1).paragraphs[0].add_run(" bold").bold = True
nested = table.cell(1, 2).add_table(rows=2, cols=2)
for i, row in enumerate(nested.rows):
    for j, cell in enumerate(row.cells):
        cell.text = f"Nested {i},{j}"
table.cell(2, 1).text = "Second row"
table.cell(2, 2).text = "Remaining cell"
paragraph = word.add_paragraph("Formula: ")
paragraph._p.append(hwp_to_omml("sum _{i=1} ^{n} {{i} over {n}}"))
word.add_paragraph("Footnote anchor with mixed sizes.").runs[0].font.size = Pt(11)
word.add_page_break()
word.add_heading("Repeated table across pages", 1)
table = word.add_table(rows=1, cols=3, style="Light Shading Accent 1")
for i, value in enumerate(["Code", "Description", "Amount"]):
    table.cell(0, i).text = value
header = OxmlElement("w:tblHeader")
table.rows[0]._tr.get_or_add_trPr().append(header)
for i in range(29):
    cells = table.add_row().cells
    for j, value in enumerate(
        [f"R{i:02}", "Review entry with two lines\nSecond paragraph content", f"{i * 12.75:.2f}"]
    ):
        cells[j].text = value
    if i % 5 == 0:
        shade = OxmlElement("w:shd")
        shade.set(qn("w:fill"), "E6F0E8")
        cells[2]._tc.get_or_add_tcPr().append(shade)
section = word.add_section(WD_SECTION_START.NEW_PAGE)
section.orientation = WD_ORIENT.LANDSCAPE
section.page_width, section.page_height = Pt(842), Pt(595)
section.header.is_linked_to_previous = False
section.header.paragraphs[0].text = "Landscape appendix"
word.add_heading("Landscape appendix", 1)
word.add_paragraph("A separate section retains its own page size and inherits even-page content.")
word.add_picture(str(folder / "diagram.png"), width=Inches(7.8))
source = folder / "source.docx"
word.save(source)
doc = Document.open(source)
with doc.edit_word_details() as edit:
    anchor = next(p for p in edit.body.blocks if getattr(p, "text", "").startswith("Footnote anchor"))
    note = anchor.add_note("A native footnote containing two paragraphs.")
    note.paragraph(0).format_text(font_size_pt=8)
    note.append_paragraph("Second note paragraph.").format_text(italic=True, font_size_pt=8)
doc.save(source)
middle = Document.open(source).save(folder / "middle.hwpx")
Document.open(middle).save(folder / "returned.docx")
print(source)
