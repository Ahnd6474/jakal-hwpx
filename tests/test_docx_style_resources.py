"""Imported Word styles must not inherit fresh Word/HWPX template fonts."""
from __future__ import annotations

from docx import Document as WordFile
from docx.oxml.ns import qn
from docx.shared import RGBColor
import pytest

from jakal_hwpx import Document, WordDocument


def structure(node):
    if node is None:
        return None
    return node.tag, dict(node.attrib), (node.text or "").strip(), [structure(child) for child in node]


@pytest.mark.parametrize("font", ["Arial", "Times New Roman"])
def test_imported_normal_overrides_unused_hwpx_template_and_preserves_rpr(tmp_path, font):
    doc = WordDocument.blank()
    with doc.edit_word_details() as edit:
        edit.format_style("Normal", font={"font_name": font, "font_size_pt": 10, "color": "123456"})
        edit.body.append_paragraph("Body")
        edit.header().paragraph(0).object_id  # Empty header inherits Normal's metrics.
    original = doc.word_document
    for index in range(3):
        common = Document.open(doc.save(tmp_path / f"middle-{index}.hwpx"))
        assert len([style for style in common.style_definitions if style.name == "Normal"]) == 2
        path = common.save(tmp_path / f"returned-{index}.docx")
        returned = WordFile(path)
        assert structure(returned.styles["Normal"].element.rPr) == structure(original.styles["Normal"].element.rPr)
        assert returned.styles["Normal"].font.name == font
        assert returned.styles["Normal"].font.size.pt == 10
        assert returned.styles["Normal"].font.color.rgb == RGBColor.from_string("123456")
        assert returned.styles["Header"].element.rPr is None
        doc = WordDocument.open(path)


def test_used_title_and_heading_keep_their_native_colour_and_theme_properties(tmp_path):
    doc = WordDocument.blank()
    with doc.edit_word_details() as edit:
        edit.format_style("Normal", font={"font_name": "Arial", "font_size_pt": 10})
        for name in ("Title", "Heading 1"):
            edit.format_style(name, font={"color": "000000"})
            edit.body.append_paragraph(name, style=name)
    original = doc.word_document
    returned = WordFile(Document.open(doc.save(tmp_path / "middle.hwpx")).save(tmp_path / "returned.docx"))
    for name in ("Title", "Heading 1"):
        assert structure(returned.styles[name].element.rPr) == structure(original.styles[name].element.rPr)
        assert returned.styles[name].font.color.rgb == RGBColor(0, 0, 0)


def test_legacy_style_records_without_rpr_use_imported_character_values(tmp_path):
    doc = WordDocument.blank()
    with doc.edit_word_details() as edit:
        edit.format_style("Normal", font={"font_name": "Arial", "font_size_pt": 10})
        edit.body.append_paragraph("Legacy metadata")
    common = Document.open(doc.save(tmp_path / "legacy.hwpx"))
    for record in common._word_resources["paragraph_styles"]:
        record.pop("rPr", None)
    returned = WordFile(common.save(tmp_path / "legacy.docx"))
    assert returned.styles["Normal"].font.name == "Arial"
    assert returned.styles["Normal"].font.size.pt == 10


def test_an_intentional_imported_style_change_still_takes_effect(tmp_path):
    doc = WordDocument.blank()
    with doc.edit_word_details() as edit:
        edit.format_style("Normal", font={"font_name": "Arial", "font_size_pt": 10})
        edit.body.append_paragraph("Edited style")
    common = Document.open(doc.save(tmp_path / "middle.hwpx"))
    paragraph = next(block for block in common.sections[0].blocks if getattr(block, "text", None) == "Edited style")
    definition = next(style for style in common.style_definitions if style.style_id == paragraph.style_id)
    character = next(style for style in common.character_styles if style.style_id == definition.char_pr_id)
    character.height, character.text_color = 1800, "#224466"
    returned = WordFile(common.save(tmp_path / "edited.docx"))
    assert returned.styles["Normal"].font.size.pt == 18
    assert returned.styles["Normal"].font.color.rgb == RGBColor.from_string("224466")
    assert returned.styles["Normal"].element.find("./" + qn("w:rPr") + "/" + qn("w:rFonts")).get(qn("w:ascii")) == "Arial"


def test_unnamed_styles_keep_their_distinct_generated_names(tmp_path):
    from jakal_hwpx._hancom_models import StyleDefinition, CharacterStyle
    doc = Document.blank()
    doc.character_styles = [CharacterStyle(style_id="1", height=1200), CharacterStyle(style_id="2", height=1800)]
    doc.style_definitions = [StyleDefinition(None, style_id="1", char_pr_id="1"), StyleDefinition(None, style_id="2", char_pr_id="2")]
    doc.append_paragraph("First").style_id = "1"
    doc.append_paragraph("Second").style_id = "2"
    returned = WordFile(doc.save(tmp_path / "unnamed.docx"))
    assert returned.styles["Imported 1"].font.size.pt == 12
    assert returned.styles["Imported 2"].font.size.pt == 18
