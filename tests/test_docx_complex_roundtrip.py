"""Regressions found by foreground-only, multi-page native-app comparisons."""
from zipfile import ZipFile
from copy import deepcopy

from docx import Document as WordFile
from docx.shared import Pt
from lxml import etree
import pytest

from jakal_hwpx import Document, HwpxDocument
from jakal_hwpx.namespaces import NS, qname


def cycle(word, tmp_path):
    source = tmp_path / "source.docx"
    word.save(source)
    middle = Document.open(source).save(tmp_path / "middle.hwpx")
    return WordFile(Document.open(middle).save(tmp_path / "returned.docx")), middle


@pytest.mark.parametrize("target", ["paragraph", "table"])
def test_manual_break_stays_a_manual_break_in_word(tmp_path, target):
    word = WordFile()
    word.add_paragraph("First page")
    word.add_page_break()
    if target == "paragraph":
        word.add_heading("Following heading", 1)
    else:
        word.add_table(rows=1, cols=1).cell(0, 0).text = "Next page table"
    result, middle = cycle(word, tmp_path)
    assert result.element.xpath(".//w:br/@w:type") == ["page"]
    assert not result.element.xpath(".//w:pPr/w:pageBreakBefore")
    assert HwpxDocument.open(middle).sections[0].root_element.xpath("./hp:p[@pageBreak='1']", namespaces=NS)


def test_empty_run_keeps_its_own_font_size(tmp_path):
    word = WordFile()
    word.add_paragraph().add_run().font.size = Pt(4)
    result, middle = cycle(word, tmp_path)
    assert result.paragraphs[0].runs[0].font.size == Pt(4)
    model = Document.open(middle)
    empty = model.sections[0].blocks[0]
    assert int(next(s for s in model.character_styles if s.style_id == empty.char_pr_id).height) == 400


def test_table_cell_widths_need_not_equal_grid_widths(tmp_path):
    word = WordFile()
    table = word.add_table(rows=3, cols=3)
    for col, size in zip(table.columns, [80, 120, 160]):
        col.width = Pt(size)
    table.cell(0, 0).merge(table.cell(0, 2)).text = "Merged"
    table.cell(1, 0).merge(table.cell(2, 0)).text = "Vertical"
    before = table._tbl.xpath("./w:tr/w:tc/w:tcPr/w:tcW/@w:w")
    result, middle = cycle(word, tmp_path)
    assert result.tables[0]._tbl.xpath("./w:tr/w:tc/w:tcPr/w:tcW/@w:w") == before
    model = Document.open(middle)
    block = model.sections[0].blocks[0]
    block.col_widths = [8000, 12500, 16000]
    resized = WordFile(model.save(tmp_path / "resized.docx")).tables[0]
    assert resized._tbl.tblGrid.gridCol_lst[1].w == Pt(125)
    assert resized.cell(1, 1).width == Pt(125)


@pytest.mark.parametrize("style", ["Table Grid", "Light Shading Accent 1"])
def test_table_style_borders_stay_inherited_and_inner_edges_are_resolved(tmp_path, style):
    word = WordFile()
    table = word.add_table(rows=4, cols=3, style=style)
    for row in table.rows:
        for cell in row.cells:
            cell.text = "line one\nline two"
    result, middle = cycle(word, tmp_path)
    # Materializing an outer bottom border on interior cells changes row height
    # and suppresses Word's border at a continuation page's bottom.
    assert not result.tables[0]._tbl.xpath("./w:tr/w:tc/w:tcPr/w:tcBorders")
    native = HwpxDocument.open(middle)
    cell = native.tables()[0].cell(1, 1)
    fill = native.header.root_element.xpath(".//hh:borderFill[@id=$id]", namespaces=NS, id=str(cell.border_fill_id))[0]
    expected = "SOLID" if style == "Table Grid" else "NONE"
    assert fill.find(qname("hh", "bottomBorder")).get("type") == expected


def native_fixture(tmp_path):
    model = Document.blank()
    model.append_character_style(style_id="1", height=1100, font_names={"latin": "Arial", "hangul": "맑은 고딕"})
    model.append_paragraph("VISIBLE_SOURCE_TEXT_731").char_pr_id = "1"
    model.append_paragraph("").char_pr_id = "1"
    model.append_table(rows=2, cols=2, cell_texts=[["A", "B"], ["C", "D"]])
    native = model.to_hwpx_document()
    native.section_settings(0).set_page_size(width=59527, height=84188)
    native.section_settings(0).set_margins(left=5669, top=4251, bottom=4252, right=5671)
    style = next(s for s in native.character_styles() if s.style_id == "1").element
    style.find(qname("hh", "spacing")).set("latin", "-3")
    style.find(qname("hh", "strikeout")).set("shape", "3D")
    table = native.tables()[0]
    table.cell(0, 0).element.find(qname("hp", "cellSz")).set("width", "14321")
    native.header.mark_modified()
    return native.save(tmp_path / "source.hwpx")


def test_native_properties_keep_units_fonts_and_cell_geometry_without_storing_text(tmp_path):
    source = native_fixture(tmp_path)
    middle = Document.open(source).save(tmp_path / "middle.docx")
    with ZipFile(middle) as package:
        record = package.read("customXml/jakal-hwpx-properties.xml")
        assert b"VISIBLE_SOURCE_TEXT_731" not in record
        assert b"word/document.xml" not in record
    word = WordFile(middle)
    assert word.paragraphs[0].runs[0].font.strike is False
    doc = Document.open(middle)
    doc.replace_text("VISIBLE_SOURCE_TEXT_731", "EDITED_CURRENT_TEXT")
    returned = HwpxDocument.open(doc.save(tmp_path / "returned.hwpx"))
    settings = returned.section_settings(0)
    assert (settings.page_width, settings.page_height) == (59527, 84188)
    assert settings.margins()["left"] == 5669
    assert settings.margins()["bottom"] == 4252
    assert "EDITED_CURRENT_TEXT" in returned.get_document_text()
    assert "VISIBLE_SOURCE_TEXT_731" not in returned.get_document_text()
    assert returned.tables()[0].cell(0, 0).element.find(qname("hp", "cellSz")).get("width") == "14321"
    paragraph = next(p for p in returned.sections[0].root_element.xpath("./hp:p", namespaces=NS) if "EDITED_CURRENT_TEXT" in "".join(p.itertext()))
    char_id = paragraph.xpath("./hp:run[hp:t]/@charPrIDRef", namespaces=NS)[0]
    char = returned.header.root_element.xpath(".//hh:charPr[@id=$id]", namespaces=NS, id=char_id)[0]
    assert char.find(qname("hh", "spacing")).get("latin") == "-3"
    assert char.find(qname("hh", "strikeout")).get("shape") == "3D"


def test_native_properties_do_not_undo_word_format_or_table_edits(tmp_path):
    source = native_fixture(tmp_path)
    word = Document.open(Document.open(source).save(tmp_path / "middle.docx"))
    with word.edit_word() as backend:
        backend.paragraphs[0].runs[0].font.size = Pt(20)
        backend.tables[0].add_row().cells[0].text = "New row"
    returned = HwpxDocument.open(word.save(tmp_path / "returned.hwpx"))
    assert returned.tables()[0].row_count == 3
    paragraph = returned.sections[0].root_element.xpath("./hp:p[hp:run/hp:t[contains(.,'VISIBLE_SOURCE')]]", namespaces=NS)[0]
    char_id = paragraph.xpath("./hp:run[hp:t]/@charPrIDRef", namespaces=NS)[0]
    char = returned.header.root_element.xpath(".//hh:charPr[@id=$id]", namespaces=NS, id=char_id)[0]
    assert char.get("height") == "2000"


def test_native_multi_section_settings_have_distinct_records(tmp_path):
    model = Document.blank()
    model.append_paragraph("Section one")
    model.append_section()
    model.append_paragraph("Section two", section_index=1)
    model.sections[0].settings.page_width = 59527
    model.sections[1].settings.page_width = 60001
    source = model.save(tmp_path / "source.hwpx")
    middle = Document.open(source).save(tmp_path / "middle.docx")
    returned = Document.open(Document.open(middle).save(tmp_path / "returned.hwpx"))
    assert [s.settings.page_width for s in returned.sections] == [59527, 60001]


def test_absent_native_stories_stay_absent_and_word_changes_are_retained(tmp_path):
    source = native_fixture(tmp_path)
    middle = Document.open(source).save(tmp_path / "middle.docx")
    doc = Document.open(middle)
    assert not doc.sections[0].header_footer_blocks
    with doc.edit_word() as backend:
        backend.sections[0].header.paragraphs[0].add_run("New Word header")
    assert any("New Word header" in h.text for h in doc.sections[0].header_footer_blocks)
    native = Document.open(doc.save(tmp_path / "changed.hwpx"))
    assert any("New Word header" in h.text for h in native.sections[0].header_footer_blocks)


def test_native_equation_properties_keep_current_operands(tmp_path):
    model = Document.blank()
    equation = model.append_equation("{1234567} over {y}", width=6712, height=2345)
    equation.base_unit = 1300
    source = model.save(tmp_path / "source.hwpx")
    middle = Document.open(source).save(tmp_path / "middle.docx")
    with ZipFile(middle) as package:
        assert b"1234567" not in package.read("customXml/jakal-hwpx-properties.xml")
    doc = Document.open(middle)
    projected = next(b for b in doc.sections[0].blocks if type(b).__name__ == "Equation")
    assert (projected.base_unit, projected.width, projected.height) == (1300, 6712, 2345)
    with doc.edit_word() as backend:
        for node in backend.element.iter("{http://schemas.openxmlformats.org/officeDocument/2006/math}t"):
            if node.text == "1234567":
                node.text = "9"
    projected = next(b for b in doc.sections[0].blocks if type(b).__name__ == "Equation")
    assert "1234567" not in projected.script and "9" in projected.script
    assert projected.base_unit == 1300
    assert projected.width != 6712


def test_native_font_definitions_and_duplicate_face_ids_are_preserved(tmp_path):
    source = native_fixture(tmp_path)
    native = HwpxDocument.open(source)
    family = native.header.root_element.xpath(".//hh:fontface[@lang='LATIN']", namespaces=NS)[0]
    duplicate = deepcopy(family[0])
    duplicate.set("id", str(len(family)))
    duplicate.set("type", "HFT")
    family.append(duplicate)
    family.set("fontCnt", str(len(family)))
    next(s for s in native.character_styles() if s.style_id == "1").element.find(qname("hh", "fontRef")).set("latin", duplicate.get("id"))
    native.header.mark_modified()
    native.save(source)
    middle = Document.open(source).save(tmp_path / "middle.docx")
    returned = HwpxDocument.open(Document.open(middle).save(tmp_path / "returned.hwpx"))
    assert returned.header.root_element.xpath(".//hh:fontface[@lang='LATIN']/hh:font[@type='HFT']", namespaces=NS)
    body = returned.sections[0].root_element.xpath("./hp:p[hp:run/hp:t[contains(.,'VISIBLE_SOURCE')]]", namespaces=NS)[0]
    char_id = body.xpath("./hp:run[hp:t]/@charPrIDRef", namespaces=NS)[0]
    char = returned.header.root_element.xpath(".//hh:charPr[@id=$id]", namespaces=NS, id=char_id)[0]
    assert char.find(qname("hh", "fontRef")).get("latin") == duplicate.get("id")


def test_native_table_geometry_does_not_restore_old_spans_after_an_edit(tmp_path):
    model = Document.open(native_fixture(tmp_path))
    table = next(b for b in model.sections[0].blocks if type(b).__name__ == "Table")
    table.cell_spans = {(0, 0): (1, 2)}
    returned = HwpxDocument.open(model.save(tmp_path / "merged.hwpx"))
    returned.strict_validate()
    assert returned.tables()[0].cell(0, 0).col_span == 2


def test_table_first_metadata_does_not_precede_the_inline_table_and_uses_small_runs(tmp_path):
    word = WordFile()
    word.add_table(rows=2, cols=2).cell(0, 0).text = "Table first"
    _, middle = cycle(word, tmp_path)
    native = HwpxDocument.open(middle)
    first = native.sections[0].root_element.xpath("./hp:p", namespaces=NS)[0]
    assert not first.xpath("./hp:run/hp:ctrl/hp:fieldBegin", namespaces=NS)
    assert first.xpath("./hp:run/hp:tbl", namespaces=NS)
    markers = first.xpath(".//hp:fieldBegin[hp:parameters/*[@name='JakalFieldKind' and text()='JAKAL_WORD_FEATURE']]", namespaces=NS)
    assert markers
    heights = {s.style_id: int(s.height) for s in native.character_styles()}
    for field in markers:
        run = field.getparent().getparent()
        assert heights[run.get("charPrIDRef")] == 1


def test_native_default_zero_is_replaced_and_keeps_its_original_format(tmp_path):
    source = HwpxDocument.blank()
    source.set_paragraph_text(0, 0, "Default zero text")
    source.character_styles()[0].set_height(800)
    named_character = source.append_character_style(style_id="1", height=700)
    named_character.element.find(qname("hh", "spacing")).set("latin", "-7")
    source.styles()[0].element.set("charPrIDRef", "1")
    source.header.mark_modified()
    source.paragraph_styles()[0].set_alignment(horizontal="LEFT")
    initial = source.save(tmp_path / "source.hwpx")
    middle = Document.open(initial).save(tmp_path / "middle.docx")
    returned = HwpxDocument.open(Document.open(middle).save(tmp_path / "returned.hwpx"))
    root = returned.header.root_element
    chars = root.xpath(".//hh:charPr[@id='0']", namespaces=NS)
    paras = root.xpath(".//hh:paraPr[@id='0']", namespaces=NS)
    assert len(chars) == len(paras) == 1
    assert chars[0].get("height") == "800"
    assert paras[0].find(qname("hh", "align")).get("horizontal") == "LEFT"
    named = root.xpath(".//hh:style[@id='0']", namespaces=NS)[0]
    named_char = root.xpath(".//hh:charPr[@id=$id]", namespaces=NS, id=named.get("charPrIDRef"))[0]
    assert named_char.get("height") == "700"
    assert named_char.find(qname("hh", "spacing")).get("latin") == "-7"


def test_native_named_styles_keep_unique_ids_and_next_style_links(tmp_path):
    model = Document.blank()
    model.append_style("Native A", style_id="1", para_pr_id="0", char_pr_id="0", next_style_id="2")
    model.append_style("Native B", style_id="2", para_pr_id="0", char_pr_id="0", next_style_id="1")
    model.append_paragraph("First").style_id = "1"
    model.append_paragraph("Second").style_id = "2"
    initial = model.save(tmp_path / "source.hwpx")
    middle = Document.open(initial).save(tmp_path / "middle.docx")
    projected = Document.open(middle)
    ids = [style.style_id for style in projected.style_definitions]
    assert len(ids) == len(set(ids))
    returned = Document.open(projected.save(tmp_path / "returned.hwpx"))
    by_name = {style.name: style for style in returned.style_definitions}
    assert by_name["Native A"].next_style_id == by_name["Native B"].style_id
    assert by_name["Native B"].next_style_id == by_name["Native A"].style_id


def test_paragraph_inheritance_is_kept_but_a_native_margin_edit_is_applied(tmp_path):
    word = WordFile()
    word.styles["Normal"].paragraph_format.space_after = Pt(5)
    word.add_paragraph("Editable paragraph")
    _, middle = cycle(word, tmp_path)
    model = Document.open(middle)
    block = model.sections[0].blocks[0]
    style = next(s for s in model.paragraph_styles if s.style_id == block.para_pr_id)
    style.margins["next"] = "900"
    returned = WordFile(model.save(tmp_path / "edited.docx"))
    assert returned.paragraphs[0].paragraph_format.space_after == Pt(9)
    assert returned.paragraphs[0].paragraph_format.space_before is None


def test_native_line_cache_is_rebuilt_instead_of_reusing_obsolete_text_positions(tmp_path):
    source = native_fixture(tmp_path)
    native = HwpxDocument.open(source)
    paragraph = native.sections[0].root_element.xpath("./hp:p", namespaces=NS)[0]
    for old in paragraph.findall(qname("hp", "linesegarray")):
        paragraph.remove(old)
    array = etree.SubElement(paragraph, qname("hp", "linesegarray"))
    etree.SubElement(array, qname("hp", "lineseg"), textpos="0", vertpos="3333", vertsize="1100", textheight="1100", baseline="900", spacing="660", horzpos="0", horzsize="40000", flags="393216")
    native.sections[0].mark_modified()
    native.save(source)
    middle = Document.open(source).save(tmp_path / "middle.docx")
    doc = Document.open(middle)
    returned = HwpxDocument.open(doc.save(tmp_path / "unchanged.hwpx"))
    assert not returned.sections[0].root_element.xpath("./hp:p/hp:linesegarray/hp:lineseg[@vertpos='3333']", namespaces=NS)
    common = doc.to_common_document()
    common.replace_text("VISIBLE_SOURCE_TEXT_731", "CHANGED")
    changed = HwpxDocument.open(common.save(tmp_path / "common-edited.hwpx"))
    assert not changed.sections[0].root_element.xpath("./hp:p/hp:linesegarray/hp:lineseg[@vertpos='3333']", namespaces=NS)
    with doc.edit_word() as backend:
        backend.paragraphs[0].runs[0].font.size = Pt(22)
    changed = HwpxDocument.open(doc.save(tmp_path / "word-edited.hwpx"))
    assert not changed.sections[0].root_element.xpath("./hp:p/hp:linesegarray/hp:lineseg[@vertpos='3333']", namespaces=NS)


@pytest.mark.parametrize("via_word", [False, True])
def test_native_layout_version_and_empty_glyph_metrics_survive(tmp_path, via_word):
    from jakal_hwpx._hancom_text_layout import glyph_runs
    source = HwpxDocument.blank()
    version = source.xml_part("version.xml").element
    version.set("xmlVersion", "1.2")
    version.set("micro", "0")
    source.append_character_style(style_id="1", height=800)
    source.append_character_style(style_id="2", height=1200)
    paragraph = source.append_paragraph("current text", char_pr_id="1").element
    run = etree.SubElement(paragraph, qname("hp", "run"), charPrIDRef="2")
    etree.SubElement(run, qname("hp", "t"))
    etree.SubElement(paragraph, qname("hp", "run"), charPrIDRef="1")
    source.sections[0].mark_modified()
    path = source.save(tmp_path / "source.hwpx")
    model = Document.open(path)
    if via_word:
        model = Document.open(model.save(tmp_path / "middle.docx"))
    returned = HwpxDocument.open(model.save(tmp_path / "returned.hwpx"))
    assert returned.xml_part("version.xml").element.attrib == version.attrib
    target = next(p for p in returned.sections[0].root_element.findall(qname("hp", "p")) if "current text" in "".join(p.itertext()))
    styles = {s.style_id: int(s.height) for s in returned.character_styles()}
    assert [(styles[r.get("charPrIDRef")], bool(r.findall(qname("hp", "t")))) for r in glyph_runs(target)] == [(800, True), (1200, True), (800, False)]


@pytest.mark.parametrize("edit", ["text", "font"])
def test_native_empty_glyph_positions_are_discarded_after_word_edits(tmp_path, edit):
    source = HwpxDocument.blank()
    source.append_character_style(style_id="1", height=2700)
    paragraph = source.append_paragraph("Editable").element
    etree.SubElement(etree.SubElement(paragraph, qname("hp", "run"), charPrIDRef="1"), qname("hp", "t"))
    source.sections[0].mark_modified()
    middle = Document.open(source.save(tmp_path / "source.hwpx")).save(tmp_path / "middle.docx")
    word = Document.open(middle)
    with word.edit_word() as backend:
        paragraph = next(p for p in backend.paragraphs if p.text == "Editable")
        if edit == "text":
            paragraph.runs[0].text = "Changed"
        else:
            paragraph.runs[0].font.size = Pt(19)
    blocks = [b for s in word.sections for b in s.blocks if getattr(b, "text", "") in {"Editable", "Changed"}]
    assert blocks
    assert all("native_empty_runs" not in getattr(b, "_word_extra", {}) for b in blocks)


def test_native_property_extension_is_outside_layout_and_preserves_word_edits(tmp_path):
    from docx.oxml import OxmlElement
    from jakal_hwpx.addons.docx.hwpx_properties import PART
    source = native_fixture(tmp_path)
    word = Document.open(Document.open(source).save(tmp_path / "middle.docx"))
    with word.edit_word() as backend:
        backend.paragraphs[0].runs[0]._r.get_or_add_rPr().append(OxmlElement("w:rtl"))
    path = word.save(tmp_path / "edited.hwpx")
    raw = HwpxDocument.open(path)
    assert raw.get_part(PART)
    assert not raw.sections[0].root_element.xpath(".//hp:fieldBegin", namespaces=NS)
    reopened = Document.from_hwpx_document(raw)
    assert not raw.sections[0].root_element.xpath(".//hp:fieldBegin", namespaces=NS)
    result = WordFile(reopened.save(tmp_path / "returned.docx"))
    assert result.paragraphs[0]._p.xpath("./w:r/w:rPr/w:rtl")
    with ZipFile(path) as package:
        assert b"VISIBLE_SOURCE_TEXT_731" not in package.read(PART)
    raw.replace_text("VISIBLE_SOURCE_TEXT_731", "Current native edit")
    edited = Document.from_hwpx_document(raw)
    assert any("attachments were discarded" in message for message in edited.import_warnings)
    assert any("Current native edit" in getattr(block, "text", "") for section in edited.sections for block in section.blocks)


def test_native_print_settings_survive_without_the_source_caret(tmp_path):
    source = HwpxDocument.blank()
    source.set_paragraph_text(0, 0, "Print layout")
    settings = source.xml_part("settings.xml").element
    item_set = etree.SubElement(settings, qname("config", "config-item-set"), name="PrintInfo")
    etree.SubElement(item_set, qname("config", "config-item"), name="PrintMethod", type="short").text = "4"
    initial = source.save(tmp_path / "source.hwpx")
    middle = Document.open(initial).save(tmp_path / "middle.docx")
    returned = HwpxDocument.open(Document.open(middle).save(tmp_path / "returned.hwpx"))
    node = returned.xml_part("settings.xml").element
    assert node.xpath("./config:config-item-set/config:config-item[@name='PrintMethod']/text()", namespaces=NS) == ["4"]


def test_native_empty_glyph_offsets_survive_repeated_cycles_and_ignore_new_text(tmp_path):
    from jakal_hwpx._hancom_text_layout import capture_empty_runs
    native = HwpxDocument.blank()
    native.append_character_style(style_id="1", height=800)
    native.append_character_style(style_id="2", height=1700)
    paragraph = native.append_paragraph("AB", char_pr_id="1").element
    etree.SubElement(etree.SubElement(paragraph, qname("hp", "run"), charPrIDRef="2"), qname("hp", "t"))
    etree.SubElement(etree.SubElement(paragraph, qname("hp", "run"), charPrIDRef="1"), qname("hp", "t")).text = "CD"
    native.sections[0].mark_modified()
    path = native.save(tmp_path / "source.hwpx")
    for index in range(3):
        middle = Document.open(path).save(tmp_path / f"middle-{index}.docx")
        path = Document.open(middle).save(tmp_path / f"returned-{index}.hwpx")
        restored = HwpxDocument.open(path)
        paragraph = next(p for p in restored.sections[0].root_element.findall(qname("hp", "p")) if "ABCD" in "".join(p.itertext()))
        record = capture_empty_runs(paragraph)
        assert [item["offset"] for item in record["runs"]] == [2]
        heights = {style.style_id: int(style.height) for style in restored.character_styles()}
        assert heights[record["runs"][0]["character"]] == 1700
    model = Document.open(path)
    model.replace_text("AB", "A longer edit")
    current = HwpxDocument.open(model.save(tmp_path / "changed.hwpx"))
    paragraph = next(p for p in current.sections[0].root_element.findall(qname("hp", "p")) if "A longer edit" in "".join(p.itertext()))
    heights = {style.style_id: int(style.height) for style in current.character_styles()}
    assert all(heights[item["character"]] != 1700 for item in capture_empty_runs(paragraph)["runs"])
