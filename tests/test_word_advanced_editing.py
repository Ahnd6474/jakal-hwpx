"""Native structural edits must remain valid after saving and HWPX roundtrips."""
from __future__ import annotations

from unittest.mock import patch
import zipfile

from docx import Document as WordFile
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
import pytest

from jakal_hwpx import Document, WordDocument


def reopen(doc, tmp_path):
    return WordFile(doc.save(tmp_path / "edited.docx"))


def test_merged_grid_row_column_edits_and_split_preserve_content(tmp_path):
    doc = WordDocument.blank()
    doc.append_table(4, 4, [[f"{r},{c}" for c in range(4)] for r in range(4)])
    with doc.edit_word_details() as edit:
        table = edit.body.table(0)
        table.merge(0, 0, 2, 1)
        before = table.cell(0, 0).paragraph(0).text
        table.add_row(index=1)
        table.add_column(index=1, width_pt=30)
        assert table._native.cell(3, 3).text == "2,2"
        assert table._native.cell(0, 0)._tc.grid_span == 3
        # Deleting the origin row promotes the original content, including its runs.
        table.delete_row(0)
        assert table.cell(0, 0).paragraph(0).text == before
        table.delete_column(0)
        table.split_cell(1, 0)
        assert not table._node.xpath("./w:tr/w:tc/w:tcPr/w:vMerge | ./w:tr/w:tc/w:tcPr/w:gridSpan")
        table.format_cell(0, 0, borders={"bottom": {"style": "double", "width_pt": 1.5, "color": "224466"}})
    table = reopen(doc, tmp_path).tables[0]
    assert len(table.rows) == 4 and len(table.columns) == 4
    assert table.cell(0, 0).paragraphs[0].text == before
    border = table.cell(0, 0)._tc.xpath("./w:tcPr/w:tcBorders/w:bottom")[0]
    assert border.get(qn("w:sz")) == "12"


def test_grid_insert_values_and_failed_operations_are_atomic(tmp_path):
    doc = WordDocument.blank()
    doc.append_table(3, 3, [["keep", "b", "c"], ["d", "e", "f"], ["g", "h", "i"]])
    with doc.edit_word_details() as edit:
        table = edit.body.table(0)
        table.merge(0, 0, 1, 0)
        before = table._node.xml
        with pytest.raises(ValueError, match="continued"):
            table.add_row(["overwrite"], index=1)
        assert table._node.xml == before
        table.add_row([None, "new", "row"], index=1)
        table.add_column(["x", "y", "z", "w"], index=3)
        assert table.cell(1, 1).paragraph(0).text == "new"
        table.delete_column(2)
        table.delete_row(3)
    result = reopen(doc, tmp_path).tables[0]
    assert (len(result.rows), len(result.columns)) == (3, 3)
    assert result.cell(1, 2).text == "y"


def test_nested_table_merges_do_not_block_parent_structure(tmp_path):
    doc = WordDocument.blank()
    with doc.edit_word_details() as edit:
        table = edit.body.insert_table(0, rows=2, cols=2)
        nested = table.cell(0, 0).insert_table(1, rows=2, cols=2)
        nested.merge(0, 0, 1, 1)
        table.add_row(["a", "b"])
        table.add_column(index=1)
    result = reopen(doc, tmp_path).tables[0]
    assert len(result.rows) == 3 and len(result.columns) == 3
    assert len(result.cell(0, 0).tables[0].rows) == 2


def test_page_columns_and_inserted_section_boundaries(tmp_path):
    doc = WordDocument.blank()
    doc.append_paragraphs(["First", "Second", "Third"])
    with doc.edit_word_details() as edit:
        edit.section().format_page(width_pt=595, height_pt=842, margins_pt={"left": 40, "right": 40, "top": 45, "bottom": 45})
        following = edit.insert_section(1, break_type="next_page")
        following.format_page(orientation="landscape", header_distance_pt=18)
        following.format_columns(count=2, gap_pt=20)
        edit.section(0).append_paragraph("Before break")
        assert [p.text for p in edit.section(1).blocks] == ["Second", "Third"]
        with pytest.raises(ValueError, match="text area"):
            following.format_page(margins_pt={"left": 1000})
    word = reopen(doc, tmp_path)
    assert len(word.sections) == 2
    assert word.sections[0].page_width == Pt(595)
    assert word.sections[1].page_width == Pt(842)
    assert word.sections[1]._sectPr.xpath("./w:cols/@w:num") == ["2"]
    assert [p.text for p in word.paragraphs if p.text] == ["First", "Before break", "Second", "Third"]


def test_extended_character_paragraph_and_style_properties(tmp_path):
    doc = WordDocument.blank()
    doc.append_paragraph("한글 Latin\tTab")
    with doc.edit_word_details() as edit:
        p = edit.body.paragraph(0)
        p.format_text(0, 2, highlight="yellow", character_spacing_pt=-.5, scale_percent=90,
                      font_names={"eastAsia": "맑은 고딕", "ascii": "Arial"}, languages={"latin": "en-US", "eastAsia": "ko-KR"})
        p.format(line_spacing_pt=18, line_spacing_rule="at_least", tab_stops=[{"position_pt": 120, "alignment": "right", "leader": "dot"}])
        edit.format_style("Normal", font={"character_spacing_pt": .1}, paragraph={"line_spacing_pt": 20})
        with pytest.raises(ValueError, match="Choose"):
            p.format(line_spacing=1.5, line_spacing_pt=18)
        with pytest.raises(ValueError, match="unique"):
            p.format(tab_stops=[{"position_pt": 10}, {"position_pt": 10}])
    word = reopen(doc, tmp_path)
    p = word.paragraphs[0]
    assert p.runs[0].text == "한글"
    assert p._p.xpath("./w:r[1]/w:rPr/w:spacing/@w:val") == ["-10"]
    assert p._p.xpath("./w:r[1]/w:rPr/w:w/@w:val") == ["90"]
    assert p._p.xpath("./w:r[1]/w:rPr/w:lang/@w:eastAsia") == ["ko-KR"]
    assert p._p.xpath("./w:pPr/w:spacing/@w:lineRule") == ["atLeast"]
    assert p._p.xpath("./w:pPr/w:tabs/w:tab/@w:pos") == ["2400"]


def test_lists_restart_continue_remove_and_roundtrip(tmp_path):
    doc = WordDocument.blank()
    doc.append_paragraphs(["one", "two", "three", "plain", "nested"])
    with doc.edit_word_details() as edit:
        listing = edit.create_list([{"format": "decimal", "text": "%1."}, {"format": "lowerLetter", "text": "%1.%2)"}])
        edit.body.paragraph(0).set_list(listing, start=3)
        restarted = edit.body.paragraph(0).listing
        edit.body.paragraph(1).set_list(restarted)
        edit.body.paragraph(2).set_list(restarted, start=8)
        edit.body.paragraph(3).set_list(listing).remove_list()
        edit.body.paragraph(4).set_list(listing, level=1)
    native = doc.save(tmp_path / "middle.hwpx")
    restored = Document.open(native).save(tmp_path / "restored.docx")
    word = WordFile(restored)
    assert word.paragraphs[0]._p.xpath("./w:pPr/w:numPr/w:numId")
    assert word.part.numbering_part.element.xpath(".//w:startOverride/@w:val") == ["3", "8"]
    assert not word.paragraphs[3]._p.xpath("./w:pPr/w:numPr/w:numId[@w:val!='0']")


def test_edit_simple_and_complex_field_codes_and_results(tmp_path):
    doc = WordDocument.blank()
    doc.append_paragraph("Date: ")
    with doc.edit_word_details() as edit:
        field = edit.body.paragraph(0).add_field('DATE \\@ "yyyy-MM-dd"', "2026-11-04")
        assert field.result == "2026-11-04"
        field.set_instruction('DATE \\@ "yyyy/MM/dd"').set_result("2026/11/04").lock()
    with doc.edit_word() as backend:
        p = backend.add_paragraph()
        for kind in ["begin", "instruction", "separate", "result", "end"]:
            r = OxmlElement("w:r")
            item = OxmlElement("w:instrText" if kind == "instruction" else "w:t" if kind == "result" else "w:fldChar")
            if kind in {"instruction", "result"}:
                item.text = "AUTHOR" if kind == "instruction" else "Old"
            else:
                item.set(qn("w:fldCharType"), kind)
            r.append(item)
            p._p.append(r)
    with doc.edit_word_details() as edit:
        assert len(edit.body.fields) == 2
        field = edit.body.fields[1]
        assert field.instruction == "AUTHOR" and field.result == "Old"
        field.set_instruction('DOCPROPERTY "Title"').set_result("New\tTitle")
        assert field.result == "New\tTitle"
    result = reopen(doc, tmp_path)
    assert result.paragraphs[0]._p.xpath("./w:fldSimple/@w:fldLock") == ["true"]
    assert result.element.xpath(".//w:instrText/text()") == ['DOCPROPERTY "Title"']
    Document.open(doc.save(tmp_path / "middle.hwpx")).save(tmp_path / "restored.docx")
    assert WordFile(tmp_path / "restored.docx").element.xpath(".//w:fldSimple/@w:instr") == ['DATE \\@ "yyyy/MM/dd"', 'DOCPROPERTY "Title"']
    assert WordFile(tmp_path / "restored.docx").element.xpath(".//w:fldSimple[1]/@w:fldLock") == ["true"]


@pytest.mark.parametrize("kind", ["footnote", "endnote"])
def test_note_edits_reopened_generic_parts_and_reference_deletion(tmp_path, kind):
    doc = WordDocument.blank()
    doc.append_paragraph("Body")
    with doc.edit_word_details() as edit:
        note = edit.body.paragraph(0).add_note("Original note", kind=kind)
        note.format_reference(superscript=True, font_size_pt=9)
    doc = WordDocument.open(doc.save(tmp_path / "source.docx"))
    with doc.edit_word_details() as edit:
        note = getattr(edit, kind)()
        note.paragraph(0).replace_text("Original", "Edited")
        note.paragraph(0).format_text(bold=True)
        note.append_paragraph("Second paragraph")
        extra = edit.body.paragraph(0).add_note("Delete this", kind=kind)
        extra.delete()
    saved = doc.save(tmp_path / "edited.docx")
    with zipfile.ZipFile(saved) as package:
        xml = package.read("word/" + kind + "s.xml").decode()
        assert "Edited note" in xml and "Second paragraph" in xml and "Delete this" not in xml
    restored = Document.open(doc.save(tmp_path / "middle.hwpx")).save(tmp_path / "restored.docx")
    with zipfile.ZipFile(restored) as package:
        assert b"Edited note" in package.read("word/" + kind + "s.xml")
    with doc.edit_word_details() as edit:
        getattr(edit, kind)().delete()
    assert not reopen(doc, tmp_path).element.xpath(".//w:" + kind + "Reference")


def test_batch_append_refreshes_once_and_noop_does_not_reload():
    doc = WordDocument.blank()
    with patch.object(doc, "_reload_word_bytes", wraps=doc._reload_word_bytes) as reload:
        result = doc.append_paragraphs(str(index) for index in range(30))
        assert len(result) == 30
        assert reload.call_count == 1
    with patch.object(doc, "_reload_word_bytes", wraps=doc._reload_word_bytes) as reload:
        with doc.edit_word_details():
            pass
        assert reload.call_count == 0


@pytest.mark.parametrize("unequal", [False, True])
def test_multisection_column_settings_survive_native_roundtrip(tmp_path, unequal):
    from jakal_hwpx import HwpxDocument
    from jakal_hwpx.namespaces import NS
    doc = WordDocument.blank()
    doc.append_paragraphs(["Page one", "Page two"])
    with doc.edit_word_details() as edit:
        edit.section().format_page(width_pt=612, height_pt=792, margins_pt={"left": 36, "right": 36})
        edit.insert_section(1, break_type="odd_page").format_columns(count=2, gap_pt=18, widths_pt=[200, 322] if unequal else None, separator=True)
    native = doc.save(tmp_path / "columns.hwpx")
    hwpx = HwpxDocument.open(native)
    hwpx.strict_validate()
    columns = hwpx.sections[1].root_element.xpath(".//hp:colPr", namespaces=NS)[0]
    assert columns.get("colCount") == "2"
    assert columns.get("sameSz") == ("0" if unequal else "1")
    restored = Document.open(native).save(tmp_path / "restored.docx")
    word = WordFile(restored)
    assert len(word.sections) == 2
    assert word.sections[1]._sectPr.xpath("./w:type/@w:val") == ["oddPage"]
    assert word.sections[1]._sectPr.xpath("./w:cols/@w:num") == ["2"]
    if unequal:
        assert word.sections[1]._sectPr.xpath("./w:cols/w:col/@w:w") == ["4000", "6440"]


def test_field_only_footer_and_continuation_borders_survive_roundtrip(tmp_path):
    from jakal_hwpx import HwpxDocument
    from jakal_hwpx.namespaces import NS
    doc = WordDocument.blank()
    with doc.edit_word_details() as edit:
        table = edit.body.insert_table(0, rows=3, cols=2, values=[["Header", "Value"], ["Merged", "1"], ["", "2"]])
        table.merge(1, 0, 2, 0)
        table.format_cell(0, 0, fill_color="E8EDF2")
        table.format_cell(1, 0, borders={"left": {"width_pt": 1, "color": "223344"}})
        edit.footer().paragraph(0).add_field("PAGE", "1")
    native = doc.save(tmp_path / "middle.hwpx")
    hwpx = HwpxDocument.open(native)
    table = hwpx.sections[0].root_element.xpath(".//hp:tbl", namespaces=NS)[0]
    base = table.get("borderFillIDRef")
    from lxml import etree
    with zipfile.ZipFile(native) as package:
        header = etree.fromstring(package.read("Contents/header.xml"))
        neutral = header.xpath(".//hh:borderFill[@id=$identity]", namespaces=NS, identity=base)[0]
        assert not neutral.xpath("./hc:fillBrush", namespaces=NS)
    result = WordFile(Document.open(native).save(tmp_path / "restored.docx"))
    assert result.sections[0].footer._element.xpath(".//w:fldSimple/@w:instr") == ["PAGE"]
    physical = result.tables[0]._tbl.tr_lst[2].tc_lst[0]
    assert physical.xpath("./w:tcPr/w:tcBorders/w:left/@w:color") == ["223344"]


def test_field_without_separator_gets_a_valid_result_run(tmp_path):
    doc = WordDocument.blank()
    with doc.edit_word() as backend:
        p = backend.add_paragraph()
        p._p.append(OxmlElement("w:r"))
        begin = OxmlElement("w:fldChar")
        begin.set(qn("w:fldCharType"), "begin")
        p._p[-1].append(begin)
        instruction = OxmlElement("w:instrText")
        instruction.text = "AUTHOR"
        p.add_run()._r.append(instruction)
        end = OxmlElement("w:fldChar")
        end.set(qn("w:fldCharType"), "end")
        p.add_run()._r.append(end)
    with doc.edit_word_details() as edit:
        field = edit.body.fields[0]
        field.set_result("New Author")
        assert field.result == "New Author"
        assert field.instruction == "AUTHOR"
    p = reopen(doc, tmp_path).paragraphs[0]
    assert p._p.xpath("./w:r/w:fldChar/@w:fldCharType") == ["begin", "separate", "end"]
    assert p._p.xpath("./w:r/w:instrText/text()") == ["AUTHOR"]


def test_nested_field_results_and_instructions_are_not_flattened(tmp_path):
    doc = WordDocument.blank()
    with doc.edit_word_details() as edit:
        p = edit.body.append_paragraph()
        outer = p.add_field('IF 1 = 1 "yes" "no"', "")
        inner = OxmlElement("w:fldSimple")
        inner.set(qn("w:instr"), "AUTHOR")
        outer._node.append(inner)
        with pytest.raises(ValueError, match="nested"):
            outer.set_result("flatten")
        outer.set_instruction('IF 2 = 2 "yes" "no"')
        assert len(p.fields) == 2
    assert reopen(doc, tmp_path).element.xpath(".//w:fldSimple/@w:instr") == ['IF 2 = 2 "yes" "no"', "AUTHOR"]


def test_note_read_only_handle_preserves_opaque_part_bytes_and_can_move_reference(tmp_path):
    from lxml import etree
    doc = WordDocument.blank()
    doc.append_paragraphs(["Original reference", "New reference"])
    with doc.edit_word_details() as edit:
        edit.body.paragraph(0).add_note("Keep note")
    path = doc.save(tmp_path / "source.docx")
    with zipfile.ZipFile(path) as package:
        members = {info.filename: package.read(info.filename) for info in package.infolist()}
    members["word/footnotes.xml"] = etree.tostring(etree.fromstring(members["word/footnotes.xml"]), pretty_print=True)
    with zipfile.ZipFile(path, "w") as package:
        for name, data in members.items():
            package.writestr(name, data)
    original = path.read_bytes()
    doc = WordDocument.open(path)
    with doc.edit_word_details() as edit:
        assert edit.footnote().paragraph(0).text == "Keep note"
    assert doc.save(tmp_path / "noop.docx").read_bytes() == original
    with doc.edit_word_details() as edit:
        edit.footnote().move_reference(edit.body.paragraph(1))
    word = reopen(doc, tmp_path)
    assert not word.paragraphs[0]._p.xpath(".//w:footnoteReference")
    assert word.paragraphs[1]._p.xpath(".//w:footnoteReference/@w:id") == ["1"]


def test_section_split_and_unlink_preserve_header_content_and_relationships(tmp_path):
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    doc = WordDocument.blank()
    doc.append_paragraphs(["Section one", "Section two"])
    with doc.edit_word() as backend:
        header = backend.sections[0].header
        header.paragraphs[0].text = "Original header"
        rid = header.part.relate_to("https://example.com/header", RT.HYPERLINK, is_external=True)
        link = OxmlElement("w:hyperlink")
        link.set(qn("r:id"), rid)
        run = OxmlElement("w:r")
        text = OxmlElement("w:t")
        text.text = " Link"
        run.append(text)
        link.append(run)
        header.paragraphs[0]._p.append(link)
    with doc.edit_word_details() as edit:
        edit.insert_section(1)
    assert doc.word_document.sections[1].header.is_linked_to_previous
    with doc.edit_word_details() as edit:
        header = edit.header(1, unlink=True)
        assert header.paragraph(0).text == "Original header Link"
        header.paragraph(0).replace_text("Original", "Independent")
    word = reopen(doc, tmp_path)
    assert word.sections[0].header.paragraphs[0].text == "Original header Link"
    assert word.sections[1].header.paragraphs[0].text == "Independent header Link"
    assert word.sections[0].header.part is not word.sections[1].header.part
    for section in word.sections:
        identity = section.header._element.xpath(".//w:hyperlink/@r:id")[0]
        assert section.header.part.rels[identity].target_ref == "https://example.com/header"


def test_direct_tab_clear_keeps_other_inherited_tab_positions(tmp_path):
    from jakal_hwpx import HwpxDocument
    from jakal_hwpx.namespaces import NS
    doc = WordDocument.blank()
    doc.append_paragraph("A\tB")
    with doc.edit_word_details() as edit:
        edit.format_style("Normal", paragraph={"tab_stops": [{"position_pt": 60}, {"position_pt": 120, "alignment": "right"}]})
        edit.body.paragraph(0).format(tab_stops=[{"position_pt": 60, "alignment": "clear"}, {"position_pt": 180, "alignment": "decimal"}])
    native = doc.save(tmp_path / "tabs.hwpx")
    hwpx = HwpxDocument.open(native)
    header = hwpx.header.root_element
    candidates = header.xpath(".//hh:tabPr[hh:tabItem[@pos='12000'] and hh:tabItem[@pos='18000']]", namespaces=NS)
    assert candidates
    assert all(not item.xpath("./hh:tabItem[@pos='6000']", namespaces=NS) for item in candidates)


@pytest.mark.parametrize("explicit", [None, False, True])
def test_word_default_widow_control_survives_native_roundtrip(tmp_path, explicit):
    doc = WordDocument.blank()
    doc.append_paragraph("Keep both lines together at the next column.")
    if explicit is not None:
        with doc.edit_word_details() as edit:
            edit.body.paragraph(0).format(widow_control=explicit)
    restored = Document.open(doc.save(tmp_path / "middle.hwpx")).save(tmp_path / "restored.docx")
    assert WordFile(restored).paragraphs[0].paragraph_format.widow_control is (explicit is not False)


def test_pure_hwp_does_not_silently_discard_new_column_settings(tmp_path):
    doc = WordDocument.blank()
    doc.append_paragraph("Two columns")
    with doc.edit_word_details() as edit:
        edit.section().format_columns(count=2)
    path = tmp_path / "kept.hwp"
    path.write_bytes(b"Existing output")
    with pytest.raises(ValueError, match="multi-column"):
        doc.save(path)
    assert path.read_bytes() == b"Existing output"


def test_first_shaded_cell_is_not_reused_as_the_whole_table_background(tmp_path):
    from lxml import etree
    from jakal_hwpx.namespaces import NS
    doc = WordDocument.blank()
    with doc.edit_word_details() as edit:
        table = edit.body.insert_table(0, rows=2, cols=1, values=[["Shaded"], ["Plain"]])
        table.format_cell(0, 0, fill_color="E8EDF2")
    native = doc.save(tmp_path / "shaded.hwpx")
    with zipfile.ZipFile(native) as package:
        header = etree.fromstring(package.read("Contents/header.xml"))
        section = etree.fromstring(package.read("Contents/section0.xml"))
    table = section.xpath(".//hp:tbl", namespaces=NS)[0]
    base = table.get("borderFillIDRef")
    assert not header.xpath(".//hh:borderFill[@id=$identity]/hc:fillBrush", namespaces=NS, identity=base)
    assert table.xpath("./hp:tr[2]/hp:tc/@borderFillIDRef", namespaces=NS) == [base]
    assert table.xpath("./hp:tr[1]/hp:tc/@borderFillIDRef", namespaces=NS) != [base]
