"""Word detail transactions preserve XML objects, package parts, and boundaries."""
from __future__ import annotations

import io
from pathlib import Path
import struct
import zipfile
import zlib

from docx import Document as WordFile
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import Inches, Pt
import pytest

from jakal_hwpx import Document, WordDocument


def png(color):
    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(b"\0" + bytes(color))) + chunk(b"IEND", b"")


def opened(word, path):
    word.save(path)
    return Document.open(path)


def test_noop_editor_preserves_original_bytes_and_handles_expire(tmp_path: Path):
    model = WordDocument.blank()
    model.append_paragraph("Kept")
    original = model.save(tmp_path / "before.docx").read_bytes()
    with model.edit_word_details() as edit:
        paragraph = edit.body.paragraph(0)
        assert paragraph.text == "Kept"
    assert model.save(tmp_path / "after.docx").read_bytes() == original
    with pytest.raises(RuntimeError, match="expired"):
        paragraph.format_text(bold=True)
    with pytest.raises(RuntimeError, match="expired"):
        _ = edit.body


def test_insert_move_delete_and_partial_format_preserve_other_package_parts(tmp_path: Path):
    word = WordFile()
    word.add_paragraph("Alpha").runs[0].italic = True
    word.add_paragraph("Beta")
    word.add_paragraph("Delete me")
    source = tmp_path / "input.docx"
    word.save(source)
    with zipfile.ZipFile(source, "a") as package:
        package.writestr("customXml/opaque.xml", b"<opaque>retained</opaque>")
        package.comment = b"kept comment"
    model = Document.open(source)
    with model.edit_word_details() as edit:
        body = edit.body
        doomed = body.paragraph(2)
        doomed.delete()
        with pytest.raises(RuntimeError, match="removed"):
            _ = doomed.text
        body.paragraph(1).move(0)
        body.insert_paragraph(1, "Inserted").format(alignment="right", space_after_pt=9)
        body.paragraph(2).format_text(1, 4, bold=True, font_size_pt=14, color="#224466")
    output = model.save(tmp_path / "output.docx")
    result = WordFile(output)
    assert [p.text for p in result.paragraphs] == ["Beta", "Inserted", "Alpha"]
    assert [r.text for r in result.paragraphs[2].runs] == ["A", "lph", "a"]
    assert result.paragraphs[2].runs[1].bold
    assert all(r.italic for r in result.paragraphs[2].runs)
    assert result.paragraphs[2].runs[1].font.size == Pt(14)
    with zipfile.ZipFile(source) as a, zipfile.ZipFile(output) as b:
        assert a.comment == b.comment
        for name in a.namelist():
            if name != "word/document.xml":
                assert a.read(name) == b.read(name), name


def test_formatting_across_hyperlinks_keeps_fields_drawings_and_math(tmp_path: Path):
    word = WordFile()
    paragraph = word.add_paragraph("AB")
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), word.part.relate_to("https://example.com", RT.HYPERLINK, is_external=True))
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "CD"
    run.append(text)
    link.append(run)
    paragraph._p.append(link)
    paragraph.add_run().add_picture(io.BytesIO(png((20, 40, 60))), width=Inches(.2))
    math = OxmlElement("m:oMath")
    math_run = OxmlElement("m:r")
    math_text = OxmlElement("m:t")
    math_text.text = "x+y"
    math_run.append(math_text)
    math.append(math_run)
    paragraph._p.append(math)
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "DOCPROPERTY Title")
    paragraph._p.append(field)
    model = opened(word, tmp_path / "input.docx")
    with model.edit_word_details() as edit:
        edit.body.paragraph(0).format_text(1, 3, bold=True)
        with pytest.raises(ValueError, match="fields"):
            edit.body.paragraph(0).set_text("must not discard field")
    result = WordFile(model.save(tmp_path / "output.docx"))
    assert result.paragraphs[0].text == "ABCD"
    assert result.element.xpath(".//w:hyperlink//w:t/text()") == ["C", "D"]
    assert result.element.xpath(".//w:hyperlink//w:r[w:rPr/w:b]/w:t/text()") == ["C"]
    assert len(result.inline_shapes) == 1
    assert result.element.xpath(".//m:t/text()") == ["x+y"]
    assert result.element.xpath(".//w:fldSimple/@w:instr") == ["DOCPROPERTY Title"]


def test_table_editing_keeps_existing_runs_and_rejects_aliased_cells(tmp_path: Path):
    word = WordFile()
    table = word.add_table(rows=2, cols=2)
    table.cell(0, 0).paragraphs[0].add_run("Header").bold = True
    table.cell(1, 0).text = "Old row"
    model = opened(word, tmp_path / "input.docx")
    with model.edit_word_details() as edit:
        table = edit.body.table(0)
        table.set_cell_text(0, 0, "Updated header")
        table.format(widths_pt=[120, 180], repeat_header=True)
        table.format_cell(0, 0, fill_color="#DDEEFF", vertical_alignment="center", padding_pt={"left": 6})
        table.format_row(0, height_pt=24, exact=True)
        table.add_row(["Inserted", "7"], index=1)
        table.delete_row(2)
        with pytest.raises(IndexError):
            table.set_cell_text(0, 2, "must not overwrite next row")
    result = WordFile(model.save(tmp_path / "output.docx"))
    table = result.tables[0]
    assert len(table.rows) == 2 and table.cell(1, 0).text == "Inserted"
    assert any(run.bold for run in table.cell(0, 0).paragraphs[0].runs)
    assert table.cell(0, 0)._tc.xpath("./w:tcPr/w:shd/@w:fill") == ["DDEEFF"]
    assert table.cell(0, 0)._tc.xpath("./w:tcPr/w:tcMar/w:left/@w:w") == ["120"]
    assert table.rows[0].height == Pt(24)
    assert table.rows[0]._tr.xpath("./w:trPr/w:tblHeader")


def test_merging_and_widths_follow_the_actual_grid(tmp_path: Path):
    model = WordDocument.blank()
    model.append_table(2, 2, [["A", "B"], ["C", "D"]])
    with model.edit_word_details() as edit:
        table = edit.body.table(0)
        table.merge(0, 0, 0, 1).format(widths_pt=[100, 150])
        table.add_row(["new row"])
    result = WordFile(model.save(tmp_path / "merged.docx"))
    assert result.tables[0].cell(0, 0).width == Pt(250)
    assert result.tables[0].cell(0, 0).text == "A\nB"
    assert result.tables[0].cell(2, 0).text == "new row"


def test_picture_replacement_does_not_change_other_instances_of_shared_media(tmp_path: Path):
    original, replacement = png((200, 100, 50)), png((30, 70, 110))
    word = WordFile()
    for _ in range(2):
        word.add_paragraph().add_run().add_picture(io.BytesIO(original), width=Inches(.5))
    model = opened(word, tmp_path / "input.docx")
    with model.edit_word_details() as edit:
        picture = edit.body.pictures[0]
        picture.replace_image(replacement).resize(width_pt=72).crop(left=.1, right=.2).rotate(30)
        picture.position(x_pt=10, y_pt=20, horizontal_from="margin", vertical_from="paragraph", wrap="square")
    result = WordFile(model.save(tmp_path / "output.docx"))
    ids = result.element.xpath(".//a:blip/@r:embed")
    assert result.part.related_parts[ids[0]].blob == replacement
    assert result.part.related_parts[ids[1]].blob == original
    assert result.element.xpath(".//wp:anchor/wp:extent/@cx") == [str(Inches(1))]
    assert result.element.xpath(".//wp:anchor/wp:positionH/wp:posOffset/text()") == [str(Pt(10))]
    assert result.element.xpath(".//a:srcRect/@l") == ["10000"]
    assert result.element.xpath(".//a:xfrm/@rot") == ["1800000"]


def test_nonfinal_section_appends_and_positioned_insertions_keep_section_properties(tmp_path: Path):
    word = WordFile()
    word.add_paragraph("First")
    word.add_section()
    word.add_paragraph("Second")
    model = opened(word, tmp_path / "input.docx")
    model.append_paragraph("Tail of first", section_index=0)
    model.append_table(1, 1, [["First section table"]], section_index=0)
    model.append_picture("image.png", png((40, 50, 60)), width=720, height=720, section_index=0)
    model.append_hyperlink("https://example.com", display_text="First section link", section_index=0)
    with model.edit_word_details() as edit:
        edit.section(1).insert_paragraph(0, "Start of second")
    result = WordFile(model.save(tmp_path / "output.docx"))
    assert len(result.sections) == 2
    texts = [[item.text for item in section.iter_inner_content() if hasattr(item, "text")] for section in result.sections]
    assert "Tail of first" in texts[0] and "First section link" in texts[0]
    assert texts[1] == ["Start of second", "Second"]
    assert model.sections[0].blocks and model.sections[1].blocks


def test_story_editing_requires_unlinking_and_keeps_previous_header(tmp_path: Path):
    word = WordFile()
    word.add_paragraph("First")
    word.sections[0].header.paragraphs[0].text = "Original header"
    word.add_section()
    word.add_paragraph("Second")
    model = opened(word, tmp_path / "input.docx")
    before = model.save(tmp_path / "before.docx").read_bytes()
    with pytest.raises(ValueError, match="unlink"):
        with model.edit_word_details() as edit:
            edit.header(1)
    assert model.save(tmp_path / "rejected.docx").read_bytes() == before
    with model.edit_word_details() as edit:
        edit.header(1, unlink=True).paragraph(0).set_text("Second only")
        edit.footer(1, variant="first", unlink=True).paragraph(0).set_text("First page footer")
    result = WordFile(model.save(tmp_path / "output.docx"))
    assert result.sections[0].header.paragraphs[0].text == "Original header"
    assert result.sections[1].header.paragraphs[0].text == "Second only"
    assert result.sections[1].first_page_footer.paragraphs[0].text == "First page footer"


@pytest.mark.parametrize("story", ["body", "header"])
def test_textbox_edits_update_the_fallback_without_keeping_stale_cache(tmp_path: Path, story):
    from test_docx_loss_regressions import textbox
    word = WordFile()
    node = textbox(word)
    if story == "header":
        word.sections[0].header.paragraphs[0]._p.append(node.getparent())
    model = opened(word, tmp_path / "input.docx")
    with model.edit_word_details() as edit:
        scope = edit.body if story == "body" else edit.header()
        scope.textbox().paragraph(0).replace_text("BOX ORIGINAL", "BOX EDITED")
    result = WordFile(model.save(tmp_path / "output.docx"))
    root = result.element if story == "body" else result.sections[0].header._element
    assert root.xpath(".//w:txbxContent//w:t/text()").count("BOX EDITED") == 2
    assert "BOX ORIGINAL" not in root.xpath(".//w:txbxContent//w:t/text()")


def test_transaction_failure_and_nested_edit_restore_original_bytes(tmp_path: Path):
    model = WordDocument.blank()
    model.append_paragraph("Kept")
    before = model.save(tmp_path / "before.docx").read_bytes()
    with pytest.raises(ValueError, match="half a point"):
        with model.edit_word_details() as edit:
            edit.body.append_paragraph("Discarded")
            edit.body.paragraph(0).format_text(font_size_pt=11.25)
    assert model.save(tmp_path / "after.docx").read_bytes() == before
    with pytest.raises(RuntimeError, match="Nested"):
        with model.edit_word_details():
            with model.edit_word():
                pass
    assert model.save(tmp_path / "nested.docx").read_bytes() == before


def test_partial_field_moves_and_deletion_are_rejected(tmp_path: Path):
    word = WordFile()
    paragraph = word.add_paragraph("Before")
    for kind in ("begin", "separate"):
        run = OxmlElement("w:r")
        field = OxmlElement("w:fldChar")
        field.set(qn("w:fldCharType"), kind)
        run.append(field)
        paragraph._p.append(run)
    paragraph = word.add_paragraph("After")
    run = OxmlElement("w:r")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run.append(end)
    paragraph._p.append(run)
    model = opened(word, tmp_path / "input.docx")
    with model.edit_word_details() as edit:
        with pytest.raises(ValueError, match="field"):
            edit.body.paragraph(0).delete()
        with pytest.raises(ValueError, match="field"):
            edit.body.paragraph(0).move(0, destination=edit.header())
    result = WordFile(model.save(tmp_path / "output.docx"))
    assert result.element.xpath(".//w:fldChar/@w:fldCharType") == ["begin", "separate", "end"]


def test_cell_movement_and_style_edits_use_the_native_backend(tmp_path: Path):
    model = WordDocument.blank()
    model.append_paragraph("Move into cell")
    model.append_table(1, 1, [["Cell"]])
    model.word_document.styles["Normal"].element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:cstheme"), "majorBidi")
    with model.edit_word_details() as edit:
        edit.format_style("Normal", font={"font_name": "Arial", "font_size_pt": 11}, paragraph={"space_after_pt": 6})
        table = edit.body.table(0)
        with pytest.raises(ValueError, match="inside itself"):
            table.move(0, destination=table.cell(0, 0))
        edit.body.paragraph(0).move(1, destination=table.cell(0, 0))
    result = WordFile(model.save(tmp_path / "output.docx"))
    assert result.tables[0].cell(0, 0).text == "Cell\nMove into cell"
    assert result.styles["Normal"].font.name == "Arial"
    assert not result.styles["Normal"].element.xpath("./w:rPr/w:rFonts/@w:cstheme")
    assert result.styles["Normal"].paragraph_format.space_after == Pt(6)


@pytest.mark.parametrize("format_only", [False, True])
@pytest.mark.parametrize("move_control", [False, True])
def test_bound_content_control_edits_update_xml_and_preserve_the_wrapper(tmp_path: Path, format_only, move_control):
    from docx.opc.part import Part
    from docx.opc.packuri import PackURI
    from lxml import etree

    word = WordFile()
    data = Part(PackURI("/customXml/customer.xml"), "application/xml", b"<customer><name><first>Alice</first></name></customer>", word.part.package)
    props = Part(PackURI("/customXml/customerProps.xml"), "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
                 b'<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="{11111111-1111-1111-1111-111111111111}"/>', word.part.package)
    data.relate_to(props, RT.CUSTOM_XML_PROPS)
    word.part.relate_to(data, RT.CUSTOM_XML)
    paragraph = word.add_paragraph("Alice")
    wrapper = OxmlElement("w:sdt")
    properties = OxmlElement("w:sdtPr")
    binding = OxmlElement("w:dataBinding")
    binding.attrib.update({qn("w:xpath"): "/customer/name", qn("w:storeItemID"): "{11111111-1111-1111-1111-111111111111}"})
    properties.append(binding)
    wrapper.append(properties)
    content = OxmlElement("w:sdtContent")
    wrapper.append(content)
    paragraph._p.addprevious(wrapper)
    content.append(paragraph._p)
    model = opened(word, tmp_path / "input.docx")
    before = model.save(tmp_path / "before.docx").read_bytes()
    with model.edit_word_details():
        pass
    assert model.save(tmp_path / "noop.docx").read_bytes() == before
    with model.edit_word_details() as edit:
        paragraph = edit.body.paragraph(0)
        with pytest.raises(ValueError, match="Wrapped"):
            paragraph.delete()
        if format_only:
            paragraph.format_text(0, 2, bold=True)
        else:
            paragraph.replace_text("Alice", "Bob")
        if move_control:
            edit.body.content_controls[0].move(0, destination=edit.header())
        edit.body.insert_paragraph(0, "Outside control")
    result = model.save(tmp_path / "output.docx")
    with zipfile.ZipFile(result) as package:
        if format_only:
            assert package.read("customXml/customer.xml") == data.blob
        else:
            assert etree.fromstring(package.read("customXml/customer.xml")).findtext("name") == "Bob"
    word = WordFile(result)
    root = word.sections[0].header._element if move_control else word.element
    assert "".join(root.xpath(".//w:sdt//w:t/text()")) == ("Alice" if format_only else "Bob")
    assert word.paragraphs[0].text == "Outside control"


def test_edited_style_absence_does_not_inherit_a_different_output_template_border(tmp_path: Path, monkeypatch):
    from jakal_hwpx.addons.docx import dependencies

    word = WordFile()
    word.add_paragraph("Plain title", style="Title")
    properties = word.styles["Title"].element.get_or_add_pPr()
    for node in list(properties.findall(qn("w:pBdr"))):
        properties.remove(node)
    model = opened(word, tmp_path / "input.docx")
    with model.edit_word_details() as edit:
        edit.format_style("Title", font={"color": "000000"}, paragraph={"space_after_pt": 12})
    middle = model.save(tmp_path / "middle.hwpx")

    def other_template(*args, **kwargs):
        result = WordFile(*args, **kwargs)
        if not args and not kwargs:
            borders = OxmlElement("w:pBdr")
            bottom = OxmlElement("w:bottom")
            bottom.set(qn("w:val"), "single")
            borders.append(bottom)
            result.styles["Title"].element.get_or_add_pPr().append(borders)
        return result

    monkeypatch.setattr(dependencies, "document_factory", lambda: other_template)
    result = WordFile(Document.open(middle).save(tmp_path / "returned.docx"))
    assert not result.styles["Title"].element.xpath("./w:pPr/w:pBdr")
    assert result.styles["Title"].paragraph_format.space_after == Pt(12)
