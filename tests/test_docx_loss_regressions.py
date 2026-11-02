"""Missing optional parts, field ranges, and editable text-box loss regressions."""
from __future__ import annotations

from pathlib import Path
from io import BytesIO
import base64

from docx import Document as WordFile
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
import pytest

from jakal_hwpx import Document, HwpxDocument
from jakal_hwpx.namespaces import NS


def run(parent, text=None, *, control=None, instruction=None):
    node = OxmlElement("w:r")
    child = OxmlElement("w:fldChar" if control else "w:instrText" if instruction is not None else "w:t")
    if control:
        child.set(qn("w:fldCharType"), control)
    else:
        child.text = instruction if instruction is not None else text
        child.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    node.append(child)
    parent.append(node)
    return node


def simple(parent, instruction, value):
    node = OxmlElement("w:fldSimple")
    node.set(qn("w:instr"), instruction)
    run(node, value)
    parent.append(node)
    return node


def cycle(word, directory, *, edit=None):
    path = directory / "source.docx"
    word.save(path)
    source = Document.open(path)
    assert source.save(directory / "untouched.docx").read_bytes() == path.read_bytes()
    native_path = source.save(directory / "middle.hwpx")
    native = HwpxDocument.open(native_path)
    native.strict_validate()
    common = Document.open(native_path)
    if edit:
        edit(common)
    returned = WordFile(common.save(directory / "returned.docx"))
    return returned, native


def commands(word):
    return word.element.xpath(".//w:instrText/text() | .//w:fldSimple/@w:instr")


def test_docx_without_optional_numbering_part_opens_and_roundtrips(tmp_path: Path):
    word = WordFile()
    word.add_paragraph("Valid Word document without lists")
    for identity, relationship in list(word.part.rels.items()):
        if relationship.reltype == RT.NUMBERING:
            word.part.drop_rel(identity)
    word.save(tmp_path / "source.docx")
    assert not any(rel.reltype == RT.NUMBERING for rel in WordFile(tmp_path / "source.docx").part.rels.values())
    word, _ = cycle(word, tmp_path)
    assert word.paragraphs[0].text == "Valid Word document without lists"


def test_nested_simple_field_keeps_inner_value_and_command_during_edits(tmp_path: Path):
    word = WordFile()
    word.core_properties.title = "NESTED_VISIBLE_VALUE"
    outer = simple(word.add_paragraph()._p, 'IF 1 = 1 "ok" "no"', "PREFIX ")
    simple(outer, "DOCPROPERTY Title", "NESTED_VISIBLE_VALUE")
    run(outer, " SUFFIX")
    for index in range(3):
        word, native = cycle(word, tmp_path, edit=lambda document: document.replace_text("PREFIX", "EDITED"))
        assert commands(word) == ['IF 1 = 1 "ok" "no"', "DOCPROPERTY Title"]
        assert word.element.xpath(".//w:t/text()") == ["EDITED ", "NESTED_VISIBLE_VALUE", " SUFFIX"]
        assert "NESTED_VISIBLE_VALUE" in native.get_document_text()
        for control in native.fields():
            if control.parameter_map().get("JakalFieldKind") == "JAKAL_WORD_FIELD_BOUNDARY":
                assert "PREFIX" not in control.parameter_map()["Record"]


def test_nested_complex_instruction_is_not_rendered_as_cached_body_text(tmp_path: Path):
    word = WordFile()
    paragraph = word.add_paragraph()._p
    run(paragraph, control="begin")
    run(paragraph, instruction="IF ")
    run(paragraph, control="begin")
    run(paragraph, instruction="MERGEFIELD status")
    run(paragraph, control="separate")
    run(paragraph, "approved")
    run(paragraph, control="end")
    run(paragraph, instruction=' = "approved" "YES" "NO"')
    run(paragraph, control="separate")
    run(paragraph, "YES")
    run(paragraph, control="end")
    for _ in range(3):
        word, native = cycle(word, tmp_path)
        assert commands(word) == ["IF ", "MERGEFIELD status", ' = "approved" "YES" "NO"']
        assert native.get_document_text().strip() == "YES"
        assert word.element.xpath(".//w:fldChar/@w:fldCharType") == ["begin", "begin", "separate", "end", "separate", "end"]


@pytest.mark.parametrize("story", ["body", "header", "cell"])
def test_field_spanning_paragraphs_keeps_code_and_updated_results(tmp_path: Path, story):
    word = WordFile()
    target = word if story == "body" else word.sections[0].header if story == "header" else word.add_table(rows=1, cols=1).cell(0, 0)
    first = target.add_paragraph()._p
    run(first, control="begin")
    run(first, instruction='TOC \\o "1-3" \\h')
    run(first, control="separate")
    run(first, "Heading one 1")
    second = target.add_paragraph()._p
    run(second, "Heading two 2")
    run(second, control="end")
    for _ in range(3):
        word, native = cycle(word, tmp_path)
        root = word.element if story != "header" else word.sections[0].header._element
        assert root.xpath(".//w:instrText/text()") == ['TOC \\o "1-3" \\h']
        assert root.xpath(".//w:t/text()") == ["Heading one 1", "Heading two 2"]
        assert root.xpath(".//w:fldChar/@w:fldCharType") == ["begin", "separate", "end"]
        for control in native.fields():
            if control.parameter_map().get("JakalFieldKind") == "JAKAL_WORD_FIELD_BOUNDARY":
                assert "Heading" not in control.parameter_map()["Record"]


def test_native_boundary_parameters_can_be_normalized_without_losing_the_range(tmp_path: Path):
    word = WordFile()
    first = word.add_paragraph()._p
    run(first, control="begin")
    run(first, instruction='TOC \\o "1-3"')
    run(first, control="separate")
    run(first, "First entry")
    second = word.add_paragraph()._p
    run(second, "Second entry")
    run(second, control="end")
    _, native = cycle(word, tmp_path)
    for control in native.fields():
        parameters = control.element.find("hp:parameters", NS)
        if parameters is not None:
            control.element.remove(parameters)
            control.section.mark_modified()
    for text in native.sections[0].root_element.xpath(".//hp:t[text()='Second entry']", namespaces=NS):
        text.text = "Updated inside Hancom"
    native.sections[0].mark_modified()
    native.save(tmp_path / "native-edited.hwpx")
    result = WordFile(Document.open(tmp_path / "native-edited.hwpx").save(tmp_path / "native-edited.docx"))
    assert commands(result) == ['TOC \\o "1-3"']
    assert result.element.xpath(".//w:t/text()") == ["First entry", "Updated inside Hancom"]


def textbox(word, *, alternate=True):
    from jakal_hwpx.addons.docx.textboxes import NS as XMLNS

    namespaces = " ".join(f'xmlns:{prefix}="{uri}"' for prefix, uri in XMLNS.items())
    drawing = f'''<w:drawing {namespaces}><wp:anchor distT="0" distB="0" distL="114300" distR="114300"
        simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
        <wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
        <wp:positionV relativeFrom="paragraph"><wp:posOffset>189865</wp:posOffset></wp:positionV>
        <wp:extent cx="2540000" cy="1016000"/><wp:wrapNone/><wp:docPr id="41" name="Text Box 1"/>
        <wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="{XMLNS['wps']}"><wps:wsp>
        <wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2540000" cy="1016000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>
        <wps:txbx><w:txbxContent><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>BOX ORIGINAL</w:t></w:r></w:p>
        <w:p><w:r><w:t>Second box paragraph</w:t></w:r></w:p></w:txbxContent></wps:txbx>
        <wps:bodyPr anchor="t" lIns="91440" rIns="91440" tIns="45720" bIns="45720"><a:noAutofit/></wps:bodyPr>
        </wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>'''
    paragraph = word.add_paragraph("BODY")
    if alternate:
        drawing = f'''<mc:AlternateContent {namespaces}><mc:Choice Requires="wps">{drawing}</mc:Choice>
            <mc:Fallback><w:pict><v:shape style="width:200pt;height:80pt"><v:textbox><w:txbxContent>
            <w:p><w:r><w:t>BOX ORIGINAL</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent>'''
    node = parse_xml(drawing)
    paragraph.add_run()._r.append(node)
    return node


@pytest.mark.parametrize("alternate", [False, True])
def test_word_textbox_is_one_editable_native_shape_and_restores_fresh_content(tmp_path: Path, alternate):
    word = WordFile()
    mode = word.settings.element.find(qn("w:compat")).find(qn("w:compatSetting"))
    mode.set(qn("w:val"), "15")
    textbox(word, alternate=alternate)
    for index in range(3):
        word, native = cycle(word, tmp_path)
        assert len(native.shapes()) == 1
        shape = native.shapes()[0]
        assert "BOX ORIGINAL" in shape.text
        assert shape.size()["width"] == 20000
        assert shape.text_margins()["left"] == 720
        assert shape.element.xpath("./hp:drawText/hp:subList/@vertAlign", namespaces=NS) == ["TOP"]
        assert word.element.xpath(".//w:txbxContent//w:t/text()") == ["BOX ORIGINAL", "Second box paragraph"]
        assert len(word.element.xpath(".//w:txbxContent//w:rPr/w:b[not(@w:val='0' or @w:val='false')]")) == 1
        assert word.settings.element.xpath("./w:compat/w:compatSetting[@w:name='compatibilityMode']/@w:val") == ["15"]
    # A native edit must replace the projected content, never resurrect a saved
    # copy of the original text from a DrawingML/VML template.
    nodes = shape.element.xpath(".//hp:t[text()='BOX ORIGINAL']", namespaces=NS)
    nodes[0].text = "EDITED INSIDE HANCOM"
    shape.section.mark_modified()
    native.save(tmp_path / "edited.hwpx")
    result = WordFile(Document.open(tmp_path / "edited.hwpx").save(tmp_path / "edited.docx"))
    assert result.element.xpath(".//w:txbxContent//w:t/text()") == ["EDITED INSIDE HANCOM", "Second box paragraph"]
    for control in native.fields():
        if control.parameter_map().get("JakalFieldKind") == "JAKAL_WORD_FEATURE":
            assert "BOX ORIGINAL" not in control.parameter_map()["Record"]


def test_textbox_rich_links_images_tables_and_fields_keep_relationships(tmp_path: Path):
    from jakal_hwpx.addons.docx.textboxes import find

    word = WordFile()
    node = textbox(word, alternate=False)
    content = find(node, ".//w:txbxContent")[0]
    paragraph = word.add_paragraph("Linked ")
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), word.part.relate_to("https://example.com/box", RT.HYPERLINK, is_external=True))
    run(link, "box link")
    paragraph._p.append(link)
    simple(paragraph._p, 'DATE \\@ "yyyy-MM-dd"', "2026-10-29")
    image = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/l9sAAAAASUVORK5CYII=")
    paragraph.add_run().add_picture(BytesIO(image))
    content.append(paragraph._p)
    table = word.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "Table inside text box"
    content.append(table._tbl)
    content.append(OxmlElement("w:p"))
    for _ in range(3):
        word, native = cycle(word, tmp_path)
        assert "Table inside text box" in native.shapes()[0].text
        assert word.element.xpath(".//w:txbxContent//w:tbl//w:t/text()") == ["Table inside text box"]
        assert word.element.xpath(".//w:txbxContent//w:hyperlink//w:t/text()") == ["box link"]
        ids = word.element.xpath(".//w:txbxContent//w:hyperlink/@r:id")
        assert word.part.rels[ids[0]].target_ref == "https://example.com/box"
        pictures = word.element.xpath(".//w:txbxContent//a:blip/@r:embed")
        assert len(pictures) == 1 and word.part.related_parts[pictures[0]].blob == image
        assert 'DATE \\@ "yyyy-MM-dd"' in commands(word)
