"""Native appearance plus reversible Word semantics for the extended addon."""
from __future__ import annotations

import base64
from io import BytesIO
from pathlib import Path
import zipfile

from docx import Document as WordFile
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches
import pytest

from jakal_hwpx import Document, HwpxDocument

PNG = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/l9sAAAAASUVORK5CYII=")


def cycle(word, directory, *, edit=None):
    word.save(directory / "source.docx")
    middle = Document.open(directory / "source.docx").save(directory / "middle.hwpx")
    HwpxDocument.open(middle).strict_validate()
    common = Document.open(middle)
    if edit:
        edit(common)
    result = common.save(directory / "result.docx")
    return WordFile(result), HwpxDocument.open(middle)


def element(parent, tag, **attrs):
    node = OxmlElement(tag)
    for key, value in attrs.items():
        node.set(qn("w:" + key), str(value))
    parent.append(node)
    return node


@pytest.mark.parametrize("floating", [False, True])
def test_picture_crop_rotation_and_anchor_are_native_and_reversible(tmp_path: Path, floating):
    word = WordFile()
    shape = word.add_paragraph().add_run().add_picture(BytesIO(PNG), width=Inches(1), height=Inches(.5))
    fill = shape._inline.graphic.graphicData.pic.blipFill
    crop = OxmlElement("a:srcRect")
    crop.attrib.update({"l": "10000", "t": "20000", "r": "15000", "b": "5000"})
    fill.insert(1, crop)
    shape._inline.graphic.graphicData.pic.spPr.xfrm.set("rot", "1800000")
    if floating:
        from jakal_hwpx.addons.docx.pictures import _position
        from jakal_hwpx._hancom_models import Picture

        picture = Picture("p.png", PNG, width=7200, height=3600,
                          layout={"treatAsChar": "0", "textWrap": "SQUARE", "horzRelTo": "PAPER",
                                  "vertRelTo": "PAGE", "horzOffset": "1800", "vertOffset": "900"},
                          out_margins={"left": 120, "right": 200, "top": 80, "bottom": 100})
        _position(shape._inline.getparent(), picture)
    result, native = cycle(word, tmp_path)
    pic = native.pictures()[0]
    assert pic.crop() == {"left": 720, "top": 720, "right": 6120, "bottom": 3420}
    assert float(pic.rotation()["angle"]) == 30
    assert pic.layout()["treatAsChar"] == ("0" if floating else "1")
    if floating:
        assert pic.layout()["horzOffset"] == "1800"
        assert result.element.xpath(".//wp:anchor/wp:positionH/wp:posOffset/text()") == [str(1800 * 127)]
    assert result.element.xpath(".//a:srcRect/@l") == ["10000"]
    assert result.element.xpath(".//a:xfrm/@rot") == ["1800000"]


@pytest.mark.parametrize("block_control", [False, True])
def test_content_control_properties_and_edited_content_survive(tmp_path: Path, block_control):
    word = WordFile()
    paragraph = word.add_paragraph("bound text")
    sdt = OxmlElement("w:sdt")
    properties = element(sdt, "w:sdtPr")
    element(properties, "w:id", val=4321)
    element(properties, "w:tag", val="customer-name")
    element(properties, "w:alias", val="Customer")
    element(properties, "w:lock", val="sdtLocked")
    element(properties, "w:text")
    content = element(sdt, "w:sdtContent")
    if block_control:
        paragraph._p.addprevious(sdt)
        content.append(paragraph._p)
        second = word.add_paragraph("second line")
        content.append(second._p)
    else:
        run = paragraph._p[-1]
        run.addprevious(sdt)
        content.append(run)
    result, _ = cycle(word, tmp_path, edit=lambda doc: doc.replace_text("bound text", "edited value"))
    controls = result.element.xpath(".//w:sdt")
    assert len(controls) == 1
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    assert controls[0].xpath("./w:sdtPr/w:tag/@w:val", namespaces=ns) == ["customer-name"]
    assert controls[0].xpath("./w:sdtPr/w:lock/@w:val", namespaces=ns) == ["sdtLocked"]
    assert "edited value" in "".join(controls[0].xpath(".//w:t/text()", namespaces=ns))
    if block_control:
        assert len(controls[0].xpath("./w:sdtContent/w:p", namespaces=ns)) == 2


def test_word_character_and_paragraph_details_survive_three_cycles(tmp_path: Path):
    word = WordFile()
    paragraph = word.add_paragraph("precise formatting")
    pr = paragraph._p.get_or_add_pPr()
    element(pr, "w:jc", val="thaiDistribute")
    element(pr, "w:bidi")
    element(pr, "w:contextualSpacing")
    rpr = paragraph.runs[0]._r.get_or_add_rPr()
    element(rpr, "w:spacing", val=17)
    element(rpr, "w:position", val=4)
    element(rpr, "w:u", val="dashDotHeavy", color="AA2255")
    element(rpr, "w:shd", val="diagCross", fill="AADDFF", color="223344")
    for index in range(3):
        directory = tmp_path / str(index)
        directory.mkdir()
        word, _ = cycle(word, directory)
        assert word.paragraphs[0]._p.xpath("./w:pPr/w:jc/@w:val") == ["thaiDistribute"]
        run = word.paragraphs[0].runs[0]
        assert run._r.xpath("./w:rPr/w:spacing/@w:val") == ["17"]
        assert run._r.xpath("./w:rPr/w:position/@w:val") == ["4"]
        assert run._r.xpath("./w:rPr/w:u/@w:val") == ["dashDotHeavy"]
        assert run._r.xpath("./w:rPr/w:shd/@w:val") == ["diagCross"]


def test_linked_picture_is_embedded_in_hwpx_and_link_restored(tmp_path: Path):
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    word = WordFile()
    shape = word.add_paragraph().add_run().add_picture(BytesIO(PNG), width=Inches(1))
    (tmp_path / "image.png").write_bytes(PNG)
    blip = shape._inline.graphic.graphicData.pic.blipFill.blip
    blip.attrib.pop(qn("r:embed"))
    blip.set(qn("r:link"), word.part.relate_to("image.png", RT.IMAGE, is_external=True))
    result, native = cycle(word, tmp_path)
    assert native.pictures()[0].binary_data() == PNG
    links = result.element.xpath(".//a:blip/@r:link")
    assert len(links) == 1 and result.part.rels[links[0]].target_ref == "image.png"


def test_image_inside_hyperlink_survives(tmp_path: Path):
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    word = WordFile()
    paragraph = word.add_paragraph()
    shape = paragraph.add_run().add_picture(BytesIO(PNG), width=Inches(.2))
    run = shape._inline.getparent().getparent()
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), word.part.relate_to("https://example.com/image", RT.HYPERLINK, is_external=True))
    run.addprevious(link)
    link.append(run)
    result, _ = cycle(word, tmp_path)
    assert len(result.element.xpath(".//w:hyperlink/w:r/w:drawing")) == 1


@pytest.mark.parametrize("edit", [False, True])
def test_custom_xml_bindings_survive_without_embedding_original_package(tmp_path: Path, edit):
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.opc.packuri import PackURI
    from docx.opc.part import Part

    word = WordFile()
    data = Part(PackURI("/customXml/customer.xml"), "application/xml", b'<customer><name>Alice</name></customer>', word.part.package)
    props = Part(PackURI("/customXml/customerProps.xml"), "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
                 b'<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="{11111111-1111-1111-1111-111111111111}"/>', word.part.package)
    data.relate_to(props, RT.CUSTOM_XML_PROPS)
    word.part.relate_to(data, RT.CUSTOM_XML)
    paragraph = word.add_paragraph("Alice")
    run = paragraph.runs[0]._r
    sdt = OxmlElement("w:sdt")
    pr = element(sdt, "w:sdtPr")
    element(pr, "w:dataBinding", xpath="/customer/name", storeItemID="{11111111-1111-1111-1111-111111111111}")
    content = element(sdt, "w:sdtContent")
    run.addprevious(sdt)
    content.append(run)
    result, _ = cycle(word, tmp_path, edit=(lambda doc: doc.replace_text("Alice", "Bob")) if edit else None)
    assert result.element.xpath(".//w:dataBinding/@w:xpath") == ["/customer/name"]
    with zipfile.ZipFile(tmp_path / "result.docx") as package:
        from lxml import etree

        assert etree.fromstring(package.read("customXml/customer.xml")).findtext("name") == ("Bob" if edit else "Alice")
        if not edit:
            assert package.read("customXml/customer.xml") == data.blob
        assert b"11111111" in package.read("customXml/customerProps.xml")
    with zipfile.ZipFile(tmp_path / "middle.hwpx") as package:
        assert not any(name.endswith(".docx") for name in package.namelist())


def test_conditional_table_style_theme_and_direct_override_are_native(tmp_path: Path):
    from docx.enum.style import WD_STYLE_TYPE
    from jakal_hwpx.hancom_document import Table

    word = WordFile()
    style = word.styles.add_style("Conditional demo", WD_STYLE_TYPE.TABLE)
    for region, fill, text_color in (("firstRow", "123456", "FFFFFF"), ("band1Horz", "CCDDEE", "112233"),
                                      ("lastRow", "334455", "FFFFFF")):
        conditional = element(style.element, "w:tblStylePr", type=region)
        cell = element(conditional, "w:tcPr")
        element(cell, "w:shd", val="clear", fill=fill)
        run = element(conditional, "w:rPr")
        element(run, "w:color", val=text_color)
    table = word.add_table(rows=4, cols=2)
    table.style = style
    look = table._tbl.tblPr.find(qn("w:tblLook"))
    for name, value in {"firstRow": "1", "lastRow": "1", "firstColumn": "0", "lastColumn": "0", "noHBand": "0", "noVBand": "1"}.items():
        look.set(qn("w:" + name), value)
    for row in table.rows:
        for cell in row.cells:
            cell.text = "Styled"
    element(table.cell(0, 1)._tc.get_or_add_tcPr(), "w:shd", val="diagCross", fill="FFCC88", color="223344")
    for index in range(3):
        directory = tmp_path / str(index)
        directory.mkdir()
        word, native = cycle(word, directory)
        block = next(block for block in Document.from_hwpx_document(native).sections[0].blocks if isinstance(block, Table))
        fills = getattr(block, "_docx_table_details")["cell_fill_colors"]
        assert fills[(0, 0)] == "#123456"
        assert fills[(0, 1)] == "#FFCC88"
        assert fills[(1, 0)] == "#CCDDEE"
        assert fills[(3, 0)] == "#334455"
        assert word.tables[0].style.name == "Conditional demo"
        assert word.tables[0].cell(0, 1)._tc.xpath("./w:tcPr/w:shd/@w:val") == ["diagCross"]
        assert len(word.styles["Conditional demo"].element.xpath("./w:tblStylePr")) == 3


def test_theme_color_tint_and_native_spacing_are_resolved(tmp_path: Path):
    from jakal_hwpx.namespaces import NS

    word = WordFile()
    run = word.add_paragraph("theme and spacing").runs[0]
    rpr = run._r.get_or_add_rPr()
    element(rpr, "w:sz", val=20)
    element(rpr, "w:color", val="000000", themeColor="accent1", themeTint="CC")
    element(rpr, "w:spacing", val=20)
    element(rpr, "w:position", val=4)
    result, native = cycle(word, tmp_path)
    pr = native.sections[0].root_element.xpath("./hp:p/hp:run[hp:t/text()]/@charPrIDRef", namespaces=NS)[0]
    style = next(style for style in native.character_styles() if style.style_id == pr)
    assert style.text_color != "#000000"
    assert style.element.xpath("./hh:spacing/@latin", namespaces=NS) == ["10"]
    assert style.element.xpath("./hh:offset/@latin", namespaces=NS) == ["20"]
    assert result.paragraphs[0].runs[0]._r.xpath("./w:rPr/w:spacing/@w:val") == ["20"]


@pytest.mark.parametrize("level", ["row", "cell"])
def test_table_row_and_cell_content_controls_survive(tmp_path: Path, level):
    word = WordFile()
    table = word.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "controlled cell"
    target = table.rows[0]._tr if level == "row" else table.cell(0, 0)._tc
    sdt = OxmlElement("w:sdt")
    element(element(sdt, "w:sdtPr"), "w:tag", val="table-control")
    content = element(sdt, "w:sdtContent")
    target.addprevious(sdt)
    content.append(target)
    result, _ = cycle(word, tmp_path)
    tag = "tr" if level == "row" else "tc"
    assert result.element.xpath(f".//w:sdt/w:sdtContent/w:{tag}")
    assert "controlled cell" in result.element.xpath(".//w:sdt//w:t/text()")


@pytest.mark.parametrize("pattern", ["diagCross", "pct20", "thinHorzStripe"])
def test_patterned_cell_shading_has_native_hwpx_fill(tmp_path: Path, pattern):
    from jakal_hwpx.namespaces import NS

    word = WordFile()
    cell = word.add_table(rows=1, cols=1).cell(0, 0)
    cell.text = "Pattern"
    element(cell._tc.get_or_add_tcPr(), "w:shd", val=pattern, fill="FFEEDD", color="112233")
    result, native = cycle(word, tmp_path)
    identity = str(native.tables()[0].cell(0, 0).border_fill_id)
    definitions = native.header.root_element.xpath(".//hh:borderFill[@id=$id]", namespaces=NS, id=identity)
    brush = definitions[0].find("hc:fillBrush", NS)
    if pattern.startswith("pct"):
        assert brush.find("hc:imgBrush", NS).get("mode") == "TILE"
        reference = brush.find("hc:imgBrush/hc:img", NS).get("binaryItemIDRef")
        assert reference.startswith("word-pattern-")
    else:
        assert brush.find("hc:winBrush", NS).get("hatchStyle") == ("CROSS_DIAGONAL" if pattern == "diagCross" else "HORIZONTAL")
    assert result.tables[0].cell(0, 0)._tc.xpath("./w:tcPr/w:shd/@w:val") == [pattern]


def test_distinct_run_features_do_not_collapse_into_first_paragraph_style(tmp_path: Path):
    word = WordFile()
    word.add_paragraph("Title", style="Title")
    word.styles["Normal"].font.name = "Arial"
    for value in ("17", "31"):
        run = word.add_paragraph("Spaced " + value).runs[0]
        element(run._r.get_or_add_rPr(), "w:spacing", val=value)
    result, _ = cycle(word, tmp_path)
    assert result.paragraphs[1].runs[0]._r.xpath("./w:rPr/w:spacing/@w:val") == ["17"]
    assert result.paragraphs[2].runs[0]._r.xpath("./w:rPr/w:spacing/@w:val") == ["31"]
    assert result.paragraphs[0].runs[0].font.name != "Arial"
