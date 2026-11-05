"""Content and geometry regressions isolated from native shapes/numbers."""
from zipfile import ZipFile
from pathlib import Path
import runpy

from docx import Document as WordFile
import pytest

from jakal_hwpx import Document, HwpxDocument
from jakal_hwpx.addons.docx.equations import M, hwp_to_omml, omml_to_hwp
from jakal_hwpx.namespaces import NS, qname


@pytest.mark.parametrize("script", [
    "v left(t right)=3t^{2}-10t+7",
    "18+15 sqrt{3}",
    "bar {{rm{AB}} it } < bar {{rm{AC}} it }`",
    "{rm{P}} it left(3,`` f left(3 right) right)`",
    "a cos `b x+3`",
    "sin A : sin C =3:5",
    "t ~ LEFT ( t GEQ 0 RIGHT )`",
    "sum_{k=1}^{n} {a_{k}} over {b_{k+1}}={1} over {2} n^{2}",
    "sum_{k=1}^{5} a_{k}`",
    "sqrt {3} of {5} times 25^{{1} over {3}}`",
    "lim _{n ``rarrow`` inf} left(sqrt{a_{n} it ^{2}+n}-a_{n} right)`",
    "n alpha_{1}-a b={q} over {p} pi`",
    "m_{1}+ sigma_{2}`",
    "angle {rm{ABC}} =it{pi} over {2}`",
])
def test_math_commands_have_stable_native_structure(script):
    original = hwp_to_omml(script)
    canonical = omml_to_hwp(original)
    assert omml_to_hwp(hwp_to_omml(canonical)) == canonical
    assert "vleft" not in canonical and "15sqrt" not in canonical


def test_native_accents_symbols_fonts_and_spaces_are_real_word_math():
    math = hwp_to_omml("bar {{rm AB} it} + hat {x} + sqrt{3} + a times b geq pi ~ `")
    ns = {"m": M}
    assert math.xpath("./m:bar/m:barPr/m:pos/@m:val", namespaces=ns) == ["top"]
    assert math.xpath("./m:bar/m:e/m:r/m:rPr/m:sty/@m:val", namespaces=ns) == ["p"]
    assert math.xpath("./m:acc/m:accPr/m:chr/@m:val", namespaces=ns) == ["̂"]
    assert math.xpath("./m:rad", namespaces=ns)
    text = "".join(math.itertext("{" + M + "}t"))
    assert "×" in text and "≥" in text and "π" in text
    assert "\u00a0" in text and "\u2005" in text
    assert not any(command in text for command in ["bar", "hat", "sqrt", "times", "geq", "rm", "~", "`"])


@pytest.mark.parametrize("body,path", [("a_{k}", "m:sSub"), ("{a_{k}} over {b_{k+1}}", "m:f")])
def test_sum_scripts_and_fractions_belong_to_the_operand(body, path):
    math = hwp_to_omml("sum_{k=1}^{n} " + body)
    ns = {"m": M}
    assert math.xpath("./m:nary/m:e/" + path, namespaces=ns)
    assert not math.xpath("./m:f | ./m:sSub | ./m:sSup", namespaces=ns)


def equation_fixture(tmp_path):
    source = HwpxDocument.blank()
    source.append_character_style(style_id="1", height=1150)
    source.append_paragraph("Before formula", char_pr_id="1")
    equation = source.append_equation("{1234567} over {y}", paragraph_index=1, width=6712, height=2345)
    equation.element.set("baseLine", "85")
    equation.element.getparent().set("charPrIDRef", "1")
    source.sections[0].mark_modified()
    return source.save(tmp_path / "source.hwpx")


@pytest.mark.parametrize("via_word", [False, True])
def test_equation_baseline_and_host_font_survive(tmp_path, via_word):
    model = Document.open(equation_fixture(tmp_path))
    if via_word:
        middle = model.save(tmp_path / "middle.docx")
        with ZipFile(middle) as package:
            assert b"1234567" not in package.read("customXml/jakal-hwpx-properties.xml")
        model = Document.open(middle)
    result = HwpxDocument.open(model.save(tmp_path / "returned.hwpx"))
    node = result.equations()[0].element
    assert node.get("baseLine") == "85"
    char = result.header.root_element.xpath(".//hh:charPr[@id=$id]", namespaces=NS, id=node.getparent().get("charPrIDRef"))[0]
    assert char.get("height") == "1150"


def test_word_operand_edit_does_not_restore_old_equation_geometry(tmp_path):
    source = equation_fixture(tmp_path)
    word = WordFile(Document.open(source).save(tmp_path / "middle.docx"))
    for text in word.element.iter("{" + M + "}t"):
        if text.text == "1234567":
            text.text = "9"
    word.save(tmp_path / "edited.docx")
    native = HwpxDocument.open(Document.open(tmp_path / "edited.docx").save(tmp_path / "edited.hwpx"))
    node = native.equations()[0].element
    assert "1234567" not in native.equations()[0].script
    assert node.get("baseLine") != "85"
    assert node.find(qname("hp", "sz")).get("width") != "6712"


def test_root_alias_and_font_directive_keep_expression_binding():
    ns = {"m": M}
    math = hwp_to_omml("sqrt {3} of {5}")
    assert math.xpath("./m:rad/m:deg/m:r/m:t/text()", namespaces=ns) == ["3"]
    assert math.xpath("./m:rad/m:e/m:r/m:t/text()", namespaces=ns) == ["5"]
    math = hwp_to_omml("a_{n} it ^{2}")
    assert math.xpath("./m:sSup/m:e/m:sSub/m:sub/m:r/m:t/text()", namespaces=ns) == ["n"]
    assert math.xpath("./m:sSup/m:sup/m:r/m:t/text()", namespaces=ns) == ["2"]
    assert omml_to_hwp(hwp_to_omml(omml_to_hwp(math))) == omml_to_hwp(math)
    math = hwp_to_omml("n alpha_{1}")
    assert math.xpath("./m:r/m:t/text()", namespaces=ns) == ["n"]
    assert math.xpath("./m:sSub/m:e/m:r/m:t/text()", namespaces=ns) == ["α"]


def test_word_script_structure_edit_invalidates_old_geometry(tmp_path):
    source = HwpxDocument.open(equation_fixture(tmp_path))
    source.equations()[0].script = "x_{y}"
    source.save(tmp_path / "source.hwpx")
    word = WordFile(Document.open(tmp_path / "source.hwpx").save(tmp_path / "middle.docx"))
    subscript = next(word.element.iter("{" + M + "}sSub"))
    subscript.tag = "{" + M + "}sSup"
    subscript.find("{" + M + "}sub").tag = "{" + M + "}sup"
    word.save(tmp_path / "edited.docx")
    native = HwpxDocument.open(Document.open(tmp_path / "edited.docx").save(tmp_path / "edited.hwpx"))
    equation = native.equations()[0]
    assert "^" in equation.script and "_" not in equation.script
    assert equation.element.get("baseLine") != "85"
    assert equation.element.find(qname("hp", "sz")).get("width") != "6712"


def test_single_column_eqalign_stays_aligned_inside_cases():
    script = "cases {eqalign {5x+a # } & & eqalign {~left(x<-2 right) # } # x^{2}-a & & x geq -2}"
    canonical = omml_to_hwp(hwp_to_omml(script))
    assert canonical.count("eqalign {") == 2
    assert "pile {" not in canonical
    assert canonical.count(" & ") == 4
    assert omml_to_hwp(hwp_to_omml(canonical)) == canonical


def test_default_italic_does_not_add_spacing_switches_between_symbols():
    canonical = omml_to_hwp(hwp_to_omml("{rm P} it left(15 leq X leq 20 right)"))
    assert canonical.count("it ") == 1
    assert "leq" in canonical
    assert "{rm {P}}" in canonical


def test_inline_table_followed_by_math_retains_its_paragraph_position(tmp_path):
    source = HwpxDocument.blank()
    source.append_table(1, 1, paragraph_index=0, cell_texts=[["TABLE"]])
    source.append_equation("x+17", paragraph_index=0)
    paragraph = source.sections[0].root_element.find(qname("hp", "p"))
    from lxml import etree
    etree.SubElement(etree.SubElement(paragraph, qname("hp", "run"), charPrIDRef="0"), qname("hp", "t")).text = "TAIL"
    source.append_paragraph("AFTER")
    source.sections[0].mark_modified()
    initial = source.save(tmp_path / "source.hwpx")
    middle = Document.open(initial).save(tmp_path / "middle.docx")
    returned = Document.open(middle).save(tmp_path / "returned.hwpx")
    tools = runpy.run_path(str(Path(__file__).resolve().parents[1] / "scripts/validate_hwpx_math_body.py"))
    assert tools["compare"](initial, returned)["structure_equal"]


def test_math_body_reference_excludes_shapes_and_numbers_without_editing_original(tmp_path):
    source = HwpxDocument.blank()
    source.set_paragraph_text(0, 0, "BODY")
    source.append_equation("x+1", paragraph_index=0)
    source.append_shape(kind="rect", text="EXCLUDED_SHAPE", paragraph_index=0)
    source.append_auto_number(kind="autoNum", number=1, number_type="PAGE", paragraph_index=0)
    source.append_header("EXCLUDED_HEADER", paragraph_index=0)
    path = source.save(tmp_path / "original.hwpx")
    original = path.read_bytes()
    tools = runpy.run_path(str(Path(__file__).resolve().parents[1] / "scripts/validate_hwpx_math_body.py"))
    removed = tools["prepare_reference"](path, tmp_path / "body.hwpx")
    assert removed["rect"] == removed["autoNum"] == removed["header"] == 1
    assert path.read_bytes() == original
    text = HwpxDocument.open(tmp_path / "body.hwpx").get_document_text()
    assert "BODY" in text and "EXCLUDED" not in text
    assert len(HwpxDocument.open(tmp_path / "body.hwpx").equations()) == 1


@pytest.mark.parametrize("via_word", [False, True])
def test_fixed_and_nonbreaking_spaces_survive_native_text_and_edits(tmp_path, via_word):
    from lxml import etree
    source = HwpxDocument.blank()
    source.set_paragraph_text(0, 0, "A")
    text = source.sections[0].root_element.xpath("./hp:p/hp:run/hp:t", namespaces=NS)[0]
    etree.SubElement(text, qname("hp", "fwSpace")).tail = "B"
    etree.SubElement(text, qname("hp", "nbSpace")).tail = "C"
    source.sections[0].mark_modified()
    initial = source.save(tmp_path / "source.hwpx")
    model = Document.open(initial)
    assert model.sections[0].blocks[0].text == "A\u2007B\u00a0C"
    if via_word:
        model = Document.open(model.save(tmp_path / "middle.docx"))
    model.replace_text("B", "EDIT")
    restored = HwpxDocument.open(model.save(tmp_path / "returned.hwpx"))
    assert restored.get_document_text() == "A\u2007EDIT\u00a0C"
    assert len(restored.sections[0].root_element.xpath(".//hp:t/hp:fwSpace | .//hp:t/hp:nbSpace", namespaces=NS)) == 2


def test_replacing_text_across_native_space_controls_keeps_their_types():
    from lxml import etree
    from jakal_hwpx._hwpx_element_helpers import _replace_text, _text_node_plain_text
    paragraph = etree.Element(qname("hp", "p"))
    text = etree.SubElement(etree.SubElement(paragraph, qname("hp", "run")), qname("hp", "t"))
    text.text = "A"
    etree.SubElement(text, qname("hp", "fwSpace")).tail = "B"
    etree.SubElement(text, qname("hp", "nbSpace")).tail = "C"
    assert _replace_text(paragraph, "B\u00a0C", "D\u00a0E") == 1
    assert _text_node_plain_text(text) == "A\u2007D\u00a0E"
    assert [etree.QName(child).localname for child in text] == ["fwSpace", "nbSpace"]
