"""Word commands become live Hancom controls without losing their Word identity."""
from __future__ import annotations

from copy import deepcopy
from pathlib import Path

from docx import Document as WordFile
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import pytest

from jakal_hwpx import Document, HwpxDocument
from jakal_hwpx.namespaces import NS, qname


def field(parent, instruction, value):
    result = OxmlElement("w:fldSimple")
    result.set(qn("w:instr"), instruction)
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = value
    run.append(text)
    result.append(run)
    parent.append(result)


@pytest.mark.parametrize(("instruction", "kind", "parameter", "expected"), [
    ('DATE \\@ "yyyy-MM-dd"', "DATE", "Command", r"\:1-02-03;0;"),
    ('DATE \\@ "yyyy/MM/dd"', "DATE", "DateFormat", "YYYY/MM/DD"),
    ('DOCPROPERTY "Title"', "SUMMERY", "Property", "$title"),
    ("AUTHOR", "SUMMERY", "Property", "$author"),
    ("REF target \\h", "CROSSREF", "RefContentType", "OBJECT_TYPE_CONTENTS"),
    ("PAGEREF target \\h", "CROSSREF", "RefContentType", "OBJECT_TYPE_PAGE"),
    ("REF target \\p", "CROSSREF", "RefContentType", "OBJECT_TYPE_UPDOWNPOS"),
    ("MERGEFIELD customer", "MAILMERGE", "FieldValue", "customer"),
    ('= SUM(LEFT) \\# "#,##0.00"', "FORMULA", "Command", "=SUM(LEFT)??%.2f,;;0"),
])
def test_native_field_types_and_commands_survive_hancom_parameter_normalization(
    tmp_path: Path, instruction, kind, parameter, expected,
):
    word = WordFile()
    if kind == "CROSSREF":
        target = word.add_paragraph()._p
        start = OxmlElement("w:bookmarkStart")
        start.set(qn("w:id"), "1")
        start.set(qn("w:name"), "target")
        target.append(start)
        run = OxmlElement("w:r")
        text = OxmlElement("w:t")
        text.text = "Target text"
        run.append(text)
        target.append(run)
        end = OxmlElement("w:bookmarkEnd")
        end.set(qn("w:id"), "1")
        target.append(end)
    paragraph = word.add_paragraph()
    field(paragraph._p, instruction, "0")
    word.save(tmp_path / "source.docx")
    for index in range(3):
        native_path = Document.open(tmp_path / "source.docx").save(tmp_path / "middle.hwpx")
        native = HwpxDocument.open(native_path)
        native.strict_validate()
        control = next(item for item in native.fields() if item.field_type == kind)
        assert control.parameter_map()[parameter] == expected
        # Hancom rebuilds its parameter set and strips addon-specific values.
        # The encoded identity in the field name must still restore Word syntax.
        parameters = control.element.find(qname("hp", "parameters"))
        if parameters is not None:
            control.element.remove(parameters)
        control.section.mark_modified()
        native.save(native_path)
        result = Document.open(native_path).save(tmp_path / "source.docx")
        word = WordFile(result)
        assert word.element.xpath(".//w:fldSimple/@w:instr") == [instruction]
        assert word.element.xpath(".//w:fldSimple//w:t/text()") == ["0"]


def test_native_page_and_total_page_fields_preserve_word_switches(tmp_path: Path):
    word = WordFile()
    commands = [r"PAGE \* roman", r"NUMPAGES \* Arabic"]
    for instruction in commands:
        field(word.add_paragraph()._p, instruction, "1")
    word.save(tmp_path / "source.docx")
    for _ in range(3):
        middle = Document.open(tmp_path / "source.docx").save(tmp_path / "middle.hwpx")
        native = HwpxDocument.open(middle)
        native.strict_validate()
        controls = native.sections[0].root_element.xpath(".//hp:autoNum", namespaces=NS)
        assert [control.get("numType") for control in controls] == ["PAGE", "TOTAL_PAGE"]
        assert controls[0].find(qname("hp", "autoNumFormat")).get("type") == "ROMAN_SMALL"
        result = Document.open(middle).save(tmp_path / "source.docx")
        assert WordFile(result).element.xpath(".//w:fldSimple/@w:instr") == commands


def test_implicit_word_time_format_is_not_replaced_with_a_calendar_date(tmp_path: Path):
    word = WordFile()
    field(word.add_paragraph()._p, "TIME", "11:42 AM")
    word.save(tmp_path / "source.docx")
    with pytest.warns(UserWarning, match="does not recalculate"):
        middle = Document.open(tmp_path / "source.docx").save(tmp_path / "middle.hwpx")
    native = HwpxDocument.open(middle)
    assert not any(control.field_type == "DATE" for control in native.fields())
    result = WordFile(Document.open(middle).save(tmp_path / "result.docx"))
    assert result.element.xpath(".//w:fldSimple/@w:instr") == ["TIME"]
    assert result.element.xpath(".//w:fldSimple//w:t/text()") == ["11:42 AM"]


def test_new_hancom_list_paragraph_inherits_active_word_counter(tmp_path: Path):
    word = WordFile()
    word.add_paragraph("One", style="List Number")
    word.add_paragraph("Two", style="List Number")
    word.save(tmp_path / "source.docx")
    path = Document.open(tmp_path / "source.docx").save(tmp_path / "middle.hwpx")
    native = HwpxDocument.open(path)
    root = native.sections[0].root_element
    previous = root.findall(qname("hp", "p"))[-1]
    # Enter in Hancom follows nextStyleIDRef. New content has no Word carrier.
    styles = {item.get("id"): item for item in native.header.root_element.findall(".//hh:style", NS)}
    original = styles[previous.get("styleIDRef")]
    following = styles[original.get("nextStyleIDRef")]
    assert following.get("paraPrIDRef") == previous.get("paraPrIDRef")
    added = deepcopy(previous)
    added.clear()
    added.attrib.update(previous.attrib)
    added.set("styleIDRef", following.get("id"))
    added.set("paraPrIDRef", following.get("paraPrIDRef"))
    run = added.makeelement(qname("hp", "run"), charPrIDRef="0")
    text = run.makeelement(qname("hp", "t"))
    text.text = "Three inserted in Hancom"
    run.append(text)
    added.append(run)
    root.append(added)
    native.sections[0].mark_modified()
    native.save(path)
    result = WordFile(Document.open(path).save(tmp_path / "result.docx"))
    assert [item.text for item in result.paragraphs] == ["One", "Two", "Three inserted in Hancom"]
    identifiers = [item._p.pPr.numPr.numId.val for item in result.paragraphs]
    assert identifiers[0] == identifiers[1] == identifiers[2]


def test_hancom_shared_fieldids_keep_distinct_begin_end_pairs():
    native = HwpxDocument.blank()
    native.append_field(field_type="CLICK_HERE", display_text="First")
    native.append_field(field_type="CLICK_HERE", display_text="Second")
    root = native.sections[0].root_element
    for control in root.xpath(".//hp:fieldBegin | .//hp:fieldEnd", namespaces=NS):
        control.set("fieldid", "627272811")
    assert not native.strict_lint_errors()
    root.xpath(".//hp:fieldEnd", namespaces=NS)[-1].set("beginIDRef", "missing")
    assert any("unknown fieldBegin" in str(issue) for issue in native.strict_lint_errors())


def test_mixed_numbering_and_distinct_bullet_levels_use_separate_native_resources(tmp_path: Path):
    word = WordFile()
    numbering = word.part.numbering_part.element
    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), "25")
    for index, (kind, label) in enumerate([("decimal", "%1."), ("bullet", "•"), ("bullet", "▪")]):
        level = OxmlElement("w:lvl")
        level.set(qn("w:ilvl"), str(index))
        for tag, value in [("start", "1"), ("numFmt", kind), ("lvlText", label)]:
            node = OxmlElement("w:" + tag)
            node.set(qn("w:val"), value)
            level.append(node)
        abstract.append(level)
    numbering.insert(0, abstract)
    number = numbering.add_num(25)
    for index in range(3):
        paragraph = word.add_paragraph(f"Level {index}")
        properties = paragraph._p.get_or_add_pPr().get_or_add_numPr()
        properties.get_or_add_numId().val = number.numId
        properties.get_or_add_ilvl().val = index
    word.save(tmp_path / "source.docx")
    path = Document.open(tmp_path / "source.docx").save(tmp_path / "middle.hwpx")
    native = HwpxDocument.open(path)
    native.strict_validate()
    heads = native.header.root_element.xpath(".//hh:paraPr/hh:heading[@type!='NONE']", namespaces=NS)
    assert {head.get("type") for head in heads} == {"NUMBER", "BULLET"}
    chars = native.header.root_element.xpath(".//hh:bullet/@char", namespaces=NS)
    assert "•" in chars and "▪" in chars
    result = WordFile(Document.open(path).save(tmp_path / "result.docx"))
    assert [p._p.pPr.numPr.ilvl.val for p in result.paragraphs] == [0, 1, 2]
    assert len({p._p.pPr.numPr.numId.val for p in result.paragraphs}) == 1
