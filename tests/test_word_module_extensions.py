"""Structured math/content editing, relation remapping, and safe refreshes."""
from __future__ import annotations

from io import BytesIO

from docx import Document as WordFile
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import pytest

from jakal_hwpx import Document, WordDocument, HwpxDocument
from jakal_hwpx.addons.docx import WordEndpoint
from jakal_hwpx.addons.docx.equations import hwp_to_omml, omml_to_hwp, capture_properties, restore_properties


def run(parent, text=None, *, control=None, instruction=None):
    node = OxmlElement("w:r")
    child = OxmlElement("w:fldChar" if control else "w:instrText" if instruction is not None else "w:t")
    if control:
        child.set(qn("w:fldCharType"), control)
    else:
        child.text = instruction if instruction is not None else text
    node.append(child)
    parent.append(node)
    return node


@pytest.mark.parametrize("script,expected", [
    ("cases {x & x>0 # -x & x<0}", "m:d"),
    ("pile {{a} over {b} # c}", "m:eqArr"),
    ("lpile {a # b}", "m:m"),
    ("rpile {a # b}", "m:m"),
    ("matrix {lcol {a # b} rcol {c # d}}", "m:m"),
    ("eqalign {x & =1 # y & =2}", "m:eqArr"),
    ("{n} choose {k}", "m:f"),
    ("binom {n} {k}", "m:f"),
    ("{x} atop {y}", "m:f"),
    ("pmatrix {a & b # c & d}", "m:d"),
    ("bmatrix {a & b}", "m:d"),
    ("dmatrix {a # b}", "m:d"),
])
def test_extended_math_native_roundtrip(tmp_path, script, expected):
    doc = WordDocument.blank()
    with doc.edit_word_details() as edit:
        edit.body.append_paragraph().add_equation(script)
    middle = doc.save(tmp_path / "middle.hwpx")
    HwpxDocument.open(middle).strict_validate()
    result = WordFile(Document.open(middle).save(tmp_path / "returned.docx"))
    assert result.element.xpath(".//" + expected)
    script2 = omml_to_hwp(result.element.xpath(".//m:oMath")[0])
    assert omml_to_hwp(hwp_to_omml(script2)) == script2
    if "atop" in script or "choose" in script or "binom" in script:
        assert result.element.xpath(".//m:fPr/m:type/@m:val") == ["noBar"]


@pytest.mark.parametrize("kind", ["limLow", "limUpp", "d"])
def test_display_substitutions_restore_current_operands_and_semantics(kind):
    math = OxmlElement("m:oMath")
    node = OxmlElement("m:" + kind)
    math.append(node)
    for name, text in ([("e", "base"), ("lim", "limit")] if kind != "d" else [("e", "one"), ("e", "two")]):
        child, r, t = OxmlElement("m:" + name), OxmlElement("m:r"), OxmlElement("m:t")
        t.text = text
        r.append(t)
        child.append(r)
        node.append(child)
    record = capture_properties(math)
    rebuilt = hwp_to_omml(omml_to_hwp(math).replace("base", "edited").replace("one", "updated"))
    restore_properties(rebuilt, record)
    assert rebuilt.find(qn("m:" + kind)) is not None
    assert "edited" in omml_to_hwp(rebuilt) if kind != "d" else "updated" in omml_to_hwp(rebuilt)
    if kind == "d":
        assert len(rebuilt[0].findall(qn("m:e"))) == 2
    assert "'base'" not in str(record) and "'limit'" not in str(record)


def complex_document():
    doc = WordDocument.blank()
    with doc.edit_word() as word:
        first = word.add_paragraph()._p
        run(first, "PREFIX ")
        run(first, control="begin")
        run(first, instruction='TOC \\o "1-3"')
        run(first, control="separate")
        run(first, "Heading one")
        second = word.add_paragraph()._p
        run(second, "Heading two ")
        inner = OxmlElement("w:fldSimple")
        inner.set(qn("w:instr"), "PAGE")
        run(inner, "2")
        second.append(inner)
        run(second, control="end")
        run(second, " SUFFIX")
    return doc


def test_complex_result_edits_preserve_outer_text_nested_fields_and_identity(tmp_path):
    doc = complex_document()
    endpoint = WordEndpoint(doc)
    state = endpoint.snapshot()
    identity = next(o["id"] for o in state["objects"] if o["kind"] == "fldChar")
    with doc.edit_word_details() as edit:
        field = edit.object(identity)
        assert field.result_fields[0].instruction == "PAGE"
        with field.edit_result() as content:
            assert content.paragraph(0).text == "Heading one"
            content.paragraph(0).format_text(bold=True)
            content.paragraph(0).replace_text("one", "ONE")
            content.fields[0].set_result("7")
            content.insert_table(1, rows=1, cols=1, values=[["Result table"]])
            content.append_paragraph("Appended")
        assert field.object_id == identity and field.instruction == 'TOC \\o "1-3"'
        assert field.result_fields[0].result == "7"
        with pytest.raises(RuntimeError, match="expired"):
            content.append_paragraph("Forbidden")
    native = WordFile(doc.save(tmp_path / "edited.docx"))
    assert native.element.xpath(".//w:instrText/text()") == ['TOC \\o "1-3"']
    texts = native.element.xpath(".//w:t/text()")
    assert texts[0] == "PREFIX " and texts[-1] == " SUFFIX"
    assert "Result table" in texts and "Appended" in texts and "7" in texts
    assert native.element.xpath(".//w:fldChar/@w:fldCharType") == ["begin", "separate", "end"]
    assert endpoint.get_object(identity)["object"]["kind"] == "fldChar"


def test_result_edit_noop_abort_and_stale_result_are_atomic():
    doc = complex_document()
    with doc.edit_word_details() as edit:
        field = edit.body.fields[0]
        before = doc._source.to_bytes(doc)
        with field.edit_result():
            pass
        assert doc._source.to_bytes(doc) == before
        with pytest.raises(ValueError, match="Abort"):
            with field.edit_result() as result:
                result.append_paragraph("Discarded")
                raise ValueError("Abort")
        assert doc._source.to_bytes(doc) == before
        with pytest.raises(RuntimeError, match="changed"):
            with field.edit_result() as result:
                result.append_paragraph("Stale")
                edit.body.paragraph(0).replace_text("PREFIX", "Changed")
        assert "Stale" not in field.result


def test_content_controls_structure_endpoint_and_hwpx_roundtrip(tmp_path):
    doc = WordDocument.blank()
    doc.append_paragraphs(["One", "Two", "Outside"])
    with doc.edit_word_details() as edit:
        control = edit.body.wrap_control(0, 2, tag="article", title="Article")
        control.insert_paragraph(1, "Middle")
        control.paragraph(0).delete()
        control.configure(tag="updated")
        other = edit.body.insert_control(2, text="Other", tag="second")
        assert control.text == "Middle\nTwo" and other.text == "Other"
        assert len(edit.body.content_controls) == 2
        assert control._node.find("./" + qn("w:sdtPr") + "/" + qn("w:id")).get(qn("w:val")) != other._node.find("./" + qn("w:sdtPr") + "/" + qn("w:id")).get(qn("w:val"))
    endpoint = WordEndpoint(doc)
    snapshot = endpoint.snapshot()
    controls = [o["id"] for o in snapshot["objects"] if o["kind"] == "sdt"]
    assert len(controls) == 2
    endpoint.apply(snapshot["version"], [{"target": controls[0], "method": "append_paragraph", "args": ["Inside"]}])
    restored = WordFile(Document.open(doc.save(tmp_path / "middle.hwpx")).save(tmp_path / "returned.docx"))
    assert len(restored.element.xpath(".//w:sdt")) == 2
    assert "Inside" in restored.element.xpath(".//w:sdtContent//w:t/text()")
    from jakal_hwpx.addons.docx.identity import ID
    assert set(controls) == {node.get(ID) for node in restored.element.xpath(".//w:sdt")}
    with doc.edit_word_details() as edit:
        edit.body.content_controls[0].unwrap()
        edit.body.content_controls[0].delete()
        assert not edit.body.content_controls
        assert "Middle" in [p.text for p in edit.body.blocks]


def test_cross_part_moves_remap_hyperlinks_and_keep_id(tmp_path):
    doc = WordDocument.blank()
    with doc.edit_word_details() as edit:
        paragraph = edit.body.append_paragraph("Body ")
        paragraph.insert_hyperlink(5, "https://example.com/moved", "Link")
        identity = paragraph.object_id
        paragraph.move(0, destination=edit.header())
        assert paragraph.object_id == identity
        assert edit.header().hyperlinks[0].target == "https://example.com/moved"
        paragraph.move(0, destination=edit.footer())
        assert edit.footer().hyperlinks[0].target == "https://example.com/moved"
        paragraph.move(0, destination=edit.body)
        assert edit.body.hyperlinks[0].target == "https://example.com/moved"
    result = WordFile(doc.save(tmp_path / "moved.docx"))
    assert result.paragraphs[0].text == "Body Link"
    assert result.part.rels[result.element.xpath(".//w:hyperlink/@r:id")[0]].target_ref == "https://example.com/moved"
    assert WordFile(BytesIO(doc._source.to_bytes(doc))).paragraphs[0].text == "Body Link"


def test_explicit_refresh_evaluates_nested_if_merge_properties_and_dates(tmp_path):
    doc = WordDocument.blank()
    with doc.edit_word() as word:
        word.core_properties.title = "Updated title"
        p = word.add_paragraph()._p
        run(p, control="begin")
        run(p, instruction="IF ")
        run(p, control="begin")
        run(p, instruction="MERGEFIELD status")
        run(p, control="separate")
        run(p, "stale")
        run(p, control="end")
        run(p, instruction=' = "approved" "YES" "NO"')
        run(p, control="separate")
        run(p, "NO")
        run(p, control="end")
    with doc.edit_word_details() as edit:
        edit.body.append_paragraph().add_field("TITLE", "Old title")
        edit.body.append_paragraph().add_field('DATE \\@ "yyyy.MM.dd"', "Old date")
        locked = edit.body.append_paragraph().add_field('IF 2 > 1 "yes" "no"', "Locked")
        locked.lock()
        edit.header().append_paragraph().add_field('TIME \\@ "HH:mm:ss"', "Old time")
        edit.footer().append_paragraph().add_field("NUMPAGES", "8")
    report = doc.update_fields(data={"status": "approved"}, now="2026-11-04T13:14:15")
    assert sum(row["status"] == "updated" for row in report) == 5
    assert sum(row["status"] == "locked" for row in report) == 1
    assert sum(row["status"] == "unsupported" for row in report) == 1
    result = WordFile(doc.save(tmp_path / "refreshed.docx"))
    assert result.element.xpath(".//w:t/text()") == ["approved", "YES", "Updated title", "2026.11.04", "Locked"]
    assert result.sections[0].header.paragraphs[-1]._p.xpath(".//w:t/text()") == ["13:14:15"]
    assert result.sections[0].footer.paragraphs[-1]._p.xpath(".//w:t/text()") == ["8"]
    before = doc._source.to_bytes(doc)
    with pytest.raises(ValueError, match="layout engine"):
        doc.update_fields(data={"status": "rejected"}, now="2026-11-05T00:00:00", strict=True)
    assert doc._source.to_bytes(doc) == before


@pytest.mark.parametrize("code,expected", [
    ('IF "ABcd" = "ab*" "matched" "missed"', "missed"),
    ('IF "ABcd" = "AB*" "matched" "missed"', "matched"),
    ('IF "A" = "a" "same" "different"', "different"),
    ('IF 2 < 10 "numeric" "lexical"', "numeric"),
    ('IF "a" <> "b" "different" "same"', "different"),
])
def test_if_comparison_matches_word_numeric_case_and_wildcard_rules(code, expected):
    doc = WordDocument.blank()
    with doc.edit_word_details() as edit:
        edit.body.append_paragraph().add_field(code, "old")
        assert edit.update_fields(strict=True)[0]["status"] == "updated"
        assert edit.body.fields[0].result == expected


def test_rebuild_toc_uses_current_headings_and_reports_missing_pagination(tmp_path):
    doc = complex_document()
    with doc.edit_word_details() as edit:
        edit.body.append_paragraph("First heading", style="Heading 1")
        edit.body.append_paragraph("Second heading", style="Heading 2")
        edit.body.append_paragraph("Excluded heading", style="Heading 4")
        field = edit.body.fields[0]
        report = field.rebuild_toc()
        assert report == {"entries": 2, "requires_layout": True}
        assert "First heading" in field.result and "Second heading" in field.result
        assert "Excluded heading" not in field.result
        assert len(edit.body.hyperlinks) == 2
        assert len(edit.bookmarks) == 2
        report = field.rebuild_toc()
        assert report["entries"] == 2 and len(edit.bookmarks) == 2
        with pytest.raises(ValueError, match="pagination"):
            field.rebuild_toc(include_page_numbers=True)
    native = WordFile(doc.save(tmp_path / "toc.docx"))
    assert native.element.xpath(".//w:hyperlink/@w:anchor")
    assert native.element.xpath(".//w:fldChar/@w:fldCharType") == ["begin", "separate", "end"]
    assert native.element.xpath(".//w:t/text()")[0] == "PREFIX "
    assert " SUFFIX" in native.element.xpath(".//w:t/text()")


def test_mixed_script_layout_switches_map_in_both_directions(tmp_path):
    doc = WordDocument.blank()
    with doc.edit_word_details() as edit:
        edit.body.append_paragraph("한글English123 줄바꿈 wrapping").format(
            word_wrap=True, auto_space_east_asian_latin=True, auto_space_east_asian_numbers=False,
            snap_to_grid=False, suppress_line_numbers=True,
        )
    style = doc.paragraph_styles[0]
    assert style.auto_spacing == {"eAsianEng": "1", "eAsianNum": "0"}
    assert style.break_setting["breakLatinWord"] == "BREAK_WORD"
    middle = doc.save(tmp_path / "layout.hwpx")
    HwpxDocument.open(middle).strict_validate()
    common = Document.open(middle)
    style = next(style for style in common.paragraph_styles if style.auto_spacing.get("eAsianEng") == "1")
    style.auto_spacing.update(eAsianEng="0", eAsianNum="1")
    style.break_setting["breakLatinWord"] = "KEEP_WORD"
    result = WordFile(common.save(tmp_path / "layout.docx"))
    pr = result.paragraphs[0]._p.pPr
    assert pr.find(qn("w:autoSpaceDE")).get(qn("w:val")) == "0"
    assert pr.find(qn("w:autoSpaceDN")).get(qn("w:val")) == "1"
    assert pr.find(qn("w:wordWrap")).get(qn("w:val")) == "0"
    assert pr.find(qn("w:snapToGrid")).get(qn("w:val")) == "0"
    assert pr.find(qn("w:suppressLineNumbers")).get(qn("w:val")) == "1"


def test_created_complex_field_and_control_move_preserve_relationships(tmp_path):
    doc = WordDocument.blank()
    with doc.edit_word_details() as edit:
        field = edit.body.append_paragraph("Before ").add_field('IF 1 = 1 "yes" "no"', "yes", complex_field=True)
        with field.edit_result() as result:
            result.append_paragraph("More")
        assert field.result == "yes\nMore"
        control = edit.body.insert_control(len(edit.body.blocks), text="Linked ", tag="moved")
        control.paragraph(0).insert_hyperlink(7, "https://example.com/control", "control")
        identity = control.object_id
        control.move(0, destination=edit.header())
        assert control.object_id == identity and control.hyperlinks[0].target == "https://example.com/control"
        control.paragraph(0).replace_text("Linked", "Edited")
        control.move(len(edit.body.blocks), destination=edit.body)
        assert control.text == "Edited control"
    native = WordFile(doc.save(tmp_path / "complex.docx"))
    assert native.element.xpath(".//w:sdtContent//w:t/text()") == ["Edited ", "control"]


def test_rich_field_result_survives_hwpx_roundtrip(tmp_path):
    doc = complex_document()
    with doc.edit_word_details() as edit:
        field = edit.body.fields[0]
        with field.edit_result() as result:
            result.paragraph(0).replace_text("one", "updated")
            result.insert_table(1, rows=1, cols=1, values=[["Nested table"]])
            result.fields[0].set_result("9")
    returned = WordFile(Document.open(doc.save(tmp_path / "rich.hwpx")).save(tmp_path / "returned.docx"))
    assert "Nested table" in returned.element.xpath(".//w:t/text()")
    assert returned.element.xpath(".//w:instrText/text()") == ['TOC \\o "1-3"']
    assert "9" in returned.element.xpath(".//w:t/text()")


def test_implicit_word_layout_defaults_stay_inherited(tmp_path):
    doc = WordDocument.blank()
    doc.append_paragraph("Keep implicit layout")
    result = WordFile(Document.open(doc.save(tmp_path / "implicit.hwpx")).save(tmp_path / "implicit.docx"))
    for root in (result.paragraphs[0]._p, result.styles["Normal"].element):
        for name in ("snapToGrid", "wordWrap", "autoSpaceDE", "autoSpaceDN", "suppressLineNumbers"):
            assert root.find("./" + qn("w:pPr") + "/" + qn("w:" + name)) is None


def test_standalone_math_uses_word_math_alignment_in_hwpx(tmp_path):
    doc = WordDocument.blank()
    with doc.edit_word_details() as edit:
        edit.body.append_paragraph().add_equation("cases {x # y}")
        edit.body.append_paragraph("Inline ").add_equation("x")
    paragraphs = [block for block in doc.sections[0].blocks if type(block).__name__ == "Paragraph"]
    styles = {style.style_id: style for style in doc.paragraph_styles}
    assert styles[paragraphs[0].para_pr_id].alignment_horizontal == "CENTER"
    assert styles[paragraphs[1].para_pr_id].alignment_horizontal == "LEFT"
    HwpxDocument.open(doc.save(tmp_path / "math-alignment.hwpx")).strict_validate()


def test_empty_header_keeps_private_anchor_and_paragraph_identity(tmp_path):
    from jakal_hwpx.addons.docx.identity import internal_name, ID
    doc = WordDocument.blank()
    with doc.edit_word_details() as edit:
        edit.body.append_paragraph("Body")
        header = edit.header().paragraph(0)
        identity = header.object_id
    returned = WordFile(Document.open(doc.save(tmp_path / "empty-header.hwpx")).save(tmp_path / "empty-header.docx"))
    paragraph = returned.sections[0].header.paragraphs[0]._p
    assert paragraph.get(ID) == identity
    assert not list(paragraph.iter(qn("w:r")))
    assert returned.styles["Header"].element.rPr is None
    assert any(internal_name(node.get(qn("w:name"))) for node in paragraph.iter(qn("w:bookmarkStart")))


def test_same_header_for_both_pages_shares_one_part_without_duplicate_ids(tmp_path):
    from jakal_hwpx.addons.docx.identity import ID
    doc = WordDocument.blank()
    with doc.edit_word_details() as edit:
        edit.body.append_paragraph("Body")
        edit.header().paragraph(0).set_text("Shared header")
        edit.header().paragraph(0).object_id
    common = Document.open(doc.save(tmp_path / "shared.hwpx"))
    common.sections[0].header_footer_blocks[0].apply_page_type = "BOTH"
    # Activate odd/even globally through another distinct story.
    from jakal_hwpx._hancom_models import HeaderFooter
    common.sections[0].header_footer_blocks.append(HeaderFooter("footer", "Even footer", apply_page_type="EVEN"))
    result = WordFile(common.save(tmp_path / "shared.docx"))
    assert result.sections[0].header.part is result.sections[0].even_page_header.part
    ids = [node.get(ID) for part in result.part.package.parts if getattr(part, "element", None) is not None for node in part.element.iter() if node.get(ID)]
    assert len(ids) == len(set(ids))


def test_modified_common_named_style_preserves_existing_direct_character_format(tmp_path):
    doc = WordDocument.blank()
    with doc.edit_word_details() as edit:
        edit.format_style("Normal", font={"font_name": "Arial", "font_size_pt": 10})
        edit.body.append_paragraph("Style edit").format_text(font_size_pt=9)
    common = Document.open(doc.save(tmp_path / "style.hwpx"))
    paragraph = next(block for block in common.sections[0].blocks if getattr(block, "text", None) == "Style edit")
    definition = next(style for style in common.style_definitions if style.style_id == paragraph.style_id)
    character = next(style for style in common.character_styles if style.style_id == definition.char_pr_id)
    character.height = 1700
    native = WordFile(common.save(tmp_path / "style.docx"))
    assert native.styles["Normal"].font.size.pt == 17
    # HWPX runs have an independent charPr; changing the named default does
    # not remove the existing run's explicit character formatting.
    assert native.paragraphs[0].runs[-1].font.size.pt == 9


def test_private_identity_bookmarks_never_become_visible_field_owners(tmp_path):
    from jakal_hwpx._hancom_models import Field, Paragraph
    from jakal_hwpx.addons.docx.identity import internal_name
    doc = WordDocument.blank()
    with doc.edit_word_details() as edit:
        p = edit.body.append_paragraph("Editable literal")
        p.object_id
    assert any(isinstance(block, Paragraph) and block.text == "Editable literal" for block in doc.sections[0].blocks)
    assert not any(isinstance(block, Field) and internal_name(block.name) for block in doc.sections[0].blocks)
    doc.replace_text("literal", "text")
    restored = Document.open(doc.save(tmp_path / "private.hwpx"))
    assert any(isinstance(block, Paragraph) and block.text == "Editable text" for block in restored.sections[0].blocks)
    assert not any(isinstance(block, Field) and internal_name(block.name) for block in restored.sections[0].blocks)
