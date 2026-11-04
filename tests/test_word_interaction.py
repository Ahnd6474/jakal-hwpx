"""Persistent interaction, rich ranges and structured math regression tests."""
from __future__ import annotations

import http.client
import json
from threading import Thread

from docx import Document as WordFile
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import pytest

from jakal_hwpx import Document, WordDocument, HwpxDocument
from jakal_hwpx.addons.docx import WordEndpoint, VersionConflict, ObjectNotFound, make_http_server
from jakal_hwpx.addons.docx.equations import hwp_to_omml, omml_to_hwp


def test_ranges_preserve_run_styles_and_replace_across_paragraphs(tmp_path):
    doc = WordDocument.blank()
    doc.append_paragraphs(["alpha beta", "gamma delta", "destination"])
    with doc.edit_word_details() as edit:
        edit.body.paragraph(0).format_text(0, 5, bold=True)
        edit.body.range(0, 6, 1, 5).replace("new\nsecond")
        assert [p.text for p in edit.body.blocks] == ["alpha new", "second delta", "destination"]
        edit.body.paragraph(0).range(6, 9).copy_to(edit.body.paragraph(2), 0)
        assert edit.body.paragraph(2).text == "newdestination"
        edit.body.paragraph(1).range(0, 6).move_to(edit.body.paragraph(2), 3)
        assert edit.body.paragraph(1).text == " delta"
        assert edit.body.paragraph(2).text == "newseconddestination"
    word = WordFile(doc.save(tmp_path / "ranges.docx"))
    assert word.paragraphs[0].runs[0].bold
    assert word.paragraphs[0].runs[0].text == "alpha"


def test_search_replace_and_reject_table_or_stale_range_boundaries():
    doc = WordDocument.blank()
    doc.append_paragraphs(["cat cat", "next cat"])
    with doc.edit_word_details() as edit:
        assert len(edit.body.find_text(r"cat\s+next", regex=True)) == 1
        assert edit.body.replace_all("cat", "DOG") == 3
        old = edit.body.paragraph(0).range(0, 3)
        edit.body.paragraph(0).replace_text("DOG", "new", count=1)
        with pytest.raises(RuntimeError, match="stale"):
            old.delete()
        edit.body.insert_table(1, rows=1, cols=1)
        with pytest.raises(ValueError, match="table"):
            edit.body.range(0, 0, 1, 2)


def test_copy_and_move_within_one_paragraph_keep_the_native_paragraph_id():
    doc = WordDocument.blank()
    doc.append_paragraph("abcdef")
    endpoint = WordEndpoint(doc)
    with doc.edit_word_details() as edit:
        p = edit.body.paragraph(0)
        identity = p.object_id
        p.range(1, 3).move_to(p, 6)
        assert p.text == "adefbc"
        p.range(4, 6).copy_to(p, 0)
        assert p.text == "bcadefbc"
        assert p.object_id == identity
    assert endpoint.get_object(identity)["object"]["text"] == "bcadefbc"


def test_bookmarks_rename_updates_links_and_segmented_ref_commands(tmp_path):
    doc = WordDocument.blank()
    doc.append_paragraphs(["Bookmark target", "Jump here", "Copy"])
    with doc.edit_word_details() as edit:
        edit.body.paragraph(0).range(0, 8).add_bookmark("target")
        edit.body.paragraph(1).range(0, 4).add_hyperlink("#target")
        edit.body.paragraph(1).add_field('REF "target" \\h', "Bookmark")
        edit.bookmarks[0].rename("renamed")
        assert edit.body.hyperlinks[0].target == "#renamed"
        assert edit.body.fields[0].instruction == "REF renamed \\h"
        edit.body.paragraph(0).range(0, 8).copy_to(edit.body.paragraph(2), 0)
        assert {b.name for b in edit.bookmarks} == {"renamed", "renamed_copy1"}
    word = WordFile(doc.save(tmp_path / "bookmarks.docx"))
    assert len(set(word.element.xpath(".//w:bookmarkStart/@w:id"))) == 2
    native = doc.save(tmp_path / "middle.hwpx")
    HwpxDocument.open(native).strict_validate()


def test_links_insert_target_edit_unlink_and_cross_story_copy(tmp_path):
    doc = WordDocument.blank()
    doc.append_paragraphs(["before after", "copy here"])
    with doc.edit_word_details() as edit:
        link = edit.body.paragraph(0).insert_hyperlink(7, "https://example.com/old", "LINK ")
        link.set_target("https://example.com/new").set_text("NEW ")
        assert link.text == "NEW "
        assert edit.body.paragraph(0).text == "before NEW after"
        edit.body.paragraph(0).range(7, 10).replace("EDIT")
        assert len(edit.body.hyperlinks) == 1 and edit.body.hyperlinks[0].text == "EDIT "
        edit.body.paragraph(0).range(7, 11).replace("NEW")
        edit.body.paragraph(0).range(7, 11).copy_to(edit.header().paragraph(0), 0)
        assert edit.header().hyperlinks[0].target == "https://example.com/new"
        edit.body.hyperlinks[0].unlink()
        assert edit.body.paragraph(0).text == "before NEW after"
    word = WordFile(doc.save(tmp_path / "links.docx"))
    assert word.sections[0].header._element.xpath(".//w:hyperlink")
    assert not word.element.xpath(".//w:hyperlink")


def test_style_catalogue_cycles_delete_reassignment_and_theme(tmp_path):
    doc = WordDocument.blank()
    doc.append_paragraph("Styled")
    with doc.edit_word_details() as edit:
        edit.create_style("Report", base="Normal", next_style="Normal", font={"bold": True}, paragraph={"space_after_pt": 12})
        edit.create_style("Child", base="Report")
        with pytest.raises(ValueError, match="cycle"):
            edit.configure_style("Report", base="Child")
        edit.body.paragraph(0).format(style="Report")
        edit.delete_style("Report", replacement="Normal")
        edit.format_theme(colors={"accent1": "224466"}, major_fonts={"latin": "Arial"})
    word = WordFile(doc.save(tmp_path / "styles.docx"))
    assert "Report" not in word.styles
    assert word.styles["Child"].base_style.name == "Normal"
    assert word.paragraphs[0].style.name == "Normal"
    native = doc.save(tmp_path / "middle.hwpx")
    restored = WordFile(Document.open(native).save(tmp_path / "restored.docx"))
    theme = next(rel.target_part.blob for rel in restored.part.rels.values() if rel.reltype.endswith("/theme"))
    assert b"224466" in theme and b"Arial" in theme


def test_table_layout_row_splitting_split_join_and_grid_guards(tmp_path):
    doc = WordDocument.blank()
    with doc.edit_word_details() as edit:
        table = edit.body.insert_table(0, rows=4, cols=2, values=[[str(i), "value"] for i in range(4)])
        table.merge(0, 0, 1, 0)
        with pytest.raises(ValueError, match="vertical"):
            table.split(1)
        following = table.split(2)
        assert len(table._native.rows) == 2 and len(following._native.rows) == 2
        table.join(following)
        assert len(table._native.rows) == 4
        table.format_row(0, allow_split=False)
        table.format_layout(wrap="around", x_pt=12, y_pt=6, distance_pt={"left": 4}, alignment="left")
    word = WordFile(doc.save(tmp_path / "table.docx"))
    assert word.tables[0]._tbl.xpath("./w:tblPr/w:tblpPr/@w:tblpX") == ["240"]
    assert word.tables[0]._tbl.xpath("./w:tr[1]/w:trPr/w:cantSplit/@w:val") == ["1"]
    restored = WordFile(Document.open(doc.save(tmp_path / "middle.hwpx")).save(tmp_path / "restored.docx"))
    assert restored.tables[0]._tbl.xpath("./w:tblPr/w:tblpPr/@w:tblpX") == ["240"]


@pytest.mark.parametrize("script", [
    "matrix {a & b # c & d}", "sum _ {i=1} ^ {n} {{x} _ {i}}", "int _ {0} ^ {1} {{x} ^ {2} dx}",
    "prod _ {k=1} ^ {n} {k}", "left [ {matrix {1 & 2 # 3 & 4}} right ]", "lim _ {x→0} {x}",
])
def test_structured_math_roundtrips_without_flattening(tmp_path, script):
    math = hwp_to_omml(script)
    canonical = omml_to_hwp(math)
    assert omml_to_hwp(hwp_to_omml(canonical)) == canonical
    doc = WordDocument.blank()
    with doc.edit_word_details() as edit:
        equation = edit.body.append_paragraph().add_equation(script)
        assert equation.script == canonical
        equation.set_script(canonical)
    restored = WordFile(Document.open(doc.save(tmp_path / "middle.hwpx")).save(tmp_path / "restored.docx"))
    assert omml_to_hwp(restored.element.xpath(".//m:oMath")[0]) == canonical


def test_math_operator_properties_survive_without_embedding_operands(tmp_path):
    doc = WordDocument.blank()
    with doc.edit_word_details() as edit:
        math = edit.body.append_paragraph().add_equation("sum _ {i=1} ^ {n} {x}")
        math._node.find(".//" + qn("m:limLoc")).set(qn("m:val"), "subSup")
    restored = WordFile(Document.open(doc.save(tmp_path / "middle.hwpx")).save(tmp_path / "restored.docx"))
    assert restored.element.xpath(".//m:naryPr/m:limLoc/@m:val") == ["subSup"]


def test_math_appearance_does_not_restore_an_old_operator_or_delimiter(tmp_path):
    from jakal_hwpx._hancom_models import Equation
    doc = WordDocument.blank()
    with doc.edit_word_details() as edit:
        equation = edit.body.append_paragraph().add_equation("sum _ {i=1} ^ {n} {x}")
        equation._node.find(".//" + qn("m:limLoc")).set(qn("m:val"), "subSup")
        equation.set_script("prod _ {k=1} ^ {m} {y}")
        assert equation._node.find(".//" + qn("m:chr")).get(qn("m:val")) == "∏"
        assert equation._node.find(".//" + qn("m:limLoc")).get(qn("m:val")) == "subSup"
        edit.body.append_paragraph().add_equation("left [ {x} right ]")
    common = Document.open(doc.save(tmp_path / "middle.hwpx"))
    equations = [block for section in common.sections for block in section.blocks if isinstance(block, Equation)]
    equations[0].script = "int _ {0} ^ {1} {x dx}"
    equations[1].script = "left ( {y} right )"
    restored = WordFile(common.save(tmp_path / "restored.docx"))
    assert restored.element.xpath(".//m:naryPr/m:chr/@m:val") == ["∫"]
    assert restored.element.xpath(".//m:dPr/m:begChr/@m:val") == ["("]


def test_display_matrix_math_grouping_is_preserved(tmp_path):
    doc = WordDocument.blank()
    with doc.edit_word() as backend:
        paragraph = backend.add_paragraph()
        display = OxmlElement("m:oMathPara")
        properties = OxmlElement("m:oMathParaPr")
        jc = OxmlElement("m:jc")
        jc.set(qn("m:val"), "center")
        properties.append(jc)
        display.append(properties)
        display.append(hwp_to_omml("matrix {1 & 2 # 3 & 4}"))
        paragraph._p.append(display)
    restored = WordFile(Document.open(doc.save(tmp_path / "middle.hwpx")).save(tmp_path / "restored.docx"))
    assert restored.element.xpath(".//m:oMathPara/m:oMath/m:m")
    assert restored.element.xpath(".//m:oMathParaPr/m:jc/@m:val") == ["center"]


def test_endpoint_persistent_ids_cas_atomic_batch_and_changed_attributes(tmp_path):
    doc = WordDocument.blank()
    doc.append_paragraphs(["Keep identity", "Other"])
    service = WordEndpoint(doc)
    initial = service.snapshot()
    p = next(item["id"] for item in initial["objects"] if item.get("text") == "Keep identity")
    response = service.apply(initial["version"], [
        {"target": "body", "method": "insert_paragraph", "args": [0, "Inserted"], "as": "$new"},
        {"target": p, "method": "format_text", "kwargs": {"bold": True}},
        {"target": "$new", "method": "set_text", "args": ["New paragraph"]},
    ])
    assert p in response["changed"] and response["created"]
    assert service.get_object(p)["object"]["text"] == "Keep identity"
    with pytest.raises(VersionConflict):
        service.apply(initial["version"], [])
    baseline = service.export()
    with pytest.raises(ObjectNotFound):
        service.apply(response["version"], [{"target": p, "method": "set_text", "args": ["Rolled back"]}, {"target": "missing", "method": "delete"}])
    assert service.export() == baseline and service.version == response["version"]
    reopened = WordEndpoint(WordDocument.open(doc.save(tmp_path / "ids.docx")))
    assert reopened.get_object(p)["object"]["text"] == "Keep identity"
    path = tmp_path / "middle.hwpx"
    path.write_bytes(service.export("hwpx"))
    restored = WordEndpoint(Document.open(Document.open(path).save(tmp_path / "restored.docx")))
    assert restored.get_object(p)["object"]["text"] == "Keep identity"


def test_identity_recovers_after_native_word_normalization_and_range_edits(tmp_path):
    from jakal_hwpx.addons.docx.identity import NS, PARA_ID, W, internal_name
    doc = WordDocument.blank()
    doc.append_paragraphs(["Original text", "Second paragraph"])
    service = WordEndpoint(doc)
    old = service.snapshot()
    identity = next(o["id"] for o in old["objects"] if o.get("text") == "Original text")
    with doc.edit_word() as backend:
        for part in backend.part.package.parts:
            root = getattr(part, "element", None)
            if root is None:
                continue
            for node in root.iter():
                for attr in list(node.attrib):
                    if attr.startswith("{" + NS + "}"):
                        node.attrib.pop(attr)
            for index, paragraph in enumerate(root.iter(W + "p")):
                paragraph.set(PARA_ID, f"{index + 1:08X}")
                # Word may hoist a bookmark start outside its paragraph.
                start = next((child for child in paragraph if child.tag == W + "bookmarkStart" and internal_name(child.get(W + "name"))), None)
                if start is not None:
                    paragraph.addprevious(start)
        backend.paragraphs[0].runs[0].text = "Changed by Word"
    reopened = WordEndpoint(WordDocument.open(doc.save(tmp_path / "word-normalized.docx")))
    assert reopened.get_object(identity)["object"]["text"] == "Changed by Word"
    with reopened.document.edit_word_details() as edit:
        assert not edit.bookmarks  # Internal anchors do not leak into the API.
        edit.body.paragraph(0).range(0, 7).replace("Updated")
        edit.body.range(0, 8, 1, 6).replace("merged")
    assert reopened.get_object(identity)["object"]["text"] == "Updated merged paragraph"


def test_endpoint_noop_keeps_revision_and_refuses_non_paragraph_ranges():
    doc = WordDocument.blank()
    doc.append_table(1, 1, [["Cell"]])
    service = WordEndpoint(doc)
    snapshot = service.snapshot()
    assert service.snapshot()["version"] == snapshot["version"]
    assert service.apply(snapshot["version"], [])["version"] == snapshot["version"]
    table = next(o["id"] for o in snapshot["objects"] if o["kind"] == "tbl")
    with pytest.raises(ValueError, match="paragraph"):
        service.apply(snapshot["version"], [{"target": table, "method": "replace_range", "kwargs": {"text": "invalid"}}])
    assert service.version == snapshot["version"]


def test_endpoint_table_split_then_join_and_section_carrier_indices():
    doc = WordDocument.blank()
    doc.append_paragraphs(["One", "Two"])
    doc.append_table(3, 2, [["a", "b"], ["c", "d"], ["e", "f"]])
    endpoint = WordEndpoint(doc)
    snapshot = endpoint.snapshot()
    table = next(o["id"] for o in snapshot["objects"] if o["kind"] == "tbl")
    endpoint.apply(snapshot["version"], [
        {"target": table, "method": "split", "args": [1], "as": "$following"},
        {"target": table, "method": "join", "kwargs": {"following": "$following"}},
    ])
    with doc.edit_word_details() as edit:
        edit.insert_section(1)
    with doc.edit_word_details() as edit:
        assert [p.text for p in edit.section(0).blocks] == ["One"]
        edit.section(0).append_paragraph("Still section one")
    with doc.edit_word_details() as edit:
        assert [p.text for p in edit.section(0).blocks] == ["One", "Still section one"]


def test_endpoint_preserves_comment_parts_without_exposing_them_as_body(tmp_path):
    import zipfile
    word = WordFile()
    run = word.add_paragraph().add_run("Body with comment")
    word.add_comment(run, text="Kept comment", author="Reviewer")
    path = tmp_path / "comments.docx"
    word.save(path)
    with zipfile.ZipFile(path) as package:
        original = package.read("word/comments.xml")
    service = WordEndpoint(WordDocument.open(path))
    snapshot = service.snapshot()
    assert any(o.get("text") == "Body with comment" for o in snapshot["objects"])
    assert not any(o.get("text") == "Kept comment" for o in snapshot["objects"])
    import io
    with zipfile.ZipFile(io.BytesIO(service.export())) as package:
        assert package.read("word/comments.xml") == original


def test_http_adapter_requests_exports_conflicts_and_unauthorized_calls():
    doc = WordDocument.blank()
    doc.append_paragraph("HTTP")
    endpoint = WordEndpoint(doc)
    server = make_http_server(endpoint, port=0, token="test-token")
    thread = Thread(target=server.serve_forever, daemon=True)
    thread.start()
    connection = http.client.HTTPConnection(*server.server_address, timeout=5)
    try:
        connection.request("GET", "/document")
        response = connection.getresponse()
        assert response.status == 401
        response.read()
        headers = {"Authorization": "Bearer test-token", "Content-Type": "application/json"}
        connection.request("GET", "/document", headers=headers)
        response = connection.getresponse()
        data = json.loads(response.read())
        target = next(item["id"] for item in data["objects"] if item.get("text") == "HTTP")
        connection.request("GET", "/version", headers=headers)
        response = connection.getresponse()
        assert response.status == 200 and json.loads(response.read())["version"] == data["version"]
        connection.request("POST", "/objects", json.dumps({"ids": [target, "missing"]}), headers)
        response = connection.getresponse()
        selected = json.loads(response.read())
        assert response.status == 200 and selected["missing"] == ["missing"]
        payload = json.dumps({"expected_version": data["version"], "compact": True, "operations": [{"target": target, "method": "set_text", "args": ["Edited"]}]})
        connection.request("POST", "/edits", payload, headers)
        response = connection.getresponse()
        assert response.status == 200
        changed = json.loads(response.read())
        assert "objects" not in changed and changed["upserted"]
        connection.request("GET", "/changes?since=" + data["version"], headers=headers)
        response = connection.getresponse()
        assert response.status == 200 and json.loads(response.read())["upserted"] == changed["upserted"]
        connection.request("POST", "/edits", payload, headers)
        response = connection.getresponse()
        assert response.status == 409
        response.read()
        connection.request("GET", "/export.docx", headers=headers)
        response = connection.getresponse()
        assert response.status == 200 and response.read().startswith(b"PK")
    finally:
        connection.close()
        server.shutdown()
        server.server_close()
        thread.join(5)


@pytest.mark.parametrize("simple", [True, False])
def test_link_identity_survives_field_representation_and_old_methods(tmp_path, simple):
    from jakal_hwpx.addons.docx.identity import NS
    doc = WordDocument.blank()
    with doc.edit_word_details() as edit:
        edit.body.append_paragraph().insert_hyperlink(0, "#reference", "Jump")
    service = WordEndpoint(doc)
    initial = service.snapshot()
    identity = next(o["id"] for o in initial["objects"] if o["kind"] == "hyperlink")
    with doc.edit_word() as backend:
        link = next(backend.element.iter(qn("w:hyperlink")))
        if simple:
            field = OxmlElement("w:fldSimple")
            field.set(qn("w:instr"), 'HYPERLINK \\l "reference" \\h')
            field.extend(list(link))
            link.addprevious(field)
        else:
            for name, text in [("fldChar", "begin"), ("instrText", ' HYPERLINK \\l "reference" \\h '), ("fldChar", "separate")]:
                run, child = OxmlElement("w:r"), OxmlElement("w:" + name)
                if name == "fldChar":
                    child.set(qn("w:fldCharType"), text)
                else:
                    child.text = text
                run.append(child)
                link.addprevious(run)
            for child in list(link):
                link.addprevious(child)
            run, end = OxmlElement("w:r"), OxmlElement("w:fldChar")
            end.set(qn("w:fldCharType"), "end")
            run.append(end)
            link.addprevious(run)
        link.getparent().remove(link)
        for node in backend.element.iter():
            for attr in list(node.attrib):
                if attr.startswith("{" + NS + "}"):
                    node.attrib.pop(attr)
    reopened = WordEndpoint(WordDocument.open(doc.save(tmp_path / "normalized.docx")))
    current = reopened.get_object(identity)
    assert current["object"]["kind"] == "hyperlink"
    assert current["object"]["target"] == "#reference"
    result = reopened.apply(current["version"], [
        {"target": identity, "method": "set_target", "args": ["https://example.com/#new"]},
        {"target": identity, "method": "set_text", "args": ["Updated"]},
    ], compact=True)
    assert identity in result["changed"] and "objects" not in result
    assert reopened.get_object(identity)["object"]["target"] == "https://example.com/#new"
    assert reopened.get_object(identity)["object"]["text"] == "Updated"
    reopened.apply(result["version"], [{"target": identity, "method": "unlink"}])
    assert reopened.get_objects([identity])["missing"] == [identity]
    with reopened.document.edit_word_details() as edit:
        assert edit.body.paragraph(0).text == "Updated"


def test_endpoint_cached_reads_compact_changes_history_and_client_copy(monkeypatch):
    doc = WordDocument.blank()
    doc.append_paragraph("Keep")
    service = WordEndpoint(doc)
    initial = service.snapshot()
    identity = next(o["id"] for o in initial["objects"] if o.get("text") == "Keep")
    prepare = service._prepare
    calls = []
    monkeypatch.setattr(service, "_prepare", lambda editor: (calls.append(True), prepare(editor)))
    for _ in range(3):
        assert service.get_object(identity)["object"]["text"] == "Keep"
        assert service.version == initial["version"]
    assert not calls
    initial["objects"].clear()
    assert service.get_objects([identity, "missing", identity])["missing"] == ["missing"]
    revision = service.version
    response = service.apply(revision, [{"target": identity, "method": "set_text", "args": ["Updated"]}], compact=True)
    delta = service.changes(revision)
    assert delta["upserted"] == response["upserted"] and not delta["deleted"]
    assert service.changes(service.version)["upserted"] == []
    delta["upserted"].clear()
    assert service.changes(revision)["upserted"]
    with pytest.raises(VersionConflict):
        service.apply(revision, [], compact=True)
    deleted = service.apply(service.version, [{"target": identity, "method": "delete"}], compact=True)
    assert identity in deleted["deleted"]
    assert identity in service.changes(revision)["deleted"]
    for index in range(16):
        service.apply(service.version, [{"target": "body", "method": "append_paragraph", "args": [str(index)]}], compact=True)
    assert service.changes(revision)["requires_snapshot"] is True
    assert "upserted" not in service.changes("unknown")


@pytest.mark.parametrize("old_count,new_count", [(1, 2), (2, 1), (2, 2)])
def test_ambiguous_normalized_links_never_reuse_ids(tmp_path, old_count, new_count):
    from jakal_hwpx.addons.docx.identity import NS
    doc = WordDocument.blank()
    with doc.edit_word_details() as edit:
        p = edit.body.append_paragraph()
        for _ in range(old_count):
            p.insert_hyperlink(0, "#reference", "Same")
    old = WordEndpoint(doc).snapshot()
    identities = {o["id"] for o in old["objects"] if o["kind"] == "hyperlink"}
    with doc.edit_word() as backend:
        paragraph = backend.paragraphs[0]._p
        for link in list(paragraph.iter(qn("w:hyperlink"))):
            paragraph.remove(link)
        for _ in range(new_count):
            field = OxmlElement("w:fldSimple")
            field.set(qn("w:instr"), 'HYPERLINK \\l "reference"')
            run, text = OxmlElement("w:r"), OxmlElement("w:t")
            text.text = "Same"
            run.append(text)
            field.append(run)
            paragraph.append(field)
        for node in backend.element.iter():
            for attr in list(node.attrib):
                if attr.startswith("{" + NS + "}"):
                    node.attrib.pop(attr)
    new = WordEndpoint(WordDocument.open(doc.save(tmp_path / "ambiguous.docx"))).snapshot()
    assert not identities.intersection(o["id"] for o in new["objects"])


@pytest.mark.parametrize("anchor,preserved", [("reference", True), ("different", False)])
def test_field_to_link_identity_requires_matching_target(tmp_path, anchor, preserved):
    from jakal_hwpx.addons.docx.identity import NS
    doc = WordDocument.blank()
    with doc.edit_word_details() as edit:
        edit.body.append_paragraph().add_field('HYPERLINK \\l "reference"', "Same caption")
    service = WordEndpoint(doc)
    old = service.snapshot()
    identity = next(o["id"] for o in old["objects"] if o["kind"] == "hyperlink")
    with doc.edit_word() as backend:
        field = next(backend.element.iter(qn("w:fldSimple")))
        link = OxmlElement("w:hyperlink")
        link.set(qn("w:anchor"), anchor)
        link.extend(list(field))
        field.addprevious(link)
        field.getparent().remove(field)
        for node in backend.element.iter():
            for attr in list(node.attrib):
                if attr.startswith("{" + NS + "}"):
                    node.attrib.pop(attr)
    reopened = WordEndpoint(WordDocument.open(doc.save(tmp_path / "reverse.docx")))
    assert (identity not in reopened.get_objects([identity])["missing"]) is preserved
    with reopened.document.edit_word_details() as edit:
        assert len(edit.body.hyperlinks) == 1
