"""Full page editor saves remain versioned, owned and validated."""
from fastapi.testclient import TestClient

from jakal_hwpx import HwpxDocument
from jakal_hwpx.namespaces import NS
from web_tests.test_api import client, raw_document, upload, wait


def save(client, identifier, raw, revision=0):
    return client.post(f"/api/documents/{identifier}/studio?revision={revision}", files={"file": ("edited.hwpx", raw, "application/hwp+zip")})


def test_studio_roundtrip_preserves_original_and_saves_format_and_structure(client, raw_document):
    identifier = upload(client, raw_document)
    document = HwpxDocument.from_bytes(raw_document)
    document.append_paragraph("직접 편집한 문단 😀")
    style = document.append_character_style()
    style.element.set("height", "1800")
    style_id = style.element.get("id")
    document.apply_style_to_paragraph(0, 0, char_pr_id=style_id)
    document.append_table(rows=1, columns=2, cell_texts=[["새 표", "새 셀"]])
    edited = document.compile()
    response = save(client, identifier, edited)
    assert response.status_code == 202, response.text
    job = wait(client, response.json()["job_id"])
    assert job["status"] == "completed", job
    base = f"/api/documents/{identifier}"
    fresh = client.get(base).json()
    assert fresh["revision"] == 1 and fresh["validated"] is None
    assert len(fresh["sections"][0]["tables"]) == 2
    revisions = client.get(base + "/revisions").json()["items"]
    assert [item["kind"] for item in revisions] == ["studio", "analyze"]
    assert client.get(base + "/original").content == raw_document
    assert client.get(base + "/source?revision=1").content == edited
    assert client.get(base + "/download?revision=1").status_code == 409
    check = client.post(base + "/validations", json={"revision": 1}).json()
    assert wait(client, check["job_id"])["status"] == "completed"
    assert client.get(base + "/download?revision=1").content == edited
    reopened = HwpxDocument.from_bytes(edited)
    assert reopened.sections[0].paragraphs()[0]._element.find("hp:run", NS).get("charPrIDRef") == style_id


def test_studio_save_preserves_nested_tables_and_their_editable_cell_text(client, raw_document):
    identifier = upload(client, raw_document)
    document = HwpxDocument.from_bytes(raw_document)
    outer = document.tables()[0]
    inner = document.append_table(rows=2, columns=2, cell_texts=[["안쪽 표", "한글 😀"], ["수식 x²", "끝"]])
    # Insert a real table control in the outer cell's own paragraph.
    cell_run = outer.cell(0, 0).element.find("hp:subList/hp:p/hp:run", NS)
    cell_run.append(inner.element)
    inner.element.find("hp:pos", NS).set("treatAsChar", "1")
    edited = document.compile()
    response = save(client, identifier, edited)
    assert response.status_code == 202
    assert wait(client, response.json()["job_id"])["status"] == "completed"
    base = f"/api/documents/{identifier}"
    check = client.post(base + "/validations", json={"revision": 1}).json()
    assert wait(client, check["job_id"])["status"] == "completed"
    reopened = HwpxDocument.from_bytes(client.get(base + "/download?revision=1").content)
    root = reopened.sections[0].root_element
    nested = root.xpath(".//hp:tc/hp:subList/hp:p/hp:run/hp:tbl", namespaces=NS)
    assert len(nested) == 1
    assert nested[0].xpath(".//hp:t/text()", namespaces=NS) == ["안쪽 표", "한글 😀", "수식 x²", "끝"]
    assert client.get(base + "/original").content == raw_document


def test_studio_invalid_reference_does_not_publish_or_replace_source(client, raw_document):
    identifier = upload(client, raw_document)
    document = HwpxDocument.from_bytes(raw_document)
    document.sections[0].paragraphs()[1].element.set("paraPrIDRef", "999999")
    response = save(client, identifier, document.compile(validate=False, auto_repair=False))
    job = wait(client, response.json()["job_id"])
    assert job["status"] == "failed" and "오류" in job["error"]
    assert client.get(f"/api/documents/{identifier}").json()["revision"] == 0
    assert client.get(f"/api/documents/{identifier}/source?revision=0").content == raw_document
    assert not list(client.app.state.store.directory(identifier).glob("studio-*.hwpx"))


def test_studio_ownership_stale_version_and_invalid_archive(client, raw_document):
    identifier = upload(client, raw_document)
    stranger = TestClient(client.app)
    assert save(stranger, identifier, raw_document).status_code == 404
    stranger.close()
    assert save(client, identifier, raw_document, revision=9).status_code == 409
    assert not list(client.app.state.store.directory(identifier).glob("studio-*.hwpx"))
    assert save(client, identifier, b"").status_code == 422
    response = save(client, identifier, b"not a ZIP")
    assert wait(client, response.json()["job_id"])["status"] == "failed"
    assert client.get(f"/api/documents/{identifier}").json()["revision"] == 0


def test_studio_protected_terms_apply_to_complete_document(client, raw_document):
    identifier = upload(client, raw_document)
    base = f"/api/documents/{identifier}"
    response = client.put(base + "/settings", json={"version": 0, "settings": {"protected_terms": ["자칼"]}})
    assert response.status_code == 200
    document = HwpxDocument.from_bytes(raw_document)
    document.tables(0)[0].cell(1, 1).set_text("다른 이름")
    response = save(client, identifier, document.compile())
    job = wait(client, response.json()["job_id"])
    assert job["status"] == "failed" and "보호 문구" in job["error"]
    assert client.get(base).json()["revision"] == 0


def test_studio_csp_keeps_editor_local_and_other_pages_unframeable(client):
    root_policy = client.get("/").headers["content-security-policy"]
    editor_policy = client.get("/studio/").headers["content-security-policy"]
    assert "frame-ancestors 'none'" in root_policy
    assert "frame-ancestors 'self'" in editor_policy
    assert "connect-src 'self'" in editor_policy
    assert "font-src 'self'" in editor_policy
