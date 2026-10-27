from __future__ import annotations

import io
import json
import time
import zipfile
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from jakal_hwpx import HwpxDocument
from jakal_hwpx.namespaces import NS
from server.app import create_app
from server.engine import cell_text, inspect_archive


@pytest.fixture
def client(tmp_path):
    with TestClient(create_app(tmp_path / "data")) as value:
        yield value


@pytest.fixture
def raw_document():
    document = HwpxDocument.blank()
    document.set_paragraph_text(0, 0, "원본 문단")
    document.append_paragraph("둘째 문단 <안내>")
    document.append_table(rows=2, columns=2, cell_texts=[["항목", "값"], ["이름", "자칼"]])
    document.append_equation("x+y", width=3200, height=1800)
    return document.compile()


def wait(client, identifier, timeout=20):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        response = client.get(f"/api/jobs/{identifier}")
        assert response.status_code == 200, response.text
        job = response.json()
        if job["status"] in {"completed", "failed"}:
            return job
        time.sleep(.04)
    pytest.fail(f"Worker did not finish within {timeout} seconds")


def upload(client, raw):
    response = client.post("/api/documents", files={"file": ("시험.hwpx", raw, "application/hwp+zip")})
    assert response.status_code == 202, response.text
    result = response.json()
    job = wait(client, result["job_id"])
    assert job["status"] == "completed", job
    return result["document_id"]


def test_full_roundtrip_and_validation_revision_gate(client, raw_document):
    identifier = upload(client, raw_document)
    base = f"/api/documents/{identifier}"
    doc = client.get(base).json()
    assert doc["sections"][0]["paragraphs"][0]["editable"]
    assert doc["sections"][0]["tables"][0]["cells"][3]["text"] == "자칼"
    assert client.get(f"{base}/download?revision=0").status_code == 409
    operations = [{"type": "paragraph", "section": 0, "index": 0, "text": "수정한 문단 & <안녕> 😀"}, {"type": "cell", "section": 0, "table": 0, "row": 1, "column": 1, "text": "홍길동\n두 번째 줄"}]
    response = client.post(f"{base}/operations", json={"revision": 0, "operations": operations})
    assert response.status_code == 202, response.text
    result = wait(client, response.json()["job_id"])
    assert result["status"] == "completed", result
    assert client.post(f"{base}/validations", json={"revision": 0}).status_code == 409
    job = client.post(f"{base}/validations", json={"revision": 1}).json()
    assert wait(client, job["job_id"])["status"] == "completed"
    exported = client.post(f"{base}/exports", json={"revision": 1})
    assert exported.status_code == 200, exported.text
    response = client.get(exported.json()["url"])
    assert response.status_code == 200
    reopened = HwpxDocument.from_bytes(response.content)
    assert reopened.sections[0].text_fragments()[0] == operations[0]["text"]
    assert cell_text(reopened.tables()[0].cell(1, 1)) == operations[1]["text"]
    assert len(reopened.equations()) == 1
    assert reopened.validation_errors() == []
    assert client.get(f"{base}/report").json()["new_issues"] == []
    assert "홍길동" in client.get(f"{base}/text").text
    store = client.app.state.store
    assert (store.directory(identifier) / "revision-0.hwpx").read_bytes() == raw_document
    # A later edit invalidates the previous validation and download URL.
    next_job = client.post(f"{base}/operations", json={"revision": 1, "operations": [{"type": "paragraph", "section": 0, "index": 1, "text": "다시 수정"}]}).json()
    assert wait(client, next_job["job_id"])["status"] == "completed"
    assert client.get(exported.json()["url"]).status_code == 409


def test_noop_download_is_byte_identical(client, raw_document):
    identifier = upload(client, raw_document)
    job = client.post(f"/api/documents/{identifier}/validations", json={"revision": 0}).json()
    assert wait(client, job["job_id"])["status"] == "completed"
    assert client.get(f"/api/documents/{identifier}/download?revision=0").content == raw_document


def test_session_isolation_and_cross_site_request(client, raw_document):
    identifier = upload(client, raw_document)
    app = client.app
    stranger = TestClient(app)
    try:
        assert stranger.get("/api/documents").json() == []
        assert stranger.get(f"/api/documents/{identifier}").status_code == 404
        assert stranger.get(f"/api/documents/{identifier}/download?revision=0").status_code == 404
        assert stranger.delete(f"/api/documents/{identifier}").status_code == 404
    finally:
        stranger.close()
    assert client.post("/api/documents", headers={"Origin": "https://other.example"}, files={"file": ("x.hwpx", raw_document)}).status_code == 403


def test_edit_failure_leaves_original_revision_intact(client, raw_document):
    identifier = upload(client, raw_document)
    base = f"/api/documents/{identifier}"
    job = client.post(f"{base}/operations", json={"revision": 0, "operations": [{"type": "paragraph", "section": 0, "index": 999, "text": "실패"}]}).json()
    assert wait(client, job["job_id"])["status"] == "failed"
    assert client.get(base).json()["revision"] == 0
    assert client.get(base).json()["sections"][0]["paragraphs"][0]["text"] == "원본 문단"


def test_complex_controls_are_readonly(client, raw_document):
    identifier = upload(client, raw_document)
    base = f"/api/documents/{identifier}"
    doc = client.get(base).json()
    readonly = [p for p in doc["sections"][0]["paragraphs"] if not p["editable"]]
    assert readonly
    job = client.post(f"{base}/operations", json={"revision": 0, "operations": [{"type": "paragraph", "section": 0, "index": readonly[0]["index"], "text": "교체"}]}).json()
    result = wait(client, job["job_id"])
    assert result["status"] == "failed"
    assert "읽기 전용" in result["error"]


def test_invalid_uploads_and_allowlisted_operations(client, raw_document):
    assert client.post("/api/documents", files={"file": ("x.doc", b"bad")}).status_code == 415
    invalid_hwp = client.post("/api/documents", files={"file": ("x.hwp", b"bad")}).json()
    assert wait(client, invalid_hwp["job_id"])["status"] == "failed"
    assert client.post("/api/documents", files={"file": ("x.hwpx", b"")}).status_code == 400
    bad = client.post("/api/documents", files={"file": ("x.hwpx", b"not a zip")}).json()
    assert wait(client, bad["job_id"])["status"] == "failed"
    identifier = upload(client, raw_document)
    op = {"type": "paragraph", "section": 0, "index": 0, "text": "test"}
    base = f"/api/documents/{identifier}"
    assert client.post(f"{base}/operations", json={"revision": 0, "operations": [op, op]}).status_code == 422
    assert client.post(f"{base}/operations", json={"revision": 0, "operations": [{"type": "exec", "code": "print(1)"}]}).status_code == 422
    assert client.post(f"{base}/operations", json={"revision": 0, "operations": [{**op, "index": -1}]}).status_code == 422


@pytest.mark.parametrize("name,content", [("../outside.xml", b"<root/>"), ("Contents/evil.xml", b'<!DOCTYPE x [<!ENTITY xxe SYSTEM "file:///secret">]><x>&xxe;</x>')])
def test_archive_preflight_rejects_paths_and_dtd(tmp_path, raw_document, name, content):
    path = tmp_path / "bad.hwpx"
    path.write_bytes(raw_document)
    with zipfile.ZipFile(path, "a") as archive:
        archive.writestr(name, content)
    with pytest.raises(ValueError):
        inspect_archive(path)


def test_delete_removes_files_and_jobs(client, raw_document):
    identifier = upload(client, raw_document)
    directory = client.app.state.store.directory(identifier)
    assert client.delete(f"/api/documents/{identifier}").status_code == 200
    assert not directory.exists()
    assert client.get(f"/api/documents/{identifier}").status_code == 404


def test_expired_documents_are_inaccessible_and_cleaned(client, raw_document):
    identifier = upload(client, raw_document)
    store = client.app.state.store
    with store.connect() as db:
        db.execute("UPDATE documents SET created=? WHERE id=?", (time.time() - 90_000, identifier))
    assert client.get(f"/api/documents/{identifier}").status_code == 404
    store.cleanup()
    assert not store.directory(identifier).exists()


def test_request_body_limit_and_upload_cleanup(client):
    response = client.post("/api/documents", files={"file": ("too-big.hwpx", b"x" * (20 * 1024 * 1024 + 1))})
    assert response.status_code == 413
    assert list(client.app.state.store.root.glob("*/revision-0.hwpx")) == []
    response = client.post("/api/documents", content=b"x", headers={"Content-Length": str(22 * 1024 * 1024)})
    assert response.status_code == 413


def test_stale_revision_and_pending_work_are_rejected(tmp_path, raw_document):
    app = create_app(tmp_path / "data", run_worker=False)
    with TestClient(app) as client:
        response = client.post("/api/documents", files={"file": ("test.hwpx", raw_document)})
        identifier = response.json()["document_id"]
        # No second job can run while initial analysis is still queued.
        response = client.post(f"/api/documents/{identifier}/validations", json={"revision": 0})
        assert response.status_code == 409


def test_interrupted_job_recovery(tmp_path):
    from server.storage import Store
    store = Store(tmp_path)
    identifier = "a" * 32
    store.create_document("owner", "test.hwpx", identifier)
    job = store.enqueue(identifier, "owner", "analyze", {})
    assert store.claim()["id"] == job
    store.recover()
    assert store.job(job, "owner")["status"] == "failed"
    assert store.document(identifier, "owner")["status"] == "failed"


def test_delete_retries_locked_files_without_orphaning_document(tmp_path, monkeypatch):
    import shutil
    from server.storage import Store
    store = Store(tmp_path)
    identifier = "b" * 32
    directory = store.directory(identifier)
    directory.mkdir()
    (directory / "revision-0.hwpx").write_bytes(b"fixture")
    store.create_document("owner", "test.hwpx", identifier)
    assert not store.delete(identifier, "stranger")
    original = shutil.rmtree
    monkeypatch.setattr(shutil, "rmtree", lambda *a, **kw: None)
    assert store.delete(identifier, "owner")
    assert store.document(identifier, "owner") is None
    assert (tmp_path / f"trash-{identifier}").exists()
    monkeypatch.setattr(shutil, "rmtree", original)
    store.cleanup()
    assert not (tmp_path / f"trash-{identifier}").exists()


def test_long_upload_name_keeps_extension(client, raw_document):
    response = client.post("/api/documents", files={"file": ("긴문서" * 70 + ".hwpx", raw_document)})
    assert response.status_code == 202
    value = response.json()
    assert wait(client, value["job_id"])["status"] == "completed"
    document = client.get(f"/api/documents/{value['document_id']}").json()
    assert len(document["name"]) == 180
    assert document["name"].endswith(".hwpx")
    assert client.get(f"/api/documents/{value['document_id']}/original").content == raw_document


@pytest.mark.parametrize("corruption", [None, "{broken json", '{"revision": 99}'])
def test_unreadable_snapshot_does_not_block_original_recovery(client, raw_document, corruption):
    identifier = upload(client, raw_document)
    path = client.app.state.store.directory(identifier) / "snapshot-0.json"
    if corruption is None:
        path.unlink()
    else:
        path.write_text(corruption, encoding="utf-8")
    response = client.get(f"/api/documents/{identifier}")
    assert response.status_code == 503
    assert "원본" in response.json()["detail"]
    assert client.get(f"/api/documents/{identifier}/original").content == raw_document
