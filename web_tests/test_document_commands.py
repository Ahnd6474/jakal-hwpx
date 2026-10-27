"""Transport-independent use cases and durable command identity."""
import json
import threading
import time
from concurrent.futures import ThreadPoolExecutor

import pytest
from fastapi.testclient import TestClient

from jakal_hwpx import HwpxDocument
from server.app import create_app
from server.document_service import DocumentService
from server.errors import WorkspaceError
from server.schemas import EditRequest
from server.storage import Store
from web_tests.test_api import client, raw_document, upload, wait


def edit(revision=0, text="공통 작업 계층에서 수정", **extra):
    return {"revision": revision, "operations": [{"type": "paragraph", "section": 0, "index": 0, "text": text}], **extra}


def test_replay_after_completion_returns_same_job_and_one_revision(client, raw_document):
    identifier = upload(client, raw_document)
    base = f"/api/documents/{identifier}"
    headers = {"Idempotency-Key": "edit-request-00000001"}
    first = client.post(base + "/operations", json=edit(settings_version=0), headers=headers)
    assert first.status_code == 202, first.text
    receipt = first.json()
    assert receipt["kind"] == "edit" and receipt["base_revision"] == 0
    assert wait(client, receipt["job_id"])["status"] == "completed"
    replay = client.post(base + "/operations", json=edit(settings_version=0), headers=headers)
    assert replay.status_code == 202 and replay.json() == receipt
    assert client.get(base).json()["revision"] == 1
    assert len(client.get(base + "/revisions").json()["items"]) == 2
    changed = client.post(base + "/operations", json=edit(text="다른 내용", settings_version=0), headers=headers)
    assert changed.status_code == 409 and changed.json()["code"] == "request_key_conflict"
    stale = client.post(base + "/operations", json=edit(), headers={"Idempotency-Key": "edit-request-00000002"})
    assert stale.status_code == 409 and stale.json()["code"] == "revision_conflict"
    assert client.get(base + "/original").content == raw_document


def test_identity_survives_store_restart_and_concurrent_submissions(tmp_path):
    store = Store(tmp_path)
    identifier = "a" * 32
    store.create_document("owner", "test.hwpx", identifier)
    with store.connect() as db:
        db.execute("UPDATE documents SET status='ready'")
    barrier = threading.Barrier(8)

    def submit(_):
        barrier.wait(timeout=10)
        return Store(tmp_path).enqueue(identifier, "owner", "edit", edit(), 0, request_key="same-request-000001")

    with ThreadPoolExecutor(max_workers=8) as pool:
        jobs = list(pool.map(submit, range(8)))
    assert len(set(jobs)) == 1
    with store.connect() as db:
        assert db.execute("SELECT count(*) FROM jobs").fetchone()[0] == 1
    restarted = Store(tmp_path)
    assert restarted.enqueue(identifier, "owner", "edit", edit(), 0, request_key="same-request-000001") == jobs[0]
    with pytest.raises(WorkspaceError) as error:
        restarted.enqueue(identifier, "stranger", "edit", edit(), 0, request_key="same-request-000001")
    assert error.value.code == "document_not_found"
    with store.connect() as db:
        db.execute("UPDATE documents SET created=?", (time.time() - store.ttl - 1,))
    with pytest.raises(WorkspaceError):
        restarted.enqueue(identifier, "owner", "edit", edit(), 0, request_key="same-request-000001")


def test_studio_replay_hashes_file_content_and_removes_unused_candidate(client, raw_document):
    identifier = upload(client, raw_document)
    document = HwpxDocument.from_bytes(raw_document)
    document.set_paragraph_text(0, 0, "원문 편집 저장")
    modified = document.compile()
    path = f"/api/documents/{identifier}/studio?revision=0&settings_version=0"
    headers = {"Idempotency-Key": "studio-request-00001"}

    def save(raw):
        return client.post(path, files={"file": ("edited.hwpx", raw)}, headers=headers)

    first = save(modified)
    assert first.status_code == 202, first.text
    assert wait(client, first.json()["job_id"])["status"] == "completed"
    assert save(modified).json() == first.json()
    conflict = save(raw_document)
    assert conflict.status_code == 409 and conflict.json()["code"] == "request_key_conflict"
    directory = client.app.state.store.directory(identifier)
    assert list(directory.glob("studio-*.hwpx")) == []
    assert (directory / "revision-1.hwpx").read_bytes() == modified


def test_settings_fence_blocks_old_policy_and_state_is_lightweight(client, raw_document, monkeypatch):
    identifier = upload(client, raw_document)
    base = f"/api/documents/{identifier}"
    settings = client.get(base).json()["settings"]
    settings["protected_terms"] = ["원본"]
    assert client.put(base + "/settings", json={"version": 0, "settings": settings}).status_code == 200
    response = client.post(base + "/operations", json=edit(settings_version=0))
    assert response.status_code == 409 and response.json()["code"] == "settings_conflict"
    response = client.post(base + "/validations", json={"revision": 0, "settings_version": 0})
    assert response.status_code == 409 and response.json()["code"] == "settings_conflict"
    response = client.post(base + "/studio?revision=0&settings_version=0", files={"file": ("edited.hwpx", raw_document)})
    assert response.status_code == 409 and response.json()["code"] == "settings_conflict"
    monkeypatch.setattr(client.app.state.documents, "view", lambda *_: pytest.fail("state must not load snapshots"))
    response = client.get(base + "/state")
    assert response.status_code == 200 and response.json()["settings_version"] == 1
    assert "sections" not in response.json() and "settings" not in response.json()
    assert response.headers["cache-control"] == "no-store"
    with TestClient(create_app(client.app.state.store.root, run_worker=False)) as stranger:
        assert stranger.get(base + "/state").status_code == 404


def test_use_case_without_http_and_invalid_keys(tmp_path):
    store = Store(tmp_path)
    identifier = "b" * 32
    store.create_document("owner", "test.hwpx", identifier)
    with store.connect() as db:
        db.execute("UPDATE documents SET status='ready'")
    service = DocumentService(store)
    body = EditRequest.model_validate(edit())
    with pytest.raises(WorkspaceError) as error:
        service.edit(identifier, "owner", body, "short")
    assert error.value.code == "invalid_request_key"
    body = EditRequest.model_validate({**edit(), "operations": edit()["operations"] * 2})
    with pytest.raises(WorkspaceError) as error:
        service.edit(identifier, "owner", body)
    assert error.value.code == "duplicate_target"
    receipt = service.edit(identifier, "owner", EditRequest.model_validate(edit()), "internal-request-0001")
    stored = store.job(receipt["job_id"], "owner")
    assert stored["kind"] == "edit" and json.loads(stored["payload"])["revision"] == 0


def test_validation_and_restoration_replay(client, raw_document):
    identifier = upload(client, raw_document)
    base = f"/api/documents/{identifier}"
    job = client.post(base + "/operations", json=edit()).json()
    assert wait(client, job["job_id"])["status"] == "completed"
    for endpoint, payload, key in [
        ("validations", {"revision": 1, "settings_version": 0}, "validation-request-01"),
        ("restorations", {"revision": 1, "source_revision": 0, "settings_version": 0}, "restoration-request-1"),
    ]:
        headers = {"Idempotency-Key": key}
        result = client.post(base + "/" + endpoint, json=payload, headers=headers)
        assert result.status_code == 202, result.text
        assert wait(client, result.json()["job_id"])["status"] == "completed"
        assert client.post(base + "/" + endpoint, json=payload, headers=headers).json() == result.json()
    assert client.get(base).json()["revision"] == 2
