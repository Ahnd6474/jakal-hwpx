from __future__ import annotations

import json
import time

import pytest
from fastapi.testclient import TestClient

from jakal_hwpx import HwpDocument
from server.engine import process
from server.storage import Store
from web_tests.test_api import client, raw_document, upload, wait


def edit(client, base, revision, text):
    response = client.post(base + "/operations", json={"revision": revision, "operations": [{"type": "paragraph", "section": 0, "index": 0, "text": text}]})
    assert response.status_code == 202, response.text
    assert wait(client, response.json()["job_id"])["status"] == "completed"


def test_history_restore_keeps_all_versions_and_requires_new_validation(client, raw_document):
    identifier = upload(client, raw_document)
    base = f"/api/documents/{identifier}"
    edit(client, base, 0, "첫 번째 수정")
    edit(client, base, 1, "두 번째 수정")
    validation = client.post(base + "/validations", json={"revision": 2})
    assert wait(client, validation.json()["job_id"])["status"] == "completed"
    history = client.get(base + "/revisions").json()
    assert [item["revision"] for item in history["items"]] == [2, 1, 0]
    assert [item["kind"] for item in history["items"]] == ["edit", "edit", "analyze"]
    assert history["next_before"] is None
    assert client.get(base + "/revisions/1").json()["sections"][0]["paragraphs"][0]["text"] == "첫 번째 수정"
    directory = client.app.state.store.directory(identifier)
    second = (directory / "revision-2.hwpx").read_bytes()
    response = client.post(base + "/restorations", json={"revision": 2, "source_revision": 0})
    assert response.status_code == 202, response.text
    result = wait(client, response.json()["job_id"])
    assert result["status"] == "completed", result
    fresh = client.get(base).json()
    assert fresh["revision"] == 3 and fresh["validated"] is None
    assert fresh["sections"][0]["paragraphs"][0]["text"] == "원본 문단"
    assert fresh["report"]["checked_revision"] == 3
    assert (directory / "revision-3.hwpx").read_bytes() == raw_document
    assert (directory / "revision-2.hwpx").read_bytes() == second
    assert client.get(base + "/original").content == raw_document
    assert client.get(base + "/download?revision=2").status_code == 409
    assert client.get(base + "/download?revision=3").status_code == 409
    history = client.get(base + "/revisions").json()["items"]
    assert [item["revision"] for item in history] == [3, 2, 1, 0]
    assert history[0]["kind"] == "restore" and history[0]["source_revision"] == 0
    validation = client.post(base + "/validations", json={"revision": 3})
    assert wait(client, validation.json()["job_id"])["status"] == "completed"
    assert client.get(base + "/download?revision=3").content == raw_document
    assert len(client.get(base + "/revisions").json()["items"]) == 4


def test_restore_honors_current_protected_phrases_and_stale_revision(client, raw_document):
    identifier = upload(client, raw_document)
    base = f"/api/documents/{identifier}"
    edit(client, base, 0, "현재 지켜야 할 문구")
    settings = {"instructions": "현재 설정 유지", "protected_terms": ["지켜야 할 문구"], "strict_validation": True}
    assert client.put(base + "/settings", json={"version": 0, "settings": settings}).status_code == 200
    assert client.post(base + "/restorations", json={"revision": 0, "source_revision": 0}).status_code == 409
    assert client.post(base + "/restorations", json={"revision": 1, "source_revision": 1}).status_code == 409
    assert client.post(base + "/restorations", json={"revision": 1, "source_revision": -1}).status_code == 422
    job = client.post(base + "/restorations", json={"revision": 1, "source_revision": 0})
    result = wait(client, job.json()["job_id"])
    assert result["status"] == "failed" and "보호 문구" in result["error"]
    fresh = client.get(base).json()
    assert fresh["revision"] == 1
    assert fresh["settings"] == settings
    assert len(client.get(base + "/revisions").json()["items"]) == 2


def test_history_owner_expiry_and_delete(client, raw_document):
    identifier = upload(client, raw_document)
    base = f"/api/documents/{identifier}"
    stranger = TestClient(client.app)
    try:
        assert stranger.get(base + "/revisions").status_code == 404
        assert stranger.get(base + "/revisions/0").status_code == 404
        assert stranger.post(base + "/restorations", json={"revision": 1, "source_revision": 0}).status_code == 404
    finally:
        stranger.close()
    assert client.get(base + "/revisions/99").status_code == 404
    assert client.get(base + "/revisions/-1").status_code == 404
    assert client.get(base + "/revisions?before=-1").status_code == 422
    store = client.app.state.store
    assert client.delete(base).status_code == 200
    with store.connect() as db:
        assert db.execute("SELECT count(*) FROM revisions WHERE document=?", (identifier,)).fetchone()[0] == 0
    identifier = upload(client, raw_document)
    with store.connect() as db:
        db.execute("UPDATE documents SET created=0 WHERE id=?", (identifier,))
    assert client.get(f"/api/documents/{identifier}/revisions").status_code == 404
    assert client.get(f"/api/documents/{identifier}/revisions/0").status_code == 404


def test_history_migration_and_cursor_only_include_published_versions(client, raw_document):
    identifier = upload(client, raw_document)
    base = f"/api/documents/{identifier}"
    edit(client, base, 0, "기존 저장본")
    store = client.app.state.store
    with store.connect() as db:
        db.execute("DROP TABLE revisions")
    migrated = Store(store.root)
    with store.connect() as db:
        owner = db.execute("SELECT owner FROM documents WHERE id=?", (identifier,)).fetchone()["owner"]
    assert [r["revision"] for r in migrated.revision_history(identifier, owner)] == [1, 0]
    with store.connect() as db:
        # Only the cursor behavior needs many rows; no large files are generated.
        db.execute("UPDATE documents SET revision=25 WHERE id=?", (identifier,))
        db.executemany("INSERT INTO revisions(document,revision,created,kind) VALUES(?,?,?,'edit')", [(identifier, n, time.time()) for n in range(2, 27)])
    first = client.get(base + "/revisions").json()
    assert [v["revision"] for v in first["items"]] == list(range(25, 5, -1))
    assert first["next_before"] == 6
    second = client.get(base + "/revisions?before=6").json()
    assert [v["revision"] for v in second["items"]] == list(range(5, -1, -1))
    assert second["next_before"] is None
    assert client.get(base + "/revisions/26").status_code == 404


def test_missing_history_snapshot_fails_restore_without_changing_current(client, raw_document):
    identifier = upload(client, raw_document)
    base = f"/api/documents/{identifier}"
    edit(client, base, 0, "현재 본문")
    directory = client.app.state.store.directory(identifier)
    (directory / "snapshot-0.json").write_text("{broken", encoding="utf-8")
    assert client.get(base + "/revisions/0").status_code == 503
    result = wait(client, client.post(base + "/restorations", json={"revision": 1, "source_revision": 0}).json()["job_id"])
    assert result["status"] == "failed"
    assert client.get(base).json()["revision"] == 1
    assert client.get(base + "/original").content == raw_document


def test_native_hwp_history_restore_preserves_original_bytes(tmp_path):
    doc = HwpDocument.blank()
    doc.append_paragraph("원본 본문")
    source = tmp_path / "revision-0.hwp"
    doc.save(source)
    raw = source.read_bytes()
    process(tmp_path, "analyze", {}, "native.hwp")
    snapshot = json.loads((tmp_path / "snapshot-0.json").read_text(encoding="utf-8"))
    p = next(p for p in snapshot["sections"][0]["paragraphs"] if p["editable"] and p["text"])
    process(tmp_path, "edit", {"revision": 0, "operations": [{"type": "paragraph", "section": 0, "index": p["index"], "text": "수정 본문"}]}, "native.hwp")
    result = process(tmp_path, "restore", {"revision": 1, "source_revision": 0, "settings": {}}, "native.hwp")
    assert result["revision"] == 2
    assert result["report"]["checked_revision"] == 2
    assert (tmp_path / "revision-2.hwp").read_bytes() == raw
    assert source.read_bytes() == raw


def test_converted_hwp_restore_keeps_conversion_metadata_and_original(client, tmp_path):
    doc = HwpDocument.blank()
    doc.append_paragraph("HWP에서 가져온 본문")
    path = tmp_path / "원본.hwp"
    doc.save(path)
    raw = path.read_bytes()
    uploaded = client.post("/api/documents", files={"file": (path.name, raw)}).json()
    assert wait(client, uploaded["job_id"])["status"] == "completed"
    base = f"/api/documents/{uploaded['document_id']}"
    converted = client.get(base + "/source?revision=0").content
    before = client.get(base).json()
    edit(client, base, 0, "변환 문서 수정")
    restored = client.post(base + "/restorations", json={"revision": 1, "source_revision": 0}).json()
    assert wait(client, restored["job_id"])["status"] == "completed"
    current = client.get(base).json()
    assert current["format"] == "hwpx" and current["revision"] == 2
    assert current["conversion"] == before["conversion"]
    assert current["format_note"] == before["format_note"]
    assert client.get(base + "/source?revision=2").content == converted
    assert client.get(base + "/original").content == raw
