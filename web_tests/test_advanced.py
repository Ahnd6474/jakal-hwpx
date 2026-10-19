from __future__ import annotations

import json
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from jakal_hwpx import HwpDocument, HwpxDocument
from server.hwp_engine import process_hwp
from web_tests.test_api import client, raw_document, upload, wait


def test_native_hwp_lifecycle(client, tmp_path):
    source = tmp_path / "원본.HWP"
    document = HwpDocument.blank()
    document.append_paragraph("편집할 본문")
    document.append_table(rows=2, cols=2, cell_texts=[["항목", "값"], ["이름", "자칼"]], section_index=0)
    document.save(source)
    original = source.read_bytes()
    uploaded = client.post("/api/documents", files={"file": (source.name, original)}).json()
    assert wait(client, uploaded["job_id"])["status"] == "completed"
    base = f"/api/documents/{uploaded['document_id']}"
    doc = client.get(base).json()
    assert doc["format"] == "hwp"
    assert "HWP 5" in doc["format_note"]
    p = next(p for p in doc["sections"][0]["paragraphs"] if p["text"] == "편집할 본문")
    assert p["editable"]
    # A no-op native export must preserve the entire CFB container exactly.
    assert wait(client, client.post(base + "/validations", json={"revision": 0}).json()["job_id"])["status"] == "completed"
    assert client.get(base + "/download?revision=0").content == original
    operations = [{"type": "paragraph", "section": 0, "index": p["index"], "text": "길이가 달라진 본문 😀 & <안내>"}, {"type": "cell", "section": 0, "table": 0, "row": 1, "column": 1, "text": "홍길동\n두 번째 줄 😀"}]
    job = client.post(base + "/operations", json={"revision": 0, "operations": operations})
    result = wait(client, job.json()["job_id"])
    assert result["status"] == "completed", result
    assert client.get(base + "/download?revision=1").status_code == 409
    assert wait(client, client.post(base + "/validations", json={"revision": 1}).json()["job_id"])["status"] == "completed"
    response = client.get(client.post(base + "/exports", json={"revision": 1}).json()["url"])
    assert response.headers["content-type"] == "application/x-hwp"
    assert ".hwp" in response.headers["content-disposition"]
    destination = tmp_path / "수정.hwp"
    destination.write_bytes(response.content)
    reopened = HwpDocument.open(destination)
    assert reopened.strict_lint_errors() == []
    fresh = client.get(base).json()
    assert next(p for p in fresh["sections"][0]["paragraphs"] if p["index"] == operations[0]["index"])["text"] == operations[0]["text"]
    assert fresh["sections"][0]["tables"][0]["cells"][3]["text"] == operations[1]["text"]
    store = client.app.state.store
    assert (store.directory(doc["id"]) / "revision-0.hwp").read_bytes() == original


@pytest.mark.parametrize("name", ["기출킬러양식.hwp", "generated_native_control_authoring.hwp"])
def test_real_hwp_noop_and_controls_preserved(tmp_path, name):
    raw = (Path(__file__).resolve().parents[1] / "examples" / "samples" / "hwp" / name).read_bytes()
    (tmp_path / "revision-0.hwp").write_bytes(raw)
    result = process_hwp(tmp_path, "analyze", {}, name)
    assert result["revision"] == 0
    result = process_hwp(tmp_path, "validate", {"revision": 0}, name)
    assert result["report"]["new_issues"] == []
    assert (tmp_path / "revision-0.hwp").read_bytes() == raw
    if name == "generated_native_control_authoring.hwp":
        from server.policies import targets
        snapshot = json.loads((tmp_path / "snapshot-0.json").read_text(encoding="utf-8"))
        target = next(t for t in targets(snapshot).values() if t["editable"] and t["text"])
        process_hwp(tmp_path, "edit", {"revision": 0, "operations": [{**target["operation"], "text": "실제 문서 본문 편집 😀"}]}, name)
        result = process_hwp(tmp_path, "validate", {"revision": 1}, name)
        assert not result["report"]["new_issues"]


def test_settings_persist_conflict_owner_and_protection(client, raw_document):
    identifier = upload(client, raw_document)
    base = f"/api/documents/{identifier}"
    settings = {"instructions": "금액을 유지하고 합니다체로 작성", "protected_terms": ["원본", "원본"], "strict_validation": False}
    assert client.put(base + "/settings", json={"version": 0, "settings": settings}).status_code == 200
    doc = client.get(base).json()
    assert doc["settings"]["protected_terms"] == ["원본"]
    assert doc["settings_version"] == 1 and doc["revision"] == 0
    assert client.put(base + "/settings", json={"version": 0, "settings": settings}).status_code == 409
    stranger = TestClient(client.app)
    try:
        assert stranger.put(base + "/settings", json={"version": 1, "settings": settings}).status_code == 404
    finally:
        stranger.close()
    op = {"type": "paragraph", "section": 0, "index": 0, "text": "고친 문단"}
    result = wait(client, client.post(base + "/operations", json={"revision": 0, "operations": [op]}).json()["job_id"])
    assert result["status"] == "failed" and "보호 문구" in result["error"]
    assert client.get(base).json()["revision"] == 0
    op["text"] = "원본 문단을 보존했습니다."
    assert wait(client, client.post(base + "/operations", json={"revision": 0, "operations": [op]}).json()["job_id"])["status"] == "completed"


def test_strict_validation_blocks_existing_errors_and_revokes_download(client):
    document = HwpxDocument.blank()
    document.sections[0].paragraphs()[0].element.set("paraPrIDRef", "99999")
    identifier = upload(client, document.compile(validate=False, auto_repair=False))
    base = f"/api/documents/{identifier}"
    doc = client.get(base).json()
    assert doc["report"]["issues"]
    assert wait(client, client.post(base + "/validations", json={"revision": 0}).json()["job_id"])["status"] == "completed"
    assert client.get(base + "/download?revision=0").status_code == 200
    settings = {"strict_validation": True}
    assert client.put(base + "/settings", json={"version": 0, "settings": settings}).status_code == 200
    assert client.get(base + "/download?revision=0").status_code == 409
    result = wait(client, client.post(base + "/validations", json={"revision": 0}).json()["job_id"])
    assert result["status"] == "failed" and "엄격 검증" in result["error"]
    assert client.get(base).json()["revision"] == 0
    assert client.put(base + "/settings", json={"version": 1, "settings": {"strict_validation": False}}).status_code == 200
    assert wait(client, client.post(base + "/validations", json={"revision": 0}).json()["job_id"])["status"] == "completed"

