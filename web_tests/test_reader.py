from fastapi.testclient import TestClient

from web_tests.test_api import client, raw_document, upload, wait


def test_reader_source_is_owned_immutable_and_revision_scoped(client, raw_document):
    identifier = upload(client, raw_document)
    base = f"/api/documents/{identifier}"
    original = client.get(base + "/source?revision=0")
    assert original.status_code == 200 and original.content == raw_document
    assert client.get(base + "/original").content == raw_document
    assert original.headers["cache-control"] == "no-store"
    assert "wasm-unsafe-eval" in original.headers["content-security-policy"]
    stranger = TestClient(client.app)
    try:
        assert stranger.get(base + "/source?revision=0").status_code == 404
        assert stranger.get(base + "/original").status_code == 404
    finally:
        stranger.close()
    assert client.get(base + "/source?revision=999").status_code == 409
    assert client.get(base + "/source?revision=-1").status_code == 409
    op = {"type": "paragraph", "section": 0, "index": 0, "text": "원문과 구분할 저장본"}
    job = client.post(base + "/operations", json={"revision": 0, "operations": [op]}).json()
    assert wait(client, job["job_id"])["status"] == "completed"
    assert client.get(base + "/source?revision=0").content == raw_document
    saved = client.get(base + "/source?revision=1")
    assert saved.status_code == 200 and saved.content != raw_document
    assert client.get(base + "/download?revision=1").status_code == 409
