from __future__ import annotations

import json
import sqlite3
import threading
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pytest

from server.storage import Store
from server.worker import Worker
from server.files import atomic_write


def queued_job(tmp_path, *, kind="analyze", payload=None):
    store = Store(tmp_path)
    identifier = "a" * 32
    store.directory(identifier).mkdir()
    store.create_document("owner", "test.hwpx", identifier)
    if kind != "analyze":
        with store.connect() as db:
            db.execute("UPDATE documents SET status='ready' WHERE id=?", (identifier,))
    store.enqueue(identifier, "owner", kind, payload or {})
    return store, store.claim()


def test_concurrent_uploads_cannot_exceed_session_quota(tmp_path):
    store = Store(tmp_path)
    for n in range(19):
        store.create_document("owner", "test.hwpx", f"{n:032x}")
    barrier = threading.Barrier(8)

    def upload(n):
        barrier.wait(timeout=5)
        try:
            store.create_document("owner", "test.hwpx", f"{n + 100:032x}")
            return True
        except ValueError:
            return False

    with ThreadPoolExecutor(max_workers=8) as pool:
        accepted = list(pool.map(upload, range(8)))
    assert sum(accepted) == 1
    assert len(store.list_documents("owner")) == 20


def test_request_write_failure_releases_document_for_retry(tmp_path, monkeypatch):
    store, job = queued_job(tmp_path)
    original = Path.write_text

    def fail_request(path, *args, **kwargs):
        if path.name.startswith("job-"):
            raise OSError("disk unavailable")
        return original(path, *args, **kwargs)

    monkeypatch.setattr(Path, "write_text", fail_request)
    Worker(store).run(job)
    assert store.job(job["id"], "owner")["status"] == "failed"
    assert store.document(job["document"], "owner")["status"] == "failed"
    assert store.enqueue(job["document"], "owner", "analyze", {})


@pytest.mark.parametrize("outcome", [{}, {"result": {"revision": 0}}, {"error": ""}])
def test_incomplete_process_output_cannot_complete_job(tmp_path, monkeypatch, outcome):
    store, job = queued_job(tmp_path)

    class Process:
        returncode = 0

        def __init__(self, args, **kwargs):
            Path(args[-1]).write_text(json.dumps(outcome), encoding="utf-8")

        def poll(self):
            return self.returncode

        def wait(self, **kwargs):
            return self.returncode

    monkeypatch.setattr("server.worker.subprocess.Popen", Process)
    Worker(store).run(job)
    assert store.job(job["id"], "owner")["status"] == "failed"
    assert store.document(job["document"], "owner")["status"] == "failed"
    assert not list(tmp_path.glob("*/job-*.json"))
    assert not list(tmp_path.glob("*/result-*.json"))


def test_recovered_job_cannot_publish_late_result(tmp_path):
    store, job = queued_job(tmp_path)
    store.recover()
    assert store.finish(job, {"revision": 9}, None) is False
    assert store.job(job["id"], "owner")["status"] == "failed"
    assert store.document(job["document"], "owner")["revision"] == 0


def test_expired_document_cannot_accept_new_work(tmp_path):
    store = Store(tmp_path)
    store.create_document("owner", "test.hwpx", "b" * 32)
    with store.connect() as db:
        db.execute("UPDATE documents SET created=0")
    with pytest.raises(ValueError, match="만료|찾을"):
        store.enqueue("b" * 32, "owner", "analyze", {})


def test_initial_job_failure_rolls_back_document_registration(tmp_path):
    store = Store(tmp_path)
    with store.connect() as db:
        db.executescript("CREATE TRIGGER fail_job BEFORE INSERT ON jobs BEGIN SELECT RAISE(ABORT, 'test failure'); END;")
    with pytest.raises(sqlite3.IntegrityError):
        store.create_document("owner", "test.hwpx", "a" * 32, enqueue_analysis=True)
    assert store.list_documents("owner") == []


def stage_result(store, job, revision):
    directory = store.directory(job["document"])
    report = {"checked_revision": revision, "issues": [], "new_issues": []}
    result = {"revision": revision, "report": report}
    (directory / f"revision-{revision}.hwpx").write_bytes(b"validated by engine")
    (directory / f"snapshot-{revision}.json").write_text(json.dumps({**result, "sections": [], "format": "hwpx"}), encoding="utf-8")
    return result


@pytest.mark.parametrize("damage", ["missing_file", "missing_snapshot", "empty_file", "wrong_report", "wrong_revision", "bad_format"])
def test_result_files_must_match_before_revision_is_published(tmp_path, damage):
    store, job = queued_job(tmp_path, kind="edit", payload={"revision": 0})
    with store.connect() as db:
        db.execute("UPDATE documents SET validated=0")
    result = stage_result(store, job, 1)
    directory = store.directory(job["document"])
    if damage == "missing_file":
        (directory / "revision-1.hwpx").unlink()
    elif damage == "missing_snapshot":
        (directory / "snapshot-1.json").unlink()
    elif damage == "empty_file":
        (directory / "revision-1.hwpx").write_bytes(b"")
    elif damage == "wrong_report":
        result["report"]["issues"] = ["different report"]
    elif damage == "wrong_revision":
        result["revision"] = 4
    else:
        snapshot = json.loads((directory / "snapshot-1.json").read_text())
        snapshot["format"] = "../other"
        (directory / "snapshot-1.json").write_text(json.dumps(snapshot))
    assert store.finish(job, result, None)
    assert store.job(job["id"], "owner")["status"] == "failed"
    document = store.document(job["document"], "owner")
    assert (document["status"], document["revision"], document["validated"]) == ("ready", 0, 0)


def test_successful_completion_is_idempotent(tmp_path):
    store, job = queued_job(tmp_path, kind="edit", payload={"revision": 0})
    result = stage_result(store, job, 1)
    assert store.finish(job, result, None)
    assert store.finish(job, None, "late failure") is False
    assert store.job(job["id"], "owner")["status"] == "completed"
    assert store.document(job["document"], "owner")["revision"] == 1


def test_stale_result_never_overwrites_a_newer_revision(tmp_path):
    store, job = queued_job(tmp_path, kind="edit", payload={"revision": 0})
    result = stage_result(store, job, 1)
    with store.connect() as db:
        db.execute("UPDATE documents SET revision=3, validated=3")
    store.finish(job, result, None)
    assert store.job(job["id"], "owner")["status"] == "failed"
    document = store.document(job["document"], "owner")
    assert (document["revision"], document["validated"]) == (3, 3)


def test_failed_atomic_replace_preserves_previous_file(tmp_path, monkeypatch):
    target = tmp_path / "snapshot.json"
    target.write_bytes(b"previous")

    def fail_replace(*args, **kwargs):
        raise OSError("file locked")

    monkeypatch.setattr(Path, "replace", fail_replace)
    with pytest.raises(OSError):
        atomic_write(target, b"incomplete replacement")
    assert target.read_bytes() == b"previous"
    assert not (tmp_path / "snapshot.json.tmp").exists()


def test_cleanup_restores_delete_rolled_back_by_crash(tmp_path):
    store, job = queued_job(tmp_path)
    store.finish(job, None, "test")
    directory = store.directory(job["document"])
    (directory / "revision-0.hwpx").write_bytes(b"immutable original")
    # Simulate a crash after the filesystem rename, before the DB commit.
    tombstone = tmp_path / f"trash-{job['document']}"
    directory.rename(tombstone)
    store.cleanup()
    assert (directory / "revision-0.hwpx").read_bytes() == b"immutable original"
    assert store.document(job["document"], "owner") is not None
    assert not tombstone.exists()


def test_restart_cleans_interrupted_files_but_keeps_queued_candidate(tmp_path, monkeypatch):
    candidate = "studio-" + "b" * 32 + ".hwpx"
    store, job = queued_job(tmp_path, kind="studio", payload={"revision": 0, "candidate": candidate})
    directory = store.directory(job["document"])
    for name in (candidate, f"job-{job['id']}.json", "revision-1.hwpx.tmp", "snapshot-1.json.tmp"):
        (directory / name).write_bytes(b"interrupted")
    other = "c" * 32
    store.create_document("owner", "queued.hwpx", other)
    store.directory(other).mkdir()
    (store.directory(other) / candidate).write_bytes(b"queued")
    queued = store.enqueue(other, "owner", "studio", {"candidate": candidate})
    worker = Worker(store)
    monkeypatch.setattr(worker.thread, "start", lambda: None)
    worker.start()
    assert store.job(job["id"], "owner")["status"] == "failed"
    assert list(directory.iterdir()) == []
    assert store.job(queued, "owner")["status"] == "queued"
    assert (store.directory(other) / candidate).read_bytes() == b"queued"


def test_job_cleanup_runs_even_when_database_finish_fails(tmp_path, monkeypatch):
    candidate = "studio-" + "b" * 32 + ".hwpx"
    store, job = queued_job(tmp_path, kind="studio", payload={"revision": 0, "candidate": candidate})
    path = store.directory(job["document"]) / candidate
    path.write_bytes(b"candidate")

    def fail(*args, **kwargs):
        raise OSError("unavailable")

    monkeypatch.setattr("server.worker.subprocess.Popen", fail)
    monkeypatch.setattr(store, "finish", fail)
    with pytest.raises(OSError):
        Worker(store).run(job)
    assert not path.exists()
    assert not list(tmp_path.glob("*/job-*.json"))


def test_slow_job_lookup_does_not_block_other_requests(tmp_path, monkeypatch):
    from fastapi.testclient import TestClient
    from server.app import create_app

    app = create_app(tmp_path, run_worker=False)
    entered, release = threading.Event(), threading.Event()

    def slow_lookup(*args):
        entered.set()
        assert release.wait(timeout=5)
        return None

    monkeypatch.setattr(app.state.store, "job", slow_lookup)
    with TestClient(app) as client, ThreadPoolExecutor(max_workers=2) as pool:
        # Establish a session before issuing concurrent requests.
        assert client.get("/api/health").status_code == 200
        pending = pool.submit(client.get, "/api/jobs/" + "a" * 32)
        assert entered.wait(timeout=2)
        health = pool.submit(client.get, "/api/health")
        try:
            assert health.result(timeout=2).status_code == 200
        finally:
            release.set()
        assert pending.result(timeout=2).status_code == 404


def test_worker_retries_terminal_state_after_temporary_database_failure(tmp_path, monkeypatch):
    store, job = queued_job(tmp_path)
    worker = Worker(store)
    real_finish = store.finish
    attempts = []
    claims = iter([job, None])

    def flaky_finish(*args, **kwargs):
        attempts.append(1)
        if len(attempts) <= 2:
            raise sqlite3.OperationalError("temporarily locked")
        result = real_finish(*args, **kwargs)
        worker.stop.set()
        return result

    def run(value):
        worker.store.finish(value, None, "test failure")

    monkeypatch.setattr(store, "claim", lambda: next(claims))
    monkeypatch.setattr(store, "finish", flaky_finish)
    monkeypatch.setattr(worker, "run", run)
    worker.loop()
    assert len(attempts) == 3
    assert store.job(job["id"], "owner")["status"] == "failed"
