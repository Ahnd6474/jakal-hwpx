from __future__ import annotations

import json
import hashlib
import re
import secrets
import shutil
import sqlite3
import time
from contextlib import contextmanager
from pathlib import Path

from .errors import WorkspaceError


class Store:
    """Durable single-host queue. No document objects are kept in API memory."""

    def __init__(self, root: Path, ttl: int = 86_400):
        self.root = root.resolve()
        self.ttl = ttl
        self.root.mkdir(parents=True, exist_ok=True)
        self.db = self.root / "app.sqlite3"
        with self.connect() as db:
            db.executescript("""
                PRAGMA journal_mode=WAL;
                CREATE TABLE IF NOT EXISTS documents (
                    id TEXT PRIMARY KEY, owner TEXT NOT NULL, name TEXT NOT NULL,
                    created REAL NOT NULL, revision INTEGER NOT NULL DEFAULT 0,
                    status TEXT NOT NULL DEFAULT 'processing', validated INTEGER,
                    error TEXT
                );
                CREATE TABLE IF NOT EXISTS jobs (
                    id TEXT PRIMARY KEY, document TEXT NOT NULL, owner TEXT NOT NULL,
                    kind TEXT NOT NULL, status TEXT NOT NULL DEFAULT 'queued',
                    created REAL NOT NULL, payload TEXT NOT NULL,
                    error TEXT, result TEXT
                );
                CREATE INDEX IF NOT EXISTS jobs_status ON jobs(status, created);
                CREATE INDEX IF NOT EXISTS jobs_document_status ON jobs(document, status);
                CREATE INDEX IF NOT EXISTS documents_owner_created ON documents(owner, created);
                CREATE TABLE IF NOT EXISTS revisions (
                    document TEXT NOT NULL, revision INTEGER NOT NULL,
                    created REAL NOT NULL, kind TEXT NOT NULL,
                    source_revision INTEGER,
                    PRIMARY KEY(document, revision)
                );
            """)
            columns = {row["name"] for row in db.execute("PRAGMA table_info(documents)")}
            if "settings" not in columns:
                db.execute("ALTER TABLE documents ADD COLUMN settings TEXT NOT NULL DEFAULT '{}'")
            if "settings_version" not in columns:
                db.execute("ALTER TABLE documents ADD COLUMN settings_version INTEGER NOT NULL DEFAULT 0")
            job_columns = {row["name"] for row in db.execute("PRAGMA table_info(jobs)")}
            for column in ("request_key", "request_hash"):
                if column not in job_columns:
                    db.execute(f"ALTER TABLE jobs ADD COLUMN {column} TEXT")
            db.execute("CREATE UNIQUE INDEX IF NOT EXISTS jobs_request_key ON jobs(document,owner,request_key) WHERE request_key IS NOT NULL")
            # Existing files and completed jobs already contain the history.
            # Backfill idempotently so an interrupted migration is safe to retry.
            db.execute("INSERT OR IGNORE INTO revisions(document,revision,created,kind) SELECT id,0,created,'analyze' FROM documents WHERE status='ready'")
            db.execute("""INSERT OR IGNORE INTO revisions(document,revision,created,kind,source_revision)
                SELECT jobs.document,json_extract(jobs.result,'$.revision'),jobs.created,jobs.kind,
                    CASE WHEN jobs.kind='restore' THEN json_extract(jobs.payload,'$.source_revision') END
                FROM jobs JOIN documents ON documents.id=jobs.document
                WHERE jobs.status='completed' AND jobs.kind IN ('edit','studio','restore','layout')
                    AND json_valid(jobs.result) AND json_type(jobs.result,'$.revision')='integer'
                    AND json_extract(jobs.result,'$.revision')<=documents.revision""")

    @contextmanager
    def connect(self):
        db = sqlite3.connect(self.db, timeout=15)
        db.execute('PRAGMA secure_delete=ON')
        db.row_factory = sqlite3.Row
        try:
            with db:
                yield db
        finally:
            db.close()

    def directory(self, document: str) -> Path:
        if len(document) != 32 or any(c not in "0123456789abcdef" for c in document):
            raise ValueError("Invalid document identifier")
        result = (self.root / document).resolve()
        if not result.is_relative_to(self.root) or result == self.root:
            raise ValueError("Document path escapes storage directory")
        return result

    def document(self, identifier: str, owner: str):
        with self.connect() as db:
            row = db.execute("SELECT * FROM documents WHERE id=? AND owner=? AND created>?", (identifier, owner, time.time() - self.ttl)).fetchone()
        return dict(row) if row else None

    def list_documents(self, owner: str):
        with self.connect() as db:
            return [dict(row) for row in db.execute("SELECT * FROM documents WHERE owner=? AND created>? ORDER BY created DESC", (owner, time.time() - self.ttl))]

    def create_document(self, owner: str, name: str, identifier: str, *, enqueue_analysis: bool = False, analysis_payload: dict | None = None):
        with self.connect() as db:
            db.execute("BEGIN IMMEDIATE")
            if db.execute("SELECT count(*) FROM documents WHERE owner=? AND created>?", (owner, time.time() - self.ttl)).fetchone()[0] >= 20:
                raise ValueError("한 세션에는 최대 20개 문서를 보관할 수 있습니다.")
            db.execute("INSERT INTO documents(id,owner,name,created) VALUES(?,?,?,?)", (identifier, owner, name, time.time()))
            if enqueue_analysis:
                # Publish the upload and its first job in one transaction. An
                # insertion failure must not leave an unprocessable document.
                job = secrets.token_hex(16)
                db.execute("INSERT INTO jobs(id,document,owner,kind,created,payload) VALUES(?,?,?,'analyze',?,?)", (job, identifier, owner, time.time(), json.dumps({'settings': {}, **(analysis_payload or {})}, ensure_ascii=False)))
                return job

    def enqueue(self, document: str, owner: str, kind: str, payload: dict, revision: int | None = None,
                *, request_key: str | None = None, fingerprint: dict | None = None,
                settings_version: int | None = None, accept_guard=None):
        if request_key is not None and not re.fullmatch(r"[A-Za-z0-9_-]{16,128}", request_key):
            raise WorkspaceError("invalid_request_key", "작업 식별자가 올바르지 않습니다.", 422)
        # Hash the submitted intent, not mutable database settings or temporary
        # candidate filenames. Replays must also work after the revision advances.
        request_hash = hashlib.sha256(json.dumps({"kind": kind, "revision": revision,
            "settings_version": settings_version, "payload": fingerprint if fingerprint is not None else payload},
            sort_keys=True, ensure_ascii=True, separators=(",", ":")).encode()).hexdigest() if request_key else None
        job = secrets.token_hex(16)
        with self.connect() as db:
            db.execute("BEGIN IMMEDIATE")
            row = db.execute("SELECT * FROM documents WHERE id=? AND owner=? AND created>?", (document, owner, time.time() - self.ttl)).fetchone()
            if not row:
                raise WorkspaceError("document_not_found", "문서를 찾을 수 없습니다.", 404)
            if request_key:
                previous = db.execute("SELECT id,request_hash FROM jobs WHERE document=? AND owner=? AND request_key=?", (document, owner, request_key)).fetchone()
                if previous:
                    if previous["request_hash"] != request_hash:
                        raise WorkspaceError("request_key_conflict", "이미 사용한 작업 식별자로 다른 내용을 요청할 수 없습니다.")
                    if accept_guard is not None:
                        accept_guard(db, row, previous["id"])
                    return previous["id"]
            if revision is not None and (row["revision"] != revision or row["status"] != "ready"):
                raise WorkspaceError("revision_conflict", "문서 버전이 변경되었습니다. 편집 내용을 보관한 뒤 최신 문서를 다시 열어 주세요.")
            if settings_version is not None and row["settings_version"] != settings_version:
                raise WorkspaceError("settings_conflict", "고급 설정이 변경되었습니다. 편집 내용을 보관한 뒤 최신 설정을 확인해 주세요.")
            if db.execute("SELECT 1 FROM jobs WHERE document=? AND status IN ('queued','running')", (document,)).fetchone():
                raise WorkspaceError("document_busy", "이 문서의 이전 작업이 끝날 때까지 기다려 주세요.")
            # The database revision is authoritative, including for internal
            # callers which omit the HTTP request's revision field.
            payload = {**payload, "revision": row["revision"], "settings": json.loads(row["settings"])}
            if accept_guard is not None:
                accept_guard(db, row, job)
            db.execute("INSERT INTO jobs(id,document,owner,kind,created,payload,request_key,request_hash) VALUES(?,?,?,?,?,?,?,?)", (job, document, owner, kind, time.time(), json.dumps(payload, ensure_ascii=False), request_key, request_hash))
        return job

    def save_settings(self, document: str, owner: str, version: int, settings: dict):
        with self.connect() as db:
            db.execute("BEGIN IMMEDIATE")
            row = db.execute("SELECT * FROM documents WHERE id=? AND owner=? AND created>?", (document, owner, time.time() - self.ttl)).fetchone()
            if not row or row["settings_version"] != version:
                raise ValueError("설정이 변경되었습니다. 문서를 다시 열어 주세요.")
            if db.execute("SELECT 1 FROM jobs WHERE document=? AND status IN ('queued','running')", (document,)).fetchone():
                raise ValueError("문서 작업이 끝난 뒤 설정을 저장해 주세요.")
            db.execute("UPDATE documents SET settings=?,settings_version=settings_version+1,validated=CASE WHEN ? THEN NULL ELSE validated END WHERE id=?", (json.dumps(settings, ensure_ascii=False), settings.get("strict_validation", False) != json.loads(row["settings"]).get("strict_validation", False), document))
        return version + 1

    def rename_document(self, document: str, owner: str, revision: int, expected_name: str, name: str):
        with self.connect() as db:
            db.execute('BEGIN IMMEDIATE')
            row = db.execute('SELECT * FROM documents WHERE id=? AND owner=? AND created>?', (document, owner, time.time() - self.ttl)).fetchone()
            if not row or row['revision'] != revision or row['name'] != expected_name or row['status'] != 'ready':
                raise ValueError('문서 이름이나 버전이 변경되었습니다. 최신 문서를 다시 확인해 주세요.')
            if db.execute("SELECT 1 FROM jobs WHERE document=? AND status IN ('queued','running')", (document,)).fetchone():
                raise ValueError('문서 작업이 끝난 뒤 이름을 변경해 주세요.')
            db.execute('UPDATE documents SET name=? WHERE id=?', (name, document))

    def copy_document(self, document: str, owner: str, revision: int, identifier: str, name: str):
        destination = self.directory(identifier)
        created = False
        try:
            with self.connect() as db:
                # Keep source deletion, revision changes and the session quota
                # fenced until the immutable revision has been copied.
                db.execute('BEGIN IMMEDIATE')
                source = db.execute('SELECT * FROM documents WHERE id=? AND owner=? AND created>?', (document, owner, time.time() - self.ttl)).fetchone()
                if not source or source['revision'] != revision or source['status'] != 'ready':
                    raise ValueError('문서 버전이 변경되었습니다. 최신 문서를 다시 확인해 주세요.')
                if db.execute("SELECT 1 FROM jobs WHERE document=? AND status IN ('queued','running')", (document,)).fetchone():
                    raise ValueError('문서 작업이 끝난 뒤 복사해 주세요.')
                if db.execute('SELECT count(*) FROM documents WHERE owner=? AND created>?', (owner, time.time() - self.ttl)).fetchone()[0] >= 20:
                    raise ValueError('한 세션에는 최대 20개 문서를 보관할 수 있습니다.')
                destination.mkdir()
                created = True
                shutil.copyfile(self.directory(document) / f'revision-{revision}.hwpx', destination / 'revision-0.hwpx')
                now = time.time()
                db.execute('INSERT INTO documents(id,owner,name,created,settings) VALUES(?,?,?,?,?)', (identifier, owner, name, now, source['settings']))
                job = secrets.token_hex(16)
                db.execute("INSERT INTO jobs(id,document,owner,kind,created,payload) VALUES(?,?,?,'analyze',?,?)", (job, identifier, owner, now, json.dumps({'settings': json.loads(source['settings'])})))
            return job
        except Exception:
            if created:
                shutil.rmtree(destination, ignore_errors=True)
            raise

    def job(self, identifier: str, owner: str):
        with self.connect() as db:
            row = db.execute("SELECT jobs.* FROM jobs JOIN documents ON documents.id=jobs.document WHERE jobs.id=? AND jobs.owner=? AND documents.owner=? AND documents.created>?", (identifier, owner, owner, time.time() - self.ttl)).fetchone()
        return dict(row) if row else None

    def revision_history(self, document: str, owner: str, before: int | None = None, limit: int = 20):
        with self.connect() as db:
            return [dict(row) for row in db.execute("""SELECT revisions.* FROM revisions
                JOIN documents ON documents.id=revisions.document
                WHERE documents.id=? AND documents.owner=? AND documents.created>?
                    AND revisions.revision<=documents.revision AND (? IS NULL OR revisions.revision<?)
                ORDER BY revisions.revision DESC LIMIT ?""", (document, owner, time.time() - self.ttl, before, before, limit))]

    def revision(self, document: str, owner: str, revision: int):
        with self.connect() as db:
            row = db.execute("""SELECT revisions.* FROM revisions JOIN documents ON documents.id=revisions.document
                WHERE documents.id=? AND documents.owner=? AND documents.created>?
                    AND revisions.revision=? AND revisions.revision<=documents.revision""", (document, owner, time.time() - self.ttl, revision)).fetchone()
        return dict(row) if row else None

    def claim(self):
        with self.connect() as db:
            db.execute("BEGIN IMMEDIATE")
            row = db.execute("SELECT * FROM jobs WHERE status='queued' ORDER BY created LIMIT 1").fetchone()
            if row:
                db.execute("UPDATE jobs SET status='running' WHERE id=?", (row["id"],))
        return {**dict(row), "status": "running"} if row else None

    def _completion_error(self, job, document, result):
        if not document or document["created"] <= time.time() - self.ttl:
            return "문서 보관 기간이 만료되었습니다."
        try:
            base = json.loads(job["payload"]).get("revision", 0)
            expected = base + (job["kind"] in {"edit", "studio", "restore", "layout"})
            if document["revision"] != base or (job["kind"] != "analyze" and document["status"] != "ready"):
                return "문서 버전이 변경되어 작업 결과를 반영하지 않았습니다."
            if not isinstance(result, dict) or type(result.get("revision")) is not int or result["revision"] != expected:
                return "문서 처리 결과의 버전이 올바르지 않습니다. 다시 시도해 주세요."
            directory = self.directory(document["id"])
            snapshot = json.loads((directory / f"snapshot-{expected}.json").read_text(encoding="utf-8"))
            report = result.get("report")
            extension = snapshot.get("format", Path(document["name"]).suffix.lower().lstrip("."))
            if (snapshot.get("revision") != expected or not isinstance(snapshot.get("sections"), list)
                    or not isinstance(report, dict) or report.get("checked_revision") != expected
                    or snapshot.get("report") != report or extension not in {"hwp", "hwpx"}
                    or (directory / f"revision-{expected}.{extension}").stat().st_size == 0):
                return "문서 처리 결과 파일이 완전하지 않습니다. 다시 시도해 주세요."
        except (OSError, ValueError, TypeError, AttributeError):
            return "문서 처리 결과 파일을 확인하지 못했습니다. 다시 시도해 주세요."
        return None

    def finish(self, job: dict, result: dict | None, error: str | None):
        with self.connect() as db:
            db.execute("BEGIN IMMEDIATE")
            # Ignore duplicate/late completion after shutdown recovery, deletion
            # or another terminal transition. Use persisted job data throughout.
            current = db.execute("SELECT * FROM jobs WHERE id=? AND document=? AND owner=? AND status='running'", (job["id"], job["document"], job["owner"])).fetchone()
            if not current:
                return False
            job = dict(current)
            document = db.execute("SELECT * FROM documents WHERE id=? AND owner=?", (job["document"], job["owner"])).fetchone()
            error = error or self._completion_error(job, document, result)
            if error:
                result = None
            db.execute("UPDATE jobs SET status=?,result=?,error=? WHERE id=?", ("failed" if error else "completed", json.dumps(result, ensure_ascii=False) if result else None, error, job["id"]))
            if error and job["kind"] == "analyze":
                db.execute("UPDATE documents SET status='failed',error=? WHERE id=?", (error, job["document"]))
            elif result:
                db.execute("UPDATE documents SET status='ready', revision=?,error=NULL WHERE id=?", (result["revision"], job["document"]))
                if job["kind"] == "validate":
                    db.execute("UPDATE documents SET validated=? WHERE id=?", (result["revision"], job["document"]))
                elif job["kind"] in {"edit", "studio", "restore", "layout"}:
                    db.execute("UPDATE documents SET validated=NULL WHERE id=?", (job["document"],))
                if job["kind"] in {"analyze", "edit", "studio", "restore", "layout"}:
                    source = json.loads(job["payload"]).get("source_revision") if job["kind"] == "restore" else None
                    db.execute("INSERT INTO revisions(document,revision,created,kind,source_revision) VALUES(?,?,?,?,?)", (job["document"], result["revision"], time.time(), job["kind"], source))
        return True

    def recover(self):
        # A stopped process never publishes a partial revision. Interrupted jobs
        # fail visibly; queued work remains durable and can run after restart.
        with self.connect() as db:
            db.execute("BEGIN IMMEDIATE")
            interrupted = [dict(row) for row in db.execute("SELECT * FROM jobs WHERE status='running'")]
            db.execute("UPDATE documents SET status='failed',error='서버 재시작으로 분석이 중단되었습니다.' WHERE id IN (SELECT document FROM jobs WHERE status='running' AND kind='analyze')")
            db.execute("UPDATE jobs SET status='failed',error='서버 재시작으로 작업이 중단되었습니다. 다시 실행해 주세요.' WHERE status='running'")
        return interrupted

    def delete(self, identifier: str, owner: str) -> bool:
        tombstone = (self.root / f"trash-{identifier}").resolve()
        if not tombstone.is_relative_to(self.root) or tombstone == self.root:
            raise ValueError("Invalid trash path")
        with self.connect() as db:
            db.execute("BEGIN IMMEDIATE")
            if not db.execute("SELECT 1 FROM documents WHERE id=? AND owner=?", (identifier, owner)).fetchone():
                return False
            if db.execute("SELECT 1 FROM jobs WHERE document=? AND owner=? AND status='running'", (identifier, owner)).fetchone():
                return False
            # Rename before removing the DB row. If a Windows file transfer locks
            # the directory, leave the document intact and let the user retry.
            source = self.directory(identifier)
            if source.exists():
                try:
                    source.rename(tombstone)
                except OSError:
                    return False
            db.execute("DELETE FROM jobs WHERE document=? AND owner=?", (identifier, owner))
            db.execute("DELETE FROM revisions WHERE document=?", (identifier,))
            db.execute("DELETE FROM documents WHERE id=? AND owner=?", (identifier, owner))
            if db.execute("SELECT 1 FROM sqlite_master WHERE name='connections'").fetchone():
                for table in ('connection_proposals', 'connection_events', 'editor_presence', 'connections'):
                    db.execute(f'DELETE FROM {table} WHERE document=?', (identifier,))
        # A locked file can survive deletion on Windows; retain the tombstone
        # name so the periodic cleanup can retry without a document DB row.
        shutil.rmtree(tombstone, ignore_errors=True)
        return True

    def cleanup(self):
        with self.connect() as db:
            if db.execute("SELECT 1 FROM sqlite_master WHERE name='connections'").fetchone():
                db.execute('DELETE FROM editor_presence WHERE expires<?', (time.time(),))
                db.execute('DELETE FROM connection_proposals WHERE document IN (SELECT id FROM documents WHERE created<?)', (time.time() - self.ttl,))
            expired = db.execute("SELECT id,owner FROM documents WHERE created<?", (time.time() - self.ttl,)).fetchall()
            # Keep expired/revoked records until all their possible documents
            # have expired, so an old credential can never be migrated again.
            if db.execute("SELECT 1 FROM sqlite_master WHERE name='sessions'").fetchone():
                db.execute('DELETE FROM sessions WHERE expires<?', (time.time() - self.ttl,))
        for row in expired:
            self.delete(row["id"], row["owner"])
        for path in self.root.glob("trash-*"):
            target = path.resolve()
            if target.is_relative_to(self.root) and target != self.root and re.fullmatch(r"trash-[a-f0-9]{32}", target.name):
                with self.connect() as db:
                    db.execute("BEGIN IMMEDIATE")
                    document = db.execute("SELECT id FROM documents WHERE id=?", (target.name[6:],)).fetchone()
                    if document:
                        # A crash between directory rename and SQLite commit
                        # rolls back deletion. Restore its files, never purge them.
                        source = self.directory(document["id"])
                        if not source.exists():
                            try:
                                target.rename(source)
                            except OSError:
                                pass
                        continue
                shutil.rmtree(target, ignore_errors=True)
