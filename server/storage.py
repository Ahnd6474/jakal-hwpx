from __future__ import annotations

import json
import re
import secrets
import shutil
import sqlite3
import time
from contextlib import contextmanager
from pathlib import Path


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
            """)
            columns = {row["name"] for row in db.execute("PRAGMA table_info(documents)")}
            if "settings" not in columns:
                db.execute("ALTER TABLE documents ADD COLUMN settings TEXT NOT NULL DEFAULT '{}'")
            if "settings_version" not in columns:
                db.execute("ALTER TABLE documents ADD COLUMN settings_version INTEGER NOT NULL DEFAULT 0")

    @contextmanager
    def connect(self):
        db = sqlite3.connect(self.db, timeout=15)
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

    def create_document(self, owner: str, name: str, identifier: str):
        with self.connect() as db:
            if db.execute("SELECT count(*) FROM documents WHERE owner=? AND created>?", (owner, time.time() - self.ttl)).fetchone()[0] >= 20:
                raise ValueError("한 세션에는 최대 20개 문서를 보관할 수 있습니다.")
            db.execute("INSERT INTO documents(id,owner,name,created) VALUES(?,?,?,?)", (identifier, owner, name, time.time()))

    def enqueue(self, document: str, owner: str, kind: str, payload: dict, revision: int | None = None):
        job = secrets.token_hex(16)
        with self.connect() as db:
            db.execute("BEGIN IMMEDIATE")
            row = db.execute("SELECT * FROM documents WHERE id=? AND owner=?", (document, owner)).fetchone()
            if not row:
                raise ValueError("문서를 찾을 수 없습니다.")
            if revision is not None and (row["revision"] != revision or row["status"] != "ready"):
                raise ValueError("문서 버전이 변경되었습니다. 문서를 다시 열어 주세요.")
            if db.execute("SELECT 1 FROM jobs WHERE document=? AND status IN ('queued','running')", (document,)).fetchone():
                raise ValueError("이 문서의 이전 작업이 끝날 때까지 기다려 주세요.")
            payload = {**payload, "settings": json.loads(row["settings"])}
            db.execute("INSERT INTO jobs(id,document,owner,kind,created,payload) VALUES(?,?,?,?,?,?)", (job, document, owner, kind, time.time(), json.dumps(payload, ensure_ascii=False)))
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

    def job(self, identifier: str, owner: str):
        with self.connect() as db:
            row = db.execute("SELECT jobs.* FROM jobs JOIN documents ON documents.id=jobs.document WHERE jobs.id=? AND jobs.owner=? AND documents.created>?", (identifier, owner, time.time() - self.ttl)).fetchone()
        return dict(row) if row else None

    def claim(self):
        with self.connect() as db:
            db.execute("BEGIN IMMEDIATE")
            row = db.execute("SELECT * FROM jobs WHERE status='queued' ORDER BY created LIMIT 1").fetchone()
            if row:
                db.execute("UPDATE jobs SET status='running' WHERE id=?", (row["id"],))
        return dict(row) if row else None

    def finish(self, job: dict, result: dict | None, error: str | None):
        with self.connect() as db:
            db.execute("UPDATE jobs SET status=?,result=?,error=? WHERE id=?", ("failed" if error else "completed", json.dumps(result, ensure_ascii=False) if result else None, error, job["id"]))
            if error and job["kind"] == "analyze":
                db.execute("UPDATE documents SET status='failed',error=? WHERE id=?", (error, job["document"]))
            elif result:
                db.execute("UPDATE documents SET status='ready', revision=?,error=NULL WHERE id=?", (result["revision"], job["document"]))
                if job["kind"] == "validate":
                    db.execute("UPDATE documents SET validated=? WHERE id=?", (result["revision"], job["document"]))
                elif job["kind"] == "edit":
                    db.execute("UPDATE documents SET validated=NULL WHERE id=?", (job["document"],))

    def recover(self):
        # A stopped process never publishes a partial revision. Interrupted jobs
        # fail visibly; queued work remains durable and can run after restart.
        with self.connect() as db:
            db.execute("UPDATE documents SET status='failed',error='서버 재시작으로 분석이 중단되었습니다.' WHERE id IN (SELECT document FROM jobs WHERE status='running' AND kind='analyze')")
            db.execute("UPDATE jobs SET status='failed',error='서버 재시작으로 작업이 중단되었습니다. 다시 실행해 주세요.' WHERE status='running'")

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
            db.execute("DELETE FROM documents WHERE id=? AND owner=?", (identifier, owner))
        # A locked file can survive deletion on Windows; retain the tombstone
        # name so the periodic cleanup can retry without a document DB row.
        shutil.rmtree(tombstone, ignore_errors=True)
        return True

    def cleanup(self):
        with self.connect() as db:
            expired = db.execute("SELECT id,owner FROM documents WHERE created<?", (time.time() - self.ttl,)).fetchall()
        for row in expired:
            self.delete(row["id"], row["owner"])
        for path in self.root.glob("trash-*"):
            target = path.resolve()
            if target.is_relative_to(self.root) and target != self.root and re.fullmatch(r"trash-[a-f0-9]{32}", target.name):
                shutil.rmtree(target, ignore_errors=True)
