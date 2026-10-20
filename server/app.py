from __future__ import annotations

import hashlib
import asyncio
import json
import os
import re
import secrets
import shutil
import time
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, Query, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from starlette.concurrency import run_in_threadpool
from .static import WebStaticFiles

from .schemas import EditRequest, RevisionRequest, RestoreRequest, SettingsRequest
from .policies import DEFAULT_SETTINGS
from .storage import Store
from .worker import Worker

REPO = Path(__file__).resolve().parents[1]
MAX_UPLOAD = 20 * 1024 * 1024
COOKIE = "jakal_session"


class BodyTooLarge(HTTPException):
    def __init__(self):
        super().__init__(413, "파일 크기는 20MB 이하여야 합니다.")


class BodyLimit:
    def __init__(self, app):
        self.app = app

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http":
            return await self.app(scope, receive, send)
        headers = dict(scope.get("headers", []))
        try:
            if int(headers.get(b"content-length", b"0")) > MAX_UPLOAD + 256 * 1024:
                return await JSONResponse({"detail": "파일 크기는 20MB 이하여야 합니다."}, status_code=413)(scope, receive, send)
        except ValueError:
            return await JSONResponse({"detail": "올바르지 않은 요청 크기입니다."}, status_code=400)(scope, receive, send)
        total = 0

        async def limited_receive():
            nonlocal total
            message = await receive()
            total += len(message.get("body", b""))
            if total > MAX_UPLOAD + 256 * 1024:
                raise BodyTooLarge()
            return message

        try:
            await self.app(scope, limited_receive, send)
        except BodyTooLarge:
            await JSONResponse({"detail": "파일 크기는 20MB 이하여야 합니다."}, status_code=413)(scope, receive, send)


def create_app(data_dir: Path | None = None, *, run_worker: bool = True) -> FastAPI:
    store = Store(data_dir or Path(os.environ.get("JAKAL_DATA_DIR", REPO / ".web-data")))
    worker = Worker(store)

    @asynccontextmanager
    async def lifespan(app):
        if run_worker:
            worker.start()
        try:
            yield
        finally:
            if run_worker:
                worker.close()

    app = FastAPI(title="Jakal Document Workspace", lifespan=lifespan, docs_url="/api/docs", redoc_url=None, openapi_url="/api/openapi.json")
    app.state.store = store
    app.state.worker = worker
    app.add_middleware(BodyLimit)

    @app.middleware("http")
    async def session_and_headers(request: Request, call_next):
        if request.method in {"POST", "PUT", "PATCH", "DELETE"}:
            origin = request.headers.get("origin")
            expected = f"{request.url.scheme}://{request.headers.get('host', '')}"
            if request.headers.get("sec-fetch-site") == "cross-site" or (origin and origin != expected):
                return JSONResponse({"detail": "다른 사이트에서 보낸 요청은 허용하지 않습니다."}, status_code=403)
        token = request.cookies.get(COOKIE, "")
        fresh = not re.fullmatch(r"[a-f0-9]{64}", token)
        if fresh:
            token = secrets.token_hex(32)
        request.state.owner = hashlib.sha256(token.encode()).hexdigest()
        response = await call_next(request)
        # Keep active sessions alive so a file uploaded near the end of an old
        # cookie's lifetime remains accessible for its own 24-hour retention.
        if not request.url.path.startswith(("/assets/", "/studio/assets/", "/fonts/", "/studio/fonts/")):
            response.set_cookie(COOKIE, token, httponly=True, samesite="strict", secure=os.environ.get("JAKAL_COOKIE_SECURE") == "1", max_age=store.ttl)
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["Referrer-Policy"] = "same-origin"
        response.headers["Content-Security-Policy"] = "default-src 'self'; script-src 'self' 'wasm-unsafe-eval'; worker-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; connect-src 'self'; object-src 'none'; frame-ancestors 'none'; base-uri 'self'; form-action 'self'"
        if request.url.path.startswith("/studio/"):
            # Only the self-hosted editor is embeddable, and only by this origin.
            response.headers["Content-Security-Policy"] = "default-src 'self'; script-src 'self' 'wasm-unsafe-eval'; worker-src 'self' blob:; style-src 'self' 'unsafe-inline'; font-src 'self' data:; img-src 'self' data: blob:; connect-src 'self'; frame-src 'self' blob:; object-src 'none'; frame-ancestors 'self'; base-uri 'self'; form-action 'self'"
            response.headers["Permissions-Policy"] = "camera=(), microphone=(), geolocation=()"
        if request.url.path.startswith("/api/"):
            response.headers["Cache-Control"] = "no-store"
        return response

    def owned(identifier: str, request: Request):
        document = store.document(identifier, request.state.owner)
        if document is None:
            raise HTTPException(404, "문서를 찾을 수 없거나 보관 기간이 만료되었습니다.")
        return document

    def view(document: dict):
        result = {key: document[key] for key in ("id", "name", "revision", "status", "validated", "created", "error")}
        result["expires"] = document["created"] + store.ttl
        if document["status"] == "ready":
            try:
                snapshot = json.loads((store.directory(document["id"]) / f"snapshot-{document['revision']}.json").read_text(encoding="utf-8"))
                if not isinstance(snapshot, dict) or snapshot.get("revision") != document["revision"]:
                    raise ValueError("Invalid snapshot revision")
                result.update(snapshot)
            except (OSError, ValueError):
                raise HTTPException(503, "저장된 문서 정보를 읽지 못했습니다. 다시 시도하거나 업로드 원본을 다운로드해 주세요.") from None
        result.setdefault("format", Path(document["name"]).suffix.lower().lstrip("."))
        result["settings"] = {**DEFAULT_SETTINGS, **json.loads(document["settings"])}
        result["settings_version"] = document["settings_version"]
        return result

    def enqueue(identifier, request, kind, payload, revision):
        owned(identifier, request)
        try:
            job = store.enqueue(identifier, request.state.owner, kind, payload, revision)
        except ValueError as exc:
            raise HTTPException(409, str(exc)) from exc
        return {"job_id": job, "document_id": identifier}

    @app.get("/api/health")
    def health():
        return {"status": "ok", "max_upload_mb": 20, "retention_hours": 24}

    @app.get("/api/documents")
    def documents(request: Request):
        return [{key: d[key] for key in ("id", "name", "status", "revision", "created")} for d in store.list_documents(request.state.owner)]

    @app.post("/api/documents", status_code=202)
    def upload(request: Request, file: UploadFile = File(...)):
        name = (file.filename or "document.hwpx").replace("\\", "/").split("/")[-1]
        name = "".join(c for c in name if c.isprintable())
        extension = Path(name).suffix.lower()
        if extension not in {".hwpx", ".hwp"}:
            raise HTTPException(415, ".hwp 또는 .hwpx 파일을 업로드해 주세요.")
        name = name[:180 - len(extension)] + name[-len(extension):] if len(name) > 180 else name
        identifier = secrets.token_hex(16)
        directory = store.directory(identifier)
        directory.mkdir()
        try:
            size = 0
            with (directory / f"revision-0{extension}").open("wb") as output:
                while chunk := file.file.read(1024 * 1024):
                    size += len(chunk)
                    if size > MAX_UPLOAD:
                        raise HTTPException(413, "파일 크기는 20MB 이하여야 합니다.")
                    output.write(chunk)
            if not size:
                raise HTTPException(400, "빈 파일은 업로드할 수 없습니다.")
            if extension == ".hwp":
                (directory / "convert-hwp").touch()
            job = store.create_document(request.state.owner, name, identifier, enqueue_analysis=True)
        except Exception as exc:
            shutil.rmtree(directory, ignore_errors=True)
            if isinstance(exc, ValueError):
                raise HTTPException(429, str(exc)) from exc
            raise
        finally:
            file.file.close()
        return {"document_id": identifier, "job_id": job}

    @app.get("/api/documents/{identifier}")
    def get_document(identifier: str, request: Request):
        return view(owned(identifier, request))

    @app.get("/api/documents/{identifier}/elements")
    def get_elements(identifier: str, request: Request):
        document = view(owned(identifier, request))
        return {"revision": document["revision"], "sections": document.get("sections", [])}

    @app.get("/api/documents/{identifier}/revisions")
    def revisions(identifier: str, request: Request, before: int | None = Query(default=None, ge=0)):
        document = owned(identifier, request)
        items = store.revision_history(identifier, request.state.owner, before, limit=21)
        return {"items": items[:20], "next_before": items[19]["revision"] if len(items) > 20 else None, "current_revision": document["revision"]}

    @app.get("/api/documents/{identifier}/revisions/{revision}")
    def revision_details(identifier: str, revision: int, request: Request):
        owned(identifier, request)
        if store.revision(identifier, request.state.owner, revision) is None:
            raise HTTPException(404, "저장 버전을 찾을 수 없습니다.")
        from .revisions import read_snapshot
        try:
            return read_snapshot(store.directory(identifier), revision)
        except ValueError as exc:
            raise HTTPException(503, str(exc)) from None

    @app.post("/api/documents/{identifier}/restorations", status_code=202)
    def restore(identifier: str, body: RestoreRequest, request: Request):
        document = owned(identifier, request)
        if body.source_revision >= document["revision"] or store.revision(identifier, request.state.owner, body.source_revision) is None:
            raise HTTPException(409, "복원할 이전 버전을 다시 선택해 주세요.")
        return enqueue(identifier, request, "restore", body.model_dump(), body.revision)

    @app.get("/api/documents/{identifier}/source")
    def source(identifier: str, revision: int, request: Request):
        document = owned(identifier, request)
        if document["status"] != "ready" or revision not in {0, document["revision"]}:
            raise HTTPException(409, "원문을 열 수 없습니다. 문서의 현재 버전을 다시 확인해 주세요.")
        extension = "." + view(document)["format"]
        filename = f"{Path(document['name']).stem}{'_수정' if revision else ''}{extension}"
        return FileResponse(store.directory(identifier) / f"revision-{revision}{extension}", media_type="application/octet-stream", filename=filename)

    @app.get("/api/documents/{identifier}/original")
    def original(identifier: str, request: Request):
        document = owned(identifier, request)
        extension = Path(document["name"]).suffix.lower()
        return FileResponse(store.directory(identifier) / f"revision-0{extension}", media_type="application/octet-stream", filename=document["name"])

    @app.post("/api/documents/{identifier}/operations", status_code=202)
    def edit(identifier: str, body: EditRequest, request: Request):
        targets = [json.dumps({k: v for k, v in op.model_dump().items() if k != "text"}, sort_keys=True) for op in body.operations]
        if len(targets) != len(set(targets)):
            raise HTTPException(422, "같은 요소를 한 작업에서 두 번 수정할 수 없습니다.")
        return enqueue(identifier, request, "edit", body.model_dump(), body.revision)

    @app.post("/api/documents/{identifier}/validations", status_code=202)
    def validate(identifier: str, body: RevisionRequest, request: Request):
        return enqueue(identifier, request, "validate", body.model_dump(), body.revision)

    @app.post("/api/documents/{identifier}/studio", status_code=202)
    def studio_save(identifier: str, request: Request, revision: int = Query(ge=0), file: UploadFile = File(...)):
        document = owned(identifier, request)
        if view(document).get("format") != "hwpx":
            raise HTTPException(409, "HWPX로 변환한 문서를 다시 열어 주세요.")
        candidate = f"studio-{secrets.token_hex(16)}.hwpx"
        path = store.directory(identifier) / candidate
        try:
            size = 0
            with path.open("xb") as output:
                while chunk := file.file.read(64 * 1024):
                    size += len(chunk)
                    if size > MAX_UPLOAD:
                        raise BodyTooLarge()
                    output.write(chunk)
            if not size:
                raise HTTPException(422, "빈 편집 파일은 저장할 수 없습니다.")
            return enqueue(identifier, request, "studio", {"revision": revision, "candidate": candidate}, revision)
        except Exception:
            path.unlink(missing_ok=True)
            raise
        finally:
            file.file.close()

    @app.put("/api/documents/{identifier}/settings")
    def update_settings(identifier: str, body: SettingsRequest, request: Request):
        owned(identifier, request)
        settings = body.settings.model_dump()
        settings["protected_terms"] = list(dict.fromkeys(t.strip() for t in settings["protected_terms"] if t.strip()))
        try:
            version = store.save_settings(identifier, request.state.owner, body.version, settings)
        except ValueError as exc:
            raise HTTPException(409, str(exc)) from exc
        return {"settings": settings, "settings_version": version}

    @app.post("/api/documents/{identifier}/exports")
    def export(identifier: str, body: RevisionRequest, request: Request):
        document = owned(identifier, request)
        if document["revision"] != body.revision or document["validated"] != body.revision:
            raise HTTPException(409, "현재 수정 버전을 먼저 검증해 주세요.")
        return {"url": f"/api/documents/{identifier}/download?revision={body.revision}"}

    @app.get("/api/jobs/{identifier}")
    async def get_job(identifier: str, request: Request, wait_ms: int = Query(default=0, ge=0, le=1500)):
        job = await run_in_threadpool(store.job, identifier, request.state.owner)
        if job is None:
            raise HTTPException(404, "작업을 찾을 수 없습니다.")
        deadline = time.monotonic() + wait_ms / 1000
        while job["status"] in {"queued", "running"} and time.monotonic() < deadline:
            if await request.is_disconnected():
                break
            await asyncio.sleep(min(.075, max(0, deadline - time.monotonic())))
            job = await run_in_threadpool(store.job, identifier, request.state.owner)
            if job is None:
                raise HTTPException(404, "작업을 찾을 수 없습니다.")
        return {"id": job["id"], "document_id": job["document"], "kind": job["kind"], "status": job["status"], "error": job["error"], "result": json.loads(job["result"]) if job["result"] else None}

    @app.get("/api/documents/{identifier}/download")
    def download(identifier: str, revision: int, request: Request):
        document = owned(identifier, request)
        if revision != document["revision"] or revision != document["validated"]:
            raise HTTPException(409, "현재 수정 버전을 검증한 후 다운로드해 주세요.")
        snapshot = view(document)
        extension = "." + snapshot["format"]
        if json.loads(document["settings"]).get("strict_validation") and snapshot["report"]["issues"]:
            raise HTTPException(409, "엄격 검증 설정에서는 오류가 없는 문서만 다운로드할 수 있습니다.")
        return FileResponse(store.directory(identifier) / f"revision-{revision}{extension}", media_type="application/x-hwp" if extension == ".hwp" else "application/hwp+zip", filename=f"{Path(document['name']).stem}_수정{extension}")

    @app.get("/api/documents/{identifier}/report")
    def report(identifier: str, request: Request):
        document = owned(identifier, request)
        snapshot = view(document)
        if "report" not in snapshot:
            raise HTTPException(409, "문서 분석이 아직 끝나지 않았습니다.")
        return JSONResponse(snapshot["report"], headers={"Content-Disposition": 'attachment; filename="validation-report.json"'})

    @app.get("/api/documents/{identifier}/text")
    def text_export(identifier: str, request: Request):
        snapshot = view(owned(identifier, request))
        if "sections" not in snapshot:
            raise HTTPException(409, "문서 분석이 아직 끝나지 않았습니다.")
        from starlette.responses import Response
        lines = []
        for section in snapshot["sections"]:
            lines.extend(p["text"] for p in section["paragraphs"])
            for table in section["tables"]:
                for row in range(table["rows"]):
                    lines.append("\t".join(c["text"] for c in table["cells"] if c["row"] == row))
        return Response("\n".join(lines), media_type="text/plain; charset=utf-8", headers={"Content-Disposition": 'attachment; filename="document.txt"'})

    @app.delete("/api/documents/{identifier}")
    def delete(identifier: str, request: Request):
        owned(identifier, request)
        if not store.delete(identifier, request.state.owner):
            raise HTTPException(409, "실행 중인 작업이나 파일 전송이 끝난 후 삭제해 주세요.")
        return {"deleted": True}

    dist = REPO / "web" / "dist"
    if dist.exists():
        app.mount("/", WebStaticFiles(directory=dist, html=True), name="web")
    else:
        @app.get("/")
        def missing_frontend():
            return JSONResponse({"detail": "프런트엔드 빌드가 필요합니다. web 폴더에서 npm install 후 npm run build를 실행해 주세요."}, status_code=503)
    return app
