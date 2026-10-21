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

from fastapi import FastAPI, File, Header, HTTPException, Query, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from starlette.concurrency import run_in_threadpool
from .static import WebStaticFiles
from .http_policy import response_headers

from .schemas import CommandRevisionRequest, CopyRequest, RenameRequest, EditRequest, RevisionRequest, RestoreRequest, SettingsRequest, TemplateRequest
from .document_names import document_name
from .document_service import DocumentService
from .errors import WorkspaceError
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
            length = int(headers.get(b"content-length", b"0"))
            if length < 0:
                raise ValueError("negative content length")
            if length > MAX_UPLOAD + 256 * 1024:
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
    service = DocumentService(store)

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
    app.state.documents = service
    app.add_middleware(BodyLimit)

    @app.exception_handler(WorkspaceError)
    async def workspace_error(request: Request, exc: WorkspaceError):
        return JSONResponse({"detail": str(exc), "code": exc.code}, status_code=exc.status,
            headers=response_headers(request.url.path))

    @app.exception_handler(Exception)
    async def unexpected_error(request: Request, exc: Exception):
        # Never expose server paths, document contents or tracebacks to the client.
        return JSONResponse({"detail": "서버에서 요청을 처리하지 못했습니다. 잠시 후 다시 확인해 주세요."}, status_code=500, headers=response_headers(request.url.path))

    @app.middleware("http")
    async def session_and_headers(request: Request, call_next):
        if request.method in {"POST", "PUT", "PATCH", "DELETE"}:
            origin = request.headers.get("origin")
            expected = f"{request.url.scheme}://{request.headers.get('host', '')}"
            if request.headers.get("sec-fetch-site") == "cross-site" or (origin and origin != expected):
                return JSONResponse({"detail": "다른 사이트에서 보낸 요청은 허용하지 않습니다."}, status_code=403, headers=response_headers(request.url.path))
        token = request.cookies.get(COOKIE, "")
        fresh = not re.fullmatch(r"[a-f0-9]{64}", token)
        if fresh:
            token = secrets.token_hex(32)
        request.state.owner = hashlib.sha256(token.encode()).hexdigest()
        response = await call_next(request)
        # Keep active sessions alive so a file uploaded near the end of an old
        # cookie's lifetime remains accessible for its own 24-hour retention.
        if not request.url.path.startswith(("/assets/", "/studio/assets/", "/fonts/", "/studio/fonts/")):
            response.set_cookie(COOKIE, token, httponly=True, samesite="strict", secure=request.url.scheme == "https" or os.environ.get("JAKAL_COOKIE_SECURE") == "1", max_age=store.ttl)
        response.headers.update(response_headers(request.url.path))
        return response

    def owned(identifier: str, request: Request):
        return service.owned(identifier, request.state.owner)

    view = service.view

    @app.get("/api/health")
    def health():
        return {"status": "ok", "max_upload_mb": 20, "retention_hours": 24}

    @app.get("/api/documents")
    def documents(request: Request):
        return [{key: d[key] for key in ("id", "name", "status", "revision", "created")} for d in store.list_documents(request.state.owner)]

    @app.get("/api/templates")
    def templates():
        from .templates import TEMPLATES
        return TEMPLATES

    @app.post("/api/documents/from-template", status_code=202)
    def from_template(body: TemplateRequest, request: Request):
        from .templates import make_template
        from .files import atomic_write
        title = "".join(c for c in body.title.strip() if c.isprintable() and c not in '/\\:*?"<>|').strip('. ')
        if not title:
            raise HTTPException(422, "문서 제목을 입력해 주세요.")
        raw = make_template(body.template, title, body.question_count)
        identifier = secrets.token_hex(16)
        directory = store.directory(identifier)
        directory.mkdir()
        try:
            atomic_write(directory / "revision-0.hwpx", raw)
            job = store.create_document(request.state.owner, title + ".hwpx", identifier, enqueue_analysis=True)
        except Exception as exc:
            shutil.rmtree(directory, ignore_errors=True)
            if isinstance(exc, ValueError):
                raise HTTPException(429, str(exc)) from exc
            raise
        return {"document_id": identifier, "job_id": job}

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

    @app.get("/api/documents/{identifier}/state")
    def document_state(identifier: str, request: Request):
        return service.state(identifier, request.state.owner)

    @app.patch('/api/documents/{identifier}/name')
    def rename_document(identifier: str, body: RenameRequest, request: Request):
        document = owned(identifier, request)
        try:
            name = document_name(body.name, Path(document['name']).suffix.lower().lstrip('.'))
        except ValueError as exc:
            raise HTTPException(422, str(exc)) from None
        try:
            store.rename_document(identifier, request.state.owner, body.revision, body.expected_name, name)
        except ValueError as exc:
            raise HTTPException(409, str(exc)) from None
        return view(owned(identifier, request))

    @app.post('/api/documents/{identifier}/copies', status_code=202)
    def copy_document(identifier: str, body: CopyRequest, request: Request):
        document = owned(identifier, request)
        if view(document)['format'] != 'hwpx':
            raise HTTPException(409, 'HWPX로 변환된 문서를 다시 열어 주세요.')
        try:
            name = document_name(body.name, 'hwpx')
        except ValueError as exc:
            raise HTTPException(422, str(exc)) from None
        target = secrets.token_hex(16)
        try:
            job = store.copy_document(identifier, request.state.owner, body.revision, target, name)
        except ValueError as exc:
            raise HTTPException(409, str(exc)) from None
        except OSError:
            raise HTTPException(503, '복사본을 저장하지 못했습니다. 원본은 유지됩니다. 저장 공간을 확인해 주세요.') from None
        return {'document_id': target, 'job_id': job}

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
    def restore(identifier: str, body: RestoreRequest, request: Request,
                request_key: str | None = Header(default=None, alias="Idempotency-Key")):
        return service.restore(identifier, request.state.owner, body, request_key)

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
    def edit(identifier: str, body: EditRequest, request: Request,
             request_key: str | None = Header(default=None, alias="Idempotency-Key")):
        return service.edit(identifier, request.state.owner, body, request_key)

    @app.post("/api/documents/{identifier}/validations", status_code=202)
    def validate(identifier: str, body: CommandRevisionRequest, request: Request,
                 request_key: str | None = Header(default=None, alias="Idempotency-Key")):
        return service.validate(identifier, request.state.owner, body, request_key)

    @app.post("/api/documents/{identifier}/studio", status_code=202)
    def studio_save(identifier: str, request: Request, revision: int = Query(ge=0), file: UploadFile = File(...),
                    settings_version: int | None = Query(default=None, ge=0),
                    request_key: str | None = Header(default=None, alias="Idempotency-Key")):
        owned(identifier, request)
        candidate = f"studio-{secrets.token_hex(16)}.hwpx"
        path = store.directory(identifier) / candidate
        try:
            size = 0
            digest = hashlib.sha256()
            with path.open("xb") as output:
                while chunk := file.file.read(64 * 1024):
                    size += len(chunk)
                    if size > MAX_UPLOAD:
                        raise BodyTooLarge()
                    output.write(chunk)
                    digest.update(chunk)
            if not size:
                raise HTTPException(422, "빈 편집 파일은 저장할 수 없습니다.")
            result = service.save_studio(identifier, request.state.owner, revision, settings_version,
                candidate, digest.hexdigest(), request_key)
            accepted = store.job(result["job_id"], request.state.owner)
            if accepted and json.loads(accepted["payload"]).get("candidate") != candidate:
                # A repeated upload got the first job. Keep its input intact and
                # remove only this retry's unused file, even if it already ran.
                path.unlink(missing_ok=True)
            return result
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
