"""Document use cases shared by the site's HTTP routes.

This module deliberately has no HTTP request, cookies, UI, or tool transport.
The caller must provide an authenticated owner; the store checks ownership and
expiry again inside the same transaction that accepts a command.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

from .errors import WorkspaceError
from .policies import DEFAULT_SETTINGS, operation_key
from .schemas import CommandRevisionRequest, EditRequest, RestoreRequest, PageLayoutRequest
from .storage import Store


class DocumentService:
    def __init__(self, store: Store):
        self.store = store

    def owned(self, identifier: str, owner: str):
        document = self.store.document(identifier, owner)
        if document is None:
            raise WorkspaceError("document_not_found", "문서를 찾을 수 없거나 보관 기간이 만료되었습니다.", 404)
        return document

    def view(self, document: dict):
        result = {key: document[key] for key in ("id", "name", "revision", "status", "validated", "created", "error")}
        result["expires"] = document["created"] + self.store.ttl
        if document["status"] == "ready":
            try:
                snapshot = json.loads((self.store.directory(document["id"]) / f"snapshot-{document['revision']}.json").read_text(encoding="utf-8"))
                if not isinstance(snapshot, dict) or snapshot.get("revision") != document["revision"]:
                    raise ValueError("Invalid snapshot revision")
                result.update(snapshot)
                result["name"] = document["name"]
            except (OSError, ValueError):
                raise WorkspaceError("snapshot_unavailable", "저장된 문서 정보를 읽지 못했습니다. 다시 시도하거나 업로드 원본을 다운로드해 주세요.", 503) from None
        result.setdefault("format", Path(document["name"]).suffix.lower().lstrip("."))
        result["settings"] = {**DEFAULT_SETTINGS, **json.loads(document["settings"])}
        result["settings_version"] = document["settings_version"]
        return result

    def state(self, identifier: str, owner: str):
        document = self.owned(identifier, owner)
        return {**{key: document[key] for key in ("id", "name", "revision", "status", "validated", "settings_version")},
                "expires": document["created"] + self.store.ttl}

    def _submit(self, identifier, owner, kind, body, request_key=None):
        self.owned(identifier, owner)
        payload = body.model_dump(exclude={"settings_version"})
        if kind == 'layout' and payload.get('background_pages') is None:
            payload.pop('background_pages', None)
        job = self.store.enqueue(identifier, owner, kind, payload, body.revision,
            request_key=request_key, settings_version=body.settings_version)
        return {"job_id": job, "document_id": identifier, "kind": kind, "base_revision": body.revision}

    def edit(self, identifier: str, owner: str, body: EditRequest, request_key: str | None = None):
        targets = [operation_key(op.model_dump()) for op in body.operations]
        if len(targets) != len(set(targets)):
            raise WorkspaceError("duplicate_target", "같은 요소를 한 작업에서 두 번 수정할 수 없습니다.", 422)
        return self._submit(identifier, owner, "edit", body, request_key)

    def validate(self, identifier: str, owner: str, body: CommandRevisionRequest, request_key: str | None = None):
        return self._submit(identifier, owner, "validate", body, request_key)

    def layout(self, identifier: str, owner: str, body: PageLayoutRequest, request_key: str | None = None):
        if body.background_pages and len(body.background_pages) > 1:
            raise WorkspaceError('multiple_backgrounds', '현재 문서에는 배경 한 쪽을 선택해 적용해 주세요. 여러 쪽 양식은 새 문서로 만들 수 있습니다.', 422)
        if self.view(self.owned(identifier, owner))["format"] != 'hwpx':
            raise WorkspaceError('conversion_required', 'HWPX로 변환한 문서를 다시 열어 주세요.')
        return self._submit(identifier, owner, 'layout', body, request_key)

    def restore(self, identifier: str, owner: str, body: RestoreRequest, request_key: str | None = None):
        document = self.owned(identifier, owner)
        if body.source_revision >= document["revision"] or self.store.revision(identifier, owner, body.source_revision) is None:
            raise WorkspaceError("restore_source_unavailable", "복원할 이전 버전을 다시 선택해 주세요.")
        return self._submit(identifier, owner, "restore", body, request_key)

    def save_studio(self, identifier: str, owner: str, revision: int, settings_version: int | None,
                    candidate: str, digest: str, request_key: str | None = None):
        # Candidate bytes are streamed and bounded by the HTTP adapter. Only the
        # generated filename reaches the worker; replay identity uses byte SHA256.
        if not re.fullmatch(r"studio-[a-f0-9]{32}\.hwpx", candidate) or not re.fullmatch(r"[a-f0-9]{64}", digest):
            raise WorkspaceError("invalid_candidate", "편집 파일 정보를 확인하지 못했습니다.", 422)
        if self.view(self.owned(identifier, owner))["format"] != "hwpx":
            raise WorkspaceError("conversion_required", "HWPX로 변환한 문서를 다시 열어 주세요.")
        job = self.store.enqueue(identifier, owner, "studio", {"revision": revision, "candidate": candidate}, revision,
            settings_version=settings_version, request_key=request_key, fingerprint={"sha256": digest})
        return {"job_id": job, "document_id": identifier, "kind": "studio", "base_revision": revision}
