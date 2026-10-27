"""Read saved history and restore a previous file as a newly checked revision."""
from __future__ import annotations

import json
from pathlib import Path

from .files import atomic_json, atomic_write


def read_snapshot(directory: Path, revision: int) -> dict:
    if type(revision) is not int or revision < 0:
        raise ValueError("저장 버전을 찾을 수 없습니다.")
    try:
        value = json.loads((directory / f"snapshot-{revision}.json").read_text(encoding="utf-8"))
        if not isinstance(value, dict) or value.get("revision") != revision or not isinstance(value.get("sections"), list) or not isinstance(value.get("report"), dict):
            raise ValueError("Invalid saved snapshot")
        return value
    except (OSError, ValueError):
        raise ValueError("이 버전의 저장 정보를 읽지 못했습니다. 업로드 원본과 다른 버전은 그대로 보관됩니다.") from None


def process_restore(directory: Path, payload: dict, name: str) -> dict:
    current, source = payload["revision"], payload["source_revision"]
    if type(source) is not int or not 0 <= source < current:
        raise ValueError("복원할 이전 버전을 다시 선택해 주세요.")
    source_view = read_snapshot(directory, source)
    current_view = read_snapshot(directory, current)
    extension = source_view.get("format", Path(name).suffix.lower().lstrip("."))
    if extension != current_view.get("format", Path(name).suffix.lower().lstrip(".")):
        raise ValueError("문서 형식이 다른 버전은 복원할 수 없습니다.")
    if extension == "hwpx":
        from .studio import process_studio
        result = process_studio(directory, payload, name, source=directory / f"revision-{source}.hwpx")
    elif extension == "hwp":
        from jakal_hwpx import HwpDocument
        from .engine import new_issues
        from .hwp_engine import hwp_snapshot, inspect_hwp

        path = directory / f"revision-{source}.hwp"
        inspect_hwp(path)
        document = HwpDocument.open(path)
        issues = [item.to_dict() for item in document.strict_lint_errors()]
        baseline = read_snapshot(directory, 0)["report"]["issues"]
        if new_issues(issues, baseline):
            raise ValueError("이전 버전에 새로운 문서 오류가 있어 복원하지 않았습니다.")
        settings = payload.get("settings", {})
        if settings.get("strict_validation") and issues:
            raise ValueError("엄격 검증 설정에 따라 오류가 있는 이전 버전은 복원할 수 없습니다.")
        revision = current + 1
        report = {**source_view["report"], "issues": issues, "new_issues": [], "checked_revision": revision}
        view = hwp_snapshot(document, revision, report, name)

        def text(snapshot):
            return "\n".join([p["text"] for s in snapshot["sections"] for p in s["paragraphs"]] + [c["text"] for s in snapshot["sections"] for t in s["tables"] for c in t["cells"]])

        old, restored = text(current_view), text(view)
        for term in settings.get("protected_terms", []):
            if old.count(term) != restored.count(term):
                raise ValueError(f"보호 문구 ‘{term}’의 개수가 변경되어 복원하지 않았습니다.")
        atomic_write(directory / f"revision-{revision}.hwp", path.read_bytes())
        atomic_json(directory / f"snapshot-{revision}.json", view)
        result = {"revision": revision, "report": report}
    else:
        raise ValueError("이 버전의 문서 형식은 복원할 수 없습니다.")
    return {**result, "restored_from": source}
