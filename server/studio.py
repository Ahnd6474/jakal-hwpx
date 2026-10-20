"""Validate a complete editor revision without replacing the immutable source."""
from __future__ import annotations

import json
import re
from pathlib import Path

from jakal_hwpx import HwpxDocument
from jakal_hwpx.namespaces import NS
from .files import atomic_json, atomic_write


def process_studio(directory: Path, payload: dict, name: str, *, source: Path | None = None) -> dict:
    from .engine import inspect_archive, issues, new_issues, paragraph_text, snapshot

    revision = payload["revision"]
    if source is None:
        candidate = payload["candidate"]
        if not re.fullmatch(r"studio-[a-f0-9]{32}\.hwpx", candidate):
            raise ValueError("올바르지 않은 편집 파일입니다.")
        path = directory / candidate
    else:
        path = source
    inspect_archive(path)
    document = HwpxDocument.open(path, validate=False)
    # Every source revision has already passed the protected-document gate.
    # Reopen it only when a rule needs its text, instead of parsing it on all saves.
    if document.is_distribution_protected:
        raise ValueError("보호 문서는 편집할 수 없습니다.")
    baseline = json.loads((directory / "snapshot-0.json").read_text(encoding="utf-8"))["report"]["issues"]
    current = issues(document)
    added = new_issues(current, baseline)
    if added:
        raise ValueError("편집 결과에 새로운 문서 오류가 있어 저장하지 않았습니다: " + added[0]["message"])
    settings = payload.get("settings", {})
    if settings.get("strict_validation") and current:
        raise ValueError("엄격 검증 설정에 따라 기존 오류가 있는 문서의 저장을 중단했습니다.")

    def all_text(doc):
        return "\n".join(paragraph_text(p) for section in doc.sections for p in section.root_element.xpath(".//hp:p", namespaces=NS))

    # Whole-document edits may insert/delete paragraphs and objects deliberately.
    # Enforce protected phrases over all body/table text, independent of indices.
    if settings.get("protected_terms"):
        before = HwpxDocument.open(directory / f"revision-{revision}.hwpx", validate=False)
        old, new = all_text(before), all_text(document)
        for term in settings["protected_terms"]:
            if old.count(term) != new.count(term):
                raise ValueError(f"보호 문구 ‘{term}’의 개수가 변경됩니다. 고급 설정을 확인해 주세요.")
    revision += 1
    report = {"status": "warning" if current else "passed", "issues": current, "new_issues": [], "baseline_count": len(baseline), "resolved_count": len(new_issues(baseline, current)), "repairs": [], "checks": ["XML 구조", "문서 참조", "편집 파일 다시 열기", "보호 문구"], "checked_revision": revision}
    view = snapshot(document, revision, report, name)
    previous = json.loads((directory / f"snapshot-{revision - 1}.json").read_text(encoding="utf-8"))
    for key in ("format", "conversion", "format_note"):
        if key in previous:
            view[key] = previous[key]
    # Keep the exact submitted ZIP; inspection must not rewrite unrelated parts.
    raw = path.read_bytes()
    atomic_write(directory / f"revision-{revision}.hwpx", raw)
    atomic_json(directory / f"snapshot-{revision}.json", view)
    return {"revision": revision, "report": report}
