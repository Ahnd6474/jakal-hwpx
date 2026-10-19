"""Shared editing rules for manual and batch edits."""
from __future__ import annotations

DEFAULT_SETTINGS = {"instructions": "", "protected_terms": [], "strict_validation": False}


def check_text(before: str, after: str, settings: dict) -> None:
    if len(after) > 100_000 or any((ord(c) < 32 and c not in "\n\t") or 0xD800 <= ord(c) <= 0xDFFF for c in after):
        raise ValueError("텍스트 길이 또는 제어 문자가 편집 범위를 벗어났습니다.")
    for term in settings.get("protected_terms", []):
        if before.count(term) != after.count(term):
            raise ValueError(f"보호 문구 ‘{term}’의 개수가 변경됩니다. 고급 설정을 확인해 주세요.")


def targets(snapshot: dict) -> dict[str, dict]:
    result = {}
    for section in snapshot["sections"]:
        for p in section["paragraphs"]:
            result[p["id"]] = {"id": p["id"], "text": p["text"], "editable": p["editable"], "label": f"섹션 {p['section'] + 1} · 문단 {p['index'] + 1}", "operation": {"type": "paragraph", "section": p["section"], "index": p["index"]}}
        for table in section["tables"]:
            for cell in table["cells"]:
                key = f"{table['id']}-{cell['row']}-{cell['column']}"
                result[key] = {"id": key, "text": cell["text"], "editable": cell["editable"], "label": f"섹션 {section['index'] + 1} · 표 {table['index'] + 1} · {cell['row'] + 1}행 {cell['column'] + 1}열", "operation": {"type": "cell", "section": section["index"], "table": table["index"], "row": cell["row"], "column": cell["column"]}}
    return result


def operation_key(op: dict) -> str:
    if op["type"] == "paragraph":
        return f"s{op['section']}-p{op['index']}"
    return f"s{op['section']}-t{op['table']}-{op['row']}-{op['column']}"


def check_operations(snapshot: dict, operations: list[dict], settings: dict) -> None:
    items = targets(snapshot)
    for op in operations:
        item = items.get(operation_key(op))
        if item is None or not item["editable"]:
            raise ValueError("읽기 전용이거나 존재하지 않는 문단 또는 셀입니다.")
        check_text(item["text"], op["text"], settings)
