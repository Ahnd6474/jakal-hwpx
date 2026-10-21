"""Allowlisted document operations, executed only inside the worker process."""
from __future__ import annotations

import hashlib
import json
import zipfile
from collections import Counter
from pathlib import Path, PurePosixPath

from lxml import etree

from jakal_hwpx import HwpxDocument
from jakal_hwpx.namespaces import NS
from jakal_hwpx.parts import XmlPart
from jakal_hwpx.elements import _replace_paragraph_text_preserving_controls
from .files import atomic_json, atomic_write

MAX_UPLOAD = 20 * 1024 * 1024
MAX_EXPANDED = 128 * 1024 * 1024
MAX_PART = 24 * 1024 * 1024
MAX_ELEMENTS = 20_000
_PARAGRAPH_TEXTS = etree.XPath(".//hp:t[ancestor::hp:p and count(ancestor::hp:p[1] | $paragraph) = 1]", namespaces={"hp": NS["hp"]})
_PARAGRAPHS = etree.XPath("descendant-or-self::hp:p", namespaces={"hp": NS["hp"]})
_RUNS = etree.XPath("./hp:run", namespaces={"hp": NS["hp"]})
_CELL_PARAGRAPHS = etree.XPath("./hp:subList/hp:p", namespaces={"hp": NS["hp"]})
_HP = "{" + NS["hp"] + "}"
_INLINE_TEXT = {_HP + "lineBreak": "\n", _HP + "tab": "\t", _HP + "nbSpace": "\u00a0", _HP + "fwSpace": "\u3000", _HP + "hyphen": "\u00ad"}


def _text_content(element) -> str:
    parts = [element.text or ""]
    for child in element:
        parts.extend((_INLINE_TEXT[child.tag] if child.tag in _INLINE_TEXT else _text_content(child), child.tail or ""))
    return "".join(parts)


def inspect_archive(path: Path) -> None:
    if path.stat().st_size > MAX_UPLOAD:
        raise ValueError("파일 크기는 20MB 이하여야 합니다.")
    with zipfile.ZipFile(path) as archive:
        entries = archive.infolist()
        if len(entries) > 4096 or sum(i.file_size for i in entries) > MAX_EXPANDED:
            raise ValueError("압축 해제 크기 또는 내부 파일 수 제한을 초과했습니다.")
        names = set()
        for item in entries:
            name = item.filename
            if name in names or PurePosixPath(name).as_posix() != name.rstrip("/") or "\\" in name or ":" in name or name.startswith("/") or ".." in PurePosixPath(name).parts:
                raise ValueError("중복되거나 안전하지 않은 내부 파일 경로입니다.")
            names.add(name)
            if item.flag_bits & 1 or item.file_size > MAX_PART:
                raise ValueError("암호화되거나 너무 큰 내부 파일이 포함되어 있습니다.")
            if item.file_size > max(1, item.compress_size) * 1000:
                raise ValueError("내부 파일의 압축률 제한을 초과했습니다.")
        required = {"mimetype", "Contents/content.hpf", "Contents/header.xml"}
        if not required.issubset(names) or archive.read("mimetype").strip() != b"application/hwp+zip":
            raise ValueError("올바른 HWPX 문서가 아닙니다.")
        # Validate all XML before the engine's permissive large-document parser.
        for item in entries:
            with archive.open(item) as stream:
                if item.filename.endswith((".xml", ".hpf", ".rels", ".rdf")):
                    parser = etree.XMLParser(resolve_entities=False, load_dtd=False, no_network=True, huge_tree=False)
                    tree = etree.parse(stream, parser)
                    if tree.docinfo.doctype:
                        raise ValueError("DTD가 포함된 XML은 지원하지 않습니다.")
                else:
                    # Drain every member to retain ZIP/CRC checks, without
                    # allocating a whole image/attachment just to discard it.
                    while stream.read(1024 * 1024):
                        pass


def fingerprint(issue: dict) -> str:
    return json.dumps({key: issue.get(key) for key in ("kind", "message", "part_path", "section_index", "paragraph_index", "cell_row", "cell_column")}, sort_keys=True, ensure_ascii=False)


def issues(document: HwpxDocument) -> list[dict]:
    found = {}
    for check in (document.validation_errors, document.xml_validation_errors, document.reference_validation_errors):
        for issue in check():
            value = issue.to_dict()
            found[fingerprint(value)] = value
    return list(found.values())


def new_issues(current: list[dict], baseline: list[dict]) -> list[dict]:
    counts = Counter(fingerprint(i) for i in baseline)
    result = []
    for item in current:
        key = fingerprint(item)
        if counts[key]:
            counts[key] -= 1
        else:
            result.append(item)
    return result


def paragraph_text(node) -> str:
    # Exclude text in table cells, headers, notes and other nested paragraphs.
    return "".join(_text_content(t) for t in _PARAGRAPH_TEXTS(node, paragraph=node))


def cell_text(cell) -> str:
    paragraphs = _CELL_PARAGRAPHS(cell.element)
    return "\n".join(paragraph_text(p) for p in paragraphs) if paragraphs else cell.text


def editable(node) -> bool:
    # An allowlist also protects controls unknown to this engine version. Section
    # and column definitions are preserved by the engine's paragraph setter.
    paragraphs = _PARAGRAPHS(node)
    if not paragraphs:
        return False
    for paragraph in paragraphs:
        if any(child.tag not in {_HP + "run", _HP + "linesegarray"} for child in paragraph):
            return False
        for run in _RUNS(paragraph):
            for child in run:
                if child.tag == f"{{{NS['hp']}}}t" and len(child) == 0:
                    continue
                if child.tag == f"{{{NS['hp']}}}secPr":
                    continue
                if child.tag == f"{{{NS['hp']}}}ctrl" and len(child) == 1 and child[0].tag == f"{{{NS['hp']}}}colPr":
                    continue
                return False
    return True


def set_cell_text(cell, text: str) -> None:
    # The public setter distributes one-line text proportionally across existing
    # paragraphs. A textarea instead treats each newline as an explicit paragraph.
    # Use it to allocate/track safely, then normalize only this plain-text cell.
    cell.set_text(text)
    paragraphs = cell.element.xpath("./hp:subList/hp:p", namespaces=NS)
    lines = text.split("\n")
    for paragraph, line in zip(paragraphs, lines):
        _replace_paragraph_text_preserving_controls(paragraph, line)
    for paragraph in paragraphs[len(lines):]:
        paragraph.getparent().remove(paragraph)


def snapshot(document: HwpxDocument, revision: int, report: dict, name: str) -> dict:
    sections = []
    total = 0
    for s, section in enumerate(document.sections):
        paragraphs = []
        for index, wrapper in enumerate(section.paragraphs()):
            # _element is a read-only inspection here; wrapper.element marks a
            # part dirty even when only reading, which would alter no-op exports.
            node = wrapper._element
            paragraphs.append({"id": f"s{s}-p{index}", "section": s, "index": index, "text": paragraph_text(node), "editable": editable(node)})
        tables = []
        for index, table in enumerate(document.tables(s)):
            if table.row_count > 1000 or table.column_count > 1000 or table.row_count * table.column_count > MAX_ELEMENTS:
                raise ValueError("표 크기가 웹 편집 한도를 초과했습니다.")
            cells = [{"row": c.row, "column": c.column, "row_span": c.row_span, "col_span": c.col_span, "text": cell_text(c), "editable": editable(c.element)} for c in table.cells()]
            occupied = set()
            for c in cells:
                if c["row"] < 0 or c["column"] < 0 or c["row_span"] < 1 or c["col_span"] < 1 or c["row"] + c["row_span"] > table.row_count or c["column"] + c["col_span"] > table.column_count:
                    raise ValueError("표의 셀 위치 또는 병합 정보가 올바르지 않습니다.")
                for r in range(c["row"], c["row"] + c["row_span"]):
                    for col in range(c["column"], c["column"] + c["col_span"]):
                        if (r, col) in occupied:
                            raise ValueError("표의 셀 병합 영역이 겹칩니다.")
                        occupied.add((r, col))
            tables.append({"id": f"s{s}-t{index}", "section": s, "index": index, "rows": table.row_count, "columns": table.column_count, "cells": cells})
        total += len(paragraphs) + sum(len(t["cells"]) for t in tables)
        if total > MAX_ELEMENTS:
            raise ValueError("문서 요소가 웹 편집 한도(20,000개)를 초과했습니다.")
        sections.append({"index": s, "paragraphs": paragraphs, "tables": tables})
    return {"revision": revision, "name": name, "title": document.metadata().title or Path(name).stem, "sections": sections, "report": report}


def structure_signature(document: HwpxDocument, *, include_binaries: bool = True) -> dict:
    """Editing text must preserve every XML element and embedded binary."""
    tags = Counter()
    binaries = {}
    for path in document.list_part_paths():
        part = document.get_part(path)
        if isinstance(part, XmlPart):
            # Layout caches are deliberately invalidated by text editing.
            tags.update(node.tag for node in part._root.iter() if isinstance(node.tag, str))
        elif include_binaries and path.startswith("BinData/"):
            binaries[path] = hashlib.sha256(part.to_bytes()).hexdigest()
    for tag in list(tags):
        if etree.QName(tag).localname in {"lineseg", "linesegarray"}:
            del tags[tag]
    return {"tags": dict(tags), "binaries": binaries}


def process(directory: Path, kind: str, payload: dict, name: str) -> dict:
    if kind == "restore":
        from .revisions import process_restore
        return process_restore(directory, payload, name)
    if kind == "studio":
        from .studio import process_studio
        return process_studio(directory, payload, name)
    conversion = None
    if Path(name).suffix.lower() == ".hwp" and (directory / "convert-hwp").exists():
        from .conversion import convert_hwp
        conversion = convert_hwp(directory)
    elif Path(name).suffix.lower() == ".hwp":
        from .hwp_engine import process_hwp
        return process_hwp(directory, kind, payload, name)
    revision = payload.get("revision", 0)
    path = directory / f"revision-{revision}.hwpx"
    inspect_archive(path)
    document = HwpxDocument.open(path, validate=False)
    if document.is_distribution_protected:
        raise ValueError("배포용·보호 문서는 이 버전에서 편집할 수 없습니다.")
    original = issues(document)
    if kind == "analyze":
        baseline = original
    else:
        baseline = json.loads((directory / "snapshot-0.json").read_text(encoding="utf-8"))["report"]["issues"]
    previous = directory / f"snapshot-{revision}.json"
    repairs = json.loads(previous.read_text(encoding="utf-8"))["report"].get("repairs", []) if kind == "validate" and previous.exists() else []
    before_structure = structure_signature(document)
    if kind == "edit":
        from .policies import check_operations
        check_operations(snapshot(document, revision, {}, name), payload["operations"], payload.get("settings", {}))
        # Resolve every target against the immutable source revision first.
        for op in payload["operations"]:
            section = op["section"]
            if op["type"] == "paragraph":
                node = document.sections[section].paragraphs()[op["index"]]._element
                if not editable(node):
                    raise ValueError("복합 개체가 있는 문단은 읽기 전용입니다.")
                document.set_paragraph_text(section, op["index"], op["text"])
            elif op["type"] == "cell":
                cell = document.tables(section)[op["table"]].cell(op["row"], op["column"])
                if not editable(cell.element):
                    raise ValueError("복합 개체가 있는 셀은 읽기 전용입니다.")
                set_cell_text(cell, op["text"])
            else:
                raise ValueError("지원하지 않는 편집 명령입니다.")
        repairs = [issue.to_dict() for issue in document.repair_stale_paragraph_layout()]
    current = issues(document)
    added = new_issues(current, baseline)
    if added:
        raise ValueError("새로운 문서 오류가 발생하여 변경을 저장하지 않았습니다: " + added[0]["message"])
    if before_structure["binaries"] != structure_signature(document)["binaries"]:
        raise ValueError("첨부 파일 보존 검사에 실패했습니다.")
    # Parent controls must not disappear. Run/text nodes may change as the
    # documented text setter flattens multi-run content within the target.
    control_tags = {tag: count for tag, count in before_structure["tags"].items() if etree.QName(tag).localname in {"tbl", "pic", "equation", "ole", "ctrl", "rect", "ellipse", "bookmark", "fieldBegin", "fieldEnd"}}
    # Only edit mutates the in-memory document before serialization. Keep the
    # independent reopened-document signature for every operation below.
    after_tags = structure_signature(document, include_binaries=False)["tags"] if kind == "edit" else before_structure["tags"]
    if any(after_tags.get(tag, 0) != count for tag, count in control_tags.items()):
        raise ValueError("문서 개체 보존 검사에 실패했습니다.")
    compiled = document.compile(validate=False, auto_repair=False) if kind == "edit" else path.read_bytes()
    reopened = HwpxDocument.from_bytes(compiled, validate=False)
    reopened_structure = structure_signature(reopened)
    if reopened_structure["binaries"] != before_structure["binaries"] or any(reopened_structure["tags"].get(tag, 0) != count for tag, count in control_tags.items()):
        raise ValueError("저장 후 문서 개체·첨부 보존 검사에 실패했습니다.")
    reopened_issues = issues(reopened)
    if payload.get("settings", {}).get("strict_validation") and reopened_issues:
        raise ValueError("엄격 검증 설정에 따라 기존 오류가 있는 문서의 처리를 중단했습니다.")
    if new_issues(reopened_issues, baseline):
        raise ValueError("저장 후 다시 열기 검사에서 새로운 오류가 발견되었습니다.")
    if kind == "edit":
        revision += 1
        # Confirm every requested text was actually saved, including Korean and
        # XML-sensitive characters, rather than trusting a successful ZIP save.
        for op in payload["operations"]:
            if op["type"] == "paragraph":
                actual = paragraph_text(reopened.sections[op["section"]].paragraphs()[op["index"]]._element)
            else:
                actual = cell_text(reopened.tables(op["section"])[op["table"]].cell(op["row"], op["column"]))
            if actual != op["text"]:
                raise ValueError("요청한 텍스트와 저장 결과가 일치하지 않습니다. 변경을 저장하지 않았습니다.")
    report = {"status": "warning" if reopened_issues else "passed", "issues": reopened_issues, "new_issues": [], "baseline_count": len(baseline), "resolved_count": len(new_issues(baseline, reopened_issues)), "repairs": repairs, "checks": ["XML 구조", "문서 참조", "개체·첨부 보존", "저장 후 다시 열기"], "checked_revision": revision}
    view = snapshot(reopened, revision, report, name)
    if conversion:
        view.update(format="hwpx", conversion=conversion, format_note="HWP를 HWPX로 변환했습니다. 표·수식·그림과 페이지 배치를 원문 보기에서 확인해 주세요. 다운로드 파일은 HWPX이며 업로드한 HWP 원본도 별도로 받을 수 있습니다.")
    if kind == "edit":
        atomic_write(directory / f"revision-{revision}.hwpx", compiled)
    # A job only becomes visible after both files exist and the DB commits.
    atomic_json(directory / f"snapshot-{revision}.json", view)
    return {"revision": revision, "report": report}
