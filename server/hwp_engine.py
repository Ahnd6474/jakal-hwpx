"""Native HWP 5 editing, deliberately limited to plain text records."""
from __future__ import annotations

import hashlib
import json
import zlib
from pathlib import Path

import olefile

from jakal_hwpx import HwpDocument
from jakal_hwpx.hwp_binary import (
    HwpBinaryFileHeader, ParagraphHeaderRecord, ParagraphTextRecord,
    TAG_LIST_HEADER, TAG_PARA_HEADER, TAG_PARA_TEXT,
    TAG_PARA_CHAR_SHAPE, TAG_PARA_LINE_SEG,
)

from .policies import check_operations
from .files import atomic_json

MAX_PART = 24 * 1024 * 1024


def inspect_hwp(path: Path) -> None:
    if path.stat().st_size > 20 * 1024 * 1024:
        raise ValueError("파일 크기는 20MB 이하여야 합니다.")
    if not olefile.isOleFile(path):
        raise ValueError("HWP 5 형식의 바이너리 문서가 아닙니다. HWP 3 문서는 한글에서 HWPX로 저장해 주세요.")
    with olefile.OleFileIO(path) as compound:
        if not compound.exists("FileHeader") or not compound.exists("DocInfo"):
            raise ValueError("올바른 HWP 문서가 아닙니다.")
        header = HwpBinaryFileHeader.from_bytes(compound.openstream("FileHeader").read(256))
        if header.signature != "HWP Document File" or header.version_bytes[-1] != 5:
            raise ValueError("HWP 5 문서만 지원합니다.")
        if header.password_protected or header.distributable:
            raise ValueError("암호·배포용 HWP는 지원하지 않습니다. 보호를 해제한 문서를 사용해 주세요.")
        paths = compound.listdir(streams=True, storages=False)
        if len(paths) > 4096:
            raise ValueError("내부 스트림 수가 한도를 초과했습니다.")
        total = 0
        body = False
        for parts in paths:
            name = "/".join(parts)
            size = compound.get_size(parts)
            if size > MAX_PART:
                raise ValueError("HWP 내부 스트림 크기가 한도를 초과했습니다.")
            raw = compound.openstream(parts).read(MAX_PART + 1)
            if name.startswith("BodyText/Section"):
                body = True
            if header.compressed and (name == "DocInfo" or name.startswith("BodyText/")):
                decoder = zlib.decompressobj(-15)
                expanded = decoder.decompress(raw, MAX_PART + 1)
                if len(expanded) > MAX_PART or decoder.unconsumed_tail or not decoder.eof:
                    raise ValueError("HWP 압축 스트림 크기 또는 형식이 올바르지 않습니다.")
                size = len(expanded)
            total += size
            if total > 128 * 1024 * 1024:
                raise ValueError("HWP 압축 해제 크기가 한도를 초과했습니다.")
        if not body:
            raise ValueError("HWP 본문 스트림을 찾을 수 없습니다.")


def text_node(header):
    return next((n for n in header.children if isinstance(n, ParagraphTextRecord)), None)


def display_text(header) -> str:
    record = text_node(header)
    if record is None:
        return ""
    # The library's control stripper works in code units; join surrogate pairs
    # before writing UTF-8 JSON. CR is the paragraph terminator, not user text.
    return record.text.encode("utf-16-le", errors="surrogatepass").decode("utf-16-le", errors="replace").removesuffix("\r")


def writable(header) -> bool:
    record = text_node(header)
    if record is None or any(n.children for n in header.children):
        return False
    allowed = {TAG_PARA_TEXT, TAG_PARA_CHAR_SHAPE, TAG_PARA_LINE_SEG}
    if any(n.tag_id not in allowed for n in header.children):
        return False
    # Preserve mixed formatting by keeping these paragraphs read-only.
    shapes = [n for n in header.children if n.tag_id == TAG_PARA_CHAR_SHAPE]
    if len(shapes) > 1 or (shapes and (len(shapes[0].payload) != 8 or shapes[0].payload[:4] != b"\0" * 4)):
        return False
    return not any(ord(c) < 32 and c not in "\r\n\t" for c in record.raw_text)


def table_groups(table, control=None):
    result = {}
    group = None
    for node in (control if control is not None else table.control_node).children:
        if node.tag_id == TAG_LIST_HEADER:
            group = None
            if len(node.payload) == 47:
                key = (int.from_bytes(node.payload[10:12], "little"), int.from_bytes(node.payload[8:10], "little"))
                group = []
                result[key] = group
        elif isinstance(node, ParagraphHeaderRecord) and group is not None:
            group.append(node)
    return result


def hwp_snapshot(document, revision, report, name):
    sections = []
    total = 0
    for s, _section in enumerate(document.sections()):
        model = document.section_model(s)
        indices = {id(p.header): p.index for p in model.paragraphs()}
        paragraphs = []
        # Nested cell/note paragraphs belong in their own control, not the body.
        for root in model.roots:
            if isinstance(root, ParagraphHeaderRecord):
                index = indices[id(root)]
                paragraphs.append({"id": f"s{s}-p{index}", "section": s, "index": index, "text": display_text(root), "editable": writable(root)})
        tables = []
        for index, table in enumerate(document.tables(s)):
            if table.row_count < 1 or table.column_count < 1 or table.row_count > 1000 or table.column_count > 1000 or table.row_count * table.column_count > 20_000:
                raise ValueError("HWP 표 크기가 편집 한도를 초과했습니다.")
            groups = table_groups(table)
            cells = []
            occupied = set()
            for cell in table.cells():
                headers = groups.get((cell.row, cell.column), [])
                if cell.row < 0 or cell.column < 0 or cell.row_span < 1 or cell.col_span < 1 or cell.row + cell.row_span > table.row_count or cell.column + cell.col_span > table.column_count:
                    raise ValueError("HWP 셀 병합 정보가 올바르지 않습니다.")
                for row in range(cell.row, cell.row + cell.row_span):
                    for column in range(cell.column, cell.column + cell.col_span):
                        if (row, column) in occupied:
                            raise ValueError("HWP 셀 병합 영역이 겹칩니다.")
                        occupied.add((row, column))
                cells.append({"row": cell.row, "column": cell.column, "row_span": cell.row_span, "col_span": cell.col_span, "text": "\n".join(display_text(h) for h in headers), "editable": len(headers) == 1 and writable(headers[0])})
            tables.append({"id": f"s{s}-t{index}", "section": s, "index": index, "rows": table.row_count, "columns": table.column_count, "cells": cells})
        total += len(paragraphs) + sum(len(t["cells"]) for t in tables)
        if total > 20_000:
            raise ValueError("문서 요소 수가 웹 편집 한도를 초과했습니다.")
        sections.append({"index": s, "paragraphs": paragraphs, "tables": tables})
    return {"format": "hwp", "revision": revision, "name": name, "title": Path(name).stem, "sections": sections, "report": report, "format_note": "HWP 5 원본 형식으로 저장합니다. 복합 개체·혼합 글자 서식·여러 문단으로 구성된 셀은 읽기 전용입니다. 인쇄 배치와 문서 미리보기는 한글에서 확인하세요."}


def set_plain_text(header, value):
    if not writable(header):
        raise ValueError("복합 개체나 혼합 서식이 있는 HWP 영역은 읽기 전용입니다.")
    record = text_node(header)
    raw = value + "\r"
    record.set_raw_text(raw)
    header.char_count = (header.char_count & 0x80000000) | (len(raw.encode("utf-16-le")) // 2)
    # Drop stale line-position caches; Hancom recalculates them when opened.
    header.children = [n for n in header.children if n.tag_id != TAG_PARA_LINE_SEG]
    tail = bytearray(header.trailing_payload)
    if len(tail) >= 6:
        tail[4:6] = b"\0\0"  # line segment count at header offset 16
    header.trailing_payload = bytes(tail)


def signature(document):
    binary = document.binary_document()
    streams = {p: hashlib.sha256(binary.read_stream(p, decompress=False)).hexdigest() for p in binary.list_stream_paths() if not p.startswith("BodyText/")}
    records = []
    for s, _ in enumerate(document.sections()):
        for node in document.section_model(s).iter_nodes():
            # Text + paragraph count/layout fields are the only mutable records.
            if node.tag_id not in {TAG_PARA_HEADER, TAG_PARA_TEXT, TAG_PARA_LINE_SEG}:
                records.append((s, node.tag_id, node.level, hashlib.sha256(node.to_record().payload).hexdigest()))
    return streams, records


def process_hwp(directory, kind, payload, name):
    from .engine import new_issues
    revision = payload.get("revision", 0)
    source = directory / f"revision-{revision}.hwp"
    inspect_hwp(source)
    document = HwpDocument.open(source)
    original = [i.to_dict() for i in document.strict_lint_errors()]
    baseline = original if kind == "analyze" else json.loads((directory / "snapshot-0.json").read_text(encoding="utf-8"))["report"]["issues"]
    before = signature(document)
    if kind == "edit":
        source_view = hwp_snapshot(document, revision, {}, name)
        check_operations(source_view, payload["operations"], payload.get("settings", {}))
        for op in payload["operations"]:
            s = op["section"]
            if op["type"] == "paragraph":
                model = document.section_model(s)
                header = model.paragraphs()[op["index"]].header
                set_plain_text(header, op["text"])
                document.binary_document().replace_section_model(s, model)
                document._invalidate_bridge()
            else:
                table = document.tables(s)[op["table"]]
                model, _paragraph, control = table._live_context()
                headers = table_groups(table, control).get((op["row"], op["column"]), [])
                if len(headers) != 1:
                    raise ValueError("여러 문단으로 구성된 HWP 셀은 읽기 전용입니다.")
                set_plain_text(headers[0], op["text"])
                table._commit(model)
        revision += 1
    temp = directory / "pending.hwp"
    try:
        if kind == "edit":
            document.save(temp)
            inspect_hwp(temp)
            reopened = HwpDocument.open(temp)
        else:
            reopened = HwpDocument.open(source)
        if signature(reopened) != before:
            raise ValueError("HWP 개체·서식·첨부 파일 보존 검사에 실패했습니다.")
        current = [i.to_dict() for i in reopened.strict_lint_errors()]
        if new_issues(current, baseline):
            raise ValueError("HWP 저장 후 새 구조 오류가 발견되었습니다. 변경을 저장하지 않았습니다.")
        if payload.get("settings", {}).get("strict_validation") and current:
            raise ValueError("엄격 검증 설정에 따라 기존 오류가 있는 문서의 처리를 중단했습니다.")
        report = {"status": "warning" if current else "passed", "issues": current, "new_issues": [], "baseline_count": len(baseline), "resolved_count": len(new_issues(baseline, current)), "repairs": [], "checks": ["HWP 바이너리 구조", "문서 참조", "개체·서식·첨부 보존", "저장 후 다시 열기"], "checked_revision": revision}
        view = hwp_snapshot(reopened, revision, report, name)
        if kind == "edit":
            from .policies import targets, operation_key
            actual = targets(view)
            for op in payload["operations"]:
                if actual[operation_key(op)]["text"] != op["text"]:
                    raise ValueError("HWP 수정 요청과 저장한 텍스트가 일치하지 않습니다.")
            temp.replace(directory / f"revision-{revision}.hwp")
        atomic_json(directory / f"snapshot-{revision}.json", view)
        return {"revision": revision, "report": report}
    finally:
        temp.unlink(missing_ok=True)
