"""Explicit field refreshes without executing code or inventing pagination."""
from __future__ import annotations

from datetime import datetime
from decimal import Decimal, InvalidOperation
import re

from .editing import WordParagraph, _tag


class _Unsupported(ValueError):
    pass


def _tokens(code):
    pattern = r'"(?:[^"\r\n]|"")*"|<=|>=|<>|[=<>]|[^\s"=<>]+'
    if re.sub(pattern + r"|\s+", "", code):
        raise _Unsupported("Unsupported field quoting.")
    return re.findall(pattern, code)


def _value(token):
    return token[1:-1].replace('""', '"') if token.startswith('"') and token.endswith('"') else token


def _instruction(field, fields, planned):
    if field._simple:
        return field.instruction
    nodes = list(field._owner.part.element.iter())
    stop = field._separator if field._separator is not None else field._end
    left, right = nodes.index(field._node), nodes.index(stop)
    starts = {item._node: item for item in fields if item is not field}
    output, index = [], left + 1
    while index < right:
        node = nodes[index]
        inner = starts.get(node)
        if inner is not None:
            text = planned.get(node, inner.result)
            output.append('"' + text.replace('"', '""') + '"')
            end = node if inner._simple else inner._end
            index = nodes.index(end) + 1
            if inner._simple:
                while index < right and any(parent is node for parent in nodes[index].iterancestors()):
                    index += 1
            continue
        if node.tag == _tag("instrText"):
            output.append(node.text or "")
        index += 1
    return "".join(output)


def _compare(left, operator, right):
    try:
        a, b = Decimal(left), Decimal(right)
        if not a.is_finite() or not b.is_finite():
            raise InvalidOperation
    except InvalidOperation:
        a, b = left, right
        if operator in {"=", "<>"}:
            if any(char in b for char in "*?") and (len(a) > 128 or len(b) > 128):
                raise _Unsupported("Wildcard comparisons beyond 128 characters require Word.")
            previous = [True] + [False] * len(a)
            for char in b:
                current = [previous[0] and char == "*"]
                for index, value in enumerate(a, 1):
                    current.append(previous[index] or current[index - 1] if char == "*" else previous[index - 1] and (char == "?" or char == value))
                previous = current
            equal = previous[-1]
            return equal if operator == "=" else not equal
        raise _Unsupported("Text ordering requires Word's locale comparison rules.")
    return {"=": lambda: a == b, "<>": lambda: a != b, "<": lambda: a < b,
            ">": lambda: a > b, "<=": lambda: a <= b, ">=": lambda: a >= b}[operator]()


def _evaluate(field, fields, planned, editor, data, now):
    tokens = _tokens(_instruction(field, fields, planned))
    if not tokens:
        raise _Unsupported("Empty field code.")
    kind = tokens.pop(0).upper()
    if len(tokens) >= 2 and tokens[-2] == r"\*" and tokens[-1].upper() == "MERGEFORMAT":
        tokens = tokens[:-2]
    if kind == "IF" and len(tokens) in {4, 5} and tokens[1] in {"=", "<>", "<", ">", "<=", ">="}:
        for token in (tokens[0], tokens[2]):
            if not token.startswith('"'):
                try:
                    Decimal(token)
                except InvalidOperation:
                    raise _Unsupported("Unquoted IF operands must be numeric; quote literal strings.") from None
        return _value(tokens[3]) if _compare(_value(tokens[0]), tokens[1], _value(tokens[2])) else _value(tokens[4]) if len(tokens) == 5 else ""
    if kind == "MERGEFIELD" and len(tokens) == 1:
        name = _value(tokens[0])
        if name not in data:
            raise _Unsupported("Merge data is missing: " + name)
        return str(data[name])
    properties = {"TITLE": "title", "SUBJECT": "subject", "AUTHOR": "author", "KEYWORDS": "keywords",
                  "COMMENTS": "comments", "LASTSAVEDBY": "last_modified_by"}
    if kind == "DOCPROPERTY" and len(tokens) == 1:
        kind = _value(tokens.pop()).upper()
    if kind in properties and not tokens:
        return str(getattr(editor._backend.core_properties, properties[kind]) or "")
    if kind in {"DATE", "TIME"}:
        formats = {"yyyy-MM-dd": "%Y-%m-%d", "yyyy/MM/dd": "%Y/%m/%d", "yyyy.MM.dd": "%Y.%m.%d",
                   "MM/dd/yyyy": "%m/%d/%Y", "dd/MM/yyyy": "%d/%m/%Y", "HH:mm": "%H:%M", "HH:mm:ss": "%H:%M:%S"}
        if len(tokens) == 2 and tokens[0] == r"\@" and _value(tokens[1]) in formats:
            return now.strftime(formats[_value(tokens[1])])
        raise _Unsupported("Date/time refresh requires an explicitly supported format.")
    if kind == "REF" and tokens and all(token == r"\h" for token in tokens[1:]):
        from .editing_navigation import bookmarks
        matches = [item for item in bookmarks(editor) if item.name == _value(tokens[0])]
        if len(matches) != 1:
            raise _Unsupported("Reference bookmark is missing or ambiguous.")
        bookmark = matches[0]
        nodes = list(bookmark._owner.part.element.iter())
        left, right = nodes.index(bookmark._node) + 1, nodes.index(bookmark._end)
        if field._node in nodes[left:right]:
            raise _Unsupported("A reference cannot read its own result.")
        output, index = [], left
        starts = {item._node: item for item in fields}
        while index < right:
            node = nodes[index]
            inner = starts.get(node)
            if inner is not None:
                output.append(planned.get(node, inner.result))
                end = inner._node if inner._simple else inner._end
                index = nodes.index(end) + 1
                if inner._simple:
                    while index < right and any(parent is node for parent in nodes[index].iterancestors()):
                        index += 1
                continue
            output.append(node.text or "" if node.tag == _tag("t") else "\t" if node.tag == _tag("tab") else "\n" if node.tag in {_tag("br"), _tag("p")} else "")
            index += 1
        return "".join(output)
    if kind == "TOC":
        raise _Unsupported("Use rebuild_toc(); page numbers require a layout engine.")
    if kind in {"PAGE", "PAGEREF", "NUMPAGES", "SECTIONPAGES"}:
        raise _Unsupported("This field requires a layout engine.")
    raise _Unsupported("Unsupported field code or switches: " + kind)


def update_fields(editor, *, data=None, now=None, strict=False):
    from .editing_navigation import containers
    from .identity import ID
    editor._check()
    if data is not None and (not isinstance(data, dict) or any(not isinstance(key, str) or not isinstance(value, (str, int, float, bool)) for key, value in data.items())):
        raise ValueError("Merge data must map string names to scalar values.")
    if now is None:
        now = datetime.now()
    elif isinstance(now, str):
        now = datetime.fromisoformat(now)
    if not isinstance(now, datetime) or not isinstance(strict, bool):
        raise ValueError("Provide a datetime/ISO timestamp and boolean strict flag.")
    owners = list(containers(editor))
    for owner in list(owners):
        index = 0
        while True:
            try:
                owners.append(owner.textbox(index))
            except IndexError:
                break
            index += 1
    fields = list({field._node: field for owner in owners for field in owner.fields}.values())
    positions = {owner.part: {node: index for index, node in enumerate(owner.part.element.iter())} for owner in owners}
    fields.sort(key=lambda field: len(list(field._node.iter())) if field._simple else positions[field._owner.part][field._end] - positions[field._owner.part][field._node])
    planned, reports = {}, []
    for field in fields:
        report = {"id": field._node.get(ID), "instruction": field.instruction}
        if field._node.get(_tag("fldLock"), "false") in {"1", "true", "on"}:
            report["status"] = "locked"
        else:
            try:
                value = _evaluate(field, fields, planned, editor, data or {}, now)
                if any(ord(char) < 32 and char not in "\t\n\r" for char in value):
                    raise ValueError("Field values must be valid XML text.")
                from lxml import etree
                etree.Element("text").text = value  # Validate before writing any result.
                if value != field.result and any(node.tag in {_tag(name) for name in ("fldChar", "fldSimple", "drawing", "pict", "tbl", "p", "hyperlink", "sdt", "customXml")} | {"{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath"} for node in field._result_nodes()):
                    raise _Unsupported("Structured results must be edited with edit_result().")
                planned[field._node] = value
                report["status"] = "unchanged" if value == field.result else "updated"
            except _Unsupported as error:
                if strict:
                    raise ValueError(str(error)) from error
                report.update(status="unsupported", reason=str(error))
        reports.append(report)
    for field in fields:
        if field._node in planned and planned[field._node] != field.result:
            field.set_result(planned[field._node])
    return reports


def rebuild_toc(field, *, include_page_numbers=False):
    """Rebuild heading text and navigation; never guess page numbers."""
    from .editing_navigation import bookmarks
    from .reader import _style_chain
    editor = field._owner._editor
    tokens = _tokens(field.instruction)
    if not tokens or tokens[0].upper() != "TOC" or field._simple:
        raise ValueError("TOC rebuilding requires a complex TOC field.")
    if include_page_numbers:
        raise ValueError("TOC page numbers require Word or Hancom pagination.")
    minimum, maximum = 1, 3
    index = 1
    while index < len(tokens):
        switch = tokens[index]
        if switch == r"\o" and index + 1 < len(tokens):
            match = re.fullmatch(r"([1-9])-([1-9])", _value(tokens[index + 1]))
            if match is None or int(match[1]) > int(match[2]):
                raise ValueError("TOC outline range must be 1..9 in increasing order.")
            minimum, maximum = int(match[1]), int(match[2])
            index += 2
        elif switch in {r"\h", r"\z", r"\u", r"\n"}:
            index += 1
        else:
            raise ValueError("Unsupported TOC switch: " + switch)
    entries, known = [], {item.name for item in bookmarks(editor)}
    result_nodes = set(field._result_nodes())
    for node in editor.body._nodes():
        if node.tag != _tag("p") or node in result_nodes:
            continue
        paragraph = WordParagraph(editor.body, node)
        formats = [node, *(style.element for style in _style_chain(paragraph._native.style))]
        level = next((int(outline.get(_tag("val"), "9")) + 1 for item in formats
                      if (outline := item.find("./" + _tag("pPr") + "/" + _tag("outlineLvl"))) is not None), 10)
        if not minimum <= level <= maximum:
            continue
        text = "".join(node.itertext(_tag("t")))
        name = "_JToc_" + paragraph.object_id
        if name not in known:
            paragraph.range().add_bookmark(name)
            known.add(name)
        entries.append((level, text, name))
    with field.edit_result() as content:
        for block in list(content.blocks):
            block.delete()
        for level, text, name in entries:
            paragraph = content.append_paragraph()
            paragraph.format(left_indent_pt=(level - minimum) * 12)
            paragraph.insert_hyperlink(0, "#" + name, text)
        if not entries:
            content.append_paragraph()
    field._node.set(_tag("dirty"), "true")
    return {"entries": len(entries), "requires_layout": r"\n" not in tokens}
