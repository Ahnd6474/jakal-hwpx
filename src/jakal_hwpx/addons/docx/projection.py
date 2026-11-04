"""Normalize Word XML into the editable view consumed by the DOCX reader."""
from __future__ import annotations

from copy import deepcopy
from typing import Any, Iterator


def accept_visible_content(root: Any) -> None:
    """Project Word's accepted view without discarding wrapped runs or tables."""
    from docx.oxml.ns import qn

    from .textboxes import select_alternatives

    select_alternatives(root)

    from .identity import NS, internal_name
    private = {node.get(qn("w:id")) for node in root.iter(qn("w:bookmarkStart")) if internal_name(node.get(qn("w:name")))}
    for node in list(root.iter()):
        if node.tag in {qn("w:bookmarkStart"), qn("w:bookmarkEnd")} and node.get(qn("w:id")) in private:
            paragraph = next(node.iterancestors(qn("w:p")), None)
            if paragraph is not None:
                paragraph.set("{" + NS + "}paragraphAnchor", "1")
            node.getparent().remove(node)

    for node in list(root.iter()):
        parent = node.getparent()
        if parent is None:
            continue
        if node.tag in {qn("w:del"), qn("w:moveFrom")}:
            parent.remove(node)
        elif node.tag in {qn("w:ins"), qn("w:moveTo"), qn("w:sdt"), qn("w:customXml")}:
            content = node.find(qn("w:sdtContent")) if node.tag == qn("w:sdt") else node
            if content is not None:
                for child in list(content):
                    node.addprevious(child)
            parent.remove(node)
    # A native BOOKMARK field is a range, unlike hp:bookmark (a point).
    # Fold simple text ranges in the projection while retaining the source ZIP.
    for paragraph in root.iter(qn("w:p")):
        for start in reversed(list(paragraph.findall(qn("w:bookmarkStart")))):
            identity = start.get(qn("w:id"))
            end = next(
                (node for node in paragraph.findall(qn("w:bookmarkEnd")) if node.get(qn("w:id")) == identity),
                None,
            )
            if end is None or paragraph.index(end) <= paragraph.index(start) + 1:
                continue
            members = list(paragraph)[paragraph.index(start) + 1 : paragraph.index(end)]
            if any(node.tag != qn("w:r") or node.xpath(".//w:drawing | .//w:object") for node in members):
                continue
            from docx.oxml import OxmlElement

            field = OxmlElement("w:fldSimple")
            field.set(qn("w:instr"), "JAKAL_BOOKMARK " + start.get(qn("w:name"), ""))
            start.addprevious(field)
            for node in members:
                field.append(node)
            paragraph.remove(start)
            paragraph.remove(end)
    from .fields import normalize_ranges

    normalize_ranges(root)


def paragraph_nodes(paragraph: Any) -> Iterator[Any]:
    """Fold same-paragraph complex fields into the simple-field representation."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    stack: list[tuple[Any, list[str], list[Any], bool]] = []
    for node in paragraph:
        controls = node.findall(qn("w:fldChar")) if node.tag == qn("w:r") else []
        instructions = node.findall(qn("w:instrText")) if node.tag == qn("w:r") else []
        for control in controls:
            kind = control.get(qn("w:fldCharType"))
            if kind == "begin":
                stack.append((OxmlElement("w:fldSimple"), [], [], False))
            elif kind == "separate" and stack:
                field, command, children, _ = stack[-1]
                stack[-1] = field, command, children, True
            elif kind == "end" and stack:
                field, command, children, _ = stack.pop()
                field.set(qn("w:instr"), "".join(command).strip())
                field.extend(deepcopy(child) for child in children)
                if stack:
                    if stack[-1][3]:
                        stack[-1][2].append(field)
                else:
                    yield field
        if instructions and stack:
            stack[-1][1].extend(item.text or "" for item in instructions)
        if controls or instructions:
            continue
        if stack:
            if stack[-1][3]:
                stack[-1][2].append(node)
        else:
            yield node
    # A field spanning paragraphs cannot be folded locally. Its cached visible
    # contents still survive rather than losing the rest of the paragraph.
    for _field, _command, children, _separated in stack:
        yield from children
