"""Hyperlinks/bookmarks with relationship remapping and reference-safe renaming."""
from __future__ import annotations

import re

from lxml import etree

from .editing import WordContainer, _tag, _attached

R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def containers(editor):
    from docx.blkcntnr import BlockItemContainer
    from .editing_stories import note_part
    yield editor.body
    for part in list(editor._backend.part.package.parts):
        root = getattr(part, "element", None)
        if root is not None and root.tag in {_tag("hdr"), _tag("ftr")}:
            yield WordContainer(editor, root, BlockItemContainer(root, part))
    for kind in ("footnote", "endnote"):
        part = note_part(editor, kind)
        if part is not None:
            for node in part.element:
                if node.tag == _tag(kind) and int(node.get(_tag("id"), "0")) > 0:
                    yield WordContainer(editor, node, BlockItemContainer(node, part))


def _name(value):
    if not isinstance(value, str) or not re.fullmatch(r"(?:[^\W\d]|_)\w{0,39}", value):
        raise ValueError("Bookmark names start with a letter/underscore and contain at most 40 word characters.")
    from .identity import internal_name
    if internal_name(value):
        raise ValueError("The _JAKAL_ identity bookmark namespace is reserved.")


def bookmarks(editor):
    output = []
    for owner in containers(editor):
        ends = {node.get(_tag("id")): node for node in owner._root.iter(_tag("bookmarkEnd"))}
        for start in owner._root.iter(_tag("bookmarkStart")):
            from .identity import internal_name
            if internal_name(start.get(_tag("name"))):
                continue
            end = ends.get(start.get(_tag("id")))
            if end is not None:
                output.append(WordBookmark(owner, start, end))
    return output


def _renamed(instruction, old, new):
    return re.sub(r'^(\s*(?:REF|PAGEREF)\s+)(?:"' + re.escape(old) + r'"|' + re.escape(old) + r')(?=\s|$)',
                  lambda match: match[1] + new, instruction, flags=re.IGNORECASE)


class WordBookmark:
    def __init__(self, owner, start, end):
        self._owner, self._node, self._end = owner, start, end

    def _check(self):
        self._owner._check()
        if not _attached(self._node, self._owner.part.element) or not _attached(self._end, self._owner.part.element):
            raise RuntimeError("Bookmark was removed.")

    @property
    def object_id(self):
        from .identity import identify
        self._check()
        return identify(self._node)

    @property
    def name(self):
        self._check()
        return self._node.get(_tag("name"))

    def rename(self, name):
        _name(name)
        old = self.name
        if name != old and any(item.name == name for item in bookmarks(self._owner._editor)):
            raise ValueError("Bookmark name already exists.")
        for owner in containers(self._owner._editor):
            for node in owner._root.iter(_tag("hyperlink")):
                if node.get(_tag("anchor")) == old:
                    node.set(_tag("anchor"), name)
            for field in owner.fields:
                updated = _renamed(field.instruction, old, name)
                if updated != field.instruction:
                    field.set_instruction(updated)
        self._node.set(_tag("name"), name)
        return self

    def delete(self):
        self._check()
        self._node.getparent().remove(self._node)
        self._end.getparent().remove(self._end)


def bookmark(selection, name, head, selected, tail, values):
    from .editing_format import element
    from .editing_ranges import _set_content
    _name(name)
    selection._validate(selected, values)
    existing = bookmarks(selection.start._owner._editor)
    if any(item.name == name for item in existing):
        raise ValueError("Bookmark name already exists.")
    identity = max([0] + [int(node.get(_tag("id"))) for owner in containers(selection.start._owner._editor)
                         for node in owner._root.iter(_tag("bookmarkStart"))]) + 1
    holder = etree.Element(_tag("p"))
    start = element(holder, "bookmarkStart", id=identity)
    start.set(_tag("name"), name)
    end = element(holder, "bookmarkEnd", id=identity)
    if len(values) == 1:
        _set_content(selection.start, head + [start] + selected[0] + [end] + tail)
    else:
        _set_content(values[0], head + [start] + selected[0])
        _set_content(values[-1], selected[-1] + [end] + tail)
    return WordBookmark(selection.start._owner, start, end)


def remap_relationships(node, source, target):
    if source is target:
        return
    mapping = {}
    for item in node.iter():
        for key, value in list(item.attrib.items()):
            if not key.startswith("{" + R + "}"):
                continue
            if value not in mapping:
                if value not in source.rels:
                    raise ValueError("Selected content contains an unresolved relationship.")
                rel = source.rels[value]
                mapping[value] = target.relate_to(rel.target_ref if rel.is_external else rel.target_part, rel.reltype, is_external=rel.is_external)
            item.set(key, mapping[value])


def clone_ids(node, editor):
    from .identity import ID
    used = {item.name for item in bookmarks(editor)}
    highest = max([0] + [int(node.get(_tag("id"))) for owner in containers(editor) for node in owner._root.iter(_tag("bookmarkStart"))])
    identities, names = {}, {}
    for item in node.iter():
        item.attrib.pop(ID, None)
        item.attrib.pop("{http://schemas.microsoft.com/office/word/2010/wordml}paraId", None)
        if item.tag == _tag("bookmarkStart"):
            old = item.get(_tag("name"))
            suffix = 1
            name = old
            while name in used:
                tail = "_copy" + str(suffix)
                name = old[:40 - len(tail)] + tail
                suffix += 1
            used.add(name)
            highest += 1
            identities[item.get(_tag("id"))] = str(highest)
            names[old] = name
            item.set(_tag("id"), str(highest))
            item.set(_tag("name"), name)
    for item in node.iter():
        if item.tag == _tag("bookmarkEnd"):
            item.set(_tag("id"), identities[item.get(_tag("id"))])
        if item.tag == _tag("hyperlink") and item.get(_tag("anchor")) in names:
            item.set(_tag("anchor"), names[item.get(_tag("anchor"))])
        for key in ("instr",):
            if item.get(_tag(key)):
                for old, new in names.items():
                    item.set(_tag(key), _renamed(item.get(_tag(key)), old, new))
        if item.tag == _tag("instrText") and item.text:
            for old, new in names.items():
                item.text = _renamed(item.text, old, new)
        if item.tag == "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr":
            values = [int(n.get("id")) for part in editor._backend.part.package.parts if getattr(part, "element", None) is not None
                      for n in part.element.iter(item.tag) if (n.get("id") or "").isdigit()]
            highest = max([highest] + values) + 1
            item.set("id", str(highest))
    from .editing_fields import find_fields
    owner = WordContainer(editor, node, editor._backend.part)
    for field in find_fields(owner, node):
        if field._simple:
            continue
        instruction = "".join(item.text or "" for item in field._instructions)
        updated = instruction
        for old, new in names.items():
            updated = _renamed(updated, old, new)
        if updated != instruction and field._instructions:
            field._instructions[0].text = updated
            for item in field._instructions[1:]:
                item.text = ""


def hyperlink(owner, target, children):
    from docx.oxml import OxmlElement
    node = OxmlElement("w:hyperlink")
    result = WordLink(owner, node)
    result._set_target(target)
    node.extend(children)
    return result


class WordLink:
    def __init__(self, owner, node):
        self._owner, self._node = owner, node

    def _check(self):
        self._owner._check()
        if not _attached(self._node, self._owner.part.element):
            raise RuntimeError("Hyperlink was removed.")

    @property
    def object_id(self):
        from .identity import identify
        self._check()
        return identify(self._node)

    @property
    def text(self):
        self._check()
        return "".join(self._node.itertext(_tag("t")))

    @property
    def target(self):
        self._check()
        anchor = self._node.get(_tag("anchor"))
        relationship = self._owner.part.rels.get(self._node.get("{" + R + "}id"))
        address = relationship.target_ref if relationship is not None else ""
        return address + ("#" + anchor if anchor else "")

    def _set_target(self, target):
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        if not isinstance(target, str) or not target or target == "#":
            raise ValueError("A hyperlink needs a URL or bookmark target.")
        if target.startswith("#"):
            _name(target[1:])
            self._node.set(_tag("anchor"), target[1:])
            self._node.attrib.pop("{" + R + "}id", None)
        else:
            self._node.set("{" + R + "}id", self._owner.part.relate_to(target, RT.HYPERLINK, is_external=True))
            self._node.attrib.pop(_tag("anchor"), None)

    def set_target(self, target):
        self._check()
        self._set_target(target)
        return self

    def set_text(self, text):
        from .preservation import _replace_slot
        self._check()
        if not isinstance(text, str):
            raise ValueError("Hyperlink text must be a string.")
        if list(self._node.iter(_tag("fldChar"))) or list(self._node.iter(_tag("fldSimple"))):
            raise ValueError("Edit hyperlink fields through their field handles.")
        texts = list(self._node.iter(_tag("t")))
        if texts:
            _replace_slot(texts[0], text)
            for node in texts[1:]:
                _replace_slot(node, "")
        else:
            from docx.text.run import Run
            from .editing_format import element
            Run(element(self._node, "r"), self._owner._parent).text = text
        return self

    def unlink(self):
        self._check()
        for child in list(self._node):
            self._node.addprevious(child)
        self._node.getparent().remove(self._node)


def link_instruction(instruction):
    """Parse literal HYPERLINK codes; leave unfamiliar/nested codes untouched."""
    tokens = re.findall(r'"[^"\r\n]*"|[^\s"]+', instruction)
    if not tokens or tokens[0].upper() != "HYPERLINK":
        return None
    if re.sub(r'"[^"\r\n]*"|[^\s"]+|\s+', '', instruction):
        return None
    address, anchor, extras = "", "", []
    index = 1
    while index < len(tokens):
        token = tokens[index]
        if token.lower() in {"\\l", "\\o", "\\t"}:
            if index + 1 == len(tokens) or tokens[index + 1].startswith("\\"):
                return None
            value = tokens[index + 1].strip('"')
            if token.lower() == "\\l":
                if anchor:
                    return None
                anchor = value
            else:
                extras.extend(tokens[index:index + 2])
            index += 2
        elif token.lower() in {"\\h", "\\m", "\\n"}:
            extras.append(token)
            index += 1
        elif token.startswith("\\") or address:
            return None
        else:
            address = token.strip('"')
            index += 1
    target = address + ("#" + anchor if anchor else "")
    return (target, extras) if target else None


def link_fields(backend):
    """Read-only field discovery, including separate textbox and note stories."""
    from types import SimpleNamespace
    from docx.blkcntnr import BlockItemContainer
    from .editing_fields import find_fields
    from .identity import STORIES
    output = {}
    editor = SimpleNamespace(_check=lambda: None)
    for part in backend.part.package.parts:
        root = getattr(part, "element", None)
        if root is None or etree.QName(root).localname not in STORIES:
            continue
        scopes = [root, *[node for node in root.iter() if node.tag in {_tag("txbxContent"), _tag("footnote"), _tag("endnote")}]]
        for scope in scopes:
            owner = WordContainer(editor, scope, BlockItemContainer(scope, part))
            for field in find_fields(owner, scope):
                parsed = link_instruction(field.instruction)
                if parsed is not None:
                    output[field._node] = (parsed[0], field.result)
    return output


class WordLinkField:
    """Keep the link API when Word stores a hyperlink as a field."""
    def __init__(self, field):
        self._field, self._node = field, field._node

    @property
    def object_id(self):
        return self._field.object_id

    @property
    def instruction(self):
        return self._field.instruction

    @property
    def result(self):
        return self._field.result

    @property
    def text(self):
        return self._field.result

    @property
    def target(self):
        parsed = link_instruction(self._field.instruction)
        if parsed is None:
            raise ValueError("Field is no longer a literal hyperlink.")
        return parsed[0]

    def set_target(self, target):
        if not isinstance(target, str) or not target or target == "#" or any(char in target for char in '"\r\n'):
            raise ValueError("A hyperlink needs a literal URL or bookmark target.")
        if target.startswith("#"):
            _name(target[1:])
        parsed = link_instruction(self._field.instruction)
        if parsed is None:
            raise ValueError("Field is no longer a literal hyperlink.")
        address, separator, anchor = target.partition("#")
        code = 'HYPERLINK' + (' "' + address + '"' if address else '')
        if separator:
            code += ' \\l "' + anchor + '"'
        self._field.set_instruction(code + (' ' + ' '.join(parsed[1]) if parsed[1] else ''))
        return self

    def set_text(self, text):
        self._field.set_result(text)
        return self

    def set_instruction(self, instruction):
        self._field.set_instruction(instruction)
        return self

    def set_result(self, text):
        return self.set_text(text)

    def lock(self, locked=True):
        self._field.lock(locked)
        return self

    def unlink(self):
        field = self._field
        field._check()
        if field._simple:
            for child in list(self._node):
                self._node.addprevious(child)
            self._node.getparent().remove(self._node)
            return
        nodes = list(field._owner.part.element.iter())
        left, right = nodes.index(field._node), nodes.index(field._end)
        boundary = nodes.index(field._separator) if field._separator is not None else right
        if any(node.tag in {_tag("fldChar"), _tag("fldSimple")} for node in nodes[left + 1:right] if node is not field._separator):
            raise ValueError("Nested hyperlink fields require explicit inner editing.")
        remove = [field._node, *field._instructions, field._end]
        if field._separator is not None:
            remove.append(field._separator)
        if any(node.tag == _tag("t") for node in nodes[left + 1:boundary]):
            raise ValueError("Hyperlink code contains unexpected content.")
        for node in remove:
            parent = node.getparent()
            parent.remove(node)
            if parent.tag == _tag("r") and all(child.tag == _tag("rPr") for child in parent):
                parent.getparent().remove(parent)
