"""Structured field-result transactions and editable content controls."""
from __future__ import annotations

from contextlib import contextmanager
from copy import deepcopy

from lxml import etree

from .editing import WordContainer, WordParagraph, _tag, _balanced, _attached


def _clean_anchors(node, private=None):
    from .identity import private_ids
    private = private_ids(node) if private is None else private
    for child in list(node.iter()):
        if child.tag in {_tag("bookmarkStart"), _tag("bookmarkEnd")} and child.get(_tag("id")) in private:
            child.getparent().remove(child)


class _ResultPart:
    def __init__(self, part, root):
        self._part, self.element = part, root
        self.created = set()

    @property
    def part(self):
        return self

    def __getattr__(self, name):
        return getattr(self._part, name)

    def relate_to(self, *args, **kwargs):
        before = set(self._part.rels)
        result = self._part.relate_to(*args, **kwargs)
        self.created.update(set(self._part.rels) - before)
        return result

    def get_or_add_image(self, *args, **kwargs):
        before = set(self._part.rels)
        result = self._part.get_or_add_image(*args, **kwargs)
        self.created.update(set(self._part.rels) - before)
        return result


class _ResultContainer(WordContainer):
    _open = True

    def _check(self):
        if not self._open:
            raise RuntimeError("This field-result edit has expired.")
        super()._check()


@contextmanager
def edit_result(field):
    """Edit a detached result container; commit only after boundary validation."""
    from docx.oxml import OxmlElement
    from docx.blkcntnr import BlockItemContainer
    from .identity import ID
    field._check()
    from .identity import private_ids
    private = private_ids(field._owner.part.element)
    scratch = OxmlElement("w:body")
    proxy = _ResultPart(field._owner.part, scratch)
    container = _ResultContainer(field._owner._editor, scratch, BlockItemContainer(scratch, proxy))
    committed = False
    if field._simple:
        paragraph = OxmlElement("w:p")
        paragraph.extend(deepcopy(child) for child in field._node)
        scratch.append(paragraph)
        original = etree.tostring(field._node)
        source = [field._node]
    else:
        if field._separator is None:
            raise ValueError("A structured result needs a field separator.")
        first = next(field._separator.iterancestors(_tag("p")), None)
        last = next(field._end.iterancestors(_tag("p")), None)
        if first is None or last is None or first.getparent() is not last.getparent():
            raise ValueError("Field result boundaries must belong to sibling paragraphs.")
        parent = first.getparent()
        source = list(parent)[parent.index(first):parent.index(last) + 1]
        original = b"".join(etree.tostring(node) for node in source)
        nodes = list(field._owner.part.element.iter())
        positions = {node: index for index, node in enumerate(nodes)}
        copied = {}
        boundaries = {field._node, field._separator, field._end, *field._instructions}
        left, right = positions[field._separator], positions[field._end]
        properties = {_tag(name) for name in ("pPr", "rPr", "tblPr", "tblGrid", "trPr", "tcPr", "sdtPr", "sdtEndPr")}

        def fragment(node, lower, upper):
            selected = lower < positions[node] < upper
            if not len(node):
                copy = deepcopy(node) if selected else None
                if selected and node in boundaries:
                    copied[node] = copy
                return copy
            copy = deepcopy(node)
            for child in list(copy):
                copy.remove(child)
            for child in node:
                if child.tag in properties:
                    copy.append(deepcopy(child))
                else:
                    item = fragment(child, lower, upper)
                    if item is not None:
                        copy.append(item)
            return copy if selected or any(child.tag not in properties for child in copy) else None

        for block in source:
            selected = fragment(block, left, right)
            scratch.append(selected if selected is not None else OxmlElement("w:p"))
        prefix = fragment(first, -1, left + 1)
        suffix = fragment(last, right - 1, len(nodes))
        # Boundary paragraph identities remain with their original paragraphs.
        # Private anchors are regenerated around the resulting paragraphs.
        for block in [scratch, prefix, suffix]:
            _clean_anchors(block, private)
    baseline = etree.tostring(scratch)
    bindings = {node.get("{urn:jakal-hwpx:word-editing:1}id", etree.tostring(node.find(_tag("sdtPr")), encoding="unicode")):
                field._owner._editor._binding_text(node) for node in scratch.iter(_tag("sdt")) if node.find("./" + _tag("sdtPr") + "/" + _tag("dataBinding")) is not None}
    try:
        yield container
        field._check()
        if etree.tostring(scratch) == baseline:
            return
        current = b"".join(etree.tostring(node) for node in source)
        if current != original:
            raise RuntimeError("Field result changed while its detached edit was open.")
        _balanced(scratch)
        if field._simple:
            blocks = list(scratch)
            if len(blocks) != 1 or blocks[0].tag != _tag("p"):
                raise ValueError("A simple field has an inline result; use a complex field for multiple blocks.")
            for child in list(field._node):
                field._node.remove(child)
            field._node.extend(child for child in list(blocks[0]) if child.tag != _tag("pPr"))
        else:
            blocks = list(scratch)
            if not blocks or blocks[0].tag != _tag("p"):
                blocks.insert(0, OxmlElement("w:p"))
            if blocks[-1].tag != _tag("p"):
                blocks.append(OxmlElement("w:p"))
            for block in blocks:
                _clean_anchors(block, private)
            first_id, last_id = first.get(ID), last.get(ID)
            if first_id:
                blocks[0].set(ID, first_id)
            if len(blocks) > 1:
                if first is not last and last_id:
                    blocks[-1].set(ID, last_id)
                elif blocks[-1].get(ID) == first_id:
                    blocks[-1].attrib.pop(ID, None)
            leading = [child for child in list(prefix) if child.tag != _tag("pPr")]
            trailing = [child for child in list(suffix) if child.tag != _tag("pPr")]
            position = 1 if len(blocks[0]) and blocks[0][0].tag == _tag("pPr") else 0
            for child in reversed(leading):
                blocks[0].insert(position, child)
            blocks[-1].extend(trailing)
            for block in blocks:
                first.addprevious(block)
            for block in source:
                parent.remove(block)
            field._node = copied.get(field._node, field._node)
            field._separator = copied.get(field._separator, field._separator)
            field._end = copied.get(field._end, field._end)
            field._instructions = [copied.get(node, node) for node in field._instructions]
        field._node.set(_tag("dirty"), "false")
        for block in source if field._simple else blocks:
            for node in block.iter(_tag("sdt")):
                if node.find("./" + _tag("sdtPr") + "/" + _tag("dataBinding")) is not None:
                    key = node.get(ID, etree.tostring(node.find(_tag("sdtPr")), encoding="unicode"))
                    field._owner._editor._bindings.append((node, field._owner.part.element, bindings.get(key, "")))
        committed = True
    finally:
        container._open = False
        if not committed:
            for identity in proxy.created:
                field._owner.part.drop_rel(identity)


def controls(owner):
    nodes = owner._nodes()
    allowed = set(nodes)
    found = []
    for node in owner._root.iter(_tag("sdt")):
        if any(item in allowed for item in node.iter()) or any(item in allowed for item in node.iterancestors()):
            found.append(WordContentControl(owner, node))
    return found


def _properties(node, *, tag="", title="", kind="rich_text"):
    from docx.oxml import OxmlElement
    if kind not in {"rich_text", "plain_text"} or not isinstance(tag, str) or not isinstance(title, str):
        raise ValueError("Content controls need string tag/title and rich_text or plain_text kind.")
    properties = OxmlElement("w:sdtPr")
    node.append(properties)
    highest = max([0] + [int(item.get(_tag("val"), "0")) for item in node.getroottree().getroot().iter(_tag("id")) if item.get(_tag("val"), "").lstrip("-").isdigit()])
    for name, value in (("alias", title), ("tag", tag), ("id", highest + 1)):
        child = OxmlElement("w:" + name)
        child.set(_tag("val"), str(value))
        properties.append(child)
    if kind == "plain_text":
        properties.append(OxmlElement("w:text"))


def wrap_control(owner, start, end=None, **properties):
    from docx.oxml import OxmlElement
    nodes = owner._nodes()
    end = start + 1 if end is None else end
    if not isinstance(start, int) or not isinstance(end, int) or not 0 <= start < end <= len(nodes):
        raise ValueError("Control range must select one or more whole blocks.")
    selected = nodes[start:end]
    if any(node.getparent() is not owner._root for node in selected):
        raise ValueError("Select blocks directly inside this container.")
    combined = OxmlElement("w:body")
    combined.extend(deepcopy(node) for node in selected)
    _balanced(combined)
    if properties.get("kind") == "plain_text" and (len(selected) != 1 or selected[0].tag != _tag("p") or list(selected[0].iter(_tag("fldChar"))) or list(selected[0].iter(_tag("fldSimple")))):
        raise ValueError("Plain text controls require one paragraph without fields.")
    node = OxmlElement("w:sdt")
    # Build properties before insertion, then allocate the native ID globally.
    _properties(node, **properties)
    native_ids = [int(item.get(_tag("val"))) for item in owner.part.element.iter(_tag("id")) if item.get(_tag("val"), "").lstrip("-").isdigit()]
    node.find("./" + _tag("sdtPr") + "/" + _tag("id")).set(_tag("val"), str(max([0, *native_ids]) + 1))
    selected[0].addprevious(node)
    content = OxmlElement("w:sdtContent")
    node.append(content)
    content.extend(selected)
    return WordContentControl(owner, node)


def insert_control(owner, index, *, text="", **properties):
    from docx.oxml import OxmlElement
    _properties(OxmlElement("w:sdt"), **properties)  # Validate before insertion.
    paragraph = owner.insert_paragraph(index, text)
    try:
        return wrap_control(owner, index, **properties)
    except BaseException:
        paragraph.delete()
        raise


class WordContentControl(WordContainer):
    def __init__(self, owner, node):
        from docx.blkcntnr import BlockItemContainer
        content = node.find(_tag("sdtContent"))
        if content is None:
            raise ValueError("Content control has no content container.")
        self._owner, self._node = owner, node
        super().__init__(owner._editor, content, BlockItemContainer(content, owner.part))

    @property
    def object_id(self):
        from .identity import identify
        self._check()
        return identify(self._node)

    def _check(self):
        self._owner._check()
        if not _attached(self._node, self.part.element):
            raise RuntimeError("Content control was removed.")

    @property
    def tag(self):
        self._check()
        node = self._node.find("./" + _tag("sdtPr") + "/" + _tag("tag"))
        return node.get(_tag("val"), "") if node is not None else ""

    @property
    def text(self):
        self._check()
        nodes = self._nodes()
        if not nodes:
            return self._editor._binding_text(self._node)
        return "\n".join(WordParagraph(self, node).text for node in nodes if node.tag == _tag("p"))

    def set_text(self, text):
        self._check()
        nodes = self._nodes()
        if not nodes:
            from .editing_navigation import WordLink
            WordLink(self._owner, self._root).set_text(text)
        elif len(nodes) == 1 and nodes[0].tag == _tag("p"):
            WordParagraph(self, nodes[0]).set_text(text)
        else:
            raise ValueError("Edit individual paragraphs/tables to preserve structured control content.")
        return self

    def configure(self, *, tag=None, title=None):
        from .editing_format import replace
        self._check()
        if any(value is not None and not isinstance(value, str) for value in (tag, title)):
            raise ValueError("Control tag/title must be strings.")
        properties = self._node.find(_tag("sdtPr"))
        if properties is None:
            from docx.oxml import OxmlElement
            properties = OxmlElement("w:sdtPr")
            self._node.insert(0, properties)
        for name, value in (("tag", tag), ("alias", title)):
            if value is not None:
                replace(properties, name, {"val": value})
        return self

    def unwrap(self):
        self._check()
        for child in list(self._root):
            self._node.addprevious(child)
        self._node.getparent().remove(self._node)

    def move(self, index, *, destination=None):
        from .editing import _Block
        from docx.blkcntnr import BlockItemContainer
        _Block.move(self, index, destination=destination)
        self._parent = BlockItemContainer(self._root, self._owner.part)
        return self

    def delete(self):
        self._check()
        _balanced(self._node)
        self._node.getparent().remove(self._node)
        self._owner._ensure_tail()
