"""Rich text ranges with explicit paragraph boundaries and guarded structure."""
from __future__ import annotations

from copy import deepcopy
import re

from lxml import etree

from .editing import WordParagraph, _tag, _balanced


def _split(paragraph, offset):
    """Split a detached copy at an editable-text offset; retain opaque XML."""
    from .preservation import _visible_slots
    native = paragraph._native
    slots = _visible_slots(native)
    if not isinstance(offset, int) or not 0 <= offset <= len(paragraph.text):
        raise IndexError("Range offset is outside the paragraph.")
    root = native._p.getroottree().getroot()
    visible = {root.find(path): (start, end, text) for path, start, end, text in slots}
    from .identity import private_ids
    private = private_ids(native._p)
    position = 0

    def divide(node):
        nonlocal position
        if node.tag in {_tag("bookmarkStart"), _tag("bookmarkEnd")} and node.get(_tag("id")) in private:
            return deepcopy(node), None
        if node in visible:
            start, end, text = visible[node]
            position = end
            cut = min(max(offset - start, 0), len(text))
            pieces = []
            for value in (text[:cut], text[cut:]):
                copy = deepcopy(node) if value else None
                if copy is not None and node.tag == _tag("t"):
                    copy.text = value
                    copy.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                pieces.append(copy)
            return pieces
        containers = {_tag(name) for name in ("r", "hyperlink", "ins", "moveTo", "sdt", "sdtContent", "customXml")}
        if node.tag not in containers:
            left = position < offset or position == offset and node.tag in {_tag("bookmarkEnd"), _tag("commentRangeEnd")}
            return (deepcopy(node), None) if left else (None, deepcopy(node))
        left, right = deepcopy(node), deepcopy(node)
        for child in list(left):
            if child.tag not in {_tag("rPr"), _tag("sdtPr"), _tag("sdtEndPr")}:
                left.remove(child)
        for child in list(right):
            if child.tag not in {_tag("rPr"), _tag("sdtPr"), _tag("sdtEndPr")}:
                right.remove(child)
        has_left = has_right = False
        for child in node:
            if child.tag in {_tag("rPr"), _tag("sdtPr"), _tag("sdtEndPr")}:
                continue
            a, b = divide(child)
            if a is not None:
                left.append(a)
                has_left = True
            if b is not None:
                right.append(b)
                has_right = True
        if has_left and has_right and node.tag in {_tag("sdt"), _tag("customXml"), _tag("ins"), _tag("moveTo")}:
            raise ValueError("A range cannot split a content control or revision wrapper.")
        return left if has_left else None, right if has_right else None

    left, right = [], []
    for child in native._p:
        if child.tag == _tag("pPr"):
            continue
        a, b = divide(child)
        if a is not None:
            left.append(a)
        if b is not None:
            right.append(b)
    return left, right


def _paragraph(owner, children, template):
    from docx.oxml import OxmlElement
    node = OxmlElement("w:p")
    if template.pPr is not None:
        properties = deepcopy(template.pPr)
        for section in list(properties.findall(_tag("sectPr"))):
            properties.remove(section)
        node.append(properties)
    node.extend(children)
    return WordParagraph(owner, node)


def _set_content(paragraph, children):
    for child in list(paragraph._node):
        if child.tag != _tag("pPr"):
            paragraph._node.remove(child)
    paragraph._node.extend(children)


class WordRange:
    def __init__(self, start, start_offset, end, end_offset):
        if start._owner._editor is not end._owner._editor or start._owner._root is not end._owner._root or start._owner.part is not end._owner.part:
            raise ValueError("A range must stay in one content container and document part.")
        self.start, self.start_offset, self.end, self.end_offset = start, start_offset, end, end_offset
        self._expected = self._text()
        self._prefix = start.text[:start_offset]

    def _paragraphs(self):
        self.start._check()
        self.end._check()
        nodes = self.start._owner._nodes()
        a, b = nodes.index(self.start._node), nodes.index(self.end._node)
        if a > b or self.start._node is self.end._node and self.start_offset > self.end_offset:
            raise ValueError("Range end precedes its start.")
        if any(node.tag != _tag("p") for node in nodes[a:b + 1]):
            raise ValueError("A text range cannot cross a table; select its cells separately.")
        if not 0 <= self.start_offset <= len(self.start.text) or not 0 <= self.end_offset <= len(self.end.text):
            raise IndexError("Range offsets are out of bounds.")
        return [WordParagraph(self.start._owner, node) for node in nodes[a:b + 1]]

    def _text(self):
        values = self._paragraphs()
        if len(values) == 1:
            return values[0].text[self.start_offset:self.end_offset]
        return "\n".join([values[0].text[self.start_offset:], *(p.text for p in values[1:-1]), values[-1].text[:self.end_offset]])

    @property
    def text(self):
        if self._text() != self._expected or self.start.text[:self.start_offset] != self._prefix:
            raise RuntimeError("This text range is stale; select the range again.")
        return self._expected

    def _pieces(self):
        _ = self.text
        values = self._paragraphs()
        from .identity import private_ids
        private = set().union(*(private_ids(p._node) for p in values))
        def cleaned(groups):
            output = []
            for group in groups:
                holder = etree.Element(_tag("p"))
                holder.extend(group)
                for node in list(holder.iter()):
                    if node.tag in {_tag("bookmarkStart"), _tag("bookmarkEnd")} and node.get(_tag("id")) in private:
                        node.getparent().remove(node)
                output.append(list(holder))
            return output
        head, first = _split(self.start, self.start_offset)
        if len(values) == 1:
            # Use a detached fragment root for the second split.
            from docx.text.paragraph import Paragraph
            from .preservation import _visible_slots
            # Split the end first instead, so the original paths remain valid.
            before_end, tail = _split(self.end, self.end_offset)
            detached = _paragraph(self.start._owner, before_end, self.start._node)
            # The proxy avoids mutating the document while validating the range.
            class Fragment:
                _native = Paragraph(detached._node, self.start._owner._parent)
                text = "".join(slot[3] for slot in _visible_slots(_native))
            _, selected = _split(Fragment(), self.start_offset)
            return head, cleaned([selected]), tail, values
        selected = [first, *[[deepcopy(child) for child in p._node if child.tag != _tag("pPr")] for p in values[1:-1]]]
        last, tail = _split(self.end, self.end_offset)
        selected.append(last)
        return head, cleaned(selected), tail, values

    def _validate(self, fragments, values):
        if len(values) > 1:
            for paragraph in values:
                if paragraph._node.getparent() is not paragraph._owner._root:
                    raise ValueError("Multi-paragraph edits cannot restructure wrapped controls.")
                if paragraph._node.find("./" + _tag("pPr") + "/" + _tag("sectPr")) is not None:
                    raise ValueError("A text range cannot remove a section boundary.")
        combined = etree.Element(_tag("p"))
        for fragment in fragments:
            combined.extend(deepcopy(child) for child in fragment)
        _balanced(combined)

    def replace(self, text):
        from docx.oxml import OxmlElement
        from docx.text.run import Run
        if not isinstance(text, str):
            raise ValueError("Replacement must be text.")
        head, selected, tail, values = self._pieces()
        self._validate(selected, values)
        fragments = []
        prototype = next((run for part in selected for node in part for run in node.iter(_tag("r"))), None)
        for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
            run = OxmlElement("w:r")
            if prototype is not None and prototype.find(_tag("rPr")) is not None:
                run.append(deepcopy(prototype.find(_tag("rPr"))))
            Run(run, self.start._native).text = line
            fragments.append([run])
        if len(values) == 1 and len(fragments) == 1 and len(selected[0]) == 1 and selected[0][0].tag == _tag("hyperlink"):
            link = selected[0][0]
            leading = head.pop() if head and head[-1].tag == link.tag and dict(head[-1].attrib) == dict(link.attrib) else None
            trailing = tail.pop(0) if tail and tail[0].tag == link.tag and dict(tail[0].attrib) == dict(link.attrib) else None
            children = (list(leading) if leading is not None else []) + fragments[0] + (list(trailing) if trailing is not None else [])
            for child in list(link):
                link.remove(child)
            link.extend(children)
            fragments = [[link]]
        self._replace_fragments(head, fragments, tail, values)
        return self.start

    def delete(self):
        return self.replace("")

    def _replace_fragments(self, head, fragments, tail, values, templates=None):
        templates = templates or values
        _set_content(self.start, head + fragments[0] + (tail if len(fragments) == 1 else []))
        previous = self.start._node
        for index, fragment in enumerate(fragments[1:], 1):
            template = templates[min(index, len(templates) - 1)]._node
            node = _paragraph(self.start._owner, fragment + (tail if index == len(fragments) - 1 else []), template)._node
            previous.addnext(node)
            previous = node
        for paragraph in values[1:]:
            paragraph._node.getparent().remove(paragraph._node)
        self.start._owner._ensure_tail()

    def copy_to(self, paragraph, offset, *, move=False):
        from .editing_navigation import clone_ids, remap_relationships
        head, selected, tail, values = self._pieces()
        self._validate(selected, values)
        if not isinstance(paragraph, WordParagraph):
            raise ValueError("Range destination must be a paragraph.")
        paragraph._check()
        if paragraph._owner._editor is not self.start._owner._editor:
            raise ValueError("Range copies must use the same active editor.")
        same = paragraph._node in {p._node for p in values}
        if same and len(values) != 1:
            raise ValueError("A multi-paragraph destination must be outside the source paragraphs.")
        if same and move and self.start_offset <= offset <= self.end_offset:
            raise ValueError("A move destination cannot lie inside its selected range.")
        before, after = _split(paragraph, offset)
        combined = etree.Element(_tag("p"))
        combined.extend(child for fragment in selected for child in fragment)
        if list(combined.iter(_tag("footnoteReference"))) or list(combined.iter(_tag("endnoteReference"))):
            if not move or paragraph._owner.part is not self.start._owner.part:
                raise ValueError("Copying note references needs a new note definition; move within the main story instead.")
        if not move:
            clone_ids(combined, self.start._owner._editor)
        remap_relationships(combined, self.start._owner.part, paragraph._owner.part)
        # Restore the paragraph grouping after the graph-wide remapping.
        selected = [list(combined)[sum(map(len, selected[:i])):sum(map(len, selected[:i + 1]))] for i in range(len(selected))]
        if move:
            self._replace_fragments(head, [[]], tail, values)
            if same:
                offset -= self.end_offset - self.start_offset if offset > self.end_offset else 0
                before, after = _split(paragraph, offset)
        destination = WordRange(paragraph, offset, paragraph, offset)
        destination._replace_fragments(before, selected, after, [paragraph], templates=values)
        return paragraph

    def move_to(self, paragraph, offset):
        return self.copy_to(paragraph, offset, move=True)

    def add_hyperlink(self, target):
        from .editing_navigation import hyperlink
        head, selected, tail, values = self._pieces()
        self._validate(selected, values)
        if len(values) != 1 or any(node.tag == _tag("hyperlink") or list(node.iter(_tag("hyperlink"))) for node in selected[0]):
            raise ValueError("Hyperlinks need one paragraph without nested hyperlinks.")
        link = hyperlink(self.start._owner, target, selected[0])
        self._replace_fragments(head, [[link._node]], tail, values)
        return link

    def add_bookmark(self, name):
        from .editing_navigation import bookmark
        head, selected, tail, values = self._pieces()
        return bookmark(self, name, head, selected, tail, values)


def find_text(container, text, *, regex=False, case_sensitive=True):
    if not isinstance(text, str) or not text:
        raise ValueError("Search text must be non-empty.")
    pattern = re.compile(text if regex else re.escape(text), 0 if case_sensitive else re.IGNORECASE)
    groups, current = [], []
    for block in container.blocks:
        if isinstance(block, WordParagraph):
            current.append(block)
        elif current:
            groups.append(current)
            current = []
    if current:
        groups.append(current)
    output = []
    for group in groups:
        joined = "\n".join(p.text for p in group)
        starts, position = [], 0
        for p in group:
            starts.append(position)
            position += len(p.text) + 1
        for match in pattern.finditer(joined):
            if match.start() == match.end():
                raise ValueError("Zero-width search results are not editable text ranges.")
            a = max(i for i, start in enumerate(starts) if start <= match.start())
            b = max(i for i, start in enumerate(starts) if start <= match.end())
            output.append(WordRange(group[a], match.start() - starts[a], group[b], match.end() - starts[b]))
    return output
