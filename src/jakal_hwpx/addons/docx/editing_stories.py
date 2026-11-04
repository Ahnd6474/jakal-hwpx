"""Section setup and genuine footnote/endnote editing."""
from __future__ import annotations

from copy import deepcopy
from lxml import etree

from .editing import WordContainer, _font, _points, _tag, _balanced
from .editing_format import element, replace


BREAKS = {"continuous": "continuous", "next_page": "nextPage", "even_page": "evenPage", "odd_page": "oddPage", "next_column": "nextColumn"}


class WordSection(WordContainer):
    """A section body with explicit page/column settings in points."""

    @property
    def object_id(self):
        from .identity import identify
        self._check()
        return identify(self._section)

    @property
    def _native_section(self):
        self._nodes()
        return next(section for section in self._editor._backend.sections if section._sectPr is self._section)

    def format_page(self, *, width_pt=None, height_pt=None, orientation=None, margins_pt=None,
                    header_distance_pt=None, footer_distance_pt=None, gutter_pt=None, break_type=None):
        from docx.enum.section import WD_ORIENT, WD_SECTION_START

        section = self._native_section
        if orientation not in {None, "portrait", "landscape"}:
            raise ValueError("Orientation must be portrait or landscape.")
        if break_type is not None and break_type not in BREAKS:
            raise ValueError("Unknown section break type.")
        margins = margins_pt or {}
        if margins.keys() - {"top", "bottom", "left", "right"}:
            raise ValueError("Page margins must use top, bottom, left, right.")
        width = _points(width_pt, positive=True) if width_pt is not None else section.page_width
        height = _points(height_pt, positive=True) if height_pt is not None else section.page_height
        if orientation is not None and width is not None and height is not None:
            if (orientation == "landscape") != (width > height):
                width, height = height, width
        converted = {edge + "_margin": _points(value) for edge, value in margins.items()}
        for name, value in (("header_distance", header_distance_pt), ("footer_distance", footer_distance_pt), ("gutter", gutter_pt)):
            if value is not None:
                converted[name] = _points(value)
        left = converted.get("left_margin", section.left_margin) or 0
        right = converted.get("right_margin", section.right_margin) or 0
        top = converted.get("top_margin", section.top_margin) or 0
        bottom = converted.get("bottom_margin", section.bottom_margin) or 0
        gutter = converted.get("gutter", section.gutter) or 0
        if width is not None and left + right + gutter >= width or height is not None and top + bottom >= height:
            raise ValueError("Page margins must leave a positive text area.")
        section.page_width, section.page_height = width, height
        if orientation is not None:
            section.orientation = WD_ORIENT.LANDSCAPE if orientation == "landscape" else WD_ORIENT.PORTRAIT
        elif width is not None and height is not None:
            section.orientation = WD_ORIENT.LANDSCAPE if width > height else WD_ORIENT.PORTRAIT
        for name, value in converted.items():
            setattr(section, name, value)
        if break_type is not None:
            section.start_type = {"continuous": WD_SECTION_START.CONTINUOUS, "next_page": WD_SECTION_START.NEW_PAGE,
                                  "even_page": WD_SECTION_START.EVEN_PAGE, "odd_page": WD_SECTION_START.ODD_PAGE,
                                  "next_column": WD_SECTION_START.NEW_COLUMN}[break_type]
        return self

    def format_columns(self, *, count=1, gap_pt=18, widths_pt=None, separator=False):
        if isinstance(count, bool) or not isinstance(count, int) or not 1 <= count <= 45:
            raise ValueError("Column count must be an integer from 1 to 45.")
        gap = _points(gap_pt)
        widths = None if widths_pt is None else [_points(value, positive=True) for value in widths_pt]
        if widths is not None and len(widths) != count:
            raise ValueError("Explicit column widths must match count.")
        section = self._native_section
        available = (section.page_width or _points(612)) - (section.left_margin or 0) - (section.right_margin or 0) - (section.gutter or 0)
        if gap * (count - 1) >= available or widths is not None and sum(widths) + gap * (count - 1) > available + 635:
            raise ValueError("Columns do not fit the section text area.")
        node = replace(self._section, "cols", {"num": count, "space": round(gap / 635), "equalWidth": "1" if widths is None else "0", "sep": "1" if separator else "0"})
        for index, width in enumerate(widths or []):
            element(node, "col", w=round(width / 635), space=round(gap / 635) if index < count - 1 else 0)
        return self


def insert_section(editor, index, break_type):
    editor._check()
    if break_type not in BREAKS:
        raise ValueError("Unknown section break type.")
    body = editor.body
    nodes = body._nodes()
    if not 0 <= index <= len(nodes):
        raise IndexError("Section insertion index is out of range.")
    # A section break needs a direct body paragraph; never cut a table or SDT.
    reference = body._reference(nodes, index) if index < len(nodes) else editor._backend.element.body.find(_tag("sectPr"))
    if index < len(nodes) and reference is not nodes[index]:
        raise ValueError("A section break cannot split a wrapped content control.")
    sections = list(editor._backend.sections)
    children = list(editor._backend.element.body)
    position = children.index(reference)
    containing = next(section for section in sections if children.index(section._sectPr if section._sectPr.getparent() is body._root else section._sectPr.getparent().getparent()) >= position)
    marker = element(body._root, "p")
    properties = element(marker, "pPr")
    copied = deepcopy(containing._sectPr)
    properties.append(copied)
    reference.addprevious(marker)
    for name in ("headerReference", "footerReference"):
        for node in list(containing._sectPr.findall(_tag(name))):
            containing._sectPr.remove(node)
    replace(containing._sectPr, "type", {"val": BREAKS[break_type]})
    return WordSection(editor, body._root, body._parent, section=containing._sectPr)


def note_part(editor, kind, *, create=False):
    from docx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
    from docx.opc.packuri import PackURI
    from docx.parts.story import StoryPart
    from docx.oxml import OxmlElement, parse_xml

    editor._check()
    if kind not in {"footnote", "endnote"}:
        raise ValueError("Note kind must be footnote or endnote.")
    doc = editor._backend.part
    reltype = RT.FOOTNOTES if kind == "footnote" else RT.ENDNOTES
    relation = next((rel for rel in doc.rels.values() if rel.reltype == reltype and not rel.is_external), None)
    if relation is None:
        if not create:
            return None
        root = OxmlElement("w:" + kind + "s")
        for identity, tag in ((-1, "separator"), (0, "continuationSeparator")):
            item = element(root, kind, id=identity, type=tag)
            element(element(element(item, "p"), "r"), tag)
        part = StoryPart(PackURI("/word/" + kind + "s.xml"), CT.WML_FOOTNOTES if kind == "footnote" else CT.WML_ENDNOTES, root, doc.package)
        doc.relate_to(part, reltype)
        return part
    old = relation.target_part
    if isinstance(old, StoryPart):
        return old
    # python-docx reads note parts as opaque Part objects. Upgrade this one
    # part in place in the relationship graph, retaining every original rel.
    part = StoryPart(old.partname, old.content_type, parse_xml(old.blob), doc.package)
    part.__dict__["rels"] = old.rels
    editor._note_upgrades.append((old, part, etree.tostring(part.element)))
    for node in part.element.iter(_tag("sdt")):
        if node.find("./" + _tag("sdtPr") + "/" + _tag("dataBinding")) is not None:
            editor._bindings.append((node, part.element, editor._binding_text(node)))
    for source in list(doc.package.parts):
        for rel in source.rels.values():
            if not rel.is_external and rel.target_part is old:
                rel._target = part
    return part


def _note_container(editor, part, node, kind):
    from docx.blkcntnr import BlockItemContainer
    return WordNote(editor, node, BlockItemContainer(node, part), kind=kind)


def note_at(editor, kind, index):
    part = note_part(editor, kind)
    if part is None:
        raise IndexError("Document has no " + kind + "s.")
    notes = [node for node in part.element if node.tag == _tag(kind) and node.get(_tag("type"), "normal") == "normal" and int(node.get(_tag("id"), "0")) > 0]
    return _note_container(editor, part, notes[index], kind)


def add_note(paragraph, text, kind):
    paragraph._check()
    if paragraph._owner.part is not paragraph._owner._editor._backend.part or any(node.tag == _tag("txbxContent") for node in paragraph._node.iterancestors()):
        raise ValueError("Note references can only be added to the main document story.")
    part = note_part(paragraph._owner._editor, kind, create=True)
    identity = max([0] + [int(node.get(_tag("id"), "0")) for node in part.element]) + 1
    node = element(part.element, kind, id=identity)
    p = element(node, "p")
    r = element(p, "r")
    element(element(r, "rPr"), "vertAlign", val="superscript")
    element(r, kind + "Ref")
    reference = paragraph._native.add_run()
    reference.font.superscript = True
    element(reference._r, kind + "Reference", id=identity)
    result = _note_container(paragraph._owner._editor, part, node, kind)
    result.paragraph(0)._native.add_run(text)
    return result


class WordNote(WordContainer):
    def __init__(self, *args, kind, **kwargs):
        super().__init__(*args, **kwargs)
        self.kind = kind

    @property
    def id(self):
        self._check()
        return int(self._root.get(_tag("id")))

    def format_reference(self, **properties):
        from docx.text.run import Run
        from docx.text.paragraph import Paragraph
        _font(None, properties)
        identity = str(self.id)
        for node in self._editor._backend.element.iter(_tag(self.kind + "Reference")):
            if node.get(_tag("id")) == identity:
                _font(Run(node.getparent(), Paragraph(node.getparent().getparent(), self._editor._backend)), properties)
        for node in self._root.iter(_tag(self.kind + "Ref")):
            _font(Run(node.getparent(), self._parent), properties)
        return self

    def delete(self):
        identity = str(self.id)
        for block in self.blocks:
            _balanced(block._node)
        for node in list(self._editor._backend.element.iter(_tag(self.kind + "Reference"))):
            if node.get(_tag("id")) == identity:
                node.getparent().remove(node)
        self._root.getparent().remove(self._root)

    def move_reference(self, paragraph):
        """Move this note's existing reference to the end of a body paragraph."""
        from docx.oxml import OxmlElement
        from .editing import WordParagraph
        if not isinstance(paragraph, WordParagraph):
            raise ValueError("A note reference destination must be a paragraph.")
        paragraph._check()
        if paragraph._owner._editor is not self._editor or paragraph._owner.part is not self._editor._backend.part or any(node.tag == _tag("txbxContent") for node in paragraph._node.iterancestors()):
            raise ValueError("Note references must stay in the same main document story.")
        identity = str(self.id)
        references = [node for node in self._editor._backend.element.iter(_tag(self.kind + "Reference")) if node.get(_tag("id")) == identity]
        if len(references) != 1:
            raise ValueError("Moving a note requires exactly one body reference.")
        reference = references[0]
        previous = reference.getparent()
        run = OxmlElement("w:r")
        properties = previous.find(_tag("rPr"))
        if properties is not None:
            run.append(deepcopy(properties))
        run.append(reference)
        paragraph._node.append(run)
        return self
