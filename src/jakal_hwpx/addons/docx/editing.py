"""Transactional Word editing without rebuilding unrelated XML or package parts.

Lengths in this API are points (1/72 inch), explicit in parameter names.
Handles belong to one edit_word_details() context and expire when it closes.
"""
from __future__ import annotations

from copy import deepcopy
from io import BytesIO
import math
from typing import Any, Sequence, Self, TYPE_CHECKING

from lxml import etree

if TYPE_CHECKING:
    from .editing_stories import WordSection, WordNote
    from .editing_fields import WordField, WordList

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _tag(name):
    return "{" + W + "}" + name


def _attached(node, root):
    return node is root or any(parent is root for parent in node.iterancestors())


def _points(value, *, positive=False):
    from docx.shared import Pt

    number = float(value)
    if not math.isfinite(number) or (number <= 0 if positive else number < 0):
        raise ValueError("Point values must be finite and " + ("positive." if positive else "non-negative."))
    return Pt(number)


def _color(value):
    value = value.lstrip("#")
    if len(value) != 6 or any(c not in "0123456789abcdefABCDEF" for c in value):
        raise ValueError("Colors require six hexadecimal RGB digits.")
    return value.upper()


def _font(target, values):
    from .editing_format import FONT_PROPERTIES, font_extras
    from docx.enum.text import WD_UNDERLINE
    from docx.oxml.ns import qn
    from docx.shared import RGBColor

    font_extras(None, values)
    extras = {key: value for key, value in values.items() if key in FONT_PROPERTIES}
    allowed = {"bold", "italic", "underline", "strike", "superscript", "subscript", "font_name", "font_size_pt", "color"} | FONT_PROPERTIES
    if values.keys() - allowed:
        raise ValueError("Unknown character properties: " + ", ".join(sorted(values.keys() - allowed)))
    if values.get("superscript") and values.get("subscript"):
        raise ValueError("Superscript and subscript cannot both be enabled.")
    converted = {key: value for key, value in values.items() if key not in FONT_PROPERTIES}
    if "font_size_pt" in values and values["font_size_pt"] is not None:
        size = float(values["font_size_pt"])
        converted["font_size_pt"] = _points(size, positive=True)
        if size * 2 != round(size * 2):
            raise ValueError("Word font sizes must be multiples of half a point.")
    if "color" in values and values["color"] is not None:
        converted["color"] = RGBColor.from_string(_color(values["color"]))
    if isinstance(values.get("underline"), str):
        choices = {name.lower(): value for name, value in WD_UNDERLINE.__members__.items()}
        choices["wave"] = WD_UNDERLINE.WAVY
        if values["underline"] not in choices:
            raise ValueError("Unsupported underline name.")
        converted["underline"] = choices[values["underline"]]
    if target is None:
        return
    font = target.font
    for key, value in converted.items():
        if key == "font_size_pt":
            font.size = value
        elif key == "color":
            font.color.rgb = value
        elif key == "font_name":
            font.name = value
            pr = target._element.get_or_add_rPr()
            fonts = pr.get_or_add_rFonts()
            fonts.attrib.pop(qn("w:cstheme"), None)
            for name in ("ascii", "hAnsi", "eastAsia", "cs"):
                fonts.attrib.pop(qn("w:" + name + "Theme"), None)
                if value is None:
                    fonts.attrib.pop(qn("w:" + name), None)
                else:
                    fonts.set(qn("w:" + name), value)
        else:
            setattr(font, key, value)
    font_extras(target, extras)


def _paragraph_format(target, values):
    from .editing_format import PARAGRAPH_PROPERTIES, paragraph_extras
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    names = {"left_indent_pt": "left_indent", "right_indent_pt": "right_indent",
             "first_line_indent_pt": "first_line_indent", "space_before_pt": "space_before", "space_after_pt": "space_after"}
    paragraph_extras(None, values)
    allowed = {*names, "alignment", "line_spacing", "keep_with_next", "keep_together", "page_break_before", "widow_control"} | PARAGRAPH_PROPERTIES
    if values.keys() - allowed:
        raise ValueError("Unknown paragraph properties: " + ", ".join(sorted(values.keys() - allowed)))
    converted = {}
    alignments = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
                  "right": WD_ALIGN_PARAGRAPH.RIGHT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
                  "distribute": WD_ALIGN_PARAGRAPH.DISTRIBUTE}
    for name, value in values.items():
        if name in PARAGRAPH_PROPERTIES:
            continue
        if name in names:
            from docx.shared import Pt
            # Negative indents are legal; spacing must remain non-negative.
            if value is not None and name.endswith("indent_pt"):
                if not math.isfinite(float(value)):
                    raise ValueError("Indents must be finite.")
                value = Pt(float(value))
            elif value is not None:
                value = _points(value)
            name = names[name]
        elif name == "alignment" and value is not None:
            if value not in alignments:
                raise ValueError("Unsupported paragraph alignment.")
            value = alignments[value]
        elif name == "line_spacing" and value is not None:
            if float(value) <= 0 or not math.isfinite(float(value)):
                raise ValueError("Line-spacing multiples must be positive and finite.")
            value = float(value)
        converted[name] = value
    if target is not None:
        for name, value in converted.items():
            setattr(target.paragraph_format, name, value)
        paragraph_extras(target, {key: value for key, value in values.items() if key in PARAGRAPH_PROPERTIES})


def _blocks(root):
    for node in root:
        if node.tag in {_tag("p"), _tag("tbl")}:
            if node.tag == _tag("p") and node.find("./" + _tag("pPr") + "/" + _tag("sectPr")) is not None and _only_properties(node):
                continue
            yield node
        elif node.tag in {_tag("sdt"), _tag("sdtContent"), _tag("customXml"), _tag("ins"), _tag("moveTo")}:
            yield from _blocks(node)


def _only_properties(node):
    from .identity import private_ids
    private = private_ids(node)
    return all(child.tag == _tag("pPr") or child.tag in {_tag("bookmarkStart"), _tag("bookmarkEnd")} and child.get(_tag("id")) in private for child in node)


def _balanced(node):
    if any(True for _ in node.iter(_tag("sectPr"))):
        raise ValueError("Section-break paragraphs cannot be deleted or moved; edit their content instead.")
    depth = 0
    for field in node.iter(_tag("fldChar")):
        kind = field.get(_tag("fldCharType"))
        depth += 1 if kind == "begin" else -1 if kind == "end" else 0
        if depth < 0:
            raise ValueError("This block ends a field that starts outside it.")
    if depth:
        raise ValueError("This block starts a field that ends outside it.")
    for start, end in (("bookmarkStart", "bookmarkEnd"), ("commentRangeStart", "commentRangeEnd")):
        if {item.get(_tag("id")) for item in node.iter(_tag(start))} != {item.get(_tag("id")) for item in node.iter(_tag(end))}:
            raise ValueError("This block crosses a bookmark or comment boundary.")


class WordEditor:
    """A transaction-scoped editor. Obtain with WordDocument.edit_word_details()."""

    def __init__(self, backend):
        self._backend = backend
        self._active = True
        self._textbox_pairs = []
        self._note_upgrades = []
        self._bindings = [(node, part.element, self._binding_text(node)) for part in backend.part.package.parts
                          if getattr(part, "element", None) is not None for node in part.element.iter(_tag("sdt"))
                          if node.find("./" + _tag("sdtPr") + "/" + _tag("dataBinding")) is not None]

    @staticmethod
    def _binding_text(node):
        content = node.find(_tag("sdtContent"))
        return "".join("\t" if item.tag == _tag("tab") else "\n" if item.tag == _tag("br") else item.text or ""
                       for item in content.iter() if item.tag in {_tag("t"), _tag("tab"), _tag("br")}) if content is not None else ""

    def _check(self):
        if not self._active:
            raise RuntimeError("This Word edit handle has expired. Open a new edit_word_details() context.")

    def _finish(self):
        for primary, fallback, before, root in self._textbox_pairs:
            if not _attached(primary, root) or etree.tostring(primary) == before:
                continue
            for child in list(fallback):
                fallback.remove(child)
            fallback.extend(deepcopy(child) for child in primary)
            for node in fallback.iterancestors():
                if etree.QName(node).localname == "AlternateContent":
                    break
                for attribute in list(node.attrib):
                    if etree.QName(attribute).localname == "gfxdata":
                        node.attrib.pop(attribute)
        changed = [node for node, root, before in self._bindings
                   if _attached(node, root) and self._binding_text(node) != before]
        if changed:
            from .features import synchronize_bindings

            synchronize_bindings(self._backend, controls=changed)
        for original, upgraded, before in self._note_upgrades:
            if etree.tostring(upgraded.element) == before:
                for part in list(self._backend.part.package.parts):
                    for relation in part.rels.values():
                        if not relation.is_external and relation.target_part is upgraded:
                            relation._target = original
        from .identity import NS as identity_namespace, stamp
        if any(getattr(part, "element", None) is not None and part.element.get("{" + identity_namespace + "}schema")
               for part in self._backend.part.package.parts):
            stamp(self._backend)

    @property
    def body(self) -> WordContainer:
        self._check()
        return WordContainer(self, self._backend.element.body, self._backend)

    def section(self, index: int = 0) -> WordSection:
        from .editing_stories import WordSection
        self._check()
        return WordSection(self, self._backend.element.body, self._backend, section=self._backend.sections[index]._sectPr)

    def insert_section(self, index: int, *, break_type: str = "next_page") -> WordSection:
        from .editing_stories import insert_section
        return insert_section(self, index, break_type)

    def create_list(self, levels: Sequence[dict[str, Any]] | None = None) -> WordList:
        from .editing_fields import create_list
        self._check()
        return create_list(self, levels)

    def update_fields(self, **options):
        from .field_updates import update_fields
        return update_fields(self, **options)

    def create_style(self, name, **properties):
        from .editing_styles import create_style
        return create_style(self, name, **properties)

    def configure_style(self, name, **properties):
        from .editing_styles import configure_style
        return configure_style(self, name, **properties)

    def delete_style(self, name, *, replacement):
        from .editing_styles import delete_style
        return delete_style(self, name, replacement=replacement)

    def format_theme(self, **properties):
        from .editing_styles import format_theme
        return format_theme(self, **properties)

    @property
    def bookmarks(self):
        from .editing_navigation import bookmarks
        self._check()
        return bookmarks(self)

    def object(self, object_id):
        from .endpoint import resolve_object
        self._check()
        return resolve_object(self, object_id)

    def footnote(self, index: int = 0) -> WordNote:
        from .editing_stories import note_at
        return note_at(self, "footnote", index)

    def endnote(self, index: int = 0) -> WordNote:
        from .editing_stories import note_at
        return note_at(self, "endnote", index)

    def header(self, section_index: int = 0, *, variant: str = "default", unlink: bool = False) -> WordContainer:
        return self._story("header", section_index, variant, unlink)

    def footer(self, section_index: int = 0, *, variant: str = "default", unlink: bool = False) -> WordContainer:
        return self._story("footer", section_index, variant, unlink)

    def _story(self, kind, index, variant, unlink):
        self._check()
        if variant not in {"default", "first", "even"}:
            raise ValueError("Story variant must be default, first, or even.")
        section = self._backend.sections[index]
        story = getattr(section, {"default": "", "first": "first_page_", "even": "even_page_"}[variant] + kind)
        if index and story.is_linked_to_previous and not unlink:
            raise ValueError("This story is inherited; pass unlink=True to edit this section independently.")
        if unlink and index and story.is_linked_to_previous:
            source = story.part
            children = [deepcopy(child) for child in source.element]
            story.is_linked_to_previous = False
            target = story.part
            relationships = {identity: target.relate_to(rel.target_ref if rel.is_external else rel.target_part, rel.reltype, is_external=rel.is_external)
                             for identity, rel in source.rels.items()}
            for child in list(target.element):
                target.element.remove(child)
            for child in children:
                for node in child.iter():
                    for key, value in list(node.attrib.items()):
                        if key.startswith("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}") and value in relationships:
                            node.set(key, relationships[value])
                target.element.append(child)
            self._bindings.extend((node, target.element, self._binding_text(node)) for node in target.element.iter(_tag("sdt"))
                                  if node.find("./" + _tag("sdtPr") + "/" + _tag("dataBinding")) is not None)
        elif unlink:
            story.is_linked_to_previous = False
        if variant == "first":
            section.different_first_page_header_footer = True
        elif variant == "even":
            self._backend.settings.odd_and_even_pages_header_footer = True
        return WordContainer(self, story._element, story)

    def format_style(self, name: str, *, font: dict[str, Any] | None = None, paragraph: dict[str, Any] | None = None) -> None:
        self._check()
        _font(None, font or {})
        _paragraph_format(None, paragraph or {})
        style = self._backend.styles[name]
        if paragraph and not hasattr(style, "paragraph_format"):
            raise ValueError("Paragraph properties require a paragraph style.")
        _font(style, font or {})
        if paragraph:
            _paragraph_format(style, paragraph)


class WordContainer:
    """Body, section, table cell, header/footer, or text-box content."""

    def __init__(self, editor, root, parent, *, section=None):
        self._editor, self._root, self._parent, self._section = editor, root, parent, section

    @property
    def part(self):
        return self._parent.part

    @property
    def object_id(self):
        from .identity import identify
        self._check()
        return identify(self._root)

    def range(self, start_paragraph, start_offset, end_paragraph, end_offset):
        from .editing_ranges import WordRange
        return WordRange(self.paragraph(start_paragraph), start_offset, self.paragraph(end_paragraph), end_offset)

    def find_text(self, text, *, regex=False, case_sensitive=True):
        from .editing_ranges import find_text
        return find_text(self, text, regex=regex, case_sensitive=case_sensitive)

    def replace_all(self, old, new, *, case_sensitive=True):
        matches = self.find_text(old, case_sensitive=case_sensitive)
        for match in reversed(matches):
            match.replace(new)
        return len(matches)

    @property
    def content_controls(self):
        from .editing_content import controls
        return controls(self)

    def wrap_control(self, start, end=None, **properties):
        from .editing_content import wrap_control
        return wrap_control(self, start, end, **properties)

    def insert_control(self, index, **properties):
        from .editing_content import insert_control
        return insert_control(self, index, **properties)

    @property
    def hyperlinks(self):
        from .editing_navigation import WordLink, WordLinkField, link_instruction
        self._check()
        links = [WordLink(self, node) for root in self._nodes() for node in root.iter(_tag("hyperlink"))]
        links.extend(WordLinkField(field) for field in self.fields if link_instruction(field.instruction) is not None)
        return links

    def _check(self):
        self._editor._check()
        if not _attached(self._root, self.part.element):
            raise RuntimeError("This Word container was removed from its document.")

    def _nodes(self):
        self._check()
        if self._section is None:
            return list(_blocks(self._root))
        sections = list(self._editor._backend.sections)
        index = next((i for i, section in enumerate(sections) if section._sectPr is self._section), None)
        if index is None:
            raise RuntimeError("This Word section was removed.")
        children = list(self._root)
        previous = sections[index - 1]._sectPr.getparent().getparent() if index else None
        ending = self._section if self._section.getparent() is self._root else self._section.getparent().getparent()
        begin = children.index(previous) + 1 if previous is not None else 0
        stop = children.index(ending) + (ending.tag != _tag("sectPr"))
        return [node for child in children[begin:stop] for node in _blocks_holder(child)]

    @property
    def blocks(self) -> list[WordParagraph | WordTable]:
        return [WordParagraph(self, node) if node.tag == _tag("p") else WordTable(self, node) for node in self._nodes()]

    def paragraph(self, index: int) -> WordParagraph:
        return [block for block in self.blocks if isinstance(block, WordParagraph)][index]

    def table(self, index: int) -> WordTable:
        return [block for block in self.blocks if isinstance(block, WordTable)][index]

    @property
    def pictures(self) -> list[WordPicture]:
        return [picture for block in self.blocks if isinstance(block, WordParagraph) for picture in block.pictures]

    @property
    def fields(self) -> list[WordField]:
        from .editing_fields import find_fields
        self._check()
        return find_fields(self)

    def textbox(self, index: int = 0) -> WordContainer:
        from .textboxes import find
        from docx.blkcntnr import BlockItemContainer

        self._check()
        boxes = []
        for node in self._nodes():
            if node.tag == _tag("p"):
                for box in find(node, ".//w:txbxContent"):
                    if not any(etree.QName(ancestor).localname == "Fallback" for ancestor in box.iterancestors()):
                        boxes.append(box)
        box = boxes[index]
        alternate = next((node for node in box.iterancestors() if etree.QName(node).localname == "AlternateContent"), None)
        if alternate is not None:
            for fallback in find(alternate, "./mc:Fallback//w:txbxContent"):
                if not any(primary is box for primary, _, _, _ in self._editor._textbox_pairs):
                    self._editor._textbox_pairs.append((box, fallback, etree.tostring(box), self.part.element))
        return WordContainer(self._editor, box, BlockItemContainer(box, self._parent))

    def _insert(self, index, node):
        nodes = self._nodes()
        if not 0 <= index <= len(nodes):
            raise IndexError("Block insertion index is out of range.")
        if index < len(nodes):
            self._reference(nodes, index).addprevious(node)
        elif self._section is not None and self._section.getparent() is not self._root:
            end = self._section.getparent().getparent()
            if not _only_properties(end):
                self._section.getparent().remove(self._section)
                end.addnext(node)
                if node.tag == _tag("p"):
                    node.get_or_add_pPr().append(self._section)
                else:
                    from docx.oxml import OxmlElement
                    carrier = OxmlElement("w:p")
                    carrier.get_or_add_pPr().append(self._section)
                    node.addnext(carrier)
            else:
                end.addprevious(node)
        else:
            ending = self._root.find(_tag("sectPr"))
            if ending is not None:
                ending.addprevious(node)
            else:
                self._root.append(node)

    def _reference(self, nodes, index):
        reference = nodes[index]
        while reference.getparent() is not self._root:
            reference = reference.getparent()
            if reference is None:
                raise ValueError("Insertion target is outside this container.")
        if reference is not nodes[index] and next(_blocks(reference), None) is not nodes[index]:
            raise ValueError("Cannot insert between blocks inside a wrapped content control from its outer container.")
        return reference

    def insert_paragraph(self, index: int, text: str = "", *, style: str | None = None) -> WordParagraph:
        from docx.oxml import OxmlElement
        from docx.text.paragraph import Paragraph

        self._nodes()  # Check scope before creating anything.
        node = OxmlElement("w:p")
        paragraph = Paragraph(node, self._parent)
        if style is not None:
            paragraph.style = style
        paragraph.add_run(text)
        self._insert(index, node)
        return WordParagraph(self, node)

    def append_paragraph(self, text: str = "", *, style: str | None = None) -> WordParagraph:
        return self.insert_paragraph(len(self._nodes()), text, style=style)

    def insert_table(self, index: int, *, rows: int, cols: int, values: Sequence[Sequence[object]] | None = None, style: str | None = None) -> WordTable:
        from docx.oxml.table import CT_Tbl
        from docx.shared import Inches
        from docx.table import Table

        if rows < 1 or cols < 1 or (values and (len(values) > rows or any(len(row) > cols for row in values))):
            raise ValueError("Table dimensions/values do not fit.")
        native = Table(CT_Tbl.new_tbl(rows, cols, Inches(6)), self._parent)
        if style is not None:
            native.style = style
        for row, entries in enumerate(values or []):
            for col, value in enumerate(entries):
                native.cell(row, col).text = str(value)
        self._insert(index, native._tbl)
        self._ensure_tail()
        return WordTable(self, native._tbl)

    def _ensure_tail(self):
        if self._root.tag in {_tag("tc"), _tag("hdr"), _tag("ftr"), _tag("txbxContent"), _tag("footnote"), _tag("endnote")} and (not len(self._root) or self._root[-1].tag != _tag("p")):
            from docx.oxml import OxmlElement
            self._root.append(OxmlElement("w:p"))


def _blocks_holder(node):
    if node.tag in {_tag("p"), _tag("tbl")}:
        if not (node.tag == _tag("p") and node.find("./" + _tag("pPr") + "/" + _tag("sectPr")) is not None and _only_properties(node)):
            yield node
    else:
        yield from _blocks(node)


class _Block:
    def __init__(self, owner, node):
        self._owner, self._node = owner, node

    @property
    def object_id(self):
        from .identity import identify
        self._check()
        return identify(self._node)

    def _check(self):
        self._owner._check()
        if not _attached(self._node, self._owner.part.element):
            raise RuntimeError("This Word object was removed from its document.")

    def delete(self) -> None:
        self._check()
        if self._node.getparent() is not self._owner._root:
            raise ValueError("Wrapped blocks require text editing or explicit raw XML structural editing.")
        _balanced(self._node)
        self._node.getparent().remove(self._node)
        self._owner._ensure_tail()

    def move(self, index: int, *, destination: WordContainer | None = None) -> Self:
        self._check()
        if self._node.getparent() is not self._owner._root:
            raise ValueError("Wrapped blocks require text editing or explicit raw XML structural editing.")
        target = destination or self._owner
        if target._editor is not self._owner._editor:
            raise ValueError("Moves must use the same active editor.")
        nodes = [node for node in target._nodes() if node is not self._node]
        if not 0 <= index <= len(nodes):
            raise IndexError("Block move index is out of range.")
        if index < len(nodes):
            target._reference(nodes, index)
        if target._root is self._node or self._node in target._root.iterancestors():
            raise ValueError("An object cannot be moved inside itself.")
        _balanced(self._node)
        if target.part is not self._owner.part:
            from .editing_navigation import remap_relationships, R
            references = {_tag(name) for name in ("footnoteReference", "endnoteReference", "footnoteRef", "endnoteRef")}
            if any(node.tag in references for node in self._node.iter()):
                raise ValueError("Note references cannot move to a different document part.")
            if any(value not in self._owner.part.rels for node in self._node.iter() for key, value in node.attrib.items() if key.startswith("{" + R + "}")):
                raise ValueError("Moved content contains an unresolved relationship.")
            remap_relationships(self._node, self._owner.part, target.part)
            highest = max([0] + [int(node.get(_tag("id"), "0")) for node in target.part.element.iter(_tag("bookmarkStart"))])
            bookmark_ids = {}
            for node in self._node.iter(_tag("bookmarkStart")):
                highest += 1
                bookmark_ids[node.get(_tag("id"))] = str(highest)
            for node in self._node.iter():
                if node.tag in {_tag("bookmarkStart"), _tag("bookmarkEnd")} and node.get(_tag("id")) in bookmark_ids:
                    node.set(_tag("id"), bookmark_ids[node.get(_tag("id"))])
            # Numbering/styles are package-wide; relationship IDs are part-local.
            for node in self._node.iter(_tag("sdt")):
                if node.find("./" + _tag("sdtPr") + "/" + _tag("dataBinding")) is not None:
                    before = next((before for existing, _, before in self._owner._editor._bindings if existing is node), self._owner._editor._binding_text(node))
                    self._owner._editor._bindings.append((node, target.part.element, before))
        self._node.getparent().remove(self._node)
        target._insert(index, self._node)
        self._owner._ensure_tail()
        self._owner = target
        return self


class WordParagraph(_Block):
    def range(self, start=0, end=None):
        from .editing_ranges import WordRange
        return WordRange(self, start, self, len(self.text) if end is None else end)

    def insert_hyperlink(self, offset, target, text):
        from .editing_ranges import _split, _set_content
        from .editing_navigation import hyperlink
        from docx.text.run import Run
        from docx.oxml import OxmlElement
        before, after = _split(self, offset)
        run = OxmlElement("w:r")
        Run(run, self._native).text = text
        link = hyperlink(self._owner, target, [run])
        _set_content(self, before + [link._node] + after)
        return link

    @property
    def equations(self):
        from .editing_math import WordEquation
        self._check()
        return [WordEquation(self._owner, node) for node in self._node.iter("{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath")]

    def add_equation(self, script):
        from .editing_math import WordEquation
        from .equations import hwp_to_omml
        node = hwp_to_omml(script)
        self._native._p.append(node)
        return WordEquation(self._owner, node)
    @property
    def fields(self) -> list[WordField]:
        from .editing_fields import find_fields
        self._check()
        return find_fields(self._owner, self._node)

    def add_field(self, instruction: str, result: str = "", *, locked: bool = False, complex_field: bool = False) -> WordField:
        from .editing_fields import add_field
        return add_field(self, instruction, result, locked, complex_field)

    def set_list(self, listing: WordList, *, level: int = 0, start: int | None = None) -> WordParagraph:
        from .editing_fields import set_list
        set_list(self, listing, level, start)
        return self

    @property
    def listing(self) -> WordList | None:
        from .editing_fields import paragraph_list
        return paragraph_list(self)

    def remove_list(self) -> WordParagraph:
        from .editing_fields import set_list
        set_list(self, None, 0, None)
        return self

    def add_note(self, text: str = "", *, kind: str = "footnote") -> WordNote:
        from .editing_stories import add_note
        return add_note(self, text, kind)

    @property
    def _native(self):
        from docx.text.paragraph import Paragraph
        self._check()
        return Paragraph(self._node, self._owner._parent)

    @property
    def text(self) -> str:
        from .preservation import _visible_slots
        return "".join(slot[3] for slot in _visible_slots(self._native))

    def replace_text(self, old: str, new: str, *, count: int = -1) -> int:
        from .preservation import _patch_paragraph
        if not old:
            raise ValueError("old must be non-empty.")
        text = self.text
        amount = text.count(old) if count < 0 else min(text.count(old), count)
        _patch_paragraph(self._node.getroottree().getroot(), self._native, text, text.replace(old, new, count))
        return amount

    def set_text(self, text: str) -> WordParagraph:
        """Replace literal text while retaining runs and non-text XML objects."""
        from .preservation import _patch_paragraph
        self._check()
        if list(self._node.iter(_tag("fldChar"))) or list(self._node.iter(_tag("fldSimple"))):
            raise ValueError("Use literal replace_text() to preserve this paragraph's fields.")
        _patch_paragraph(self._node.getroottree().getroot(), self._native, self.text, text)
        return self

    def format(self, *, style: str | None = None, **properties: Any) -> WordParagraph:
        _paragraph_format(None, properties)
        paragraph = self._native
        if style is not None:
            paragraph.style = style
        _paragraph_format(paragraph, properties)
        return self

    def format_text(self, start: int = 0, end: int | None = None, **properties: Any) -> WordParagraph:
        from docx.text.run import Run
        from .preservation import _visible_slots

        _font(None, properties)
        paragraph = self._native
        root = self._node.getroottree().getroot()
        slots = _visible_slots(paragraph)
        length = sum(len(slot[3]) for slot in slots)
        end = length if end is None else end
        if not 0 <= start <= end <= length:
            raise IndexError("Text formatting range is out of bounds.")
        if not properties:
            return self
        nodes = {root.find(path): (a, b, text) for path, a, b, text in slots}
        runs = {node.getparent() for node, (a, b, _) in nodes.items() if a < end and b > start}
        # Split only touched runs; retain non-text children, relationships,
        # fields and surrounding wrappers at their original position.
        for original in runs:
            clones = []
            for child in original:
                if child.tag == _tag("rPr"):
                    continue
                segments = [(deepcopy(child), False)]
                if child in nodes:
                    a, b, text = nodes[child]
                    left, right = max(a, start), min(b, end)
                    if left < right:
                        segments = []
                        for lo, hi, selected in ((a, left, False), (left, right, True), (right, b, False)):
                            if lo >= hi:
                                continue
                            copy = deepcopy(child)
                            if child.tag == _tag("t"):
                                copy.text = text[lo - a:hi - a]
                                copy.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                            segments.append((copy, selected))
                for item, selected in segments:
                    clone = original.makeelement(original.tag, attrib=dict(original.attrib))
                    if original.rPr is not None:
                        clone.append(deepcopy(original.rPr))
                    clone.append(item)
                    if selected:
                        _font(Run(clone, paragraph), properties)
                    clone.tail = None
                    clones.append(clone)
            for clone in clones:
                original.addprevious(clone)
            if clones:
                clones[-1].tail = original.tail
            original.getparent().remove(original)
        return self

    @property
    def pictures(self) -> list[WordPicture]:
        self._check()
        return [WordPicture(self._owner, node) for node in self._node.iter(_tag("drawing"))
                if node.xpath(".//pic:pic") and not any(etree.QName(a).localname == "Fallback" for a in node.iterancestors())]

    def add_picture(self, data: bytes, *, width_pt: float, height_pt: float | None = None, alt_text: str | None = None) -> WordPicture:
        width = _points(width_pt, positive=True)
        height = _points(height_pt, positive=True) if height_pt is not None else None
        shape = self._native.add_run().add_picture(BytesIO(data), width=width, height=height)
        if alt_text is not None:
            shape._inline.docPr.set("descr", alt_text)
        return WordPicture(self._owner, shape._inline.getparent())


class WordTable(_Block):
    @property
    def _native(self):
        from docx.table import Table
        self._check()
        return Table(self._node, self._owner._parent)

    def cell(self, row: int, col: int) -> WordContainer:
        cell = self._cell(row, col)
        return WordContainer(self._owner._editor, cell._tc, cell)

    def _cell(self, row, col):
        table = self._native
        if not 0 <= row < len(table.rows) or not 0 <= col < len(table.columns):
            raise IndexError("Table cell is out of range.")
        native_row = table.rows[row]
        offset = col - native_row.grid_cols_before
        if not 0 <= offset < len(native_row.cells):
            raise IndexError("This grid position has no cell in the row.")
        return native_row.cells[offset]

    def set_cell_text(self, row: int, col: int, text: str) -> WordTable:
        cell = self.cell(row, col)
        paragraphs = [block for block in cell.blocks if isinstance(block, WordParagraph)]
        if len(cell.blocks) != 1 or len(paragraphs) != 1:
            raise ValueError("Use cell().paragraph() for cells containing multiple paragraphs or nested tables.")
        paragraph = paragraphs[0]
        from .preservation import _patch_paragraph
        if list(paragraph._node.iter(_tag("fldChar"))) or list(paragraph._node.iter(_tag("fldSimple"))):
            raise ValueError("Use literal replace_text() to preserve this cell's fields.")
        _patch_paragraph(self._node.getroottree().getroot(), paragraph._native, paragraph.text, text)
        return self

    def format(self, *, style: str | None = None, widths_pt: Sequence[float] | None = None, repeat_header: bool | None = None) -> WordTable:
        from docx.oxml import OxmlElement
        table = self._native
        widths = [_points(value, positive=True) for value in widths_pt] if widths_pt is not None else None
        if widths is not None and len(widths) != len(table.columns):
            raise ValueError("Column widths must match the table grid.")
        if style is not None:
            table.style = style
        if widths is not None:
            table.autofit = False
            for column, width in zip(table.columns, widths):
                column.width = width
            for row in table.rows:
                seen = set()
                for col, cell in enumerate(row.cells, start=row.grid_cols_before):
                    if cell._tc in seen:
                        continue
                    seen.add(cell._tc)
                    span = int(cell._tc.grid_span)
                    cell.width = sum(widths[col:col + span])
        if repeat_header is not None:
            pr = table.rows[0]._tr.get_or_add_trPr()
            for node in list(pr.findall(_tag("tblHeader"))):
                pr.remove(node)
            if repeat_header:
                pr.append(OxmlElement("w:tblHeader"))
        return self

    def format_cell(self, row: int, col: int, *, fill_color: str | None = None, vertical_alignment: str | None = None, padding_pt: dict[str, float] | None = None, borders: dict | None = None) -> WordTable:
        from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
        from docx.oxml import OxmlElement
        colors = _color(fill_color) if fill_color is not None else None
        alignments = {"top": WD_CELL_VERTICAL_ALIGNMENT.TOP, "center": WD_CELL_VERTICAL_ALIGNMENT.CENTER,
                      "bottom": WD_CELL_VERTICAL_ALIGNMENT.BOTTOM}
        if vertical_alignment is not None and vertical_alignment not in alignments:
            raise ValueError("Unsupported cell vertical alignment.")
        padding = dict(padding_pt or {})
        if padding.keys() - {"left", "right", "top", "bottom"}:
            raise ValueError("Padding keys must be left, right, top, or bottom.")
        margins = {key: round(_points(value) / 635) for key, value in padding.items()}
        cell = self._cell(row, col)
        if borders is not None:
            from .editing_format import cell_borders
            from docx.table import _Cell
            cell_borders(cell, borders)
            # A visible merged cell is represented by multiple physical cells.
            # Word reads side/bottom borders from the continuation rows too.
            if cell._tc.vMerge == "restart":
                following = cell._tc._tr_below
                while following is not None:
                    continued = next((tc for tc in following.tc_lst if tc.grid_offset == cell._tc.grid_offset), None)
                    if continued is None or continued.vMerge != "continue" or continued.grid_span != cell._tc.grid_span:
                        break
                    cell_borders(_Cell(continued, self._native), borders)
                    following = continued._tr_below
        pr = cell._tc.get_or_add_tcPr()
        if colors is not None:
            for node in list(pr.findall(_tag("shd"))):
                pr.remove(node)
            node = OxmlElement("w:shd")
            node.attrib.update({_tag("val"): "clear", _tag("fill"): colors})
            pr.append(node)
        if vertical_alignment is not None:
            cell.vertical_alignment = alignments[vertical_alignment]
        if margins:
            node = pr.find(_tag("tcMar"))
            if node is None:
                node = OxmlElement("w:tcMar")
                pr.append(node)
            for key, value in margins.items():
                for previous in list(node.findall(_tag(key))):
                    node.remove(previous)
                edge = OxmlElement("w:" + key)
                edge.attrib.update({_tag("w"): str(value), _tag("type"): "dxa"})
                node.append(edge)
        return self

    def format_row(self, index: int, *, height_pt: float | None = None, exact: bool = False, allow_split: bool | None = None) -> WordTable:
        from docx.enum.table import WD_ROW_HEIGHT_RULE
        height = _points(height_pt, positive=True) if height_pt is not None else None
        row = self._native.rows[index]
        row.height = height
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY if height is not None and exact else WD_ROW_HEIGHT_RULE.AT_LEAST if height is not None else None
        if allow_split is not None:
            from .editing_format import replace
            replace(row._tr.get_or_add_trPr(), "cantSplit", {"val": "0" if allow_split else "1"})
        return self

    def format_layout(self, **properties) -> WordTable:
        from .editing_table_layout import format_layout
        format_layout(self, **properties)
        return self

    def split(self, row_index):
        from .editing_table_layout import split_table
        return split_table(self, row_index)

    def join(self, following, *, adopt_style=False):
        from .editing_table_layout import join_tables
        join_tables(self, following, adopt_style=adopt_style)
        return self

    def add_row(self, values: Sequence[object] | None = None, *, index: int | None = None) -> WordTable:
        from .editing_tables import insert_axis
        insert_axis(self, "row", index, values)
        return self

    def delete_row(self, index: int) -> WordTable:
        from .editing_tables import delete_axis
        delete_axis(self, "row", index)
        return self

    def add_column(self, values: Sequence[object] | None = None, *, index: int | None = None, width_pt: float = 72) -> WordTable:
        from .editing_tables import insert_axis
        insert_axis(self, "column", index, values, width_pt)
        return self

    def delete_column(self, index: int) -> WordTable:
        from .editing_tables import delete_axis
        delete_axis(self, "column", index)
        return self

    def split_cell(self, row: int, col: int) -> WordTable:
        from .editing_tables import split_cell
        split_cell(self, row, col)
        return self

    def merge(self, top: int, left: int, bottom: int, right: int) -> WordTable:
        table = self._native
        if top < 0 or left < 0 or bottom < top or right < left or bottom >= len(table.rows) or right >= len(table.columns):
            raise IndexError("Invalid merge rectangle.")
        for row in range(top, bottom + 1):
            for col in range(left, right + 1):
                _balanced(self._cell(row, col)._tc)
        self._cell(top, left).merge(self._cell(bottom, right))
        return self


class WordPicture(_Block):
    def resize(self, *, width_pt: float | None = None, height_pt: float | None = None) -> WordPicture:
        self._check()
        extent = self._node.xpath("./wp:inline/wp:extent | ./wp:anchor/wp:extent")[0]
        old_width, old_height = int(extent.get("cx")), int(extent.get("cy"))
        if width_pt is None and height_pt is None:
            return self
        width = _points(width_pt, positive=True) if width_pt is not None else round(_points(height_pt, positive=True) * old_width / old_height)
        height = _points(height_pt, positive=True) if height_pt is not None else round(width * old_height / old_width)
        for node in self._node.xpath("./wp:inline/wp:extent | ./wp:anchor/wp:extent | .//pic:spPr/a:xfrm/a:ext"):
            node.set("cx", str(width))
            node.set("cy", str(height))
        return self

    def replace_image(self, data: bytes) -> WordPicture:
        from docx.oxml.ns import qn
        self._check()
        rid, _image = self._owner.part.get_or_add_image(BytesIO(data))
        for blip in self._node.xpath(".//pic:blipFill/a:blip"):
            blip.set(qn("r:embed"), rid)
            blip.attrib.pop(qn("r:link"), None)
        return self

    def crop(self, *, left: float = 0, top: float = 0, right: float = 0, bottom: float = 0) -> WordPicture:
        from docx.oxml import OxmlElement
        values = dict(left=left, top=top, right=right, bottom=bottom)
        if any(not math.isfinite(float(value)) or not 0 <= value < 1 for value in values.values()) or left + right >= 1 or top + bottom >= 1:
            raise ValueError("Crop fractions must leave a positive visible image rectangle.")
        self._check()
        fill = self._node.xpath(".//pic:blipFill")[0]
        for node in list(fill.findall("{" + A + "}srcRect")):
            fill.remove(node)
        crop = OxmlElement("a:srcRect")
        for edge, key in (("left", "l"), ("top", "t"), ("right", "r"), ("bottom", "b")):
            crop.set(key, str(round(float(values[edge]) * 100000)))
        fill.insert(1, crop)
        return self

    def rotate(self, degrees: float) -> WordPicture:
        if not math.isfinite(float(degrees)):
            raise ValueError("Rotation must be finite.")
        self._check()
        self._node.xpath(".//pic:spPr/a:xfrm")[0].set("rot", str(round(float(degrees) % 360 * 60000)))
        return self

    def position(self, *, x_pt: float = 0, y_pt: float = 0, horizontal_from: str = "margin", vertical_from: str = "paragraph", wrap: str = "square") -> WordPicture:
        from ..._hancom_models import Picture
        from .pictures import _position
        horizontal = {"page": "PAPER", "margin": "PAGE", "column": "COLUMN"}
        vertical = {"page": "PAPER", "margin": "PAGE", "paragraph": "PARA", "line": "LINE"}
        wrapping = {"square": "SQUARE", "top_and_bottom": "TOP_AND_BOTTOM", "front": "IN_FRONT_OF_TEXT", "behind": "BEHIND_TEXT"}
        if horizontal_from not in horizontal or vertical_from not in vertical or wrap not in wrapping:
            raise ValueError("Unsupported picture reference frame or wrap mode.")
        if not all(math.isfinite(float(value)) for value in (x_pt, y_pt)):
            raise ValueError("Picture offsets must be finite.")
        self._check()
        model = Picture("", b"", layout={"treatAsChar": "0", "horzRelTo": horizontal[horizontal_from],
            "vertRelTo": vertical[vertical_from], "horzOffset": str(round(float(x_pt) * 100)),
            "vertOffset": str(round(float(y_pt) * 100)), "textWrap": wrapping[wrap]},
            out_margins={edge: round(int(self._node[0].get("dist" + suffix, "0")) / 127)
                         for edge, suffix in (("left", "L"), ("right", "R"), ("top", "T"), ("bottom", "B"))})
        _position(self._node, model)
        return self

    def delete(self) -> None:
        self._check()
        parent = self._node.getparent()
        parent.remove(self._node)
