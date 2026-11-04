"""Native list definitions and separate field-code/result editing."""
from __future__ import annotations

from copy import deepcopy

from .editing import _tag, _points, _attached
from .editing_format import element, replace


def _numbering(editor, create=False):
    from docx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
    from docx.opc.packuri import PackURI
    from docx.parts.numbering import NumberingPart
    from docx.oxml import OxmlElement
    part = next((rel.target_part for rel in editor._backend.part.rels.values() if rel.reltype == RT.NUMBERING and not rel.is_external), None)
    if part is None and create:
        part = NumberingPart(PackURI("/word/numbering.xml"), CT.WML_NUMBERING, OxmlElement("w:numbering"), editor._backend.part.package)
        editor._backend.part.relate_to(part, RT.NUMBERING)
    return part.element if part is not None else None


def _next_id(root, tag, attribute, minimum=0):
    return max([minimum - 1] + [int(node.get(_tag(attribute))) for node in root if node.tag == _tag(tag)]) + 1


class WordList:
    """An actual numbering instance, reusable to continue the same list."""
    def __init__(self, editor, node):
        self._editor, self._node = editor, node

    @property
    def object_id(self):
        from .identity import identify
        _ = self.id
        return identify(self._node)

    @property
    def id(self):
        self._editor._check()
        root = _numbering(self._editor)
        if root is None or not _attached(self._node, root):
            raise RuntimeError("Numbering instance was removed.")
        return int(self._node.get(_tag("numId")))

    def restart(self, start=1, *, level=0):
        if not isinstance(level, int) or not 0 <= level <= 8 or not isinstance(start, int) or isinstance(start, bool) or not 0 <= start <= 2147483647:
            raise ValueError("List restart needs level 0..8 and a non-negative integer start.")
        _ = self.id
        root = _numbering(self._editor)
        node = deepcopy(self._node)
        node.set(_tag("numId"), str(_next_id(root, "num", "numId", 1)))
        for override in list(node.findall(_tag("lvlOverride"))):
            if override.get(_tag("ilvl")) == str(level):
                # Keep a per-level format override, replacing only its start.
                node_start = replace(override, "startOverride", {"val": start})
                definition = override.find(_tag("lvl"))
                if definition is not None:
                    definition.addprevious(node_start)
                break
        else:
            override = element(node, "lvlOverride", ilvl=level)
            element(override, "startOverride", val=start)
        root.append(node)
        return WordList(self._editor, node)


def create_list(editor, levels=None):
    from docx.oxml import OxmlElement
    levels = [{"format": "decimal", "text": "%1."}] if levels is None else list(levels)
    if not 1 <= len(levels) <= 9:
        raise ValueError("Lists need between 1 and 9 levels.")
    definitions = []
    for index, value in enumerate(levels):
        if value.keys() - {"format", "text", "start", "left_indent_pt", "hanging_pt", "font_name", "suffix"}:
            raise ValueError("Unknown list level property.")
        fmt = value.get("format", "decimal")
        if fmt not in {"decimal", "decimalZero", "lowerRoman", "upperRoman", "lowerLetter", "upperLetter", "bullet"}:
            raise ValueError("Unsupported list number format.")
        start = value.get("start", 1)
        if not isinstance(start, int) or isinstance(start, bool) or not 0 <= start <= 2147483647:
            raise ValueError("List start must be a non-negative integer.")
        text = value.get("text", "•" if fmt == "bullet" else "%" + str(index + 1) + ".")
        if not isinstance(text, str) or not text:
            raise ValueError("List marker text must be non-empty.")
        suffix = value.get("suffix", "tab")
        if suffix not in {"tab", "space", "nothing"}:
            raise ValueError("List suffix must be tab, space, or nothing.")
        left = _points(value.get("left_indent_pt", 36 * (index + 1)))
        hanging = _points(value.get("hanging_pt", 18))
        node = OxmlElement("w:lvl")
        node.set(_tag("ilvl"), str(index))
        element(node, "start", val=start)
        element(node, "numFmt", val=fmt)
        element(node, "suff", val=suffix)
        element(node, "lvlText", val=text)
        element(node, "lvlJc", val="left")
        pr = element(node, "pPr")
        element(element(pr, "tabs"), "tab", val="num", pos=round(left / 635))
        element(pr, "ind", left=round(left / 635), hanging=round(hanging / 635))
        if value.get("font_name"):
            element(element(node, "rPr"), "rFonts", ascii=value["font_name"], hAnsi=value["font_name"], eastAsia=value["font_name"])
        definitions.append(node)
    root = _numbering(editor, True)
    identity = _next_id(root, "abstractNum", "abstractNumId")
    abstract = OxmlElement("w:abstractNum")
    abstract.set(_tag("abstractNumId"), str(identity))
    element(abstract, "multiLevelType", val="singleLevel" if len(levels) == 1 else "multilevel")
    abstract.extend(definitions)
    first_num = root.find(_tag("num"))
    if first_num is None:
        root.append(abstract)
    else:
        first_num.addprevious(abstract)
    instance = element(root, "num", numId=_next_id(root, "num", "numId", 1))
    element(instance, "abstractNumId", val=identity)
    return WordList(editor, instance)


def paragraph_list(paragraph):
    native = paragraph._native
    properties = [native._p]
    style, seen = native.style, set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        properties.append(style.element)
        style = style.base_style
    identity = next((node.get(_tag("val")) for pr in properties if (node := pr.find("./" + _tag("pPr") + "/" + _tag("numPr") + "/" + _tag("numId"))) is not None), None)
    root = _numbering(paragraph._owner._editor)
    if root is None or identity in {None, "0"}:
        return None
    node = next((node for node in root if node.tag == _tag("num") and node.get(_tag("numId")) == identity), None)
    return WordList(paragraph._owner._editor, node) if node is not None else None


def set_list(paragraph, listing, level, start):
    paragraph._check()
    if isinstance(level, bool) or not isinstance(level, int) or not 0 <= level <= 8:
        raise ValueError("List levels range from 0 to 8.")
    if listing is not None:
        if not isinstance(listing, WordList) or listing._editor is not paragraph._owner._editor:
            raise ValueError("Use a WordList from the same active editor.")
        identity = listing.id
        root = _numbering(listing._editor)
        abstract_id = listing._node.find(_tag("abstractNumId")).get(_tag("val"))
        abstract = next(node for node in root if node.tag == _tag("abstractNum") and node.get(_tag("abstractNumId")) == abstract_id)
        if not any(node.get(_tag("ilvl")) == str(level) for node in abstract.findall(_tag("lvl"))) and not any(node.get(_tag("ilvl")) == str(level) and node.find(_tag("lvl")) is not None for node in listing._node.findall(_tag("lvlOverride"))):
            raise ValueError("This numbering definition has no requested level.")
        if start is not None:
            listing = listing.restart(start, level=level)
            identity = listing.id
    else:
        if start is not None:
            raise ValueError("Removing numbering cannot restart it.")
        identity = 0
    pr = paragraph._native._p.get_or_add_pPr()
    node = replace(pr, "numPr", {})
    element(node, "ilvl", val=level)
    element(node, "numId", val=identity)


def _instruction(value):
    if not isinstance(value, str) or not value.strip() or any(ord(char) < 32 and char not in "\t\n\r" for char in value):
        raise ValueError("A field instruction must be non-empty XML text.")


def _story(node):
    return next((parent for parent in node.iterancestors() if parent.tag in {_tag("txbxContent"), _tag("footnote"), _tag("endnote")}), node.getroottree().getroot())


def find_fields(owner, scope=None):
    nodes = owner._nodes() if scope is None else [scope]
    entries = [item for node in nodes for item in node.iter()]
    allowed_story = _story(owner._root) if owner._root.tag not in {_tag("txbxContent"), _tag("footnote"), _tag("endnote")} else owner._root
    output, stack = [], []
    for node in entries:
        if _story(node) is not allowed_story:
            continue
        if node.tag == _tag("fldSimple"):
            output.append(WordField(owner, node))
        elif node.tag == _tag("fldChar"):
            kind = node.get(_tag("fldCharType"))
            if kind == "begin":
                handle = WordField(owner, node)
                output.append(handle)
                stack.append(handle)
            elif stack and kind == "separate":
                stack[-1]._separator = node
            elif stack and kind == "end":
                stack.pop()._end = node
        elif node.tag == _tag("instrText") and stack and stack[-1]._separator is None:
            stack[-1]._instructions.append(node)
    return [handle for handle in output if handle._simple or handle._end is not None]


def add_field(paragraph, instruction, result, locked, complex_field=False):
    _instruction(instruction)
    paragraph._check()
    if not isinstance(result, str) or not isinstance(complex_field, bool):
        raise ValueError("Field result must be text and complex_field must be boolean.")
    if complex_field:
        holder = paragraph._native._p
        begin = element(element(holder, "r"), "fldChar", fldCharType="begin", dirty="true", fldLock="true" if locked else "false")
        code = element(element(holder, "r"), "instrText")
        code.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        code.text = instruction
        separator = element(element(holder, "r"), "fldChar", fldCharType="separate")
        from docx.text.run import Run
        Run(element(holder, "r"), paragraph._native).text = result
        end = element(element(holder, "r"), "fldChar", fldCharType="end")
        field = WordField(paragraph._owner, begin)
        field._instructions, field._separator, field._end = [code], separator, end
        return field
    node = element(paragraph._native._p, "fldSimple", instr=instruction, dirty="true", fldLock="true" if locked else "false")
    from docx.text.run import Run
    Run(element(node, "r"), paragraph._native).text = result
    return WordField(paragraph._owner, node)


class WordField:
    def __init__(self, owner, node):
        self._owner, self._node = owner, node
        self._simple = node.tag == _tag("fldSimple")
        self._instructions = []
        self._separator = self._end = None

    def _check(self):
        self._owner._check()
        if not _attached(self._node, self._owner.part.element) or not self._simple and (self._end is None or not _attached(self._end, self._owner.part.element)):
            raise RuntimeError("Field was removed or its boundaries changed.")

    @property
    def object_id(self):
        from .identity import identify
        self._check()
        return identify(self._node)

    @property
    def instruction(self):
        self._check()
        return self._node.get(_tag("instr"), "") if self._simple else "".join(node.text or "" for node in self._instructions)

    def set_instruction(self, instruction):
        _instruction(instruction)
        self._check()
        if self._simple:
            self._node.set(_tag("instr"), instruction)
        else:
            if not self._instructions:
                raise ValueError("Field has no literal instruction; edit its nested code explicitly.")
            # A nested code expression must not be replaced by a literal while
            # retaining its nested executable fields.
            all_nodes = list(self._owner.part.element.iter())
            left = all_nodes.index(self._node)
            right = all_nodes.index(self._separator if self._separator is not None else self._end)
            if any(node.tag in {_tag("fldSimple"), _tag("fldChar")} for node in all_nodes[left + 1:right]):
                raise ValueError("This instruction contains nested fields; edit the inner field instead.")
            self._instructions[0].text = instruction
            self._instructions[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            for node in self._instructions[1:]:
                node.text = ""
        self._node.set(_tag("dirty"), "true")
        return self

    def edit_result(self):
        from .editing_content import edit_result
        return edit_result(self)

    def rebuild_toc(self, **options):
        from .field_updates import rebuild_toc
        return rebuild_toc(self, **options)

    @property
    def result_fields(self):
        nodes = set(self._result_nodes())
        return [field for field in self._owner.fields if field._node in nodes]

    def replace_result_text(self, old, new):
        with self.edit_result() as content:
            return content.replace_all(old, new)

    def _result_nodes(self):
        self._check()
        if self._simple:
            return list(self._node.iter())[1:]
        if self._separator is None:
            return []
        nodes = list(self._owner.part.element.iter())
        return nodes[nodes.index(self._separator) + 1:nodes.index(self._end)]

    @property
    def result(self):
        return "".join(node.text or "" if node.tag == _tag("t") else "\t" if node.tag == _tag("tab") else "\n" if node.tag in {_tag("br"), _tag("p")} else "" for node in self._result_nodes())

    def set_result(self, text):
        from docx.text.run import Run
        if not isinstance(text, str):
            raise ValueError("Field result must be a string.")
        nodes = self._result_nodes()
        if any(node.tag in {_tag("fldChar"), _tag("fldSimple"), _tag("drawing"), _tag("pict"), _tag("tbl"), _tag("p"), "{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath"} for node in nodes):
            raise ValueError("Rich/nested field results require editing their inner content.")
        text_nodes = [node for node in nodes if node.tag in {_tag("t"), _tag("tab"), _tag("br")}]
        if text_nodes:
            first = text_nodes[0]
            holder = first.getparent()
            if holder.tag != _tag("r"):
                raise ValueError("Field result requires explicit structured content editing.")
            from docx.oxml import OxmlElement
            run = OxmlElement("w:r")
            Run(run, self._owner._parent).text = text
            for child in list(run):
                first.addprevious(child)
            for node in text_nodes:
                node.getparent().remove(node)
        elif self._simple:
            Run(element(self._node, "r"), self._owner._parent).text = text
        else:
            if self._separator is None:
                separator_run = element(self._end.getparent().getparent(), "r")
                self._end.getparent().addprevious(separator_run)
                separator = element(separator_run, "fldChar", fldCharType="separate")
                self._separator = separator
            run = element(self._end.getparent().getparent(), "r")
            self._end.getparent().addprevious(run)
            Run(run, self._owner._parent).text = text
        self._node.set(_tag("dirty"), "false")
        return self

    def lock(self, locked=True):
        self._check()
        self._node.set(_tag("fldLock"), "true" if locked else "false")
        return self
