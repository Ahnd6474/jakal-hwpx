"""Adapt python-docx objects to the existing Hancom document model.

This module is loaded only for DOCX input, so python-docx remains optional.
"""
from __future__ import annotations

import warnings
import json
import re
from io import BytesIO
from copy import deepcopy
from pathlib import Path
from typing import TYPE_CHECKING, Any

from ..._hancom_models import (
    CharacterStyle,
    AutoNumber,
    Bookmark,
    Equation,
    Field,
    HancomBlock,
    HancomMetadata,
    HancomSection,
    HeaderFooter,
    Hyperlink,
    Note,
    Paragraph,
    ParagraphStyle,
    Picture,
    SectionSettings,
    StyleDefinition,
    Table,
)
from .projection import accept_visible_content as _accept_visible_content
from .projection import paragraph_nodes as _paragraph_nodes

if TYPE_CHECKING:
    from ...hancom_document import HancomDocument


def populate_from_docx(target: HancomDocument, path: Path) -> None:
    from .dependencies import document_factory

    Document = document_factory()

    original = Document(str(path))
    # Work on a projection. Accepting visible revisions/content controls must
    # never change the Word backend or the untouched source package.
    source = Document(BytesIO(path.read_bytes()))
    from docx.table import Table as DocxTable

    reader = _DocxReader(target)
    math_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
    justification = source.settings.element.find("./" + math_ns + "mathPr/" + math_ns + "defJc")
    reader.math_justification = justification.get(math_ns + "val", "centerGroup") if justification is not None else "centerGroup"
    from .features import capture_controls, capture_resources

    reader.source_path = path
    reader.controls = capture_controls(source.element)
    target._word_resources = capture_resources(source)
    target._word_feature_model = True
    reader.source_paths = {projected: raw.getroottree().getelementpath(raw)
                           for projected, raw in zip(source.element.iter(), original.element.iter())}
    # numbering.xml is optional. python-docx's accessor tries to create it
    # through an unimplemented factory when a valid, list-free file omits it.
    reader.numbering_root = next((rel.target_part.element for rel in source.part.rels.values()
                                  if rel.reltype.endswith("/numbering") and not rel.is_external), None)
    for relationship in source.part.rels.values():
        if relationship.reltype.endswith("/theme") and not relationship.is_external:
            from lxml import etree

            reader.theme = etree.fromstring(relationship.target_part.blob)
            break
    props = original.core_properties
    target.metadata = HancomMetadata(
        title=props.title or None,
        language=props.language or None,
        creator=props.author or None,
        subject=props.subject or None,
        description=props.comments or None,
        lastsaveby=props.last_modified_by or None,
        created=props.created.isoformat() if props.created else None,
        modified=props.modified.isoformat() if props.modified else None,
        keyword=props.keywords or None,
    )
    defaults = source.styles.element.xpath("./w:docDefaults/w:rPrDefault/w:rPr")
    if defaults:
        reader.default_run_properties = defaults[0]
    defaults = source.styles.element.xpath("./w:docDefaults/w:pPrDefault/w:pPr")
    if defaults:
        from docx.oxml import OxmlElement
        from docx.text.parfmt import ParagraphFormat

        default_paragraph = OxmlElement("w:p")
        default_paragraph.append(deepcopy(defaults[0]))
        reader.default_paragraph_format = ParagraphFormat(default_paragraph)
    reader.check_unsupported(source.element)
    _accept_visible_content(source.element)
    inherited_headers: dict[tuple[str, str], str] = {}
    for source_section in source.sections:
        section = HancomSection(settings=_section_settings(source_section))
        for kind in ("footnote", "endnote"):
            if source.element.xpath(".//w:" + kind + "Reference"):
                from docx.oxml.ns import qn

                properties = source_section._sectPr.find(qn("w:" + kind + "Pr"))
                if properties is None:
                    properties = source.settings.element.find(qn("w:" + kind + "Pr"))
                default_format = "decimal" if kind == "footnote" else "lowerRoman"
                format_node = properties.find(qn("w:numFmt")) if properties is not None else None
                format_name = format_node.get(qn("w:val"), default_format) if format_node is not None else default_format
                settings = {"autoNumFormat": {"type": _NOTE_FORMATS.get(format_name, "DIGIT"), "supscript": "0"},
                            "noteLine": {"length": "14400", "width": "0.12 mm"},
                            "noteSpacing": {"betweenNotes": "0", "belowLine": "0", "aboveLine": "850"}}
                setattr(section.settings, kind + "_pr", settings)
                reader.note_formats[kind] = _NOTE_FORMATS.get(format_name, "DIGIT")
        target.sections.append(section)
        for paragraph_index, item in enumerate(source_section.iter_inner_content()):
            if isinstance(item, DocxTable):
                block = reader.table(item)
                _mark_source(block, paragraph_index * 100000, 0)
                setattr(block, "_docx_source_path", reader.source_paths.get(item._tbl))
                if reader.controls.get(item._tbl):
                    setattr(block, "_word_extra", {**getattr(block, "_word_extra", {}), "controls": reader.controls[item._tbl]})
                section.blocks.append(block)
            else:
                # A bare section-break carrier is structural, not an extra blank
                # body paragraph. This matches Word's native import into Hancom.
                if item._p.xpath("./w:pPr/w:sectPr") and not item._p.xpath("./*[not(self::w:pPr)]"):
                    continue
                section.blocks.extend(reader.paragraph(item, paragraph_index))
        reader.headers(source_section, section, inherited_headers, source.settings)
    if not target.sections:
        target.sections.append(HancomSection())
    definitions = {style.style_id: style for style in target.style_definitions}
    for style in source.styles:
        if style.style_id in reader._named_styles and hasattr(style, "next_paragraph_style"):
            following = style.next_paragraph_style
            if following is not None:
                definitions[reader._named_styles[style.style_id]].next_style_id = reader._named_styles.get(following.style_id)
    from .preservation import capture_docx_source

    capture_docx_source(target, original, path)


def _units(length: int) -> int:
    # 914400 EMU/inch, 7200 HWP units/inch.
    return round(int(length) / 127)


def _cell_margins(node: Any) -> dict[str, int]:
    """Read direct OOXML cell margins (twips) into HWP's 1/7200-inch units."""
    from docx.oxml.ns import qn

    margins: dict[str, int] = {}
    if node is None:
        return margins
    for side, names in (
        ("left", ("left", "start")), ("right", ("right", "end")),
        ("top", ("top",)), ("bottom", ("bottom",)),
    ):
        for name in names:
            edge = node.find(qn("w:" + name))
            if edge is None or edge.get(qn("w:type"), "dxa") != "dxa":
                continue
            width = edge.get(qn("w:w"))
            if width is not None:
                margins[side] = int(width) * 5
                break
    return margins


def _border_styles(node: Any) -> dict[str, dict[str, str | int]]:
    from docx.oxml.ns import qn

    borders: dict[str, dict[str, str | int]] = {}
    if node is None:
        return borders
    for side, names in (
        ("left", ("left", "start")), ("right", ("right", "end")),
        ("top", ("top",)), ("bottom", ("bottom",)),
    ):
        for name in names:
            edge = node.find(qn("w:" + name))
            if edge is None:
                continue
            borders[side] = {
                "style": edge.get(qn("w:val"), "single"),
                "size": int(edge.get(qn("w:sz"), "4")),
                "color": edge.get(qn("w:color"), "000000"),
            }
            break
    return borders


def _section_settings(section: Any) -> SectionSettings:
    from docx.oxml.ns import qn
    from .features import xml
    settings = SectionSettings(
        page_width=_units(section.page_width) if section.page_width is not None else None,
        page_height=_units(section.page_height) if section.page_height is not None else None,
        landscape="LANDSCAPE" if section.orientation == 1 else "WIDELY",
    )
    for name, attribute in (
        ("left", "left_margin"), ("right", "right_margin"),
        ("top", "top_margin"), ("bottom", "bottom_margin"),
        ("header", "header_distance"), ("footer", "footer_distance"), ("gutter", "gutter"),
    ):
        value = getattr(section, attribute)
        if value is not None:
            settings.margins[name] = _units(value)
    # Word measures both the body and header/footer from the paper edge. HWP's
    # top/bottom margin is the outer margin, with a separate reserved distance
    # from the header/footer to the body. Preserve the total body offset.
    for edge, story, total, distance in (
        ("top", "header", section.top_margin, section.header_distance),
        ("bottom", "footer", section.bottom_margin, section.footer_distance),
    ):
        if total is not None and distance is not None:
            outer = min(_units(total), _units(distance))
            settings.margins[edge] = outer
            settings.margins[story] = _units(total) - outer
    cols = section._sectPr.find(qn("w:cols"))
    if cols is not None:
        explicit = cols.findall(qn("w:col"))
        settings.columns = {"count": len(explicit) or int(cols.get(qn("w:num"), "1")),
                            "gap": int(cols.get(qn("w:space"), "0")) * 5,
                            "equal": cols.get(qn("w:equalWidth"), "1") not in {"0", "false", "off"},
                            "separator": cols.get(qn("w:sep"), "0") in {"1", "true", "on"},
                            "widths": [int(node.get(qn("w:w"), "0")) * 5 for node in explicit],
                            "gaps": [int(node.get(qn("w:space"), cols.get(qn("w:space"), "0"))) * 5 for node in explicit]}
    # Only non-body section properties travel through the reversible carrier.
    setattr(settings, "_word_extra", {name: xml(section._sectPr.find(qn("w:" + name))) for name in ("type", "footnotePr", "endnotePr", "pgNumType", "titlePg")})
    return settings


def _mark_source(block: HancomBlock | HeaderFooter, index: int, order: int) -> None:
    # Use the same grouping metadata as the HWP/HWPX readers.
    setattr(block, "source_paragraph_index", index)
    setattr(block, "source_order", order)


def _style_chain(style: Any):
    seen: set[str] = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        yield style
        style = style.base_style


def _list_number(value: int, kind: str) -> str:
    if kind in {"upperLetter", "lowerLetter"} and value > 0:
        text = ""
        while value:
            value, digit = divmod(value - 1, 26)
            text = chr(ord("A") + digit) + text
        return text.lower() if kind == "lowerLetter" else text
    if kind in {"upperRoman", "lowerRoman"} and 0 < value < 4000:
        text = ""
        for number, letters in ((1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
                                (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")):
            count, value = divmod(value, number)
            text += letters * count
        return text.lower() if kind == "lowerRoman" else text
    return str(value).zfill(2) if kind == "decimalZero" else str(value)


_NOTE_FORMATS = {"decimal": "DIGIT", "lowerRoman": "ROMAN_SMALL", "upperRoman": "ROMAN_CAPITAL",
                 "lowerLetter": "LATIN_SMALL", "upperLetter": "LATIN_CAPITAL"}


class _DocxReader:
    def __init__(self, target: HancomDocument) -> None:
        self.target = target
        self.math_justification = "centerGroup"
        self.default_run_properties: Any = None
        self.default_paragraph_format: Any = None
        self._paragraph_styles: dict[str, str] = {}
        self._character_styles: dict[str, str] = {}
        self._named_styles: dict[str, str] = {}
        self.theme: Any = None
        self.numbering_root: Any = None
        self._list_counters: dict[tuple[str, int], int] = {}
        self._list_definition: dict[str, str] = {}
        self._list_level: Any = None
        self.source_paths: dict[Any, str] = {}
        self.note_formats: dict[str, str] = {}
        self.note_counts: dict[str, int] = {}
        self.pending_breaks: dict[Any, str] = {}
        self.source_path = Path(".")
        self.controls: dict[Any, list[dict[str, Any]]] = {}
        self.table_context: dict[str, Any] = {}

    def warn(self, message: str) -> None:
        if message not in self.target.import_warnings:
            self.target.import_warnings.append(message)
            if not getattr(self.target, "_preserves_word_details", False):
                warnings.warn(message, UserWarning, stacklevel=3)

    def check_unsupported(self, element: Any) -> None:
        for expression, feature in (
            (".//w:altChunk", "embedded document content"),
            (".//w:commentRangeStart | .//w:commentReference", "comments"),
            (".//w:pict[not(.//w:txbxContent)] | .//w:object", "legacy drawings or embedded objects"),
        ):
            if element.xpath(expression):
                self.warn(f"DOCX import does not preserve {feature}.")
        if element.xpath(".//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo"):
            self.warn("Tracked revisions are projected as the accepted visible content; revision history remains in the original DOCX.")

    def paragraph(self, source: Any, index: int) -> list[HancomBlock]:
        from docx.drawing import Drawing
        from docx.text.run import Run as DocxRun
        from docx.text.hyperlink import Hyperlink as DocxHyperlink
        from lxml import etree
        from docx.oxml.ns import qn

        following = source._p.getnext()
        run_children = [child for run in source._p if run.tag == qn("w:r") for child in run if child.tag != qn("w:rPr")]
        if (following is not None and following.tag in {qn("w:p"), qn("w:tbl")}
                and len(run_children) == 1 and run_children[0].tag == qn("w:br")
                and run_children[0].get(qn("w:type")) in {"page", "column"}
                and all(child.tag in {qn("w:pPr"), qn("w:r")} for child in source._p)):
            previous = self.pending_breaks.get(source._p.getparent())
            self.pending_breaks[source._p.getparent()] = run_children[0].get(qn("w:type"))
            if previous:
                carrier = Paragraph("", para_pr_id=self.paragraph_style(source), style_id=self.named_style(source))
                setattr(carrier, "_hwpx_paragraph_attributes", {"pageBreak" if previous == "page" else "columnBreak": "1"})
                _mark_source(carrier, index * 100000, 0)
                setattr(carrier, "_docx_source_path", self.source_paths.get(source._p))
                return [carrier]
            return []
        pending_break = self.pending_breaks.pop(source._p.getparent(), None)

        para_id = self.paragraph_style(source)
        math_nodes = [node for node in source._p if node.tag in {qn("m:oMath"), qn("m:oMathPara")}]
        if math_nodes and all(node.tag in {qn("w:pPr"), qn("w:bookmarkStart"), qn("w:bookmarkEnd"), qn("m:oMath"), qn("m:oMathPara")} or
                              node.tag == qn("w:r") and all(child.tag in {qn("w:rPr"), qn("w:t")} for child in node) and not "".join(node.itertext(qn("w:t"))).strip()
                              for node in source._p):
            justification = math_nodes[0].find("./" + qn("m:oMathParaPr") + "/" + qn("m:jc"))
            value = justification.get(qn("m:val")) if justification is not None else getattr(self, "math_justification", "centerGroup")
            style = deepcopy(next(style for style in self.target.paragraph_styles if style.style_id == para_id))
            style.alignment_horizontal = {"left": "LEFT", "right": "RIGHT", "center": "CENTER", "centerGroup": "CENTER"}.get(value, style.alignment_horizontal)
            para_id = self.register_paragraph_style(style)
        style_id = self.named_style(source)
        blocks: list[HancomBlock] = []
        segment = 0
        segments: dict[int, int] = {}

        def append(block: HancomBlock) -> None:
            segments[id(block)] = segment
            blocks.append(block)

        def add_text(text: str, char_id: str | None) -> None:
            if not text:
                return
            controls = self.controls.get(node, [])
            if (blocks and isinstance(blocks[-1], Paragraph) and blocks[-1].char_pr_id == char_id
                    and getattr(blocks[-1], "_word_extra", {}).get("controls", []) == controls):
                blocks[-1].text += text
            else:
                block = Paragraph(text, style_id=style_id, para_pr_id=para_id, char_pr_id=char_id)
                if controls:
                    setattr(block, "_word_extra", {"controls": controls})
                append(block)

        word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        math_ns = "http://schemas.openxmlformats.org/officeDocument/2006/math"
        for node in _paragraph_nodes(source._p):
            tag = etree.QName(node)
            if tag.namespace == word_ns and tag.localname == "bookmarkStart":
                name = node.get("{" + word_ns + "}name", "")
                from .identity import internal_name
                if internal_name(name):
                    continue
                if name and name != "_GoBack":
                    append(Bookmark(name))
                continue
            if tag.namespace == word_ns and tag.localname == "fldSimple":
                append(self.field(node, source))
                continue
            if tag.namespace == math_ns and tag.localname in {"oMath", "oMathPara"}:
                from .equations import omml_to_hwp

                unsupported = None
                try:
                    script = omml_to_hwp(node)
                except ValueError as error:
                    unsupported = str(error)
                    script = "".join(node.xpath(".//m:t/text()", namespaces={"m": math_ns}))
                    self.warn(unsupported)
                if script or unsupported:
                    equation = Equation(script=script)
                    from .identity import ID
                    from .equations import capture_properties
                    inner = node.find("{" + math_ns + "}oMath") if tag.localname == "oMathPara" else node
                    setattr(equation, "_word_extra", {"identity": inner.get(ID) if inner is not None else None,
                                                     "math": capture_properties(node), "display": tag.localname == "oMathPara"})
                    if unsupported:
                        setattr(equation, "_unsupported_word_math", unsupported)
                    append(equation)
                continue
            if tag.namespace != word_ns or tag.localname not in {"r", "hyperlink"}:
                continue
            item = DocxRun(node, source) if tag.localname == "r" else DocxHyperlink(node, source)
            if isinstance(item, DocxHyperlink):
                target = item.url or (f"#{item.fragment}" if item.fragment else "")
                if target:
                    hyperlink = Hyperlink(target, display_text=item.text)
                    styled = [Paragraph(run.text, char_pr_id=self.character_style(source, run)) for run in item.runs]
                    setattr(hyperlink, "_docx_hyperlink_runs", (item.text, styled))
                    images = []
                    for run in item.runs:
                        for drawing in run._r.xpath("./w:drawing"):
                            picture = self.picture(Drawing(drawing, run))
                            if picture is not None:
                                getattr(picture, "_word_extra")["hyperlink"] = target
                                images.append(picture)
                    append(hyperlink)
                    for picture in images:
                        append(picture)
                else:
                    add_text(item.text, None)
                for run in item.runs:
                    self.character_style(source, run)
                continue
            char_id = self.character_style(source, item)
            for content in item._r:
                name = etree.QName(content).localname
                if name == "t":
                    add_text(content.text or "", char_id)
                elif name == "tab":
                    add_text("\t", char_id)
                elif name in {"br", "cr"}:
                    break_type = content.get("{" + word_ns + "}type", "textWrapping")
                    if break_type in {"page", "column"}:
                        segment += 1
                        carrier = Paragraph("", style_id=style_id, para_pr_id=para_id, char_pr_id=char_id)
                        setattr(carrier, "_hwpx_paragraph_attributes", {"pageBreak" if break_type == "page" else "columnBreak": "1"})
                        append(carrier)
                    else:
                        add_text("\n", char_id)
                elif name in {"drawing", "pict"}:
                    from .textboxes import read_textbox

                    textbox = read_textbox(content, source, self)
                    if textbox is not None:
                        append(textbox)
                        continue
                    if name == "pict":
                        continue
                    picture = self.picture(Drawing(content, item))
                    if picture is not None:
                        append(picture)
                elif name in {"footnoteReference", "endnoteReference"}:
                    note = self.note(content, source)
                    if note is not None:
                        note.reference_run_properties = self.reference_run_properties(source, item)
                        setattr(note, "_docx_char_pr_id", char_id)
                        append(note)
        if not blocks:
            from docx.oxml import OxmlElement

            char_id = self.character_style(source, DocxRun(OxmlElement("w:r"), source))
            append(Paragraph("", style_id=style_id, para_pr_id=para_id, char_pr_id=char_id))
        # A leading empty paragraph carries alignment for picture/link-only paragraphs.
        elif not isinstance(blocks[0], Paragraph):
            blocks.insert(0, Paragraph("", style_id=style_id, para_pr_id=para_id))
        marker = self._numbering_marker(source)
        if marker:
            para_id = self.list_paragraph_style(source, para_id)
            for block in blocks:
                if isinstance(block, Paragraph):
                    block.para_pr_id = para_id
            # Native Word list formats exceed HWP's list vocabulary. A visible
            # label field provides the same marker while carrying just the list
            # definition needed to recreate an active Word list on return.
            label = marker + self._list_definition.get("Suffix", "\t")
            field = Field("JAKAL_LIST_LABEL", label, parameters={**self._list_definition, "RenderedLabel": label})
            carrier = Paragraph("", style_id=style_id, para_pr_id=para_id,
                                char_pr_id=next((b.char_pr_id for b in blocks if isinstance(b, Paragraph)), None))
            setattr(field, "_docx_hyperlink_runs", (label, [Paragraph(label, char_pr_id=carrier.char_pr_id)]))
            blocks[0:0] = [carrier, field]
            from .native_effects import NUMBER_FORMATS

            if self._list_definition.get("Format") not in {"bullet", *NUMBER_FORMATS}:
                self.warn("This Word list number format retains its displayed label; no equivalent native Hancom counter is available.")
        for order, block in enumerate(blocks):
            _mark_source(block, index * 100000 + segments.get(id(block), 0), order)
            setattr(block, "_docx_source_path", self.source_paths.get(source._p))
            if isinstance(block, Paragraph):
                from .features import xml

                from .identity import ID, PARA_ID, NS, private_ids
                setattr(block, "_word_extra", {**getattr(block, "_word_extra", {}), "pPr": xml(source._p.pPr),
                                              "identity": source._p.get(ID), "paragraph_identity": source._p.get(PARA_ID),
                                              "paragraph_anchor": bool(private_ids(source._p)) or source._p.get("{" + NS + "}paragraphAnchor") == "1",
                                              "no_runs": not any(True for _ in source._p.iter(qn("w:r")))})
                paragraph_style = next((style for style in self.target.paragraph_styles if style.style_id == block.para_pr_id), None)
                if paragraph_style is not None:
                    getattr(block, "_word_extra")["layout_switches"] = dict(getattr(paragraph_style, "_word_extra", {}).get("layout_switches", {}))
                character = next((style for style in self.target.character_styles if style.style_id == block.char_pr_id), None)
                if character is not None and getattr(character, "_word_extra", None):
                    getattr(block, "_word_extra")["character"] = deepcopy(getattr(character, "_word_extra"))
                controls = getattr(block, "_word_extra", {}).get("controls", self.controls.get(source._p, []))
                if controls:
                    getattr(block, "_word_extra")["controls"] = controls
        if pending_break:
            setattr(blocks[0], "_hwpx_paragraph_attributes", {"pageBreak" if pending_break == "page" else "columnBreak": "1"})
        return blocks

    def field(self, node: Any, source: Any) -> Field | AutoNumber:
        from docx.oxml.ns import qn
        from docx.text.run import Run

        instruction = node.get(qn("w:instr"), "").strip()
        command = instruction.split(maxsplit=1)[0].upper() if instruction else ""
        if command == "JAKAL_FIELD_BOUNDARY":
            data = instruction.split(maxsplit=1)[1]
            record = json.loads(data)
            if record["side"] == "begin":
                kind = record.get("instruction", "").strip().split(maxsplit=1)
                self.warn("Complex Word field " + (kind[0] if kind else "(empty)")
                          + " retains its visible content and field structure; HWPX does not recalculate this outer field.")
            return Field("JAKAL_WORD_FIELD_BOUNDARY", "", parameters={"Record": data})
        cached = "".join(Run(child, source).text for child in node if child.tag == qn("w:r"))
        if command == "JAKAL_BOOKMARK":
            result = Field("BOOKMARK", cached, name=instruction.split(maxsplit=1)[1], parameters={"Prop": 2})
            styled = [Paragraph(Run(child, source).text, char_pr_id=self.character_style(source, Run(child, source)))
                      for child in node if child.tag == qn("w:r")]
            setattr(result, "_docx_hyperlink_runs", (cached, styled))
            return result
        if command in {"PAGE", "NUMPAGES"}:
            result_number = AutoNumber(kind="autoNum", number=int(cached.strip()) if cached.strip().isdigit() else 1,
                                       number_type="PAGE" if command == "PAGE" else "TOTAL_PAGE")
            setattr(result_number, "_word_extra", {"instruction": instruction, "attributes": {key: value for key, value in node.attrib.items() if key != qn("w:instr")}})
            format_match = re.search(r"\\\*\s+(ROMAN|roman|ALPHABETIC|alphabetic|Arabic)", instruction)
            if format_match:
                setattr(result_number, "_docx_number_format", {"ROMAN": "ROMAN_CAPITAL", "roman": "ROMAN_SMALL",
                    "ALPHABETIC": "LATIN_CAPITAL", "alphabetic": "LATIN_SMALL", "Arabic": "DIGIT"}[format_match[1]])
            first = node.find(qn("w:r"))
            if first is not None:
                setattr(result_number, "_docx_char_pr_id", self.character_style(source, Run(first, source)))
            return result_number
        kind = {"MERGEFIELD": "MAILMERGE", "REF": "CROSSREF", "PAGEREF": "CROSSREF",
                "DOCPROPERTY": "DOCPROPERTY", "DATE": "DATE", "TIME": "DATE", "=": "FORMULA"}.get(command, "JAKAL_WORD_FIELD")
        result = Field(kind, cached, parameters={"WordInstruction": instruction})
        setattr(result, "_word_extra", {"attributes": {key: value for key, value in node.attrib.items() if key != qn("w:instr")}})
        styled = [Paragraph(Run(child, source).text, char_pr_id=self.character_style(source, Run(child, source)))
                  for child in node if child.tag == qn("w:r")]
        setattr(result, "_docx_hyperlink_runs", (cached, styled))
        from .native_effects import native_field

        if native_field(instruction) is None:
            self.warn("Word field " + (command or "(empty)") + " retains its cached display in HWPX and restores its instruction in DOCX; HWPX does not recalculate it.")
        return result

    def note(self, node: Any, source: Any) -> Note | None:
        from docx.oxml import OxmlElement, parse_xml
        from docx.oxml.ns import qn
        from docx.blkcntnr import BlockItemContainer
        from types import SimpleNamespace
        from docx.text.run import Run

        kind = "footnote" if node.tag == qn("w:footnoteReference") else "endnote"
        relationship = next((rel for rel in source.part.rels.values() if rel.reltype.endswith("/" + kind + "s") and not rel.is_external), None)
        if relationship is None:
            self.warn("Missing Word " + kind + " part; reference could not be resolved.")
            return None
        note_id = node.get(qn("w:id"), "0")
        root = parse_xml(relationship.target_part.blob)
        definition = next((child for child in root if child.get(qn("w:id")) == note_id), None)
        if definition is None:
            self.warn("Missing Word " + kind + " definition " + note_id + ".")
            return None
        content = OxmlElement("w:hdr")
        content.extend(deepcopy(child) for child in definition)
        _accept_visible_content(content)
        note_part = relationship.target_part
        # Generic note parts have no style API. Delegate styles to the document,
        # while image relationships must resolve against the actual note part.
        class NotePart:
            related_parts = note_part.related_parts
            def __getattr__(self, name):
                return getattr(source.part, name)
        container = BlockItemContainer(content, SimpleNamespace(part=NotePart()))
        nested = self.container_blocks(container)
        text = "\n".join(p.text for p in container.paragraphs)
        self.note_counts[kind] = self.note_counts.get(kind, 0) + 1
        marker = AutoNumber(kind="autoNum", number=self.note_counts[kind], number_type=kind.upper())
        setattr(marker, "_docx_number_format", self.note_formats.get(kind, "DIGIT"))
        setattr(marker, "_docx_superscript", False)
        marker_properties = None
        for paragraph in container.paragraphs:
            references = paragraph._p.xpath("./w:r[w:" + kind + "Ref]")
            if references:
                run = Run(references[0], paragraph)
                marker_properties = self.reference_run_properties(paragraph, run)
                setattr(marker, "_docx_char_pr_id", self.character_style(paragraph, run))
                break
        first = next((block for block in nested if isinstance(block, Paragraph)), Paragraph(""))
        carrier = deepcopy(first)
        carrier.text = ""
        for block in (carrier, marker):
            setattr(block, "nested_paragraph_index", 0)
        setattr(carrier, "nested_source_order", -2)
        setattr(marker, "nested_source_order", -1)
        nested[0:0] = [carrier, marker]
        result = Note("footNote" if kind == "footnote" else "endNote", text, number=self.note_counts[kind])
        result.marker_run_properties = marker_properties
        result.nested_blocks.extend(nested)
        return result

    def reference_run_properties(self, paragraph: Any, run: Any, *, font_names: dict | None = None) -> str:
        """Flatten run/style/default properties so references survive new styles/themes."""
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from lxml import etree

        properties = [run._r.rPr]
        properties.extend(style.element.rPr for style in _style_chain(run.style))
        properties.extend(style.element.rPr for style in _style_chain(paragraph.style))
        properties.append(self.table_context.get("rPr"))
        properties.append(self.default_run_properties)
        resolved = OxmlElement("w:rPr")
        for source in properties:
            if source is None:
                continue
            for child in source:
                if child.tag in {qn("w:rStyle"), qn("w:rPrChange")}:
                    continue
                existing = resolved.find(child.tag)
                if existing is None:
                    resolved.append(deepcopy(child))
                elif child.tag in {qn("w:rFonts"), qn("w:lang")}:
                    for name, value in child.attrib.items():
                        if child.tag == qn("w:rFonts"):
                            attribute = etree.QName(name).localname
                            partner = attribute.removesuffix("Theme") if attribute.endswith("Theme") else attribute + "Theme"
                            if attribute == "cstheme":
                                partner = "cs"
                            if existing.get(qn("w:" + partner)) is not None:
                                continue
                        if name not in existing.attrib:
                            existing.set(name, value)
        if font_names is None:
            style_id = self.character_style(paragraph, run)
            font_names = next(style.font_names for style in self.target.character_styles if style.style_id == style_id)
        fonts = resolved.find(qn("w:rFonts"))
        if font_names:
            if fonts is None:
                fonts = OxmlElement("w:rFonts")
                resolved.insert(0, fonts)
            for attribute in ("ascii", "hAnsi", "eastAsia", "cs"):
                name = qn("w:" + attribute)
                theme_name = qn("w:" + attribute + "Theme")
                theme = fonts.get(theme_name)
                if fonts.get(name) is None and theme and self.theme is not None:
                    kind = "majorFont" if theme.startswith("major") else "minorFont"
                    script = 'a:font[@script="Hang"]' if attribute == "eastAsia" else "a:cs" if attribute == "cs" else "a:latin"
                    names = self.theme.xpath(f".//a:fontScheme/a:{kind}/{script}/@typeface", namespaces={"a": "http://schemas.openxmlformats.org/drawingml/2006/main"})
                    if names and names[0]:
                        fonts.set(name, names[0])
                if fonts.get(name) is not None:
                    fonts.attrib.pop(theme_name, None)
        return etree.tostring(resolved, encoding="unicode")

    def _numbering_marker(self, source: Any) -> str | None:
        from docx.oxml.ns import qn

        numbering = self.numbering_root
        if numbering is None:
            return None
        properties = [source._p, *(style.element for style in _style_chain(source.style))]
        number_properties = [node for element in properties if (node := element.find(f"./{qn('w:pPr')}/{qn('w:numPr')}")) is not None]
        number_id_node = next((node for prop in number_properties if (node := prop.find(qn("w:numId"))) is not None), None)
        if number_id_node is None:
            return None
        number_id = number_id_node.get(qn("w:val"), "0")
        if number_id == "0":
            return None
        level_node = next((node for prop in number_properties if (node := prop.find(qn("w:ilvl"))) is not None), None)
        level = int(level_node.get(qn("w:val"), "0")) if level_node is not None else 0
        number = next(
            (node for node in numbering if node.tag == qn("w:num") and node.get(qn("w:numId")) == number_id),
            None,
        )
        if number is None:
            return None
        abstract_ref = number.find(qn("w:abstractNumId"))
        if abstract_ref is None:
            return None
        abstract_id = abstract_ref.get(qn("w:val"))
        abstract = next(
            (node for node in numbering if node.tag == qn("w:abstractNum") and node.get(qn("w:abstractNumId")) == abstract_id),
            None,
        )
        if abstract is None:
            return None
        overrides = {int(item.get(qn("w:ilvl"), "0")): item for item in number if item.tag == qn("w:lvlOverride")}
        levels = {int(item.get(qn("w:ilvl"), "0")): item for item in abstract if item.tag == qn("w:lvl")}
        for item_level, override in overrides.items():
            definition = override.find(qn("w:lvl"))
            if definition is not None:
                levels[item_level] = definition
        level_definition = levels.get(level)
        if level_definition is None:
            return None
        self._list_level = level_definition
        format_node = level_definition.find(qn("w:numFmt"))
        text_node = level_definition.find(qn("w:lvlText"))
        if format_node is None or text_node is None:
            return None
        number_format = format_node.get(qn("w:val"), "")
        marker = text_node.get(qn("w:val"), "")
        from lxml import etree
        root = etree.Element(qn("w:numbering"), nsmap={"w": qn("w:p")[1:].split("}")[0]})
        root.extend([deepcopy(abstract), deepcopy(number)])
        suffix = level_definition.find(qn("w:suff"))
        self._list_definition = {"NumberingXML": etree.tostring(root, encoding="unicode"), "Level": str(level),
                                 "Format": number_format,
                                 "Suffix": {"space": " ", "nothing": ""}.get(suffix.get(qn("w:val")) if suffix is not None else "", "\t")}
        if number_format == "bullet":
            return marker.replace("\uf0b7", "•").replace("\uf0a7", "▪") or "•"
        if number_format in {"decimal", "decimalZero", "lowerRoman", "upperRoman", "lowerLetter", "upperLetter"}:
            key = (number_id, level)
            start_node = level_definition.find(qn("w:start"))
            override = overrides.get(level)
            if override is not None and override.find(qn("w:startOverride")) is not None:
                start_node = override.find(qn("w:startOverride"))
            start = int(start_node.get(qn("w:val"), "1")) if start_node is not None else 1
            value = self._list_counters.get(key, start - 1) + 1
            self._list_counters[key] = value
            for counter_key in list(self._list_counters):
                if counter_key[0] == number_id and counter_key[1] > level:
                    definition = levels.get(counter_key[1])
                    restart = definition.find(qn("w:lvlRestart")) if definition is not None else None
                    if restart is not None and int(restart.get(qn("w:val"), "0")) != level + 1:
                        continue
                    del self._list_counters[counter_key]
            for parent_level in range(level + 1):
                parent = levels.get(parent_level)
                parent_format = parent.find(qn("w:numFmt")) if parent is not None else None
                parent_start = parent.find(qn("w:start")) if parent is not None else None
                start_value = int(parent_start.get(qn("w:val"), "1")) if parent_start is not None else 1
                value = self._list_counters.get((number_id, parent_level), start_value)
                format_name = parent_format.get(qn("w:val"), "decimal") if parent_format is not None else "decimal"
                marker = marker.replace(f"%{parent_level + 1}", _list_number(value, format_name))
            return marker
        self.warn(f"DOCX list format {number_format!r} is not yet materialized in HWP/HWPX.")
        return None

    def paragraph_style(self, source: Any, *, include_direct: bool = True) -> str:
        from docx.enum.text import WD_TAB_ALIGNMENT
        from docx.oxml.ns import qn
        formats = [style.paragraph_format for style in _style_chain(source.style)]
        if include_direct:
            formats.insert(0, source.paragraph_format)
        if self.table_context.get("pPr") is not None:
            from docx.oxml import OxmlElement
            from docx.text.parfmt import ParagraphFormat

            table_paragraph = OxmlElement("w:p")
            table_paragraph.append(deepcopy(self.table_context["pPr"]))
            formats.append(ParagraphFormat(table_paragraph))
        if self.default_paragraph_format is not None:
            formats.append(self.default_paragraph_format)

        def effective(name: str):
            return next((getattr(fmt, name) for fmt in formats if getattr(fmt, name) is not None), None)

        alignment = effective("alignment")
        alignment_name = {0: "LEFT", 1: "CENTER", 2: "RIGHT", 3: "JUSTIFY", 4: "DISTRIBUTE",
                          5: "JUSTIFY", 7: "JUSTIFY", 8: "JUSTIFY", 9: "DISTRIBUTE"}.get(alignment, "LEFT")
        if alignment is not None and alignment_name is None:
            self.warn("DOCX import does not preserve this paragraph alignment; using the default.")
        style = ParagraphStyle(alignment_horizontal=alignment_name)
        # These affect mixed Korean/Latin wrapping even when font and margins
        # agree. Resolve inheritance rather than retaining them only as metadata.
        def explicit_switch(name):
            for fmt in formats:
                root = fmt._element
                node = root.find("./" + qn("w:pPr") + "/" + qn("w:" + name))
                if node is None:
                    node = root.find(qn("w:" + name))
                if node is not None:
                    return node.get(qn("w:val"), "1").lower() not in {"0", "false", "off"}
            return None

        for native, key in (("autoSpaceDE", "eAsianEng"), ("autoSpaceDN", "eAsianNum")):
            value = explicit_switch(native)
            if value is not None:
                style.auto_spacing[key] = "1" if value else "0"
        for native, key in (("snapToGrid", "snap_to_grid"), ("suppressLineNumbers", "suppress_line_numbers")):
            value = explicit_switch(native)
            if value is not None:
                setattr(style, key, "1" if value else "0")
        wrapping = explicit_switch("wordWrap")
        if wrapping is not None:
            style.break_setting["breakLatinWord"] = "BREAK_WORD" if wrapping else "KEEP_WORD"
        setattr(style, "_word_extra", {"layout_switches": {
            "autoSpaceDE": style.auto_spacing.get("eAsianEng", "0"), "autoSpaceDN": style.auto_spacing.get("eAsianNum", "0"),
            "wordWrap": "1" if style.break_setting.get("breakLatinWord") == "BREAK_WORD" else "0",
            "snapToGrid": style.snap_to_grid if style.snap_to_grid is not None else "1",
            "suppressLineNumbers": style.suppress_line_numbers if style.suppress_line_numbers is not None else "0",
        }})
        for attribute, key in (
            ("left_indent", "left"), ("right_indent", "right"), ("first_line_indent", "intent"),
            ("space_before", "prev"), ("space_after", "next"),
        ):
            value = effective(attribute)
            if value is not None:
                style.margins[key] = str(_units(value))
        spacing = effective("line_spacing")
        if isinstance(spacing, float):
            style.line_spacing = round(spacing * 100)
            style.line_spacing_type = "PERCENT"
        elif spacing is not None:
            style.line_spacing = _units(spacing)
            style.line_spacing_type = "AT_LEAST" if effective("line_spacing_rule") == 3 else "FIXED"
        for attribute, key in (
            ("keep_with_next", "keepWithNext"), ("keep_together", "keepLines"),
            ("page_break_before", "pageBreakBefore"), ("widow_control", "widowOrphan"),
        ):
            value = effective(attribute)
            if value is None and attribute == "widow_control":
                # Word defaults to widow/orphan protection; HWPX defaults to
                # off. Leaving this unspecified splits paragraphs at a column
                # or page boundary after a seemingly faithful roundtrip.
                value = True
            if value is not None:
                style.break_setting[key] = "1" if value else "0"
        # Tab positions inherit independently. A direct clear at one position
        # must not erase other inherited tabs or become a native left tab.
        inherited_tabs: dict[int, Any] = {}
        for fmt in reversed(formats):
            for tab in fmt.tab_stops:
                if tab.alignment == WD_TAB_ALIGNMENT.CLEAR:
                    inherited_tabs.pop(int(tab.position), None)
                else:
                    inherited_tabs[int(tab.position)] = tab
        tabs = [inherited_tabs[position] for position in sorted(inherited_tabs)]
        tab_stops = [{"pos": _units(tab.position), "type": {0: "LEFT", 1: "CENTER", 2: "RIGHT", 3: "DECIMAL", 4: "BAR"}.get(tab.alignment, "LEFT"),
                      "leader": {0: "NONE", 1: "DOT", 2: "DASH", 3: "SOLID", 4: "DOUBLE_SLIM"}.get(tab.leader, "NONE")} for tab in tabs]
        setattr(style, "_docx_tab_stops", tab_stops)
        if include_direct:
            from docx.oxml.ns import qn
            from docx.text.paragraph import Paragraph as WordParagraph

            contextual = next((node for element in [source._p, *(s.element for s in _style_chain(source.style))]
                               if (node := element.find(f"./{qn('w:pPr')}/{qn('w:contextualSpacing')}")) is not None), None)
            if contextual is not None and contextual.get(qn("w:val"), "1") not in {"0", "false", "off"}:
                for neighbor, edge in ((source._p.getprevious(), "prev"), (source._p.getnext(), "next")):
                    if neighbor is not None and neighbor.tag == qn("w:p") and WordParagraph(neighbor, source._parent).style == source.style:
                        style.margins[edge] = "0"
        return self.register_paragraph_style(style)

    def register_paragraph_style(self, style: ParagraphStyle) -> str:
        style.style_id = None
        tab_stops = getattr(style, "_docx_tab_stops", [])
        key = repr(style) + repr(tab_stops)
        if key not in self._paragraph_styles:
            style.style_id = str(len(self._paragraph_styles) + 1)
            self.target.paragraph_styles.append(style)
            self._paragraph_styles[key] = style.style_id
        return self._paragraph_styles[key]

    def list_paragraph_style(self, source: Any, para_id: str) -> str:
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.text.parfmt import ParagraphFormat

        style = deepcopy(next(style for style in self.target.paragraph_styles if style.style_id == para_id))
        properties = self._list_level.find(qn("w:pPr"))
        if properties is None:
            return para_id
        carrier = OxmlElement("w:p")
        carrier.append(deepcopy(properties))
        fmt = ParagraphFormat(carrier)
        for attribute, key in (("left_indent", "left"), ("right_indent", "right"), ("first_line_indent", "intent")):
            if getattr(source.paragraph_format, attribute) is None and getattr(fmt, attribute) is not None:
                style.margins[key] = str(_units(getattr(fmt, attribute)))
        self._list_definition["WordMargins"] = json.dumps(style.margins)
        style.margins["left"] = str(int(style.margins.get("left", "0")) + min(int(style.margins.get("intent", "0")), 0))
        setattr(style, "_docx_list_layout", True)
        tabs = list(getattr(style, "_docx_tab_stops", []))
        for tab in properties.findall(f"./{qn('w:tabs')}/{qn('w:tab')}"):
            position = int(tab.get(qn("w:pos"), "0")) * 5
            if not any(stop["pos"] == position for stop in tabs):
                tabs.append({"pos": position, "type": "LEFT", "leader": "NONE"})
        if not tabs and self._list_definition.get("Suffix") == "\t":
            tabs.append({"pos": int(style.margins.get("left", "3600")), "type": "LEFT", "leader": "NONE"})
        setattr(style, "_docx_tab_stops", sorted(tabs, key=lambda stop: stop["pos"]))
        return self.register_paragraph_style(style)

    def named_style(self, source: Any) -> str | None:
        style = source.style
        if style is None:
            return None
        if style.style_id not in self._named_styles:
            from docx.oxml import OxmlElement
            from docx.text.run import Run

            style_id = str(len(self._named_styles) + 1)
            self._named_styles[style.style_id] = style_id
            char_id = self.character_style(source, Run(OxmlElement("w:r"), source))
            self.target.style_definitions.append(StyleDefinition(
                style.name, style_id=style_id, para_pr_id=self.paragraph_style(source, include_direct=False),
                char_pr_id=char_id,
            ))
            from .features import character_signature
            character = next((item for item in self.target.character_styles if item.style_id == char_id), None)
            if character is not None:
                self.target._word_resources.setdefault("style_signatures", {})[style.name] = character_signature(character)
        return self._named_styles[style.style_id]

    def character_style(self, paragraph: Any, run: Any) -> str | None:
        from docx.oxml.ns import qn

        properties = [run._r.rPr]
        properties.extend(style.element.rPr for style in _style_chain(run.style))
        properties.extend(style.element.rPr for style in _style_chain(paragraph.style))
        properties.append(self.table_context.get("rPr"))
        properties.append(self.default_run_properties)
        properties = [prop for prop in properties if prop is not None]

        def value(tag: str) -> str | None:
            for prop in properties:
                element = prop.find(qn(f"w:{tag}"))
                if element is not None:
                    return element.get(qn("w:val"))
            return None

        from .formatting import color as resolve_color

        color_node = next((node for prop in properties if (node := prop.find(qn("w:color"))) is not None), None)
        color = resolve_color(color_node, self.theme)
        size = value("sz")
        height = int(size) * 50 if size is not None else None
        def enabled(tag: str) -> bool | None:
            for prop in properties:
                node = prop.find(qn(f"w:{tag}"))
                if node is not None:
                    return node.get(qn("w:val"), "1").lower() not in {"0", "false", "off", "none"}
            return None

        fonts = {}
        for lang, attribute in (("latin", "ascii"), ("hangul", "eastAsia")):
            for prop in properties:
                node = prop.find(qn("w:rFonts"))
                if node is None:
                    continue
                face = node.get(qn("w:" + attribute))
                theme = node.get(qn("w:" + attribute + "Theme"))
                if not face and theme and self.theme is not None:
                    kind = "majorFont" if theme.startswith("major") else "minorFont"
                    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
                    font = "a:latin" if lang == "latin" else 'a:font[@script="Hang"]'
                    names = self.theme.xpath(f".//a:fontScheme/a:{kind}/{font}/@typeface", namespaces=ns)
                    face = names[0] if names else None
                if face:
                    fonts[lang] = face
                    break
        for lang in ("hanja", "japanese"):
            if "hangul" in fonts:
                fonts[lang] = fonts["hangul"]
        for lang in ("other", "symbol", "user"):
            if "latin" in fonts:
                fonts[lang] = fonts["latin"]
        underline = value("u")
        underline_shapes = {"single": "SOLID", "double": "DOUBLE_SLIM", "dotted": "DOT", "dash": "DASH", "wave": "WAVE"}
        if underline not in (None, "none", *underline_shapes):
            self.warn(f"DOCX underline style {underline!r} is approximated in model conversion.")
        highlight = value("highlight")
        shading = next((node for prop in properties if (node := prop.find(qn("w:shd"))) is not None), None)
        text_scale = value("w")
        font_ratios = {}
        if text_scale is not None:
            scale = int(text_scale)
            if 50 <= scale <= 200:
                font_ratios = {lang: scale for lang in ("hangul", "latin", "hanja", "japanese", "other", "symbol", "user")}
            else:
                self.warn("DOCX character width outside the HWP 50–200% range is not preserved.")
        shades = {"yellow": "#FFFF00", "green": "#00FF00", "cyan": "#00FFFF", "magenta": "#FF00FF", "blue": "#0000FF", "red": "#FF0000", "darkBlue": "#000080", "darkCyan": "#008080", "darkGreen": "#008000", "darkMagenta": "#800080", "darkRed": "#800000", "darkYellow": "#808000", "darkGray": "#808080", "lightGray": "#C0C0C0", "black": "#000000", "white": "#FFFFFF", "none": "none"}
        style = CharacterStyle(
            text_color=color, height=height, font_names=fonts, font_ratios=font_ratios,
            bold=enabled("b"), italic=enabled("i"),
            underline={"type": "BOTTOM" if underline not in (None, "none") else "NONE", "shape": underline_shapes.get(underline or "", "SOLID"), "color": color or "#000000"},
            strikeout={"shape": "SOLID" if enabled("strike") or enabled("dstrike") else "NONE", "color": color or "#000000"},
            shade_color=shades.get(highlight) if highlight else resolve_color(shading, self.theme, fill=True),
            script_position={"superscript": "SUPERSCRIPT", "subscript": "SUBSCRIPT"}.get(value("vertAlign") or "", "NORMAL"),
        )
        resolved = self.reference_run_properties(paragraph, run, font_names=fonts)
        from dataclasses import asdict

        setattr(style, "_word_extra", {"rPr": resolved, "mapped": {key: asdict(style)[key] for key in (
            "bold", "italic", "height", "font_names", "font_ratios", "text_color", "shade_color", "underline", "strikeout", "script_position")}})
        key = json.dumps({"color": color, "height": height, "fonts": fonts, "ratios": font_ratios, "bold": style.bold, "italic": style.italic, "underline": style.underline, "strikeout": style.strikeout, "shade": style.shade_color, "position": style.script_position, "rPr": resolved}, sort_keys=True)
        if key not in self._character_styles:
            style_id = str(len(self._character_styles) + 1)
            style.style_id = style_id
            self.target.character_styles.append(style)
            self._character_styles[key] = style_id
        return self._character_styles[key]

    def picture(self, drawing: Any) -> Picture | None:
        from .pictures import read_picture

        result = read_picture(drawing, self.source_path)
        if result is None:
            self.warn("DOCX import does not preserve charts, SmartArt, or other non-picture drawings.")
        return result

    def table(self, source: Any) -> Table:
        from .tables import read_table

        return read_table(self, source)

    def container_blocks(self, container: Any) -> list[HancomBlock]:
        from docx.table import Table as WordTable

        blocks: list[HancomBlock] = []
        for index, item in enumerate(container.iter_inner_content()):
            values = [self.table(item)] if isinstance(item, WordTable) else self.paragraph(item, index)
            for order, block in enumerate(values):
                setattr(block, "nested_paragraph_index", getattr(block, "source_paragraph_index", index * 100000))
                setattr(block, "nested_source_order", order)
            blocks.extend(values)
        return blocks

    def headers(self, source: Any, section: HancomSection, inherited: dict[tuple[str, str], str], settings: Any) -> None:
        if source.different_first_page_header_footer:
            section.settings.visibility.update(hideFirstHeader="1", hideFirstFooter="1")
        for kind in ("header", "footer"):
            for variant, prefix in (("default", ""), ("even", "even_page_"), ("first", "first_page_")):
                story = getattr(source, prefix + kind)
                key = (kind, variant)
                if story.is_linked_to_previous and key not in inherited:
                    continue  # Reading absent stories must not create Word parts.
                if not story.is_linked_to_previous:
                    self.check_unsupported(story._element)
                    from .features import capture_controls

                    self.controls.update(capture_controls(story._element))
                    _accept_visible_content(story._element)
                    inherited[key] = "\n".join(paragraph.text for paragraph in story.paragraphs)
                text = inherited.get(key, "")
                active = variant == "default" or (
                    settings.odd_and_even_pages_header_footer if variant == "even" else source.different_first_page_header_footer
                )
                if not active:
                    continue
                page_type = "FIRST" if variant == "first" else "EVEN" if variant == "even" else ("ODD" if settings.odd_and_even_pages_header_footer else "BOTH")
                header = HeaderFooter(kind, text, apply_page_type=page_type)
                header.nested_blocks = list(self.container_blocks(story))
                _mark_source(header, 0, -10 + len(section.header_footer_blocks))
                section.header_footer_blocks.append(header)
