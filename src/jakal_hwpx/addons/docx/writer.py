"""Portable shared document -> Word export; no desktop application is required."""
from __future__ import annotations

from io import BytesIO
from copy import deepcopy
from typing import Any, TYPE_CHECKING
import warnings

from ..._hancom_models import (
    AutoNumber, Bookmark, CharacterStyle, Equation, Field, HeaderFooter, Hyperlink,
    Memo, Note, Paragraph, ParagraphStyle, Picture, Shape, Table,
)

if TYPE_CHECKING:
    from ...hancom_document import Document


def build_docx(document: Document) -> bytes:
    from .dependencies import document_factory

    WordFile = document_factory()

    word = WordFile()
    writer = _Writer(document, word)
    for name, value in (
        ("title", document.metadata.title), ("subject", document.metadata.subject),
        ("author", document.metadata.creator), ("comments", document.metadata.description),
        ("keywords", document.metadata.keyword), ("language", document.metadata.language),
        ("last_modified_by", document.metadata.lastsaveby),
    ):
        if value is not None:
            if len(value) <= 255:
                setattr(word.core_properties, name, value)
            else:
                # python-docx's convenience setters impose a 255-character
                # limit. Core-property XML can retain longer source metadata
                # (HWP commonly exposes its entire preview text as a title).
                from docx.oxml import OxmlElement
                from docx.oxml.ns import qn

                tag = {"title": "dc:title", "subject": "dc:subject", "author": "dc:creator",
                       "comments": "dc:description", "keywords": "cp:keywords", "language": "dc:language",
                       "last_modified_by": "cp:lastModifiedBy"}[name]
                node = word.core_properties._element.find(qn(tag))
                if node is None:
                    node = OxmlElement(tag)
                    word.core_properties._element.append(node)
                node.text = value
        else:
            setattr(word.core_properties, name, "")
    from datetime import datetime, timezone
    from docx.oxml.ns import qn

    for name, value in (("created", document.metadata.created), ("modified", document.metadata.modified)):
        if value:
            try:
                timestamp = datetime.fromisoformat(value)
                if timestamp.tzinfo is not None:
                    timestamp = timestamp.astimezone(timezone.utc)
                setattr(word.core_properties, name, timestamp)
            except ValueError:
                writer.warn("Invalid " + name + " metadata date could not be represented in Word.")
        else:
            node = word.core_properties._element.find(qn("dcterms:" + name))
            if node is not None:
                word.core_properties._element.remove(node)
    for index, source in enumerate(document.sections):
        section = word.sections[0] if index == 0 else word.add_section()
        writer.section_settings(section, source.settings)
        writer.stories(section, source)
        writer.blocks(word, source.blocks)
        if not source.blocks:
            word.add_paragraph()
    if not document.sections:
        word.add_paragraph()
    stream = BytesIO()
    word.save(stream)
    return stream.getvalue()


def _length(value: Any) -> Any:
    from docx.shared import Emu

    return Emu(round(float(value) * 127))


def _element(parent: Any, tag: str, **attributes: Any) -> Any:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    node = OxmlElement("w:" + tag)
    for key, value in attributes.items():
        node.set(qn("w:" + key), str(value))
    parent.append(node)
    return node


def _color(value: str | None) -> str | None:
    value = (value or "").lstrip("#")
    return value.upper() if len(value) == 6 and all(c in "0123456789abcdefABCDEF" for c in value) else None


class _Writer:
    def __init__(self, source: Document, word: Any) -> None:
        self.source = source
        self.word = word
        self.characters = {style.style_id: _word_character_style(style) for style in source.character_styles}
        self.paragraphs = {style.style_id: _word_paragraph_style(style) for style in source.paragraph_styles}
        self.styles = {style.style_id: style for style in source.style_definitions}
        from docx.enum.style import WD_STYLE_TYPE

        self.word_styles = {}
        for style in source.style_definitions:
            kind = WD_STYLE_TYPE.CHARACTER if style.style_type == "CHAR" else WD_STYLE_TYPE.PARAGRAPH
            name = style.name or "Imported " + str(style.style_id)
            target = word.styles[name] if name in word.styles else word.styles.add_style(name, kind)
            if target.type == kind:
                self.word_styles[style.style_id] = target
        for style in source.style_definitions:
            target = self.word_styles.get(style.style_id)
            following = self.word_styles.get(style.next_style_id)
            if target is not None and following is not None and target.type == following.type == WD_STYLE_TYPE.PARAGRAPH:
                target.next_paragraph_style = following
            if target is not None and style.char_pr_id is not None:
                from docx.oxml import OxmlElement
                from docx.text.run import Run

                run = Run(OxmlElement("w:r"), word)
                self.character_style(run, Paragraph("", char_pr_id=style.char_pr_id))
                if run._r.rPr is not None:
                    previous = target.element.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr")
                    if previous is not None:
                        target.element.remove(previous)
                    target.element.append(deepcopy(run._r.rPr))
        self.warned: set[str] = set()
        self.bookmark_id = 0
        self.note_parts: dict[str, Any] = {}
        self.list_ids: dict[str, int] = {}

    def warn(self, message: str) -> None:
        if message not in self.warned:
            warnings.warn("DOCX export: " + message, UserWarning, stacklevel=3)
            self.warned.add(message)

    def section_settings(self, section: Any, settings: Any) -> None:
        from docx.enum.section import WD_ORIENT

        width, height = settings.page_width, settings.page_height
        if settings.landscape == "NARROWLY":
            width, height = height, width
        if width is not None:
            section.page_width = _length(width)
        if height is not None:
            section.page_height = _length(height)
        section.orientation = WD_ORIENT.LANDSCAPE if section.page_width > section.page_height else WD_ORIENT.PORTRAIT
        margins = settings.margins
        native = getattr(settings, "_docx_native_margins", None)
        if native is not None and native[0] == margins:
            margins = native[1]
        for edge in ("left", "right"):
            if edge in margins:
                setattr(section, edge + "_margin", _length(margins[edge]))
        for edge, story in (("top", "header"), ("bottom", "footer")):
            if edge in margins or story in margins:
                setattr(section, edge + "_margin", _length(margins.get(edge, 0) + margins.get(story, 0)))
                setattr(section, story + "_distance", _length(margins.get(edge, 0)))
        if "gutter" in margins:
            section.gutter = _length(margins["gutter"])
        if settings.page_numbers:
            self.warn("Page-number placement is not mapped; use an external converter for native controls.")
        for kind in ("footnote", "endnote"):
            properties = getattr(settings, kind + "_pr")
            if properties:
                name = {"DIGIT": "decimal", "ROMAN_SMALL": "lowerRoman", "ROMAN_CAPITAL": "upperRoman",
                        "LATIN_SMALL": "lowerLetter", "LATIN_CAPITAL": "upperLetter"}.get(properties.get("autoNumFormat", {}).get("type", "DIGIT"), "decimal")
                target = _element(section._sectPr, kind + "Pr")
                _element(target, "numFmt", val=name)

    def stories(self, section: Any, source: Any) -> None:
        # Every section gets its own stories, including empty stories. Otherwise
        # Word would silently inherit a previous section's header/footer.
        odd_even = any(story.apply_page_type in {"ODD", "EVEN"} for story in source.header_footer_blocks)
        if odd_even:
            self.word.settings.odd_and_even_pages_header_footer = True
        for kind in ("header", "footer"):
            for prefix in ("", "even_page_", "first_page_"):
                story = getattr(section, prefix + kind)
                story.is_linked_to_previous = False
        for block in source.header_footer_blocks:
            prefixes = ["first_page_" if block.apply_page_type == "FIRST" else "even_page_" if block.apply_page_type == "EVEN" else ""]
            if block.apply_page_type == "BOTH":
                prefixes.append("even_page_")
            for prefix in prefixes:
                story = getattr(section, prefix + block.kind)
                if block.nested_blocks:
                    self.blocks(story, block.nested_blocks, first=story.paragraphs[0], nested=True)
                else:
                    story.paragraphs[0].add_run(block.text)
        if (any(block.apply_page_type == "FIRST" for block in source.header_footer_blocks)
                or source.settings.visibility.get("hideFirstHeader") == "1" or source.settings.visibility.get("hideFirstFooter") == "1"):
            section.different_first_page_header_footer = True
            for kind in ("header", "footer"):
                if source.settings.visibility.get("hideFirst" + kind.capitalize()) != "1":
                    ordinary = getattr(section, kind)
                    first = getattr(section, "first_page_" + kind)
                    from copy import deepcopy

                    for child in list(first._element):
                        first._element.remove(child)
                    for child in ordinary._element:
                        first._element.append(deepcopy(child))

    def paragraph_style(self, paragraph: Any, block: Paragraph) -> None:
        from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING

        definition = self.styles.get(block.style_id)
        target_style = self.word_styles.get(block.style_id)
        if target_style is not None and target_style.type == 1:
            paragraph.style = target_style
        style = self.paragraphs.get(block.para_pr_id or (definition.para_pr_id if definition else None))
        attributes = getattr(block, "_hwpx_paragraph_attributes", {})
        if style is None:
            if attributes.get("pageBreak") == "1":
                paragraph.paragraph_format.page_break_before = True
            return
        fmt = paragraph.paragraph_format
        alignment = {
            "LEFT": WD_ALIGN_PARAGRAPH.LEFT, "RIGHT": WD_ALIGN_PARAGRAPH.RIGHT,
            "CENTER": WD_ALIGN_PARAGRAPH.CENTER, "JUSTIFY": WD_ALIGN_PARAGRAPH.JUSTIFY,
            "DISTRIBUTE": WD_ALIGN_PARAGRAPH.DISTRIBUTE, "DISTRIBUTE_SPACE": WD_ALIGN_PARAGRAPH.DISTRIBUTE,
        }.get(style.alignment_horizontal or "")
        if alignment is not None:
            paragraph.alignment = alignment
        margins = style.margins
        native = getattr(style, "_docx_paragraph_margins", None)
        if native is not None and native[0] == margins:
            margins = native[1]
        for key, attribute in (("left", "left_indent"), ("right", "right_indent"), ("intent", "first_line_indent"),
                               ("prev", "space_before"), ("next", "space_after")):
            if key in margins:
                setattr(fmt, attribute, _length(margins[key]))
        if style.line_spacing is not None:
            spacing = float(style.line_spacing)
            kind = style.line_spacing_type or "PERCENT"
            if kind == "PERCENT":
                fmt.line_spacing = spacing / 100
            elif kind in {"FIXED", "AT_LEAST"}:
                fmt.line_spacing = _length(spacing)
                fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY if kind == "FIXED" else WD_LINE_SPACING.AT_LEAST
            else:
                self.warn("Unsupported paragraph line-spacing mode: " + kind)
        for key, attribute in (("keepWithNext", "keep_with_next"), ("keepLines", "keep_together"),
                               ("pageBreakBefore", "page_break_before"), ("widowOrphan", "widow_control")):
            if key in style.break_setting:
                setattr(fmt, attribute, str(style.break_setting[key]).lower() in {"1", "true"})
        if attributes.get("pageBreak") == "1":
            fmt.page_break_before = True
        if attributes.get("columnBreak") == "1":
            from docx.enum.text import WD_BREAK

            paragraph.add_run().add_break(WD_BREAK.COLUMN)
        if hasattr(style, "_docx_tab_stops"):
            from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER

            fmt.tab_stops.clear_all()
            for stop in style._docx_tab_stops:
                alignment = {"LEFT": WD_TAB_ALIGNMENT.LEFT, "CENTER": WD_TAB_ALIGNMENT.CENTER, "RIGHT": WD_TAB_ALIGNMENT.RIGHT,
                             "DECIMAL": WD_TAB_ALIGNMENT.DECIMAL, "BAR": WD_TAB_ALIGNMENT.BAR}.get(stop["type"], WD_TAB_ALIGNMENT.LEFT)
                leader = {"NONE": WD_TAB_LEADER.SPACES, "DOT": WD_TAB_LEADER.DOTS, "DASH": WD_TAB_LEADER.DASHES,
                          "SOLID": WD_TAB_LEADER.LINES, "DOUBLE_SLIM": WD_TAB_LEADER.HEAVY}.get(stop["leader"], WD_TAB_LEADER.SPACES)
                fmt.tab_stops.add_tab_stop(_length(stop["pos"]), alignment, leader)
        if style.heading.get("type", "NONE") != "NONE":
            self.warn("Automatic numbering/bullets are not mapped to Word list definitions.")

    def character_style(self, run: Any, block: Paragraph) -> None:
        from docx.oxml.ns import qn
        from docx.shared import Pt, RGBColor

        definition = self.styles.get(block.style_id)
        style: CharacterStyle | None = self.characters.get(block.char_pr_id or (definition.char_pr_id if definition else None))
        if style is None:
            return
        run.bold, run.italic = style.bold, style.italic
        if style.height is not None:
            run.font.size = Pt(float(style.height) / 100)
        latin = style.font_names.get("latin") or next(iter(style.font_names.values()), None)
        if latin:
            run.font.name = latin
        if style.font_names.get("hangul"):
            run._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), style.font_names["hangul"])
        color = _color(style.text_color)
        if color:
            run.font.color.rgb = RGBColor.from_string(color)
        shade = _color(style.shade_color)
        if shade:
            _element(run._r.get_or_add_rPr(), "shd", val="clear", fill=shade)
        if style.underline:
            shape = style.underline.get("shape", "SOLID")
            value = "none" if style.underline.get("type", "BOTTOM") == "NONE" else {
                "SOLID": "single", "DOUBLE_SLIM": "double", "DASH": "dash", "DOT": "dotted", "WAVE": "wave",
            }.get(shape, "single")
            attrs = {"val": value}
            if _color(style.underline.get("color")):
                attrs["color"] = str(_color(style.underline["color"]))
            _element(run._r.get_or_add_rPr(), "u", **attrs)
        if style.strikeout:
            run.font.strike = style.strikeout.get("shape", "NONE") != "NONE"
        if style.script_position is not None:
            run.font.superscript = style.script_position == "SUPERSCRIPT"
            if style.script_position == "SUBSCRIPT":
                run.font.subscript = True
        if style.font_ratios.get("latin"):
            _element(run._r.get_or_add_rPr(), "w", val=style.font_ratios["latin"])

    def blocks(self, container: Any, blocks: list[Any], *, first: Any = None, nested: bool = False) -> None:
        paragraph = None
        group: Any = object()
        from ...hancom_document import _is_owned_text_shadow_paragraph

        for position, block in enumerate(blocks):
            if isinstance(block, HeaderFooter):
                continue
            if getattr(block, "source_text_owned_by_container", False):
                continue
            if isinstance(block, Table):
                attributes = getattr(block, "_hwpx_paragraph_attributes", {})
                for key, kind in (("pageBreak", "PAGE"), ("columnBreak", "COLUMN")):
                    if attributes.get(key) == "1":
                        from docx.enum.text import WD_BREAK

                        container.add_paragraph().add_run().add_break(getattr(WD_BREAK, kind))
                if first is not None and not first.text:
                    first._p.getparent().remove(first._p)
                self.table(container, block)
                first = container.paragraphs[-1] if hasattr(container, "_tc") and not container.paragraphs[-1].text else None
                paragraph, group = None, object()
                continue
            index = getattr(block, "nested_paragraph_index" if nested else "source_paragraph_index", None)
            if paragraph is None or index is None or index != group:
                paragraph = first if first is not None else container.add_paragraph()
                first = None
                group = index
                if isinstance(block, Paragraph):
                    self.paragraph_style(paragraph, block)
            following = blocks[position + 1] if position + 1 < len(blocks) else None
            if nested or not _is_owned_text_shadow_paragraph(block, following):
                self.inline(paragraph, block)

    def inline(self, paragraph: Any, block: Any) -> None:
        from docx.oxml.ns import qn
        from docx.opc.constants import RELATIONSHIP_TYPE

        if isinstance(block, Paragraph):
            rich = getattr(block, "_docx_runs", None)
            if rich and "".join(text for text, _ in rich) == block.text:
                for text, char_id in rich:
                    self.character_style(paragraph.add_run(text), Paragraph(text, style_id=block.style_id, char_pr_id=char_id))
            else:
                self.character_style(paragraph.add_run(block.text), block)
        elif isinstance(block, Picture):
            try:
                shape = paragraph.add_run().add_picture(BytesIO(block.data), width=_length(block.width), height=_length(block.height))
            except Exception as error:
                raise ValueError(f"Cannot export picture {block.name!r} to Word; use a supported image format or external converter.") from error
            shape._inline.docPr.set("descr", block.shape_comment or block.name)
            crop = block.crop
            cropped = bool(crop and (
                crop.get("left", 0) or crop.get("top", 0) or
                crop.get("right", 0) not in {0, block.original_width or block.width} or
                crop.get("bottom", 0) not in {0, block.original_height or block.height}
            ))
            if cropped or float(block.rotation.get("angle", "0")) or block.layout.get("treatAsChar") == "0":
                self.warn("Picture crop, rotation and floating placement are exported as an inline image.")
        elif isinstance(block, Hyperlink):
            from docx.text.run import Run

            link = _element(paragraph._p, "hyperlink")
            if block.target.startswith("#"):
                link.set(qn("w:anchor"), block.target[1:])
            else:
                link.set(qn("r:id"), paragraph.part.relate_to(block.target, RELATIONSHIP_TYPE.HYPERLINK, is_external=True))
            rich = getattr(block, "_docx_hyperlink_runs", None)
            segments = rich[1] if rich and rich[0] == block.display_text else [Paragraph(block.display_text if block.display_text is not None else block.target)]
            for segment in segments:
                run = Run(_element(link, "r"), paragraph)
                run.text = segment.text
                self.character_style(run, segment)
        elif isinstance(block, Bookmark):
            self.bookmark_id += 1
            _element(paragraph._p, "bookmarkStart", id=self.bookmark_id, name=block.name)
            _element(paragraph._p, "bookmarkEnd", id=self.bookmark_id)
        elif isinstance(block, Field):
            if block.field_type == "JAKAL_LIST_LABEL" and block.parameters.get("RenderedLabel") == block.display_text:
                self.list_label(paragraph, block)
                return
            instruction = block.parameters.get("WordInstruction")
            if not instruction:
                instruction = {
                    "MAILMERGE": "MERGEFIELD " + (block.merge_field_name or ""),
                    "CROSSREF": "REF " + (block.bookmark_name or ""),
                    "DOCPROPERTY": "DOCPROPERTY " + (block.document_property_name or ""),
                    "DATE": "DATE", "FORMULA": "= " + (block.expression or ""),
                }.get(block.semantic_field_type)
            if instruction:
                field = _element(paragraph._p, "fldSimple", instr=instruction, dirty="false")
                from docx.text.run import Run

                rich = getattr(block, "_docx_hyperlink_runs", None)
                segments = rich[1] if rich and rich[0] == block.display_text else [Paragraph(block.display_text or "")]
                for segment in segments:
                    run = Run(_element(field, "r"), paragraph)
                    run.text = segment.text
                    self.character_style(run, segment)
            else:
                paragraph.add_run(block.display_text or "")
                self.warn("Unknown field retains its displayed text; no equivalent Word instruction is available.")
        elif isinstance(block, AutoNumber):
            if block.kind == "autoNum" and block.number_type == "PAGE":
                field = _element(paragraph._p, "fldSimple", instr="PAGE", dirty="false")
                from docx.text.run import Run

                run = Run(_element(field, "r"), paragraph)
                run.text = str(block.number)
                self.character_style(run, Paragraph("", char_pr_id=getattr(block, "_docx_char_pr_id", None)))
            else:
                paragraph.add_run(str(block.number))
                self.warn("This automatic number has no equivalent Word field and retains its displayed number.")
        elif isinstance(block, Equation):
            from .equations import hwp_to_omml

            paragraph._p.append(hwp_to_omml(block.script))
        elif isinstance(block, Note):
            self.note(paragraph, block)
        elif isinstance(block, (Shape, Memo)):
            paragraph.add_run(block.text)
            self.warn(type(block).__name__ + " is exported as body text; its native control/placement is not preserved.")
        else:
            raise ValueError(f"DOCX export does not support {type(block).__name__}; supply an external converter.")

    def note(self, paragraph: Any, block: Note) -> None:
        from docx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
        from docx.opc.packuri import PackURI
        from docx.parts.story import StoryPart
        from docx.oxml import OxmlElement
        from docx.text.paragraph import Paragraph as WordParagraph

        kind = "endnote" if block.kind == "endNote" else "footnote"
        if kind not in self.note_parts:
            root = OxmlElement("w:" + kind + "s")
            # Word's required separator stories occupy IDs -1 and 0.
            for note_id, separator in ((-1, "separator"), (0, "continuationSeparator")):
                item = _element(root, kind, id=note_id, type=separator)
                _element(_element(_element(item, "p"), "r"), separator)
            part = StoryPart(PackURI("/word/" + kind + "s.xml"),
                           CT.WML_ENDNOTES if kind == "endnote" else CT.WML_FOOTNOTES,
                           root, self.word.part.package)
            self.word.part.relate_to(part, RT.ENDNOTES if kind == "endnote" else RT.FOOTNOTES)
            self.note_parts[kind] = part
        part = self.note_parts[kind]
        note_id = len(part.element) - 1
        reference = _element(paragraph._p, "r")
        self.note_run_properties(reference, block.reference_run_properties)
        _element(reference, kind + "Reference", id=note_id)
        item = _element(part.element, kind, id=note_id)
        # Build through a real story container so tables, hyperlinks, math and
        # pictures resolve and serialize against the note's own relationships.
        from docx.blkcntnr import BlockItemContainer
        from types import SimpleNamespace

        root = OxmlElement("w:hdr")
        container = BlockItemContainer(root, SimpleNamespace(part=part))
        first = container.add_paragraph()
        marker = _element(first._p, "r")
        self.note_run_properties(marker, block.marker_run_properties)
        _element(marker, kind + "Ref")
        if block.nested_blocks:
            nested = [item for item in block.nested_blocks if not (isinstance(item, AutoNumber) and item.number_type in {"FOOTNOTE", "ENDNOTE"})]
            self.blocks(container, nested, first=first, nested=True)
        else:
            WordParagraph(first._p, container).add_run(block.text)
        item.extend(list(root))

    @staticmethod
    def note_run_properties(run: Any, properties: str | None) -> None:
        if properties is None:
            # Native HWP/HWPX notes without a Word formatting carrier retain the
            # previous superscript default. An explicit empty rPr means normal.
            _element(_element(run, "rPr"), "vertAlign", val="superscript")
            return
        from lxml import etree
        from docx.oxml.ns import qn

        node = etree.fromstring(properties.encode("utf-8"), etree.XMLParser(resolve_entities=False))
        if node.tag != qn("w:rPr"):
            raise ValueError("A note reference format must contain Word run properties.")
        run.insert(0, node)

    def list_label(self, paragraph: Any, block: Field) -> None:
        from docx.oxml import parse_xml
        from docx.oxml.ns import qn

        raw = str(block.parameters.get("NumberingXML", ""))
        if raw not in self.list_ids:
            root = parse_xml(raw)
            abstract = root.find(qn("w:abstractNum"))
            number = root.find(qn("w:num"))
            if abstract is None or number is None:
                raise ValueError("The preserved Word list definition is incomplete.")
            target = self.word.part.numbering_part.element
            abstract_id = max((int(node.get(qn("w:abstractNumId"), "0")) for node in target if node.tag == qn("w:abstractNum")), default=-1) + 1
            number_id = max((int(node.get(qn("w:numId"), "0")) for node in target if node.tag == qn("w:num")), default=0) + 1
            abstract.set(qn("w:abstractNumId"), str(abstract_id))
            number.set(qn("w:numId"), str(number_id))
            reference = number.find(qn("w:abstractNumId"))
            if reference is None:
                raise ValueError("The preserved Word list has no abstract definition reference.")
            reference.set(qn("w:val"), str(abstract_id))
            # Abstract definitions must precede num instances in OOXML.
            existing_numbers = target.findall(qn("w:num"))
            if existing_numbers:
                existing_numbers[0].addprevious(abstract)
            else:
                target.append(abstract)
            target.append(number)
            self.list_ids[raw] = number_id
        properties = paragraph._p.get_or_add_pPr().get_or_add_numPr()
        properties.get_or_add_ilvl().val = int(block.parameters.get("Level", "0"))
        properties.get_or_add_numId().val = self.list_ids[raw]
        if "WordMargins" in block.parameters:
            import json

            margins = json.loads(str(block.parameters["WordMargins"]))
            for key, attribute in (("left", "left_indent"), ("right", "right_indent"), ("intent", "first_line_indent"),
                                   ("prev", "space_before"), ("next", "space_after")):
                if key in margins:
                    setattr(paragraph.paragraph_format, attribute, _length(margins[key]))

    def table(self, container: Any, block: Table) -> None:
        from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT

        details = getattr(block, "_docx_table_details", {})
        widths = block.col_widths or details.get("col_widths")
        heights = block.row_heights or details.get("row_heights")
        fills = block.cell_fill_colors if block.cell_fill_colors is not None else details.get("cell_fill_colors", {})
        border_styles = block.cell_border_styles if block.cell_border_styles is not None else details.get("cell_border_styles", {})
        if container is self.word or hasattr(container, "_tc"):
            table = container.add_table(rows=block.rows, cols=block.cols)
        else:
            table = container.add_table(rows=block.rows, cols=block.cols, width=_length(sum(widths or [36000])))
        if widths:
            table.autofit = False
            for column, width in zip(table.columns, widths):
                column.width = _length(width)
            for row in table.rows:
                for cell, width in zip(row.cells, widths):
                    cell.width = _length(width)
        for (row, col), (rows, cols) in (block.cell_spans or {}).items():
            table.cell(row, col).merge(table.cell(row + rows - 1, col + cols - 1))
        seen: set[Any] = set()
        for row in range(block.rows):
            if heights and row < len(heights) and heights[row] > 0:
                table.rows[row].height = _length(heights[row])
            if row == 0 and block.repeat_header:
                _element(table.rows[row]._tr.get_or_add_trPr(), "tblHeader")
            for col in range(block.cols):
                cell = table.cell(row, col)
                if cell._tc in seen:
                    continue
                seen.add(cell._tc)
                value = block.cell_texts[row][col] if row < len(block.cell_texts) and col < len(block.cell_texts[row]) else ""
                rich = getattr(block, "_docx_cell_content", {}).get((row, col))
                if rich and rich[0] == value:
                    self.blocks(cell, rich[1], first=cell.paragraphs[0], nested=True)
                else:
                    cell.paragraphs[0].add_run(value)
                pr = cell._tc.get_or_add_tcPr()
                fill = _color(fills.get((row, col)))
                if fill:
                    _element(pr, "shd", val="clear", fill=fill)
                margins = dict(block.table_margins or {})
                margins.update((block.cell_margins or {}).get((row, col), {}))
                if margins:
                    node = _element(pr, "tcMar")
                    for side, margin_value in margins.items():
                        if side in {"left", "right", "top", "bottom"}:
                            _element(node, side, w=round(margin_value / 5), type="dxa")
                alignment = (block.cell_vertical_aligns or {}).get((row, col), "CENTER")
                cell.vertical_alignment = {"TOP": WD_CELL_VERTICAL_ALIGNMENT.TOP, "CENTER": WD_CELL_VERTICAL_ALIGNMENT.CENTER,
                                           "BOTTOM": WD_CELL_VERTICAL_ALIGNMENT.BOTTOM}.get(alignment, WD_CELL_VERTICAL_ALIGNMENT.CENTER)
                borders = border_styles.get((row, col), {})
                if borders:
                    node = _element(pr, "tcBorders")
                    for side, values in borders.items():
                        if side not in {"left", "right", "top", "bottom"}:
                            continue
                        value = str(values.get("style", values.get("type", "SOLID")))
                        shape = {"SOLID": "single", "NONE": "nil", "DASH": "dashed", "DOT": "dotted", "DOUBLE_SLIM": "double",
                                 "single": "single", "none": "nil", "nil": "nil", "dashed": "dashed", "dotted": "dotted", "double": "double"}.get(value, "single")
                        width = float(str(values.get("width", "0.1")).replace("mm", ""))
                        size = int(values["size"]) if "size" in values else max(2, round(width * 72 / 25.4 * 8))
                        _element(node, side, val=shape, sz=size, color=_color(str(values.get("color", "000000"))) or "000000")


def _word_character_style(style: CharacterStyle) -> CharacterStyle:
    """Decode native COLORREFs without changing legacy HWP writer behavior."""
    payload = style.native_hwp_payload
    if payload is None:
        return style
    from ..._hancom_styles import _apply_hwp_character_style_payload

    baseline = CharacterStyle()
    _apply_hwp_character_style_payload(baseline, payload)
    result = deepcopy(style)
    for attribute, offset in (("text_color", 52), ("shade_color", 60)):
        if len(payload) >= offset + 4 and getattr(style, attribute) == getattr(baseline, attribute):
            value = payload[offset:offset + 4]
            setattr(result, attribute, "none" if value == b"\xff\xff\xff\xff" else "#" + value[:3].hex().upper())
    if style.underline == baseline.underline and len(payload) >= 60:
        result.underline["color"] = "#" + payload[56:59].hex().upper()
    return result


def _word_paragraph_style(style: ParagraphStyle) -> ParagraphStyle:
    """Map native HWP bit fields and doubled paragraph distances to Word.

    The legacy common writer uses its own attribute layout. Read the original
    native payload here, preserving any explicitly edited shared properties.
    """
    payload = style.native_hwp_payload
    if payload is None or len(payload) < 28:
        return style
    from ..._hancom_styles import _apply_hwp_paragraph_style_payload

    baseline = ParagraphStyle()
    _apply_hwp_paragraph_style_payload(baseline, payload)
    result = deepcopy(style)
    attributes = int.from_bytes(payload[:4], "little")
    if style.alignment_horizontal == baseline.alignment_horizontal:
        result.alignment_horizontal = {0: "JUSTIFY", 1: "LEFT", 2: "RIGHT", 3: "CENTER", 4: "DISTRIBUTE", 5: "DISTRIBUTE_SPACE"}.get((attributes >> 2) & 7)
    if style.margins == baseline.margins:
        result.margins = {
            key: str(int.from_bytes(payload[offset:offset + 4], "little", signed=True) / 2)
            for key, offset in (("left", 4), ("right", 8), ("intent", 12), ("prev", 16), ("next", 20))
        }
    if style.break_setting == baseline.break_setting:
        result.break_setting = {key: "1" if attributes & (1 << bit) else "0"
                                for key, bit in (("widowOrphan", 16), ("keepWithNext", 17), ("keepLines", 18), ("pageBreakBefore", 19))}
    if style.line_spacing_type is None:
        kind = (int.from_bytes(payload[46:50], "little") & 31) if len(payload) >= 54 else attributes & 3
        result.line_spacing_type = {0: "PERCENT", 1: "FIXED", 2: "BETWEEN", 3: "AT_LEAST"}.get(kind, "PERCENT")
    return result
