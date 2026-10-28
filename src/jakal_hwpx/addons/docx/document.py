"""Word specialization retaining the complete original DOCX package."""
from __future__ import annotations

from contextlib import contextmanager
from copy import deepcopy
from pathlib import Path
from typing import Any
import io

from ...hancom_document import Document, HancomConverter, Paragraph, Table, Picture, Hyperlink
from ...document import HwpxDocument
from ...hwp_document import HwpDocument
from .preservation import DocxSource


class WordDocument(Document):
    """Use shared editing methods while retaining Word-specific XML and parts.

    DOCX output preserves the original package. HWP/HWPX output uses the
    common model unless an external converter is explicitly supplied.
    """

    _preserves_word_details = True

    def __init__(self, **kwargs) -> None:
        super().__init__(**kwargs)
        self._docx_source: DocxSource | None = None

    @property
    def _source(self) -> DocxSource:
        if self._docx_source is None:
            raise ValueError("WordDocument has no DOCX backend. Use open() or blank().")
        return self._docx_source

    @classmethod
    def open(cls, path: str | Path, *, converter: HancomConverter | None = None) -> WordDocument:
        return cls.read_docx(path, converter=converter)

    @classmethod
    def blank(cls, *, converter: HancomConverter | None = None) -> WordDocument:
        from .dependencies import document_factory

        WordFile = document_factory()

        instance = cls(sections=[], source_format="docx", converter=converter)
        source = instance._ensure_temp_dir() / "blank.docx"
        WordFile().save(source)
        from .reader import populate_from_docx

        populate_from_docx(instance, source)
        return instance

    @property
    def word_document(self) -> Any:
        """The python-docx backend. Prefer edit_word() to refresh common views."""
        return self._source.source

    def _reload_word_bytes(self, raw: bytes) -> None:
        from .reader import populate_from_docx

        converter, directory = self._converter, self._ensure_temp_dir()
        path = directory / "word-backend.docx"
        path.write_bytes(raw)
        Document.__init__(self, sections=[], source_format="docx", converter=converter)
        self._temp_dir = directory
        populate_from_docx(self, path)

    def refresh_from_word(self) -> None:
        """Synchronize native backend edits and pending common model text edits."""
        self._reload_word_bytes(self._source.to_bytes(self))

    @contextmanager
    def edit_word(self):
        """Edit any python-docx detail and refresh the shared model on success."""
        self.refresh_from_word()
        original = self._source.raw
        try:
            yield self.word_document
            self.refresh_from_word()
        except Exception:
            self._reload_word_bytes(original)
            raise

    def append_paragraph(self, text: str, *, section_index: int = 0) -> Paragraph:
        with self.edit_word() as word:
            if section_index != len(word.sections) - 1:
                raise ValueError("Word paragraph append currently targets the final section.")
            word.add_paragraph(text)
        return next(block for block in reversed(self.sections[section_index].blocks) if isinstance(block, Paragraph))

    def append_section(self):
        with self.edit_word() as word:
            word.add_section()
        return self.sections[-1]

    def append_table(self, rows: int, cols: int, cell_texts=None, *, section_index: int = 0, **kwargs) -> Table:
        with self.edit_word() as word:
            if section_index != len(word.sections) - 1:
                raise ValueError("Word table append currently targets the final section.")
            if any(value is not None for value in kwargs.values()):
                raise ValueError("Use edit_word() for Word table geometry and formatting.")
            table = word.add_table(rows=rows, cols=cols)
            for row, values in enumerate(cell_texts or []):
                for col, value in enumerate(values):
                    table.cell(row, col).text = value
        return next(block for block in reversed(self.sections[section_index].blocks) if isinstance(block, Table))

    def append_picture(self, name: str, data: bytes, *, extension=None, width=7200, height=7200, shape_comment=None, section_index=0, **kwargs) -> Picture:
        from docx.shared import Emu

        with self.edit_word() as word:
            if section_index != len(word.sections) - 1:
                raise ValueError("Word picture append currently targets the final section.")
            if any(value is not None for value in kwargs.values()):
                raise ValueError("Use edit_word() for advanced Word picture placement.")
            picture = word.add_paragraph().add_run().add_picture(io.BytesIO(data), width=Emu(width * 127), height=Emu(height * 127))
            if shape_comment:
                picture._inline.docPr.set("descr", shape_comment)
        return next(block for block in reversed(self.sections[section_index].blocks) if isinstance(block, Picture))

    def append_hyperlink(self, target: str, *, display_text=None, section_index=0, **kwargs) -> Hyperlink:
        from docx.opc.constants import RELATIONSHIP_TYPE
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        if not target or target == "#":
            raise ValueError("Word hyperlinks require a non-empty target or bookmark name.")
        with self.edit_word() as word:
            if section_index != len(word.sections) - 1 or kwargs:
                raise ValueError("Use edit_word() for advanced Word hyperlinks.")
            paragraph = word.add_paragraph()
            node = OxmlElement("w:hyperlink")
            if target.startswith("#"):
                node.set(qn("w:anchor"), target[1:])
            else:
                rid = paragraph.part.relate_to(target, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
                node.set(qn("r:id"), rid)
            run = OxmlElement("w:r")
            text = OxmlElement("w:t")
            text.text = display_text if display_text is not None else target
            run.append(text)
            node.append(run)
            paragraph._p.append(node)
        return next(block for block in reversed(self.sections[section_index].blocks) if isinstance(block, Hyperlink))

    def _append_word_story(self, kind: str, text: str, section_index: int, apply_page_type: str):
        if apply_page_type not in {"BOTH", "ODD", "EVEN", "FIRST"}:
            raise ValueError("apply_page_type must be BOTH, ODD, EVEN, or FIRST.")
        with self.edit_word() as word:
            section = word.sections[section_index]
            if apply_page_type in {"ODD", "EVEN"}:
                word.settings.odd_and_even_pages_header_footer = True
            if apply_page_type == "FIRST":
                section.different_first_page_header_footer = True
            prefixes = ["first_page_" if apply_page_type == "FIRST" else "even_page_" if apply_page_type == "EVEN" else ""]
            if apply_page_type == "BOTH" and word.settings.odd_and_even_pages_header_footer:
                prefixes.append("even_page_")
            for prefix in prefixes:
                story = getattr(section, prefix + kind)
                if section_index and story.is_linked_to_previous:
                    raise ValueError("Unlink the inherited Word header/footer in edit_word() before appending to a later section.")
                if len(story.paragraphs) == 1 and not story.paragraphs[0].text:
                    story.paragraphs[0].add_run(text)
                else:
                    story.add_paragraph(text)
        return next(block for block in self.sections[section_index].header_footer_blocks if block.kind == kind and (block.apply_page_type == apply_page_type or apply_page_type == "BOTH"))

    def append_header(self, text: str, *, section_index: int = 0, apply_page_type: str = "BOTH"):
        return self._append_word_story("header", text, section_index, apply_page_type)

    def append_footer(self, text: str, *, section_index: int = 0, apply_page_type: str = "BOTH"):
        return self._append_word_story("footer", text, section_index, apply_page_type)

    def replace_text(self, old: str, new: str, *, count: int = -1) -> int:
        from docx.text.paragraph import Paragraph as WordParagraph
        from .preservation import replace_word_match, _visible_slots

        if not old:
            raise ValueError("old must be non-empty.")
        replacements = 0
        remaining = count
        with self.edit_word() as word:
            root = word.part.element
            for node in list(root.xpath(".//w:p")):
                if remaining == 0:
                    break
                paragraph = WordParagraph(node, word)
                current = "".join(slot[3] for slot in _visible_slots(paragraph))
                changed = current.count(old) if remaining < 0 else min(current.count(old), remaining)
                if changed:
                    positions = []
                    offset = 0
                    for _ in range(changed):
                        offset = current.index(old, offset)
                        positions.append(offset)
                        offset += len(old)
                    for offset in reversed(positions):
                        replace_word_match(root, paragraph, offset, offset + len(old), new)
                    replacements += changed
                    if remaining >= 0:
                        remaining -= changed
        return replacements

    def to_docx_document(self, *, converter: HancomConverter | None = None, preserve_formatting: bool | None = None) -> WordDocument:
        return WordDocument.open(self.write_to_docx(self._ensure_temp_dir() / "exported.docx"))

    def write_to_docx(self, path: str | Path, *, converter: HancomConverter | None = None, preserve_formatting: bool | None = None) -> Path:
        raw = self._source.to_bytes(self)
        target = Path(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(raw)
        return target

    def _convert_word(self, format_name: str, converter: HancomConverter | None = None) -> Path:
        selected_converter = converter or self._converter
        if selected_converter is None:
            raise ValueError("Formatting-preserving DOCX conversion requires an explicit converter.")
        directory = self._ensure_temp_dir()
        source = self.write_to_docx(directory / "word-source.docx")
        output = directory / ("word-output." + format_name.lower())
        if output.exists():
            output.unlink()
        converted = Path(selected_converter(source, output, format_name))
        if not converted.is_file():
            raise ValueError("The DOCX converter did not produce an output file.")
        return converted

    def to_hwpx_document(self, *, preserve_formatting: bool | None = None, converter: HancomConverter | None = None) -> HwpxDocument:
        if preserve_formatting is False or (preserve_formatting is None and converter is None and self._converter is None):
            return self.to_common_document().to_hwpx_document(converter=converter)
        return HwpxDocument.open(self._convert_word("HWPX", converter))

    def to_hwp_document(self, *, preserve_formatting: bool | None = None, converter: HancomConverter | None = None) -> HwpDocument:
        if preserve_formatting is False or (preserve_formatting is None and converter is None and self._converter is None):
            return self.to_common_document().to_hwp_document(converter=converter)
        return HwpDocument.open(self._convert_word("HWP", converter), converter=converter or self._converter)

    def to_common_document(self) -> Document:
        """Explicitly copy the projected model for portable semantic conversion."""
        import warnings

        for message in self.import_warnings:
            warnings.warn(message, UserWarning, stacklevel=2)
        result = Document(metadata=deepcopy(self.metadata), sections=deepcopy(self.sections), source_format="docx", converter=self._converter)
        for section in result.sections:
            for block in section.blocks:
                if isinstance(block, Paragraph) and block.list_marker:
                    block.text = f"{block.list_marker}\t{block.text}"
        for name in ("style_definitions", "paragraph_styles", "character_styles", "numbering_definitions", "bullet_definitions", "memo_shape_definitions", "import_warnings"):
            setattr(result, name, deepcopy(getattr(self, name)))
        return result
