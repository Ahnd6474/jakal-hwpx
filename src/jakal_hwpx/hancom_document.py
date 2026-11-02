from __future__ import annotations

import tempfile
import zipfile
from pathlib import Path
from typing import Callable

from lxml import etree

from ._hwpx_element_helpers import _set_text_node_content, _text_node_plain_text  # noqa: F401 - compatibility re-exports

from ._hancom_models import (  # noqa: F401 - compatibility re-exports
    _HWP_NATIVE_FIELD_TYPE_BY_SEMANTIC_TYPE,
    _HWPX_CHART_CARRIER_PREFIX,
    _HWPX_FORM_FIELD_TYPE,
    _HWPX_MEMO_FIELD_TYPE,
    AutoNumber,
    AutoNumberNode,
    Bookmark,
    BookmarkNode,
    BulletDefinition,
    BulletDefinitionNode,
    CharacterStyle,
    CharacterStyleNode,
    Chart,
    Equation,
    EquationNode,
    Field,
    FieldNode,
    Form,
    HancomBlock,
    HancomMetadata,
    HancomSection,
    HeaderFooter,
    HeaderFooterNode,
    Hyperlink,
    HyperlinkNode,
    Memo,
    MemoShapeDefinition,
    MemoShapeDefinitionNode,
    Note,
    NoteNode,
    NumberingDefinition,
    NumberingDefinitionNode,
    Ole,
    OleNode,
    Paragraph,
    ParagraphNode,
    ParagraphStyle,
    ParagraphStyleNode,
    Picture,
    PictureNode,
    SectionSettings,
    SectionSettingsNode,
    Shape,
    ShapeNode,
    StyleDefinition,
    StyleDefinitionNode,
    Table,
    TableNode,
    _semantic_field_type_from_native,
)
from ._hancom_styles import (  # noqa: F401 - compatibility re-exports
    _DEFAULT_HWP_BULLET_PAYLOAD,
    _DEFAULT_HWP_CHAR_SHAPE_PAYLOAD,
    _DEFAULT_HWP_NUMBERING_PAYLOAD,
    _DEFAULT_HWP_PARA_SHAPE_PAYLOAD,
    _HWP_CHAR_FONT_ORDER,
    _HWP_DOCINFO_COUNT_INDEX_BY_TAG,
    _HWP_NUMBERING_LEVEL_COUNT,
    _HWP_PARA_ALIGN_CODES,
    _HWP_PARA_ALIGN_NAMES,
    _HWP_STYLE_TYPE_CODES,
    _HWP_STYLE_TYPE_NAMES,
    _apply_hwp_bullet_payload,
    _apply_hwp_character_style_payload,
    _apply_hwp_numbering_payload,
    _apply_hwp_paragraph_style_payload,
    _assign_missing_hwp_style_ids,
    _bit_dict_to_int,
    _build_hwp_bullet_payload,
    _build_hwp_character_style_payload,
    _build_hwp_colorref,
    _build_hwp_memo_shape_payload,
    _build_hwp_numbering_payload,
    _build_hwp_paragraph_style_payload,
    _build_hwp_placeholder_character_styles,
    _build_hwp_placeholder_paragraph_styles,
    _build_hwp_style_definition_payload,
    _coerce_bool,
    _coerce_hwp_docinfo_word,
    _coerce_int,
    _ensure_hwp_docinfo_record_capacity,
    _extract_character_style,
    _extract_hwp_bullet_definitions,
    _extract_hwp_memo_shape_definitions,
    _extract_hwp_numbering_definitions,
    _extract_hwp_style_definition_trailing_payload,
    _extract_hwp_style_definitions,
    _extract_hwpx_attribute_map,
    _extract_hwpx_margin_map,
    _extract_paragraph_style,
    _extract_style_definition,
    _merge_hwp_style_defaults_into_paragraphs,
    _parse_hwp_colorref,
    _parse_hwp_style_definition_payload,
    _parse_optional_int,
    _seed_hwp_payload,
    _sync_hwp_docinfo_slot_payloads,
    _sync_hwp_docinfo_styles,
    _sync_hwp_memo_shape_payloads,
)
from .document import HwpxDocument
from .hwp_binary import (  # noqa: F401 - compatibility re-exports
    TAG_CTRL_DATA,
    TAG_CTRL_HEADER,
    TAG_PARA_CHAR_SHAPE,
    TAG_LIST_HEADER,
    TAG_PARA_HEADER,
    TAG_PAGE_DEF,
    ParagraphHeaderRecord,
    SectionParagraphModel,
)
from .hwp_document import (  # noqa: F401 - compatibility re-exports
    HwpArcShapeObject,
    HwpConnectLineShapeObject,
    HwpContainerShapeObject,
    HwpCurveShapeObject,
    HwpDocument,
    HwpEllipseShapeObject,
    HwpLineShapeObject,
    HwpPolygonShapeObject,
    HwpRectangleShapeObject,
    HwpTextArtShapeObject,
)
from .namespaces import NS, qname  # noqa: F401 - compatibility re-exports

HancomConverter = Callable[[str | Path, str | Path, str], Path]

class Document:
    def __init__(
        self,
        *,
        metadata: HancomMetadata | None = None,
        sections: list[HancomSection] | None = None,
        source_format: str | None = None,
        converter: HancomConverter | None = None,
    ) -> None:
        self.metadata = metadata or HancomMetadata()
        self.sections = sections if sections is not None else [HancomSection()]
        self.source_format = source_format
        self._converter = converter
        self._temp_dir: Path | None = None
        self.style_definitions: list[StyleDefinition] = []
        self.paragraph_styles: list[ParagraphStyle] = []
        self.character_styles: list[CharacterStyle] = []
        self.numbering_definitions: list[NumberingDefinition] = []
        self.bullet_definitions: list[BulletDefinition] = []
        self.memo_shape_definitions: list[MemoShapeDefinition] = []
        self.import_warnings: list[str] = []
        self._word_feature_model = False
        self._word_resources: dict = {}

    @classmethod
    def blank(cls, *, converter: HancomConverter | None = None) -> "HancomDocument":
        return cls(converter=converter)

    @classmethod
    def open(cls, path: str | Path, *, converter: HancomConverter | None = None) -> "HancomDocument":
        candidate = Path(path)
        suffix = candidate.suffix.lower()
        if suffix == ".docx":
            return cls.read_docx(candidate, converter=converter)
        if zipfile.is_zipfile(candidate):
            with zipfile.ZipFile(candidate) as package:
                if "word/document.xml" in package.namelist():
                    return cls.read_docx(candidate, converter=converter)
            return cls.read_hwpx(candidate, converter=converter)
        if suffix == ".hwp":
            return cls.read_hwp(candidate, converter=converter)
        raise ValueError(f"Unsupported Hancom document path: {candidate}. Expected .hwp, .hwpx, or .docx.")

    def bridge(self, *, converter: HancomConverter | None = None):
        from .bridge import HwpHwpxBridge

        return HwpHwpxBridge.from_hancom(self, converter=converter or self._converter)

    @classmethod
    def read_hwpx(cls, path: str | Path, *, converter: HancomConverter | None = None) -> "HancomDocument":
        instance = cls.from_hwpx_document(HwpxDocument.open(path), converter=converter)
        instance._capture_docx_conversion_source(Path(path))
        return instance

    @classmethod
    def read_hwp(cls, path: str | Path, *, converter: HancomConverter | None = None) -> "HancomDocument":
        instance = cls.from_hwp_document(HwpDocument.open(path, converter=converter), converter=converter)
        instance._capture_docx_conversion_source(Path(path))
        return instance

    def _docx_export_state(self):
        return tuple(getattr(self, key) for key in (
            "metadata", "sections", "style_definitions", "paragraph_styles", "character_styles",
            "numbering_definitions", "bullet_definitions", "memo_shape_definitions",
        ))

    def _capture_docx_conversion_source(self, path: Path) -> None:
        from copy import deepcopy

        self._docx_conversion_source = (path.read_bytes(), self.source_format, deepcopy(self._docx_export_state()))

    @classmethod
    def read_docx(cls, path: str | Path, *, converter: HancomConverter | None = None) -> "HancomDocument":
        """Read a Word document into the shared model (requires the ``docx`` extra).

        Unsupported features are reported through UserWarning and import_warnings.
        The original DOCX package is retained for format-preserving DOCX output.
        """
        from .addons.docx.reader import populate_from_docx
        from .addons.docx import WordDocument

        factory = WordDocument if cls is Document else cls
        instance = factory(sections=[], source_format="docx", converter=converter)
        populate_from_docx(instance, Path(path))
        return instance

    @classmethod
    def from_hwpx_document(
        cls,
        document: HwpxDocument,
        *,
        converter: HancomConverter | None = None,
    ) -> "HancomDocument":
        instance = cls(
            metadata=HancomMetadata.from_hwpx_metadata(document.metadata()),
            sections=[],
            source_format="hwpx",
            converter=converter,
        )
        instance.style_definitions = [_extract_style_definition(style) for style in document.styles()]
        instance.paragraph_styles = [_extract_paragraph_style(style) for style in document.paragraph_styles()]
        font_lookup = {
            (family.get("lang", "").lower(), font.get("id", "")): font.get("face", "")
            for family in document.header.readonly_root_element.xpath(".//hh:fontface", namespaces=NS)
            for font in family
        }
        instance.character_styles = [_extract_character_style(style, font_lookup=font_lookup) for style in document.character_styles()]
        instance._hwpx_ref_resources = [etree.tostring(node) for node in document.header.readonly_root_element.xpath("./hh:refList/*", namespaces=NS)
                                        if etree.QName(node).localname not in {"fontfaces", "charProperties", "paraProperties", "styles"}]
        instance.memo_shape_definitions = [_extract_hwpx_memo_shape(memo_shape) for memo_shape in document.memo_shapes()]
        for section in document.sections:
            instance.sections.append(_extract_hwpx_section(section))
        if not instance.sections:
            instance.sections.append(HancomSection())
        if _has_word_feature_carriers(instance.sections):
            from .addons.docx.features import restore_carriers

            restore_carriers(instance)
        return instance

    @classmethod
    def from_hwp_document(
        cls,
        document: HwpDocument,
        *,
        converter: HancomConverter | None = None,
    ) -> "HancomDocument":
        effective_converter = converter or getattr(document, "_converter", None)
        docinfo_model = document.docinfo_model()
        style_definitions = _extract_hwp_style_definitions(docinfo_model)
        sections = [_extract_hwp_section(section) for section in document.sections()]
        if not sections:
            sections = [HancomSection()]
        preview_text = document.preview_text().strip()
        title = preview_text or (document.source_path.stem if document.source_path else None)
        metadata = HancomMetadata(title=title)
        instance = cls(metadata=metadata, sections=sections, source_format="hwp", converter=effective_converter)
        instance.style_definitions = style_definitions
        instance.paragraph_styles = _build_hwp_placeholder_paragraph_styles(docinfo_model, sections, style_definitions)
        instance.character_styles = _build_hwp_placeholder_character_styles(docinfo_model, style_definitions)
        instance.numbering_definitions = _extract_hwp_numbering_definitions(docinfo_model)
        instance.bullet_definitions = _extract_hwp_bullet_definitions(docinfo_model)
        instance.memo_shape_definitions = _extract_hwp_memo_shape_definitions(docinfo_model)
        _merge_hwp_style_defaults_into_paragraphs(instance)
        return instance

    def append_section(self) -> HancomSection:
        section = HancomSection()
        self.sections.append(section)
        return section

    def append_paragraph(self, text: str, *, section_index: int = 0) -> Paragraph:
        section = self._ensure_section(section_index)
        block = Paragraph(text)
        section.blocks.append(block)
        return block

    def replace_text(self, old: str, new: str, *, count: int = -1) -> int:
        """Replace projected text; format specializations override this operation."""
        if not old:
            raise ValueError("old must be non-empty.")
        changed = 0
        remaining = count
        for section in self.sections:
            for block in section.blocks:
                targets = []
                if isinstance(block, Paragraph):
                    targets.append((block, "text"))
                elif isinstance(block, Hyperlink):
                    targets.append((block, "display_text"))
                elif isinstance(block, Table):
                    for row in block.cell_texts:
                        for index, text in enumerate(row):
                            amount = text.count(old) if remaining < 0 else min(text.count(old), remaining)
                            row[index] = text.replace(old, new, remaining)
                            changed += amount
                            if remaining >= 0:
                                remaining -= amount
                for target, attribute in targets:
                    text = getattr(target, attribute) or ""
                    amount = text.count(old) if remaining < 0 else min(text.count(old), remaining)
                    setattr(target, attribute, text.replace(old, new, remaining))
                    changed += amount
                    if remaining >= 0:
                        remaining -= amount
        return changed

    def append_table(
        self,
        *,
        rows: int,
        cols: int,
        cell_texts: list[list[str]] | None = None,
        row_heights: list[int] | None = None,
        col_widths: list[int] | None = None,
        cell_spans: dict[tuple[int, int], tuple[int, int]] | None = None,
        cell_border_fill_ids: dict[tuple[int, int], int] | None = None,
        table_border_fill_id: int = 1,
        section_index: int = 0,
    ) -> Table:
        section = self._ensure_section(section_index)
        block = Table(
            rows=rows,
            cols=cols,
            cell_texts=cell_texts or _blank_cell_texts(rows, cols),
            row_heights=row_heights,
            col_widths=col_widths,
            cell_spans=cell_spans,
            cell_border_fill_ids=cell_border_fill_ids,
            table_border_fill_id=table_border_fill_id,
        )
        section.blocks.append(block)
        return block

    def append_picture(
        self,
        name: str,
        data: bytes,
        *,
        extension: str | None = None,
        width: int = 7200,
        height: int = 7200,
        original_width: int | None = None,
        original_height: int | None = None,
        current_width: int | None = None,
        current_height: int | None = None,
        section_index: int = 0,
    ) -> Picture:
        section = self._ensure_section(section_index)
        block = Picture(
            name=name,
            data=data,
            extension=extension,
            width=width,
            height=height,
            original_width=original_width,
            original_height=original_height,
            current_width=current_width,
            current_height=current_height,
        )
        section.blocks.append(block)
        return block

    def append_hyperlink(
        self,
        target: str,
        *,
        display_text: str | None = None,
        metadata_fields: list[str | int] | None = None,
        section_index: int = 0,
    ) -> Hyperlink:
        section = self._ensure_section(section_index)
        block = Hyperlink(target=target, display_text=display_text, metadata_fields=list(metadata_fields or []))
        section.blocks.append(block)
        return block

    def append_bookmark(self, name: str, *, section_index: int = 0) -> Bookmark:
        section = self._ensure_section(section_index)
        block = Bookmark(name=name)
        section.blocks.append(block)
        return block

    def append_field(
        self,
        *,
        field_type: str,
        display_text: str | None = None,
        name: str | None = None,
        parameters: dict[str, str | int] | None = None,
        editable: bool = False,
        dirty: bool = False,
        section_index: int = 0,
    ) -> Field:
        section = self._ensure_section(section_index)
        normalized_field_type = field_type.strip()
        native_field_type = normalized_field_type if normalized_field_type.startswith("%") else None
        block = Field(
            field_type=_semantic_field_type_from_native(normalized_field_type, parameters)
            if native_field_type is not None
            else normalized_field_type,
            display_text=display_text,
            name=name,
            parameters=dict(parameters or {}),
            editable=editable,
            dirty=dirty,
            native_field_type=native_field_type,
        )
        section.blocks.append(block)
        return block

    def append_mail_merge_field(
        self,
        field_name: str,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        section_index: int = 0,
    ) -> Field:
        block = self.append_field(
            field_type="MAILMERGE",
            display_text=display_text,
            name=field_name,
            parameters={"FieldName": field_name, "MergeField": field_name},
            editable=editable,
            dirty=dirty,
            section_index=section_index,
        )
        block.configure_mail_merge(field_name, display_text=display_text)
        return block

    def append_calculation_field(
        self,
        expression: str,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        section_index: int = 0,
    ) -> Field:
        block = self.append_field(
            field_type="FORMULA",
            display_text=display_text,
            parameters={"Expression": expression, "Command": expression},
            editable=editable,
            dirty=dirty,
            section_index=section_index,
        )
        block.configure_calculation(expression, display_text=display_text)
        return block

    def append_cross_reference(
        self,
        bookmark_name: str,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        section_index: int = 0,
    ) -> Field:
        block = self.append_field(
            field_type="CROSSREF",
            display_text=display_text,
            name=bookmark_name,
            parameters={"BookmarkName": bookmark_name, "Path": bookmark_name},
            editable=editable,
            dirty=dirty,
            section_index=section_index,
        )
        block.configure_cross_reference(bookmark_name, display_text=display_text)
        return block

    def append_doc_property_field(
        self,
        property_name: str,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        section_index: int = 0,
    ) -> Field:
        block = self.append_field(
            field_type="DOCPROPERTY",
            display_text=display_text,
            name=property_name,
            parameters={"FieldName": property_name},
            editable=editable,
            dirty=dirty,
            section_index=section_index,
        )
        block.configure_doc_property(property_name, display_text=display_text)
        return block

    def append_date_field(
        self,
        *,
        display_text: str | None = None,
        editable: bool = False,
        dirty: bool = False,
        section_index: int = 0,
    ) -> Field:
        block = self.append_field(
            field_type="DATE",
            display_text=display_text,
            editable=editable,
            dirty=dirty,
            section_index=section_index,
        )
        block.configure_date(display_text=display_text)
        return block

    def append_auto_number(
        self,
        *,
        number: int | str = 1,
        number_type: str = "PAGE",
        kind: str = "newNum",
        section_index: int = 0,
    ) -> AutoNumber:
        section = self._ensure_section(section_index)
        block = AutoNumber(kind=kind, number=number, number_type=number_type)
        section.blocks.append(block)
        return block

    def append_note(
        self,
        text: str,
        *,
        kind: str = "footNote",
        number: int | None = None,
        section_index: int = 0,
        reference_run_properties: str | None = None,
        marker_run_properties: str | None = None,
    ) -> Note:
        section = self._ensure_section(section_index)
        block = Note(kind=kind, text=text, number=number,
                     reference_run_properties=reference_run_properties, marker_run_properties=marker_run_properties)
        section.blocks.append(block)
        return block

    def append_form(
        self,
        label: str = "",
        *,
        form_type: str = "INPUT",
        name: str | None = None,
        value: str | None = None,
        checked: bool = False,
        items: list[str] | None = None,
        editable: bool = True,
        locked: bool = False,
        placeholder: str | None = None,
        section_index: int = 0,
    ) -> Form:
        section = self._ensure_section(section_index)
        block = Form(
            label=label,
            form_type=form_type,
            name=name,
            value=value,
            checked=checked,
            items=list(items or []),
            editable=editable,
            locked=locked,
            placeholder=placeholder,
        )
        section.blocks.append(block)
        return block

    def append_memo(
        self,
        text: str,
        *,
        author: str | None = None,
        memo_id: str | None = None,
        anchor_id: str | None = None,
        order: int | None = None,
        visible: bool = True,
        section_index: int = 0,
    ) -> Memo:
        section = self._ensure_section(section_index)
        block = Memo(
            text=text,
            author=author,
            memo_id=memo_id,
            anchor_id=anchor_id,
            order=order,
            visible=visible,
        )
        section.blocks.append(block)
        return block

    def append_chart(
        self,
        title: str = "",
        *,
        chart_type: str = "BAR",
        categories: list[str] | None = None,
        series: list[dict[str, object]] | None = None,
        data_ref: str | None = None,
        legend_visible: bool = True,
        width: int = 12000,
        height: int = 3200,
        shape_comment: str | None = None,
        section_index: int = 0,
    ) -> Chart:
        section = self._ensure_section(section_index)
        block = Chart(
            title=title,
            chart_type=chart_type,
            categories=list(categories or []),
            series=[dict(value) for value in (series or [])],
            data_ref=data_ref,
            legend_visible=legend_visible,
            width=width,
            height=height,
            shape_comment=shape_comment,
        )
        section.blocks.append(block)
        return block

    def append_footnote(
        self,
        text: str,
        *,
        number: int | None = None,
        section_index: int = 0,
    ) -> Note:
        return self.append_note(text, kind="footNote", number=number, section_index=section_index)

    def append_endnote(
        self,
        text: str,
        *,
        number: int | None = None,
        section_index: int = 0,
    ) -> Note:
        return self.append_note(text, kind="endNote", number=number, section_index=section_index)

    def append_equation(
        self,
        script: str,
        *,
        width: int = 4800,
        height: int = 2300,
        shape_comment: str | None = None,
        text_color: str = "#000000",
        base_unit: int = 1100,
        font: str = "HYhwpEQ",
        section_index: int = 0,
    ) -> Equation:
        section = self._ensure_section(section_index)
        block = Equation(
            script=script,
            width=width,
            height=height,
            shape_comment=shape_comment,
            text_color=text_color,
            base_unit=base_unit,
            font=font,
        )
        section.blocks.append(block)
        return block

    def append_shape(
        self,
        *,
        kind: str = "rect",
        text: str = "",
        width: int = 12000,
        height: int = 3200,
        fill_color: str = "#FFFFFF",
        line_color: str = "#000000",
        line_width: int = 33,
        shape_comment: str | None = None,
        original_width: int | None = None,
        original_height: int | None = None,
        current_width: int | None = None,
        current_height: int | None = None,
        section_index: int = 0,
    ) -> Shape:
        section = self._ensure_section(section_index)
        block = Shape(
            kind=kind,
            text=text,
            width=width,
            height=height,
            fill_color=fill_color,
            line_color=line_color,
            shape_comment=shape_comment,
            line_width=line_width,
            original_width=original_width,
            original_height=original_height,
            current_width=current_width,
            current_height=current_height,
        )
        section.blocks.append(block)
        return block

    def append_ole(
        self,
        name: str,
        data: bytes,
        *,
        width: int = 42001,
        height: int = 13501,
        shape_comment: str | None = None,
        object_type: str = "EMBEDDED",
        draw_aspect: str = "CONTENT",
        has_moniker: bool = False,
        eq_baseline: int = 0,
        line_color: str = "#000000",
        line_width: int = 0,
        original_width: int | None = None,
        original_height: int | None = None,
        current_width: int | None = None,
        current_height: int | None = None,
        section_index: int = 0,
    ) -> Ole:
        section = self._ensure_section(section_index)
        block = Ole(
            name=name,
            data=data,
            width=width,
            height=height,
            shape_comment=shape_comment,
            object_type=object_type,
            draw_aspect=draw_aspect,
            has_moniker=has_moniker,
            eq_baseline=eq_baseline,
            line_color=line_color,
            line_width=line_width,
            original_width=original_width,
            original_height=original_height,
            current_width=current_width,
            current_height=current_height,
        )
        section.blocks.append(block)
        return block

    def append_header(
        self,
        text: str,
        *,
        apply_page_type: str = "BOTH",
        section_index: int = 0,
    ) -> HeaderFooter:
        section = self._ensure_section(section_index)
        block = HeaderFooter(kind="header", text=text, apply_page_type=apply_page_type)
        section.header_footer_blocks.append(block)
        return block

    def append_footer(
        self,
        text: str,
        *,
        apply_page_type: str = "BOTH",
        section_index: int = 0,
    ) -> HeaderFooter:
        section = self._ensure_section(section_index)
        block = HeaderFooter(kind="footer", text=text, apply_page_type=apply_page_type)
        section.header_footer_blocks.append(block)
        return block

    def append_style(
        self,
        name: str,
        *,
        style_id: str | None = None,
        english_name: str | None = None,
        style_type: str = "PARA",
        para_pr_id: str | None = None,
        char_pr_id: str | None = None,
        next_style_id: str | None = None,
        lang_id: str = "1042",
        lock_form: bool = False,
        native_hwp_payload: bytes | None = None,
    ) -> StyleDefinition:
        node = StyleDefinition(
            name=name,
            style_id=style_id,
            english_name=english_name,
            style_type=style_type,
            para_pr_id=para_pr_id,
            char_pr_id=char_pr_id,
            next_style_id=next_style_id,
            lang_id=lang_id,
            lock_form=lock_form,
            native_hwp_payload=native_hwp_payload,
        )
        self.style_definitions.append(node)
        return node

    def append_paragraph_style(
        self,
        *,
        style_id: str | None = None,
        alignment_horizontal: str | None = None,
        alignment_vertical: str | None = None,
        line_spacing: int | str | None = None,
        line_spacing_type: str | None = None,
        margins: dict[str, str] | None = None,
        tab_pr_id: str | None = None,
        snap_to_grid: str | None = None,
        condense: str | None = None,
        font_line_height: str | None = None,
        suppress_line_numbers: str | None = None,
        checked: str | None = None,
        heading: dict[str, str] | None = None,
        break_setting: dict[str, str] | None = None,
        auto_spacing: dict[str, str] | None = None,
    ) -> ParagraphStyle:
        node = ParagraphStyle(
            style_id=style_id,
            alignment_horizontal=alignment_horizontal,
            alignment_vertical=alignment_vertical,
            line_spacing=line_spacing,
            line_spacing_type=line_spacing_type,
            margins=dict(margins or {}),
            tab_pr_id=tab_pr_id,
            snap_to_grid=snap_to_grid,
            condense=condense,
            font_line_height=font_line_height,
            suppress_line_numbers=suppress_line_numbers,
            checked=checked,
            heading=dict(heading or {}),
            break_setting=dict(break_setting or {}),
            auto_spacing=dict(auto_spacing or {}),
        )
        self.paragraph_styles.append(node)
        return node

    def append_character_style(
        self,
        *,
        style_id: str | None = None,
        text_color: str | None = None,
        shade_color: str | None = None,
        height: int | str | None = None,
        use_font_space: str | None = None,
        use_kerning: str | None = None,
        sym_mark: str | None = None,
        font_refs: dict[str, str] | None = None,
        font_names: dict[str, str] | None = None,
        font_ratios: dict[str, int] | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: dict[str, str] | None = None,
        strikeout: dict[str, str] | None = None,
        script_position: str | None = None,
    ) -> CharacterStyle:
        node = CharacterStyle(
            style_id=style_id,
            text_color=text_color,
            shade_color=shade_color,
            height=height,
            use_font_space=use_font_space,
            use_kerning=use_kerning,
            sym_mark=sym_mark,
            font_refs=dict(font_refs or {}),
            font_names=dict(font_names or {}),
            font_ratios=dict(font_ratios or {}),
            bold=bold,
            italic=italic,
            underline=dict(underline or {}),
            strikeout=dict(strikeout or {}),
            script_position=script_position,
        )
        self.character_styles.append(node)
        return node

    def append_numbering_definition(
        self,
        *,
        style_id: str | None = None,
        formats: list[str] | None = None,
        para_heads: list[int] | None = None,
        para_head_flags: list[dict[str, bool]] | None = None,
        width_adjusts: list[int] | None = None,
        text_offsets: list[int] | None = None,
        char_pr_ids: list[int | None] | None = None,
        start_numbers: list[int] | None = None,
        unknown_short: int = 0,
        unknown_short_bits: dict[str, bool] | None = None,
        native_hwp_payload: bytes | None = None,
    ) -> NumberingDefinition:
        node = NumberingDefinition(
            style_id=style_id,
            formats=list(formats or []),
            para_heads=list(para_heads or []),
            para_head_flags=[dict(value) for value in (para_head_flags or [])],
            width_adjusts=list(width_adjusts or []),
            text_offsets=list(text_offsets or []),
            char_pr_ids=list(char_pr_ids or []),
            start_numbers=list(start_numbers or []),
            unknown_short=unknown_short,
            unknown_short_bits=dict(unknown_short_bits or {}),
            native_hwp_payload=native_hwp_payload,
        )
        self.numbering_definitions.append(node)
        return node

    def append_bullet_definition(
        self,
        *,
        style_id: str | None = None,
        flags: int | None = None,
        flag_bits: dict[str, bool] | None = None,
        bullet_char: str = "\uf0a1",
        width_adjust: int = 0,
        text_offset: int = 50,
        char_pr_id: int | None = None,
        unknown_tail: bytes | None = None,
        native_hwp_payload: bytes | None = None,
    ) -> BulletDefinition:
        node = BulletDefinition(
            style_id=style_id,
            flags=flags,
            flag_bits=dict(flag_bits or {}),
            bullet_char=bullet_char,
            width_adjust=width_adjust,
            text_offset=text_offset,
            char_pr_id=char_pr_id,
            unknown_tail=unknown_tail,
            native_hwp_payload=native_hwp_payload,
        )
        self.bullet_definitions.append(node)
        return node

    def append_memo_shape_definition(
        self,
        *,
        memo_shape_id: str | None = None,
        width: int | None = None,
        line_width: int | None = None,
        line_type: str | None = None,
        line_color: str | None = None,
        fill_color: str | None = None,
        active_color: str | None = None,
        memo_type: str | None = None,
    ) -> MemoShapeDefinition:
        if memo_shape_id is None:
            used_ids = {
                int(node.memo_shape_id)
                for node in self.memo_shape_definitions
                if node.memo_shape_id is not None and str(node.memo_shape_id).isdigit()
            }
            next_id = 0
            while next_id in used_ids:
                next_id += 1
            memo_shape_id = str(next_id)
        node = MemoShapeDefinition(
            memo_shape_id=memo_shape_id,
            width=width,
            line_width=line_width,
            line_type=line_type,
            line_color=line_color,
            fill_color=fill_color,
            active_color=active_color,
            memo_type=memo_type,
        )
        self.memo_shape_definitions.append(node)
        return node

    def to_hwpx_document(self, *, preserve_formatting: bool | None = None, converter: HancomConverter | None = None) -> HwpxDocument:
        if preserve_formatting:
            raise ValueError("Formatting-preserving conversion requires a WordDocument.")
        if getattr(self, "_word_feature_model", False) and not getattr(self, "_word_carriers_prepared", False):
            from .addons.docx.features import carrier_model

            return carrier_model(self).to_hwpx_document(converter=converter)
        document = HwpxDocument.blank()
        self.metadata.apply_to_hwpx_document(document)
        _apply_styles(document, self)
        while len(document.sections) < len(self.sections):
            document.add_section(text="")
        for section_index in range(len(self.sections)):
            _reset_hwpx_section(document, section_index)
        for section_index, section in enumerate(self.sections):
            uses_source_paragraphs = _section_uses_source_paragraphs(section)
            _apply_section_settings(
                document,
                section_index,
                section.settings,
                apply_page_numbers=not uses_source_paragraphs,
            )
            if not uses_source_paragraphs:
                _apply_header_footer_blocks(document, section_index, section.header_footer_blocks)
            _write_section_to_hwpx(document, section_index, section)
            if uses_source_paragraphs and section.settings.page_numbers:
                _apply_page_numbers_to_source_paragraphs(document, section_index, section.settings)
        for section in document.sections:
            for text_node in section.root_element.xpath(".//hp:t", namespaces=NS):
                if not len(text_node) and any(char in (text_node.text or "") for char in "\t\r\n"):
                    _set_text_node_content(text_node, text_node.text or "")
                    section.mark_modified()
        if getattr(self, "_word_feature_model", False):
            from .addons.docx.native_effects import activate_lists

            activate_lists(document)
        return document

    def write_to_hwpx(self, path: str | Path, *, validate: bool = True, preserve_formatting: bool | None = None, converter: HancomConverter | None = None) -> Path:
        return self.to_hwpx_document(preserve_formatting=preserve_formatting, converter=converter).save(path, validate=validate)

    def to_docx_document(self, *, converter: HancomConverter | None = None, preserve_formatting: bool | None = None):
        """Export the current shared model as an editable WordDocument.

        The default path is portable Python. An explicitly supplied converter
        receives a current HWPX snapshot and the ``DOCX`` output format.
        """
        from .addons.docx.writer import build_docx

        effective_converter = converter or self._converter
        directory = self._ensure_temp_dir()
        output = directory / "exported.docx"
        if preserve_formatting is False or (preserve_formatting is None and effective_converter is None):
            output.write_bytes(build_docx(self))
        else:
            if effective_converter is None:
                raise ValueError("Formatting-preserving DOCX export requires an explicit converter.")
            original = getattr(self, "_docx_conversion_source", None)
            if original is not None and original[2] == self._docx_export_state():
                source = directory / ("docx-source." + original[1])
                source.write_bytes(original[0])
            else:
                source = self.to_hwpx_document().save(directory / "docx-source.hwpx")
            output.unlink(missing_ok=True)
            output = Path(effective_converter(source, output, "DOCX"))
            if not output.is_file():
                raise ValueError("The DOCX converter did not produce an output file.")
        return Document.read_docx(output)

    def write_to_docx(self, path: str | Path, *, converter: HancomConverter | None = None, preserve_formatting: bool | None = None) -> Path:
        return self.to_docx_document(converter=converter, preserve_formatting=preserve_formatting).write_to_docx(path)

    def save(
        self,
        path: str | Path,
        *,
        validate: bool = True,
        converter: HancomConverter | None = None,
        preserve_formatting: bool | None = None,
    ) -> Path:
        target = Path(path)
        suffix = target.suffix.lower()
        if suffix == ".docx":
            return self.write_to_docx(target, converter=converter, preserve_formatting=preserve_formatting)
        if suffix == ".hwpx":
            return self.write_to_hwpx(target, validate=validate, preserve_formatting=preserve_formatting, converter=converter)
        if suffix == ".hwp":
            return self.write_to_hwp(target, converter=converter, preserve_formatting=preserve_formatting)
        raise ValueError(f"Unsupported Hancom document path: {target}. Expected .hwp, .hwpx, or .docx.")

    def to_hwp_document(self, *, converter: HancomConverter | None = None, preserve_formatting: bool | None = None) -> HwpDocument:
        if preserve_formatting:
            raise ValueError("Formatting-preserving conversion requires a WordDocument.")
        effective_converter = converter or self._converter
        document = HwpDocument.blank(converter=effective_converter)
        document.ensure_section_count(max(len(self.sections), 1))
        preview_text = self.metadata.title or self.metadata.subject or self.metadata.description
        if preview_text:
            document.set_preview_text(preview_text)
        _sync_hwp_docinfo_styles(document, self)
        docinfo = document.docinfo_model()
        paragraph_attribute_limits = (len(docinfo.para_shape_records()), len(docinfo.style_records()))
        for section_index, section in enumerate(self.sections):
            width, height, landscape = section.settings.page_width, section.settings.page_height, section.settings.landscape
            if landscape == "NARROWLY":
                width, height, landscape = height, width, "LANDSCAPE"
            document.apply_section_settings(
                section_index=section_index,
                page_width=width,
                page_height=height,
                landscape=landscape,
                margins=section.settings.margins or None,
                visibility=section.settings.visibility or None,
                grid=section.settings.grid or None,
                start_numbers=section.settings.start_numbers or None,
                numbering_shape_id=section.settings.numbering_shape_id,
                memo_shape_id=section.settings.memo_shape_id,
            )
            if section.settings.page_border_fills:
                document.apply_section_page_border_fills(
                    section.settings.page_border_fills,
                    section_index=section_index,
                )
            if section.settings.page_numbers:
                document.apply_section_page_numbers(
                    section.settings.page_numbers,
                    section_index=section_index,
                )
            if section.settings.footnote_pr or section.settings.endnote_pr:
                document.apply_section_note_settings(
                    section_index=section_index,
                    footnote_pr=section.settings.footnote_pr or None,
                    endnote_pr=section.settings.endnote_pr or None,
                )
            header_written = False
            footer_written = False
            for block in section.header_footer_blocks:
                if block.apply_page_type == "FIRST":
                    if block.nested_blocks:
                        import warnings

                        warnings.warn("HWP export: first-page stories retain their text in a page-relative shape; use HWPX or an explicit native converter for rich story content.", UserWarning, stacklevel=2)
                    _append_hancom_block_to_hwp(document, _first_page_story_shape(block, section.settings), section_index=section_index)
                    continue
                if block.kind == "header" and not header_written:
                    header = document.append_header(block.text, apply_page_type=block.apply_page_type, section_index=section_index)
                    if header is not None and block.nested_blocks and hasattr(header, "set_nested_blocks"):
                        header.set_nested_blocks(block.nested_blocks)
                    header_written = True
                elif block.kind == "footer" and not footer_written:
                    footer = document.append_footer(block.text, apply_page_type=block.apply_page_type, section_index=section_index)
                    if footer is not None and block.nested_blocks and hasattr(footer, "set_nested_blocks"):
                        footer.set_nested_blocks(block.nested_blocks)
                    footer_written = True
            binary_document = document.binary_document()
            batch_active = False
            try:
                for block in _hwp_effect_blocks(section.blocks):
                    uses_direct_section_records = isinstance(block, (Table, Picture, Hyperlink))
                    if uses_direct_section_records and batch_active:
                        binary_document.end_section_model_batch()
                        batch_active = False
                    elif not uses_direct_section_records and not batch_active:
                        binary_document.begin_section_model_batch()
                        batch_active = True
                    _append_hancom_block_to_hwp(
                        document,
                        block,
                        section_index=section_index,
                        paragraph_attribute_limits=paragraph_attribute_limits,
                    )
            finally:
                if batch_active:
                    binary_document.end_section_model_batch()
        return document

    def write_to_hwp(self, path: str | Path, *, converter: HancomConverter | None = None, preserve_formatting: bool | None = None) -> Path:
        document = self.to_hwp_document(converter=converter, preserve_formatting=preserve_formatting)
        return document.save(path)

    def _ensure_section(self, section_index: int) -> HancomSection:
        while len(self.sections) <= section_index:
            self.sections.append(HancomSection())
        return self.sections[section_index]

    def _ensure_temp_dir(self) -> Path:
        if self._temp_dir is None:
            self._temp_dir = Path(tempfile.mkdtemp(prefix="jakal_hancom_document_"))
        return self._temp_dir


from ._hancom_conversion_common import (  # noqa: E402, F401 - compatibility re-exports
    _blank_cell_texts,
    _normalize_hwp_text,
    _drop_surrogate_codepoints,
    _xml_safe_scalar,
    _normalize_str_map,
    _normalize_int_map,
    _block_owned_text,
    _block_owned_text_for_write,
    _is_owned_text_shadow_paragraph,
    _normalize_hwp_setting_scalar,
    _has_word_feature_carriers,
    _json_list_parameter,
    _nested_blocks,
    _has_nested_blocks,
    _media_type_for_extension,
)
from ._hancom_hwp_read import (  # noqa: E402, F401 - compatibility re-exports
    _extract_hwp_section,
    _hwp_effect_blocks,
    _extract_hwp_paragraph_blocks,
    _extract_supported_nested_hwp_blocks,
    _extract_hwp_nested_paragraph_blocks,
    _extract_hwp_section_settings,
    _build_hwp_paragraph_block,
    _tokenize_hwp_paragraph_raw_text,
    _find_matching_control_index,
    _attach_hwp_docx_table,
    _extract_hwp_table_column_widths,
    _apply_hwp_shape_specific_fields,
    _parse_hwp_field_parameters,
    _resolve_hwp_paragraph_attributes,
    _extract_hwp_header_footer,
    _extract_hwp_header_footer_nested_blocks,
)
from ._hancom_hwpx_graphics import (  # noqa: E402, F401 - compatibility re-exports
    _optional_hwpx_int_attr,
    _extract_hwpx_graphic_size_fields,
    _extract_hwpx_graphic_size_attributes,
    _apply_hwpx_graphic_size_node,
    _remove_hwpx_child_nodes,
    _apply_hwpx_optional_graphic_nodes,
    _hwpx_layout_kwargs,
    _hwpx_rotation_kwargs,
    _hwpx_line_style_kwargs,
    _extract_hwpx_shape_points,
    _extract_hwpx_shape_specific_fields,
    _NESTED_TEXT_CONTROL_NAMES,
    _extract_hwpx_direct_text,
    _extract_hwpx_shape_direct_text,
    _set_hwpx_shape_points,
    _apply_hwpx_shape_specific_fields,
    _apply_hwpx_table_block,
    _ensure_hwpx_table_border_fill,
    _apply_hwpx_picture_block,
    _apply_hwpx_equation_block,
    _apply_hwpx_shape_block,
    _ensure_hwpx_shape_draw_text,
    _apply_hwpx_ole_block,
    _extract_hwpx_table,
    _extract_hwpx_picture,
    _extract_hwpx_note,
    _extract_hwpx_memo,
    _extract_hwpx_memo_shape,
    _extract_hwpx_equation,
    _extract_hwpx_chart,
    _extract_hwpx_shape,
    _extract_hwpx_ole,
)
from ._hancom_hwpx_read import (  # noqa: E402, F401 - compatibility re-exports
    _extract_hwpx_section,
    _hwpx_paragraph_has_header_footer,
    _sort_hwpx_blocks_by_source_order,
    _sort_hwpx_nested_blocks_by_source_order,
    _mark_hwpx_nested_paragraph,
    _extract_hwpx_text_container_blocks,
    _extract_hwpx_nested_paragraph_blocks,
    _extract_missing_hwpx_descendant_blocks,
    _extract_hwpx_caption_text,
    _CAPTION_TEXT_BLOCK_NAMES,
    _hwpx_node_is_inside_table_cell,
    _HWPX_TEXT_CONTAINER_NAMES,
    _HWPX_NESTED_INLINE_BLOCK_NAMES,
    _hwpx_node_is_inside_text_container,
    _clear_nested_owned_text,
    _mark_hwpx_text_owned_by_container,
    _hwpx_owned_inline_text,
    _hwpx_node_key,
    _extract_hwpx_run_node,
    _hwpx_chart_is_covered_by_carrier_shape,
    _source_order_for_element,
    _mark_hwpx_source_paragraph,
    _build_hwpx_paragraph_block,
    _extract_hwpx_paragraph_char_pr_id,
    _extract_bookmark,
    _extract_field,
    _extract_hwpx_form_field,
    _extract_hwpx_form,
    _extract_hwpx_memo_field,
    _memo_requires_hwpx_carrier,
    _merge_hwpx_memo_with_carrier,
    _merge_hwpx_memos_with_carriers,
    _find_matching_hwpx_memo_carrier,
    _normalize_hwpx_memo_pairing_key,
    _extract_hwpx_chart_carrier,
    _extract_auto_number,
    _extract_header_footer_blocks,
    _source_paragraph_position_for_element,
    _source_paragraph_index_for_element,
    _build_table_wrapper,
    _build_picture_wrapper,
    _build_field_wrapper,
    _build_bookmark_wrapper,
    _build_auto_number_wrapper,
    _build_note_wrapper,
    _build_memo_wrapper,
    _build_form_wrapper,
    _build_equation_wrapper,
    _build_shape_wrapper,
    _build_chart_wrapper,
    _build_ole_wrapper,
)
from ._hancom_hwpx_write import (  # noqa: E402, F401 - compatibility re-exports
    _build_hwpx_form_parameters,
    _build_hwpx_memo_parameters,
    _build_hwpx_chart_carrier_comment,
    _reset_hwpx_section,
    _section_uses_source_paragraphs,
    _write_grouped_section_to_hwpx,
    _ensure_hwpx_source_paragraph,
    _append_text_to_hwpx_paragraph,
    _is_supported_nested_hwpx_block,
    _build_hwpx_nested_sublist,
    _append_nested_block_to_hwpx_paragraph,
    _append_nested_field_to_hwpx_paragraph,
    _hwpx_hyperlink_parameters,
    _apply_hwpx_hyperlink_runs,
    _hwpx_field_kwargs,
    _WORD_EFFECT_FIELD_PREFIX,
    _word_effect_from_field_name,
    _apply_hwpx_nested_text_blocks,
    _append_header_footer_block_to_hwpx,
    _FIRST_PAGE_STORY_MARKER,
    _first_page_story_shape,
    _restore_first_page_stories,
    _append_control_host_paragraph,
    _apply_hwpx_paragraph_style_block,
    _resolve_hwpx_style_id,
    _resolve_hwpx_para_pr_id,
    _resolve_hwpx_char_pr_id,
)
from ._hancom_hwpx_settings import (  # noqa: E402, F401 - compatibility re-exports
    _extract_page_border_fills,
    _extract_page_numbers_with_positions,
    _extract_page_numbers,
    _is_direct_section_paragraph_control,
    _extract_note_pr,
    _extract_line_number_shape,
    _extract_section_settings,
    _apply_section_settings,
    _apply_page_border_fills,
    _apply_line_number_shape,
    _apply_note_settings,
    _append_note_pr,
    _apply_page_numbers,
    _apply_page_numbers_to_source_paragraphs,
    _page_number_xml,
    _apply_header_footer_blocks,
)
from ._hancom_hwpx_styles import (  # noqa: E402, F401 - compatibility re-exports
    _apply_styles,
)
# Backward-compatible name; new application code uses Document.
HancomDocument = Document

# Direction-specific converters are imported late so they can reuse the shared
# format-neutral helper layer without exposing circular imports to callers.
from ._hancom_from_hwp import _extract_hwp_control  # noqa: E402, F401
from ._hancom_from_hwpx import _extract_hwpx_paragraph_blocks  # noqa: E402, F401
from ._hancom_to_hwp import _append_hancom_block_to_hwp  # noqa: E402, F401
from ._hancom_to_hwpx import (  # noqa: E402, F401
    _append_non_paragraph_block_to_hwpx,
    _write_section_to_hwpx,
)
