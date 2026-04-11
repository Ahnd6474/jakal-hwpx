from __future__ import annotations

import tempfile
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from .document import HwpxDocument
from .hwp_binary import (
    _build_table_cell_list_header_payload,
    _build_table_record_payload,
    _build_hyperlink_ctrl_payload,
    _build_hyperlink_para_text_payload,
    _TableCellSpec,
    _TABLE_CELL_CHAR_SHAPE,
    _TABLE_CELL_LINE_SEG,
    _TABLE_CELL_PARA_HEADER_TAIL,
    DEFAULT_HWP_TABLE_CELL_HEIGHT,
    DEFAULT_HWP_TABLE_CELL_WIDTH,
    DocInfoModel,
    HwpBinaryDocument,
    HwpBinaryFileHeader,
    HwpDocumentProperties,
    HwpParagraph,
    HwpStreamCapacity,
    ParagraphHeaderRecord,
    ParagraphTextRecord,
    RecordNode,
    SectionModel,
    SectionParagraphModel,
    TAG_CTRL_DATA,
    TAG_CTRL_HEADER,
    TAG_EQEDIT,
    TAG_LIST_HEADER,
    TAG_PARA_CHAR_SHAPE,
    TAG_PARA_HEADER,
    TAG_PARA_LINE_SEG,
    TAG_SHAPE_COMPONENT,
    TAG_SHAPE_COMPONENT_ARC,
    TAG_SHAPE_COMPONENT_CONTAINER,
    TAG_SHAPE_COMPONENT_CURVE,
    TAG_SHAPE_COMPONENT_ELLIPSE,
    TAG_SHAPE_COMPONENT_LINE,
    TAG_SHAPE_COMPONENT_PICTURE,
    TAG_SHAPE_COMPONENT_OLE,
    TAG_SHAPE_COMPONENT_POLYGON,
    TAG_SHAPE_COMPONENT_RECTANGLE,
    TAG_SHAPE_COMPONENT_TEXTART,
    TAG_TABLE,
)
from .hwp_pure_profile import HwpPureProfile, append_feature_from_profile

HancomConverter = Callable[[str | Path, str | Path, str], Path]


class HwpDocument:
    def __init__(self, binary_document: HwpBinaryDocument, *, converter: HancomConverter | None = None):
        self._binary_document = binary_document
        self._converter = converter
        self._bridge_document: HwpxDocument | None = None
        self._bridge_temp_dir: Path | None = None
        self._pure_profile: HwpPureProfile | None = None

    @classmethod
    def blank(cls, *, converter: HancomConverter | None = None) -> "HwpDocument":
        instance = cls.blank_from_profile(HwpPureProfile.load_bundled().root, converter=converter)
        return instance

    @classmethod
    def open(cls, path: str | Path, *, converter: HancomConverter | None = None) -> "HwpDocument":
        return cls(HwpBinaryDocument.open(path), converter=converter)

    @classmethod
    def blank_from_profile(
        cls,
        profile_root: str | Path,
        *,
        converter: HancomConverter | None = None,
    ) -> "HwpDocument":
        profile = HwpPureProfile.load(profile_root)
        instance = cls(HwpBinaryDocument.open(profile.base_path), converter=converter)
        instance._pure_profile = profile
        instance._binary_document.reset_body_sections_to_blank()
        return instance

    @property
    def source_path(self) -> Path:
        return self._binary_document.source_path

    def binary_document(self) -> HwpBinaryDocument:
        return self._binary_document

    def file_header(self) -> HwpBinaryFileHeader:
        return self._binary_document.file_header()

    def document_properties(self) -> HwpDocumentProperties:
        return self._binary_document.document_properties()

    def docinfo_model(self) -> DocInfoModel:
        return self._binary_document.docinfo_model()

    def list_stream_paths(self) -> list[str]:
        return self._binary_document.list_stream_paths()

    def bindata_stream_paths(self) -> list[str]:
        return self._binary_document.bindata_stream_paths()

    def stream_capacity(self, path: str) -> HwpStreamCapacity:
        data = self._binary_document.read_stream(path, decompress=self._binary_document._stream_is_compressed(path))
        return self._binary_document.stream_capacity(
            path,
            data,
            compress=self._binary_document._stream_is_compressed(path),
        )

    def preview_text(self) -> str:
        return self._binary_document.preview_text()

    def set_preview_text(self, value: str) -> None:
        self._invalidate_bridge()
        self._binary_document.set_preview_text(value)

    def sections(self) -> list["HwpSection"]:
        return [HwpSection(self, index) for index, _path in enumerate(self._binary_document.section_stream_paths())]

    def section(self, index: int) -> "HwpSection":
        return HwpSection(self, index)

    def section_model(self, section_index: int = 0) -> SectionModel:
        return self._binary_document.section_model(section_index)

    def ensure_section_count(self, section_count: int) -> None:
        self._invalidate_bridge()
        self._binary_document.ensure_section_count(section_count)

    def apply_section_settings(
        self,
        *,
        section_index: int = 0,
        page_width: int | None = None,
        page_height: int | None = None,
        landscape: str | None = None,
        margins: dict[str, int] | None = None,
        visibility: dict[str, str] | None = None,
        grid: dict[str, int] | None = None,
        start_numbers: dict[str, str] | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.set_section_page_settings(
            section_index,
            page_width=page_width,
            page_height=page_height,
            landscape=landscape,
            margins=margins,
        )
        self._binary_document.set_section_definition_settings(
            section_index,
            visibility=visibility,
            grid=grid,
            start_numbers=start_numbers,
        )

    def apply_section_page_border_fills(
        self,
        page_border_fills: list[dict[str, str | int]],
        *,
        section_index: int = 0,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.set_section_page_border_fills(section_index, page_border_fills)

    def apply_section_page_numbers(
        self,
        page_numbers: list[dict[str, str]],
        *,
        section_index: int = 0,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.set_section_page_numbers(section_index, page_numbers)

    def apply_section_note_settings(
        self,
        *,
        section_index: int = 0,
        footnote_pr: dict[str, object] | None = None,
        endnote_pr: dict[str, object] | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.set_section_note_settings(
            section_index,
            footnote_pr=footnote_pr,
            endnote_pr=endnote_pr,
        )

    def paragraphs(self, section_index: int = 0) -> list["HwpParagraphObject"]:
        return self.section(section_index).paragraphs()

    def controls(self, section_index: int | None = None) -> list["HwpControlObject"]:
        sections = [self.section(section_index)] if section_index is not None else self.sections()
        controls: list[HwpControlObject] = []
        for section in sections:
            controls.extend(section.controls())
        return controls

    def tables(self, section_index: int | None = None) -> list["HwpTableObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpTableObject)]

    def pictures(self, section_index: int | None = None) -> list["HwpPictureObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpPictureObject)]

    def hyperlinks(self, section_index: int | None = None) -> list["HwpHyperlinkObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpHyperlinkObject)]

    def bookmarks(self, section_index: int | None = None) -> list["HwpBookmarkObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpBookmarkObject)]

    def notes(self, section_index: int | None = None) -> list["HwpNoteObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpNoteObject)]

    def page_numbers(self, section_index: int | None = None) -> list["HwpPageNumObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpPageNumObject)]

    def fields(self, section_index: int | None = None) -> list["HwpFieldObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpFieldObject)]

    def shapes(self, section_index: int | None = None) -> list["HwpShapeObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpShapeObject)]

    def oles(self, section_index: int | None = None) -> list["HwpOleObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpOleObject)]

    def equations(self, section_index: int | None = None) -> list["HwpEquationObject"]:
        return [control for control in self.controls(section_index) if isinstance(control, HwpEquationObject)]

    def get_document_text(self) -> str:
        return self._binary_document.get_document_text()

    def replace_text_same_length(
        self,
        old: str,
        new: str,
        *,
        section_index: int | None = None,
        count: int = -1,
    ) -> int:
        return self._binary_document.replace_text_same_length(old, new, section_index=section_index, count=count)

    def append_paragraph(
        self,
        text: str,
        *,
        section_index: int = 0,
        para_shape_id: int = 0,
        style_id: int = 0,
        split_flags: int = 0,
        control_mask: int = 0,
    ) -> HwpParagraphObject:
        self._invalidate_bridge()
        paragraph = self._binary_document.append_paragraph(
            text,
            section_index=section_index,
            para_shape_id=para_shape_id,
            style_id=style_id,
            split_flags=split_flags,
            control_mask=control_mask,
        )
        return HwpParagraphObject(
            self,
            HwpParagraph(
                index=paragraph.index,
                section_index=paragraph.section_index,
                char_count=paragraph.header.char_count,
                control_mask=paragraph.header.control_mask,
                para_shape_id=paragraph.header.para_shape_id,
                style_id=paragraph.header.style_id,
                split_flags=paragraph.header.split_flags,
                raw_text=paragraph.raw_text,
                text=paragraph.text,
                text_record_offset=-1,
                text_record_size=len(paragraph.raw_text.encode("utf-16-le")),
            ),
        )

    def to_hwpx_document(self, *, force_refresh: bool = False) -> HwpxDocument:
        if self._bridge_document is not None and not force_refresh:
            return self._bridge_document

        from .hancom_document import HancomDocument

        self._bridge_document = HancomDocument.from_hwp_document(self, converter=self._converter).to_hwpx_document()
        return self._bridge_document

    def save_as_hwpx(self, path: str | Path) -> Path:
        target_path = Path(path).expanduser().resolve()
        target_path.parent.mkdir(parents=True, exist_ok=True)
        self.to_hwpx_document(force_refresh=False).save(target_path)
        return target_path

    def save(self, path: str | Path | None = None) -> Path:
        if path is not None:
            target_path = Path(path).expanduser().resolve()
            if target_path.suffix.lower() == ".hwpx":
                return self.save_as_hwpx(target_path)
        else:
            target_path = self.source_path

        if self._bridge_document is None:
            return self._binary_document.save(path)

        self._sync_from_bridge_document()
        self._binary_document.save(target_path)
        self._binary_document.source_path = target_path
        return target_path

    def load_pure_profile(self, profile_root: str | Path) -> HwpPureProfile:
        self._pure_profile = HwpPureProfile.load(profile_root)
        return self._pure_profile

    def append_table(
        self,
        cell_text: str | None = None,
        *,
        rows: int = 1,
        cols: int = 1,
        cell_texts: Sequence[str] | Sequence[Sequence[str]] | None = None,
        row_heights: Sequence[int] | None = None,
        col_widths: Sequence[int] | None = None,
        cell_spans: dict[tuple[int, int], tuple[int, int]] | None = None,
        cell_border_fill_ids: dict[tuple[int, int], int] | None = None,
        table_border_fill_id: int = 1,
        section_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        if self._pure_profile is None:
            self._pure_profile = HwpPureProfile.load_bundled()
        self._binary_document.append_table(
            cell_text,
            rows=rows,
            cols=cols,
            cell_texts=cell_texts,
            row_heights=row_heights,
            col_widths=col_widths,
            cell_spans=cell_spans,
            cell_border_fill_ids=cell_border_fill_ids,
            table_border_fill_id=table_border_fill_id,
            profile_root=self._pure_profile.root,
            section_index=section_index,
        )

    def append_picture(
        self,
        image_bytes: bytes | None = None,
        *,
        extension: str | None = None,
        section_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        if self._pure_profile is None:
            self._pure_profile = HwpPureProfile.load_bundled()
        self._binary_document.append_picture(
            image_bytes,
            extension=extension,
            profile_root=self._pure_profile.root,
            section_index=section_index,
        )

    def append_hyperlink(
        self,
        url: str | None = None,
        *,
        text: str | None = None,
        metadata_fields: Sequence[str | int] | None = None,
        section_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        if self._pure_profile is None:
            self._pure_profile = HwpPureProfile.load_bundled()
        self._binary_document.append_hyperlink(
            url,
            text=text,
            metadata_fields=metadata_fields,
            profile_root=self._pure_profile.root,
            section_index=section_index,
        )

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
        paragraph_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_field(
            field_type=field_type,
            display_text=display_text,
            name=name,
            parameters=parameters,
            editable=editable,
            dirty=dirty,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_bookmark(
        self,
        name: str,
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_bookmark(name, section_index=section_index, paragraph_index=paragraph_index)

    def append_note(
        self,
        text: str,
        *,
        kind: str = "footNote",
        number: int | str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_note(
            text,
            kind=kind,
            number=number,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_footnote(
        self,
        text: str,
        *,
        number: int | str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self.append_note(
            text,
            kind="footNote",
            number=number,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_endnote(
        self,
        text: str,
        *,
        number: int | str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self.append_note(
            text,
            kind="endNote",
            number=number,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_auto_number(
        self,
        *,
        number: int | str = 1,
        number_type: str = "PAGE",
        kind: str = "newNum",
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_auto_number(
            number=number,
            number_type=number_type,
            kind=kind,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_header(
        self,
        text: str,
        *,
        section_index: int = 0,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_header(text, section_index=section_index)

    def append_footer(
        self,
        text: str,
        *,
        section_index: int = 0,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_footer(text, section_index=section_index)

    def append_equation(
        self,
        script: str,
        *,
        width: int = 4800,
        height: int = 2300,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_equation(
            script,
            width=width,
            height=height,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_shape(
        self,
        *,
        kind: str = "rect",
        text: str = "",
        width: int = 12000,
        height: int = 3200,
        fill_color: str = "#FFFFFF",
        line_color: str = "#000000",
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_shape(
            kind=kind,
            text=text,
            width=width,
            height=height,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_ole(
        self,
        name: str,
        data: bytes,
        *,
        width: int = 42001,
        height: int = 13501,
        section_index: int = 0,
        paragraph_index: int | None = None,
    ) -> None:
        self._invalidate_bridge()
        self._binary_document.append_ole(
            name,
            data,
            width=width,
            height=height,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def add_embedded_bindata(self, data: bytes, *, extension: str, storage_id: int | None = None) -> tuple[int, str]:
        self._invalidate_bridge()
        return self._binary_document.add_embedded_bindata(data, extension=extension, storage_id=storage_id)

    def remove_embedded_bindata(self, storage_id: int) -> bool:
        self._invalidate_bridge()
        return self._binary_document.remove_embedded_bindata(storage_id)

    def append_table_pure(self) -> None:
        self.append_table()

    def append_picture_pure(self) -> None:
        self.append_picture()

    def append_hyperlink_pure(self) -> None:
        self.append_hyperlink()

    def bridge(self):
        from .bridge import HwpHwpxBridge

        return HwpHwpxBridge.from_hwp(self, converter=self._converter)

    def __getattr__(self, name: str):
        if name.startswith("_"):
            raise AttributeError(name)
        bridge_document = self.to_hwpx_document()
        return getattr(bridge_document, name)

    def _ensure_bridge_workspace(self) -> Path:
        if self._bridge_temp_dir is None:
            self._bridge_temp_dir = Path(tempfile.mkdtemp(prefix="jakal_hwp_bridge_"))
        return self._bridge_temp_dir

    def _append_feature_pure(self, feature: str) -> None:
        self._invalidate_bridge()
        if self._pure_profile is None:
            self._pure_profile = HwpPureProfile.load_bundled()
        append_feature_from_profile(self._binary_document, self._pure_profile, feature)

    def _mutate_via_hwpx(self, mutate: Callable[[HwpxDocument], object]) -> object:
        bridge_document = self.to_hwpx_document(force_refresh=False)
        result = mutate(bridge_document)
        self._bridge_document = bridge_document
        self._sync_from_bridge_document()
        return result

    def _sync_from_bridge_document(self) -> None:
        if self._bridge_document is None:
            return
        from .hancom_document import HancomDocument

        rebuilt = HancomDocument.from_hwpx_document(self._bridge_document, converter=self._converter).to_hwp_document(
            converter=self._converter
        )
        self._binary_document = rebuilt.binary_document()

    def _invalidate_bridge(self) -> None:
        self._bridge_document = None


@dataclass(frozen=True)
class HwpSection:
    document: HwpDocument
    index: int

    def paragraphs(self) -> list["HwpParagraphObject"]:
        return [HwpParagraphObject(self.document, paragraph) for paragraph in self.document.binary_document().paragraphs(self.index)]

    def paragraph(self, index: int) -> "HwpParagraphObject":
        return self.paragraphs()[index]

    def records(self):
        return self.document.binary_document().section_records(self.index)

    def model(self) -> SectionModel:
        return self.document.binary_document().section_model(self.index)

    def controls(self) -> list["HwpControlObject"]:
        section_model = self.model()
        controls: list[HwpControlObject] = []
        for paragraph in section_model.paragraphs():
            control_ordinal = 0
            for control_node in paragraph.header.children:
                if control_node.tag_id != TAG_CTRL_HEADER:
                    continue
                wrapper = _build_control_wrapper(self.document, section_model, paragraph, control_node, control_ordinal)
                if wrapper is not None:
                    controls.append(wrapper)
                control_ordinal += 1
        return controls

    def tables(self) -> list["HwpTableObject"]:
        return [control for control in self.controls() if isinstance(control, HwpTableObject)]

    def pictures(self) -> list["HwpPictureObject"]:
        return [control for control in self.controls() if isinstance(control, HwpPictureObject)]

    def hyperlinks(self) -> list["HwpHyperlinkObject"]:
        return [control for control in self.controls() if isinstance(control, HwpHyperlinkObject)]

    def bookmarks(self) -> list["HwpBookmarkObject"]:
        return [control for control in self.controls() if isinstance(control, HwpBookmarkObject)]

    def notes(self) -> list["HwpNoteObject"]:
        return [control for control in self.controls() if isinstance(control, HwpNoteObject)]

    def page_numbers(self) -> list["HwpPageNumObject"]:
        return [control for control in self.controls() if isinstance(control, HwpPageNumObject)]

    def fields(self) -> list["HwpFieldObject"]:
        return [control for control in self.controls() if isinstance(control, HwpFieldObject)]

    def shapes(self) -> list["HwpShapeObject"]:
        return [control for control in self.controls() if isinstance(control, HwpShapeObject)]

    def oles(self) -> list["HwpOleObject"]:
        return [control for control in self.controls() if isinstance(control, HwpOleObject)]

    def equations(self) -> list["HwpEquationObject"]:
        return [control for control in self.controls() if isinstance(control, HwpEquationObject)]

    def replace_text_same_length(self, old: str, new: str, *, count: int = -1) -> int:
        return self.document.replace_text_same_length(old, new, section_index=self.index, count=count)

    def append_paragraph(
        self,
        text: str,
        *,
        para_shape_id: int = 0,
        style_id: int = 0,
        split_flags: int = 0,
        control_mask: int = 0,
    ) -> "HwpParagraphObject":
        return self.document.append_paragraph(
            text,
            section_index=self.index,
            para_shape_id=para_shape_id,
            style_id=style_id,
            split_flags=split_flags,
            control_mask=control_mask,
        )

    def append_field(
        self,
        *,
        field_type: str,
        display_text: str | None = None,
        name: str | None = None,
        parameters: dict[str, str | int] | None = None,
        editable: bool = False,
        dirty: bool = False,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_field(
            field_type=field_type,
            display_text=display_text,
            name=name,
            parameters=parameters,
            editable=editable,
            dirty=dirty,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_bookmark(self, name: str, *, paragraph_index: int | None = None) -> None:
        self.document.append_bookmark(name, section_index=self.index, paragraph_index=paragraph_index)

    def append_note(
        self,
        text: str,
        *,
        kind: str = "footNote",
        number: int | str | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_note(
            text,
            kind=kind,
            number=number,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_footnote(
        self,
        text: str,
        *,
        number: int | str | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        self.append_note(text, kind="footNote", number=number, paragraph_index=paragraph_index)

    def append_endnote(
        self,
        text: str,
        *,
        number: int | str | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        self.append_note(text, kind="endNote", number=number, paragraph_index=paragraph_index)

    def append_auto_number(
        self,
        *,
        number: int | str = 1,
        number_type: str = "PAGE",
        kind: str = "newNum",
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_auto_number(
            number=number,
            number_type=number_type,
            kind=kind,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_header(self, text: str) -> None:
        self.document.append_header(text, section_index=self.index)

    def append_footer(self, text: str) -> None:
        self.document.append_footer(text, section_index=self.index)

    def append_equation(
        self,
        script: str,
        *,
        width: int = 4800,
        height: int = 2300,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_equation(
            script,
            width=width,
            height=height,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_shape(
        self,
        *,
        kind: str = "rect",
        text: str = "",
        width: int = 12000,
        height: int = 3200,
        fill_color: str = "#FFFFFF",
        line_color: str = "#000000",
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_shape(
            kind=kind,
            text=text,
            width=width,
            height=height,
            fill_color=fill_color,
            line_color=line_color,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )

    def append_ole(
        self,
        name: str,
        data: bytes,
        *,
        width: int = 42001,
        height: int = 13501,
        paragraph_index: int | None = None,
    ) -> None:
        self.document.append_ole(
            name,
            data,
            width=width,
            height=height,
            section_index=self.index,
            paragraph_index=paragraph_index,
        )


class HwpParagraphObject:
    def __init__(self, document: HwpDocument, paragraph: HwpParagraph):
        self._document = document
        self._section_index = paragraph.section_index
        self._index = paragraph.index

    @property
    def document(self) -> HwpDocument:
        return self._document

    @property
    def index(self) -> int:
        return self._index

    @property
    def section_index(self) -> int:
        return self._section_index

    def _snapshot(self) -> HwpParagraph:
        return self._document.binary_document().paragraphs(self._section_index)[self._index]

    @property
    def text(self) -> str:
        return self._snapshot().text

    @property
    def raw_text(self) -> str:
        return self._snapshot().raw_text

    @property
    def char_count(self) -> int:
        return self._snapshot().char_count

    @property
    def control_mask(self) -> int:
        return self._snapshot().control_mask

    @property
    def para_shape_id(self) -> int:
        return self._snapshot().para_shape_id

    @property
    def style_id(self) -> int:
        return self._snapshot().style_id

    @property
    def split_flags(self) -> int:
        return self._snapshot().split_flags

    @property
    def has_hidden_controls(self) -> bool:
        snapshot = self._snapshot()
        return snapshot.raw_text != snapshot.text

    def set_text_same_length(self, value: str) -> None:
        self._document.binary_document().set_paragraph_text_same_length(self._section_index, self._index, value)

    def replace_text_same_length(self, old: str, new: str, *, count: int = -1) -> int:
        return self._document.binary_document().replace_paragraph_text_same_length(
            self._section_index,
            self._index,
            old,
            new,
            count=count,
        )

    def set_text(self, value: str) -> None:
        model = self._document.section_model(self._section_index)
        paragraph = model.paragraphs()[self._index]
        text_record = paragraph.text_record()
        if text_record is None:
            raise ValueError("HWP paragraph does not contain a para text record.")
        text_record.set_raw_text(value)
        paragraph.header.char_count = len(value)
        self._document.binary_document().replace_section_model(self._section_index, model)
        self._document._invalidate_bridge()

    def replace_text(self, old: str, new: str, *, count: int = -1) -> int:
        model = self._document.section_model(self._section_index)
        paragraph = model.paragraphs()[self._index]
        text_record = paragraph.text_record()
        if text_record is None or not old:
            return 0
        raw_text = text_record.raw_text
        replaced = raw_text.count(old) if count < 0 else min(raw_text.count(old), count)
        if replaced == 0:
            return 0
        updated = raw_text.replace(old, new) if count < 0 else raw_text.replace(old, new, count)
        text_record.set_raw_text(updated)
        paragraph.header.char_count = len(updated)
        self._document.binary_document().replace_section_model(self._section_index, model)
        self._document._invalidate_bridge()
        return replaced


class HwpControlObject:
    def __init__(
        self,
        document: HwpDocument,
        paragraph: SectionParagraphModel,
        control_node: RecordNode,
        control_ordinal: int,
    ):
        self._document = document
        self._paragraph = paragraph
        self._control_ordinal = control_ordinal
        self._initial_control_node = control_node

    @property
    def document(self) -> HwpDocument:
        return self._document

    @property
    def section_index(self) -> int:
        return self._paragraph.section_index

    @property
    def paragraph_index(self) -> int:
        return self._paragraph.index

    @property
    def control_node(self) -> RecordNode:
        return self._live_context()[2]

    @property
    def control_id(self) -> str:
        payload = self.control_node.payload
        if len(payload) < 4:
            return ""
        return payload[:4][::-1].decode("latin1", errors="replace")

    def descendant_nodes(self) -> list[RecordNode]:
        return list(self.control_node.iter_descendants())

    def _live_context(self) -> tuple[SectionModel, SectionParagraphModel, RecordNode]:
        section_model = self._document.section_model(self.section_index)
        paragraph = section_model.paragraphs()[self.paragraph_index]
        controls = [child for child in paragraph.header.children if child.tag_id == TAG_CTRL_HEADER]
        control_node = controls[self._control_ordinal]
        return section_model, paragraph, control_node

    def _commit(self, section_model: SectionModel) -> None:
        self._document.binary_document().replace_section_model(self.section_index, section_model)
        self._document._invalidate_bridge()


@dataclass(frozen=True)
class HwpTableCellObject:
    row: int
    column: int
    col_span: int
    row_span: int
    width: int
    height: int
    border_fill_id: int
    margins: tuple[int, int, int, int]
    text: str


class HwpTableObject(HwpControlObject):
    @property
    def row_count(self) -> int:
        payload = self._table_record().payload
        return int.from_bytes(payload[4:6], "little")

    @property
    def column_count(self) -> int:
        payload = self._table_record().payload
        return int.from_bytes(payload[6:8], "little")

    @property
    def row_heights(self) -> list[int]:
        payload = self._table_record().payload
        row_count = self.row_count
        start = 18
        return [int.from_bytes(payload[start + index * 2 : start + index * 2 + 2], "little") for index in range(row_count)]

    @property
    def table_border_fill_id(self) -> int:
        payload = self._table_record().payload
        offset = 18 + self.row_count * 2
        return int.from_bytes(payload[offset : offset + 2], "little")

    def cells(self) -> list[HwpTableCellObject]:
        result: list[HwpTableCellObject] = []
        children = self.control_node.children
        for index, node in enumerate(children):
            if node.tag_id != TAG_LIST_HEADER or len(node.payload) != 47:
                continue
            payload = node.payload
            text_node = children[index + 1] if index + 1 < len(children) and children[index + 1].tag_id == TAG_PARA_HEADER else None
            result.append(
                HwpTableCellObject(
                    row=int.from_bytes(payload[10:12], "little"),
                    column=int.from_bytes(payload[8:10], "little"),
                    col_span=int.from_bytes(payload[12:14], "little"),
                    row_span=int.from_bytes(payload[14:16], "little"),
                    width=int.from_bytes(payload[16:20], "little"),
                    height=int.from_bytes(payload[20:24], "little"),
                    margins=(
                        int.from_bytes(payload[24:26], "little"),
                        int.from_bytes(payload[26:28], "little"),
                        int.from_bytes(payload[28:30], "little"),
                        int.from_bytes(payload[30:32], "little"),
                    ),
                    border_fill_id=int.from_bytes(payload[32:34], "little"),
                    text=_collect_text_from_node(text_node) if text_node is not None else "",
                )
            )
        result.sort(key=lambda cell: (cell.row, cell.column))
        return result

    def cell(self, row: int, column: int) -> HwpTableCellObject:
        for cell in self.cells():
            if cell.row == row and cell.column == column:
                return cell
        raise IndexError(f"No HWP table cell at row={row}, column={column}.")

    def cell_text_matrix(self) -> list[list[str]]:
        matrix = [["" for _ in range(self.column_count)] for _ in range(self.row_count)]
        for cell in self.cells():
            if cell.row < self.row_count and cell.column < self.column_count:
                matrix[cell.row][cell.column] = cell.text
        return matrix

    def set_cell_text(self, row: int, column: int, text: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        children = control_node.children
        for index, node in enumerate(children):
            if node.tag_id != TAG_LIST_HEADER or len(node.payload) != 47:
                continue
            payload = node.payload
            current_column = int.from_bytes(payload[8:10], "little")
            current_row = int.from_bytes(payload[10:12], "little")
            if current_row != row or current_column != column:
                continue
            para_node = children[index + 1] if index + 1 < len(children) and children[index + 1].tag_id == TAG_PARA_HEADER else None
            if not isinstance(para_node, ParagraphHeaderRecord):
                raise ValueError("HWP table cell does not contain a writable paragraph header.")
            text_record = next((child for child in para_node.children if isinstance(child, ParagraphTextRecord)), None)
            if text_record is None:
                raise ValueError("HWP table cell does not contain a para text record.")
            text_record.set_raw_text(f"{text}\r")
            para_node.char_count = len(text) + 1
            self._commit(section_model)
            return
        raise IndexError(f"No HWP table cell at row={row}, column={column}.")

    def set_row_heights(self, heights: Sequence[int]) -> None:
        section_model, _paragraph, control_node = self._live_context()
        values = [int(value) for value in heights]
        if len(values) != self.row_count:
            raise ValueError(f"row heights length must match row_count={self.row_count}.")
        record = self._table_record(control_node)
        payload = bytearray(record.payload)
        start = 18
        for index, value in enumerate(values):
            payload[start + index * 2 : start + index * 2 + 2] = value.to_bytes(2, "little")
        record.payload = bytes(payload)
        self._commit(section_model)

    def set_table_border_fill_id(self, border_fill_id: int) -> None:
        section_model, _paragraph, control_node = self._live_context()
        record = self._table_record(control_node)
        payload = bytearray(record.payload)
        offset = 18 + self.row_count * 2
        payload[offset : offset + 2] = int(border_fill_id).to_bytes(2, "little")
        record.payload = bytes(payload)
        self._commit(section_model)

    def append_row(self) -> list[HwpTableCellObject]:
        section_model, _paragraph, control_node = self._live_context()
        row_heights = list(self.row_heights)
        new_row_index = self.row_count
        row_heights.append(row_heights[-1] if row_heights else DEFAULT_HWP_TABLE_CELL_HEIGHT)
        column_widths = self._column_widths()
        current_specs = self._cell_specs()
        for column_index, column_width in enumerate(column_widths):
            current_specs.append(
                _TableCellSpec(
                    row=new_row_index,
                    col=column_index,
                    row_span=1,
                    col_span=1,
                    width=column_width,
                    height=row_heights[-1],
                    text="",
                    border_fill_id=self.table_border_fill_id,
                    margins=self._default_cell_margins(),
                )
            )
        self._replace_table_cells(
            section_model,
            control_node,
            row_count=self.row_count + 1,
            column_count=self.column_count,
            row_heights=row_heights,
            cell_specs=current_specs,
            table_border_fill_id=self.table_border_fill_id,
        )
        refreshed = self.document.section(self.section_index).tables()[self._table_index_in_section()]
        new_row_index = refreshed.row_count - 1
        return [cell for cell in refreshed.cells() if cell.row == new_row_index]

    def merge_cells(self, start_row: int, start_column: int, end_row: int, end_column: int) -> HwpTableCellObject:
        if not (0 <= start_row <= end_row < self.row_count):
            raise IndexError("row merge range is out of bounds.")
        if not (0 <= start_column <= end_column < self.column_count):
            raise IndexError("column merge range is out of bounds.")

        section_model, _paragraph, control_node = self._live_context()
        row_heights = list(self.row_heights)
        column_widths = self._column_widths()
        merged_specs: list[_TableCellSpec] = []
        anchor: _TableCellSpec | None = None

        for spec in self._cell_specs():
            within_range = (
                start_row <= spec.row <= end_row
                and start_column <= spec.col <= end_column
            )
            if not within_range:
                merged_specs.append(spec)
                continue
            if spec.row == start_row and spec.col == start_column:
                anchor = _TableCellSpec(
                    row=start_row,
                    col=start_column,
                    row_span=end_row - start_row + 1,
                    col_span=end_column - start_column + 1,
                    width=sum(column_widths[start_column : end_column + 1]),
                    height=sum(row_heights[start_row : end_row + 1]),
                    text=spec.text,
                    border_fill_id=spec.border_fill_id,
                    margins=spec.margins,
                )
            elif spec.row + spec.row_span - 1 > end_row or spec.col + spec.col_span - 1 > end_column:
                raise ValueError("merge range intersects an existing spanned cell.")

        if anchor is None:
            raise IndexError("No anchor cell exists at the merge start coordinate.")

        merged_specs.append(anchor)
        merged_specs.sort(key=lambda item: (item.row, item.col))
        self._replace_table_cells(
            section_model,
            control_node,
            row_count=self.row_count,
            column_count=self.column_count,
            row_heights=row_heights,
            cell_specs=merged_specs,
            table_border_fill_id=self.table_border_fill_id,
        )
        refreshed = self.document.section(self.section_index).tables()[self._table_index_in_section()]
        return refreshed.cell(start_row, start_column)

    def _table_index_in_section(self) -> int:
        section = self.document.section(self.section_index)
        tables = section.tables()
        matches = [
            table
            for table in tables
            if table.paragraph_index == self.paragraph_index and table._control_ordinal == self._control_ordinal
        ]
        if matches:
            return tables.index(matches[0])
        for index, table in enumerate(tables):
            if table.paragraph_index == self.paragraph_index:
                return index
        raise ValueError("Could not locate the current HWP table within its section.")

    def _table_record(self, control_node: RecordNode | None = None) -> RecordNode:
        target = self.control_node if control_node is None else control_node
        for node in target.iter_descendants():
            if node.tag_id == TAG_TABLE:
                return node
        raise ValueError("HWP table control does not contain a table record.")

    def _cell_specs(self) -> list[_TableCellSpec]:
        return [
            _TableCellSpec(
                row=cell.row,
                col=cell.column,
                row_span=cell.row_span,
                col_span=cell.col_span,
                width=cell.width,
                height=cell.height,
                text=cell.text,
                border_fill_id=cell.border_fill_id,
                margins=cell.margins,
            )
            for cell in self.cells()
        ]

    def _column_widths(self) -> list[int]:
        widths = [0] * self.column_count
        for cell in self.cells():
            if cell.col_span == 1 and widths[cell.column] == 0:
                widths[cell.column] = cell.width
        fallback = max((value for value in widths if value > 0), default=DEFAULT_HWP_TABLE_CELL_WIDTH)
        for cell in self.cells():
            if cell.col_span <= 1:
                continue
            missing = [index for index in range(cell.column, cell.column + cell.col_span) if widths[index] == 0]
            if not missing:
                continue
            distributed = max(cell.width // cell.col_span, 1)
            for index in missing:
                widths[index] = distributed
        return [value or fallback for value in widths]

    def _default_cell_margins(self) -> tuple[int, int, int, int]:
        existing_cells = self.cells()
        if existing_cells:
            return existing_cells[0].margins
        return (510, 510, 141, 141)

    def _replace_table_cells(
        self,
        section_model: SectionModel,
        control_node: RecordNode,
        *,
        row_count: int,
        column_count: int,
        row_heights: Sequence[int],
        cell_specs: Sequence[_TableCellSpec],
        table_border_fill_id: int,
    ) -> None:
        table_record = self._table_record(control_node)
        table_record.payload = _build_table_record_payload(
            row_count,
            column_count,
            default_row_height=1,
            border_fill_id=table_border_fill_id,
        )
        payload = bytearray(table_record.payload)
        row_sizes_offset = 18
        for row_index, row_height in enumerate(row_heights):
            offset = row_sizes_offset + row_index * 2
            payload[offset : offset + 2] = int(row_height).to_bytes(2, "little", signed=False)
        table_record.payload = bytes(payload)

        new_children: list[RecordNode] = [table_record]
        for cell in sorted(cell_specs, key=lambda item: (item.row, item.col)):
            list_header = RecordNode(tag_id=TAG_LIST_HEADER, level=2, payload=_build_table_cell_list_header_payload(cell))
            raw_text = f"{cell.text}\r"
            char_count = len(raw_text)
            paragraph_header = ParagraphHeaderRecord(
                level=2,
                char_count=0x80000000 | char_count,
                control_mask=int.from_bytes(_TABLE_CELL_PARA_HEADER_TAIL[0:4], "little"),
                para_shape_id=int.from_bytes(_TABLE_CELL_PARA_HEADER_TAIL[4:6], "little"),
                style_id=_TABLE_CELL_PARA_HEADER_TAIL[6],
                split_flags=_TABLE_CELL_PARA_HEADER_TAIL[7],
                trailing_payload=_TABLE_CELL_PARA_HEADER_TAIL[8:],
            )
            paragraph_header.add_child(
                ParagraphTextRecord(
                    level=3,
                    raw_text=raw_text,
                )
            )
            paragraph_header.add_child(
                RecordNode(
                    tag_id=TAG_PARA_CHAR_SHAPE,
                    level=3,
                    payload=_TABLE_CELL_CHAR_SHAPE,
                )
            )
            paragraph_header.add_child(
                RecordNode(
                    tag_id=TAG_PARA_LINE_SEG,
                    level=3,
                    payload=_TABLE_CELL_LINE_SEG,
                )
            )
            new_children.extend((list_header, paragraph_header))

        control_node.children = []
        for child in new_children:
            control_node.add_child(child)
        self._commit(section_model)


class HwpPictureObject(HwpControlObject):
    @property
    def storage_id(self) -> int:
        payload = self._shape_picture_record().payload
        return int.from_bytes(payload[71:73], "little")

    def bindata_path(self) -> str:
        prefix = f"BinData/BIN{self.storage_id:04d}."
        for stream_path in self.document.bindata_stream_paths():
            if stream_path.startswith(prefix):
                return stream_path
        raise KeyError(f"Could not find BinData stream for storage id {self.storage_id}.")

    @property
    def extension(self) -> str | None:
        suffix = Path(self.bindata_path()).suffix
        return suffix.lstrip(".") or None

    def size(self) -> dict[str, int]:
        shape_component = self._shape_component_record()
        if shape_component is None or len(shape_component.payload) < 28:
            return {"width": 0, "height": 0}
        return {
            "width": int.from_bytes(shape_component.payload[20:24], "little", signed=False),
            "height": int.from_bytes(shape_component.payload[24:28], "little", signed=False),
        }

    def binary_data(self) -> bytes:
        return self.document.binary_document().read_stream(self.bindata_path(), decompress=False)

    def replace_binary(self, data: bytes, *, extension: str | None = None) -> None:
        if extension is None or extension.lower().lstrip(".") == (self.extension or "").lower():
            self.document.binary_document().add_stream(self.bindata_path(), data)
        else:
            section_model, _paragraph, control_node = self._live_context()
            storage_id, _ = self.document.add_embedded_bindata(data, extension=extension)
            payload = bytearray(self._shape_picture_record(control_node).payload)
            payload[71:73] = int(storage_id).to_bytes(2, "little")
            self._shape_picture_record(control_node).payload = bytes(payload)
            self._commit(section_model)
            return
        self.document._invalidate_bridge()

    def _shape_picture_record(self, control_node: RecordNode | None = None) -> RecordNode:
        target = self.control_node if control_node is None else control_node
        for node in target.iter_descendants():
            if node.tag_id == TAG_SHAPE_COMPONENT_PICTURE:
                return node
        raise ValueError("HWP picture control does not contain a shape picture record.")

    def _shape_component_record(self, control_node: RecordNode | None = None) -> RecordNode | None:
        target = self.control_node if control_node is None else control_node
        for node in target.iter_descendants():
            if node.tag_id == TAG_SHAPE_COMPONENT:
                return node
        return None


class HwpBookmarkObject(HwpControlObject):
    @property
    def name(self) -> str:
        data_node = next((child for child in self.control_node.children if child.tag_id == TAG_CTRL_DATA), None)
        if data_node is None or len(data_node.payload) < 12:
            return ""
        name_size = int.from_bytes(data_node.payload[10:12], "little", signed=False)
        encoded_name = data_node.payload[12 : 12 + name_size * 2]
        return encoded_name.decode("utf-16-le", errors="ignore")

    def rename(self, value: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        data_node = next((child for child in control_node.children if child.tag_id == TAG_CTRL_DATA), None)
        if data_node is None:
            raise ValueError("HWP bookmark control does not contain ctrl data.")
        prefix = data_node.payload[:10]
        payload = bytearray(prefix)
        payload.extend(len(value).to_bytes(2, "little", signed=False))
        payload.extend(value.encode("utf-16-le"))
        data_node.payload = bytes(payload)
        self._commit(section_model)


class HwpFieldObject(HwpControlObject):
    @property
    def field_type(self) -> str:
        return self.control_id


class HwpHyperlinkObject(HwpFieldObject):
    @property
    def command(self) -> str:
        payload = self.control_node.payload
        prefix_size = 9
        text_size = int.from_bytes(payload[prefix_size : prefix_size + 2], "little")
        start = prefix_size + 2
        end = start + text_size * 2
        return payload[start:end].decode("utf-16-le", errors="ignore")

    @property
    def url(self) -> str:
        command = self.command
        return command.split(";", 1)[0].replace("\\:", ":")

    @property
    def metadata_fields(self) -> list[str]:
        parts = self.command.split(";")
        if parts and parts[-1] == "":
            parts = parts[:-1]
        return parts[1:]

    @property
    def display_text(self) -> str:
        return self._paragraph.text.replace("\r", "")

    def set_url(self, url: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        control_node.payload = _build_hyperlink_ctrl_payload(url, metadata_fields=self.metadata_fields)
        self._commit(section_model)

    def set_metadata_fields(self, metadata_fields: Sequence[str | int]) -> None:
        section_model, _paragraph, control_node = self._live_context()
        control_node.payload = _build_hyperlink_ctrl_payload(self.url, metadata_fields=metadata_fields)
        self._commit(section_model)

    def set_display_text(self, text: str) -> None:
        section_model, paragraph, _control_node = self._live_context()
        text_record = paragraph.text_record()
        if text_record is None:
            raise ValueError("HWP hyperlink paragraph does not contain a para text record.")
        raw_text = _build_hyperlink_para_text_payload(text).decode("utf-16-le", errors="ignore")
        text_record.set_raw_text(raw_text)
        paragraph.header.char_count = len(raw_text)
        self._commit(section_model)


class HwpAutoNumberObject(HwpControlObject):
    @property
    def kind(self) -> str:
        return "newNum" if self.control_id == "nwno" else "autoNum"

    @property
    def number(self) -> str | None:
        payload = self.control_node.payload
        if self.control_id == "nwno" and len(payload) >= 10:
            return str(int.from_bytes(payload[8:10], "little"))
        if self.control_id == "atno":
            return "1"
        return None

    @property
    def number_type(self) -> str:
        return "PAGE"

    def set_number(self, value: int | str) -> None:
        if self.control_id != "nwno":
            raise ValueError("Only newNum controls support explicit numeric rewrites in the native HWP wrapper.")
        section_model, _paragraph, control_node = self._live_context()
        payload = bytearray(control_node.payload)
        payload[8:10] = int(value).to_bytes(2, "little", signed=False)
        control_node.payload = bytes(payload)
        self._commit(section_model)


class HwpHeaderFooterObject(HwpControlObject):
    @property
    def kind(self) -> str:
        return "header" if self.control_id == "head" else "footer"

    @property
    def text(self) -> str:
        return _collect_text_from_node(self.control_node).replace("\r", "").strip()

    def set_text(self, value: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        paragraph_headers = [node for node in control_node.children if isinstance(node, ParagraphHeaderRecord)]
        if not paragraph_headers:
            raise ValueError("HWP header/footer control does not contain a writable paragraph header.")
        target = paragraph_headers[0]
        text_record = next((child for child in target.children if isinstance(child, ParagraphTextRecord)), None)
        raw_text = f"{value}\r" if value else ""
        if text_record is None:
            text_record = ParagraphTextRecord(level=target.level + 1, raw_text=raw_text)
            target.children.insert(0, text_record)
        else:
            text_record.set_raw_text(raw_text)
        target.char_count = len(raw_text)
        target.sync_payload()
        self._commit(section_model)


class HwpNoteObject(HwpControlObject):
    @property
    def kind(self) -> str:
        return "footNote" if self.control_id == "fn  " else "endNote"

    @property
    def number(self) -> str | None:
        return None

    @property
    def text(self) -> str:
        return _collect_text_from_node(self.control_node).replace("\r", "").strip()

    def set_text(self, value: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        paragraph_headers = [node for node in control_node.children if isinstance(node, ParagraphHeaderRecord)]
        if not paragraph_headers:
            raise ValueError("HWP note control does not contain a writable paragraph header.")
        target = paragraph_headers[0]
        text_record = next((child for child in target.children if isinstance(child, ParagraphTextRecord)), None)
        raw_text = f"{value}\r" if value else ""
        if text_record is None:
            text_record = ParagraphTextRecord(level=target.level + 1, raw_text=raw_text)
            target.children.insert(0, text_record)
        else:
            text_record.set_raw_text(raw_text)
        target.char_count = 0x80000000 | len(raw_text)
        target.sync_payload()
        self._commit(section_model)


class HwpPageNumObject(HwpControlObject):
    @property
    def pos(self) -> str:
        payload = self.control_node.payload
        attributes = int.from_bytes(payload[4:8].ljust(4, b"\x00"), "little")
        positions = {
            0: "NONE",
            1: "TOP_LEFT",
            2: "TOP_CENTER",
            3: "TOP_RIGHT",
            4: "BOTTOM_LEFT",
            5: "BOTTOM_CENTER",
            6: "BOTTOM_RIGHT",
            7: "OUTSIDE_TOP",
            8: "OUTSIDE_BOTTOM",
            9: "INSIDE_TOP",
            10: "INSIDE_BOTTOM",
        }
        return positions.get((attributes >> 8) & 0x0F, "BOTTOM_CENTER")

    @property
    def format_type(self) -> str:
        payload = self.control_node.payload
        attributes = int.from_bytes(payload[4:8].ljust(4, b"\x00"), "little")
        formats = {
            0: "DIGIT",
            1: "CIRCLED_DIGIT",
            2: "ROMAN_CAPITAL",
            3: "ROMAN_SMALL",
            4: "LATIN_CAPITAL",
            5: "LATIN_SMALL",
        }
        return formats.get(attributes & 0xFF, "DIGIT")

    @property
    def side_char(self) -> str:
        payload = self.control_node.payload
        if len(payload) < 16:
            return "-"
        return payload[14:16].decode("utf-16-le", errors="ignore") or "-"


class HwpEquationObject(HwpControlObject):
    @property
    def script(self) -> str:
        eqedit_record = self._eqedit_record()
        if eqedit_record is None:
            return ""
        decoded = eqedit_record.payload[4:].decode("utf-16-le", errors="ignore")
        script = decoded.split("Equation Version", 1)[0].replace("\x00", "").strip()
        if "ь" in script:
            script = script.split("ь", 1)[0]
        script = "".join(character for character in script if ord(character) >= 32 or character in "\t\n")
        if script.startswith("* "):
            script = script[2:]
        return script.rstrip("`").strip()

    def _eqedit_record(self, control_node: RecordNode | None = None) -> RecordNode | None:
        target = self.control_node if control_node is None else control_node
        for node in target.iter_descendants():
            if node.tag_id == TAG_EQEDIT:
                return node
        return None

    def set_script(self, script: str) -> None:
        section_model, _paragraph, control_node = self._live_context()
        eqedit_record = self._eqedit_record(control_node)
        if eqedit_record is None:
            raise ValueError("HWP equation control does not contain an eqedit record.")
        decoded = eqedit_record.payload[4:].decode("utf-16-le", errors="ignore")
        prefix = "* " if decoded.startswith("* ") else ""
        marker_index = decoded.find("Equation Version")
        suffix = decoded[marker_index:] if marker_index >= 0 else ""
        eqedit_record.payload = eqedit_record.payload[:4] + (prefix + script + suffix).encode("utf-16-le")
        self._commit(section_model)


class HwpShapeObject(HwpControlObject):
    @property
    def kind(self) -> str:
        return _shape_kind_from_control_node(self.control_node)

    @property
    def text(self) -> str:
        return self._paragraph.text.replace("\r", "")

    def descendant_tag_ids(self) -> list[int]:
        return [node.tag_id for node in self.control_node.iter_descendants()]

    def size(self) -> dict[str, int]:
        record = self._shape_component_record()
        if record is None or len(record.payload) < 28:
            return {"width": 0, "height": 0}
        return {
            "width": int.from_bytes(record.payload[20:24], "little", signed=False),
            "height": int.from_bytes(record.payload[24:28], "little", signed=False),
        }

    def _shape_component_record(self, control_node: RecordNode | None = None) -> RecordNode | None:
        target = self.control_node if control_node is None else control_node
        for node in target.iter_descendants():
            if node.tag_id == TAG_SHAPE_COMPONENT:
                return node
        return None

    def set_size(self, *, width: int | None = None, height: int | None = None) -> None:
        section_model, _paragraph, control_node = self._live_context()
        record = self._shape_component_record(control_node)
        if record is None:
            raise ValueError("HWP shape control does not contain a shape component record.")
        payload = bytearray(record.payload)
        if width is not None:
            payload[20:24] = int(width).to_bytes(4, "little")
        if height is not None:
            payload[24:28] = int(height).to_bytes(4, "little")
        record.payload = bytes(payload)
        self._commit(section_model)


class HwpOleObject(HwpShapeObject):
    @property
    def kind(self) -> str:
        return "ole"

    def bindata_path(self) -> str:
        bindata_path = next(
            (stream_path for stream_path in self.document.bindata_stream_paths() if stream_path.endswith(".ole")),
            None,
        )
        if bindata_path is None:
            raise ValueError("No OLE BinData stream is available.")
        return bindata_path

    def binary_data(self) -> bytes:
        return self.document.binary_document().read_stream(self.bindata_path(), decompress=False)

    def replace_binary(self, data: bytes, *, extension: str | None = None) -> None:
        self.document.binary_document().add_stream(self.bindata_path(), data)
        self.document._invalidate_bridge()


def _build_control_wrapper(
    document: HwpDocument,
    section_model: SectionModel,
    paragraph: SectionParagraphModel,
    control_node: RecordNode,
    control_ordinal: int,
) -> HwpControlObject | None:
    control_id = _control_id(control_node)
    if control_id == "tbl ":
        return HwpTableObject(document, paragraph, control_node, control_ordinal)
    if control_id == "bokm":
        return HwpBookmarkObject(document, paragraph, control_node, control_ordinal)
    if control_id in {"fn  ", "en  "}:
        return HwpNoteObject(document, paragraph, control_node, control_ordinal)
    if control_id == "%hlk":
        return HwpHyperlinkObject(document, paragraph, control_node, control_ordinal)
    if control_id.startswith("%"):
        return HwpFieldObject(document, paragraph, control_node, control_ordinal)
    if control_id in {"atno", "nwno"}:
        return HwpAutoNumberObject(document, paragraph, control_node, control_ordinal)
    if control_id in {"head", "foot"}:
        return HwpHeaderFooterObject(document, paragraph, control_node, control_ordinal)
    if control_id == "pgnp":
        return HwpPageNumObject(document, paragraph, control_node, control_ordinal)
    if control_id == "eqed":
        return HwpEquationObject(document, paragraph, control_node, control_ordinal)
    if control_id == "gso " and _control_has_descendant(control_node, TAG_SHAPE_COMPONENT_OLE):
        return HwpOleObject(document, paragraph, control_node, control_ordinal)
    if control_id == "gso " and any(node.tag_id == TAG_SHAPE_COMPONENT_PICTURE for node in control_node.iter_descendants()):
        return HwpPictureObject(document, paragraph, control_node, control_ordinal)
    if control_id == "gso ":
        return HwpShapeObject(document, paragraph, control_node, control_ordinal)
    return None


def _control_id(control_node: RecordNode) -> str:
    payload = control_node.payload
    if len(payload) < 4:
        return ""
    return payload[:4][::-1].decode("latin1", errors="replace")


def _collect_text_from_node(node: RecordNode | None) -> str:
    if node is None:
        return ""
    texts: list[str] = []
    for descendant in node.iter_descendants():
        if isinstance(descendant, ParagraphTextRecord):
            texts.append(descendant.text.replace("\r", ""))
    return "".join(texts)


def _control_has_descendant(control_node: RecordNode, tag_id: int) -> bool:
    return any(node.tag_id == tag_id for node in control_node.iter_descendants())


def _shape_kind_from_control_node(control_node: RecordNode) -> str:
    tag_to_kind = {
        TAG_SHAPE_COMPONENT_LINE: "line",
        TAG_SHAPE_COMPONENT_RECTANGLE: "rect",
        TAG_SHAPE_COMPONENT_ELLIPSE: "ellipse",
        TAG_SHAPE_COMPONENT_ARC: "arc",
        TAG_SHAPE_COMPONENT_POLYGON: "polygon",
        TAG_SHAPE_COMPONENT_CURVE: "curve",
        TAG_SHAPE_COMPONENT_OLE: "ole",
        TAG_SHAPE_COMPONENT_PICTURE: "picture",
        TAG_SHAPE_COMPONENT_CONTAINER: "container",
        TAG_SHAPE_COMPONENT_TEXTART: "textart",
    }
    for node in control_node.iter_descendants():
        kind = tag_to_kind.get(node.tag_id)
        if kind is not None:
            return kind
    return "shape"
