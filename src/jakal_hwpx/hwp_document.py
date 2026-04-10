from __future__ import annotations

import tempfile
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from ._hancom import convert_document
from .document import HwpxDocument
from .hwp_binary import (
    DocInfoModel,
    HwpBinaryDocument,
    HwpBinaryFileHeader,
    HwpDocumentProperties,
    HwpParagraph,
    HwpStreamCapacity,
    SectionModel,
)
from .hwp_pure_profile import HwpPureProfile, append_feature_from_profile

HancomConverter = Callable[[str | Path, str | Path, str], Path]


class HwpDocument:
    def __init__(self, binary_document: HwpBinaryDocument, *, converter: HancomConverter | None = None):
        self._binary_document = binary_document
        self._converter = converter or convert_document
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
        self._binary_document.set_preview_text(value)

    def sections(self) -> list["HwpSection"]:
        return [HwpSection(self, index) for index, _path in enumerate(self._binary_document.section_stream_paths())]

    def section(self, index: int) -> "HwpSection":
        return HwpSection(self, index)

    def section_model(self, section_index: int = 0) -> SectionModel:
        return self._binary_document.section_model(section_index)

    def paragraphs(self, section_index: int = 0) -> list["HwpParagraphObject"]:
        return self.section(section_index).paragraphs()

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

        workspace = self._ensure_bridge_workspace()
        source_hwp = workspace / "bridge_source.hwp"
        output_hwpx = workspace / "bridge_output.hwpx"
        self._binary_document.save_copy(source_hwp)
        self._converter(source_hwp, output_hwpx, "HWPX")
        self._bridge_document = HwpxDocument.open(output_hwpx)
        return self._bridge_document

    def save_as_hwpx(self, path: str | Path) -> Path:
        target_path = Path(path).expanduser().resolve()
        target_path.parent.mkdir(parents=True, exist_ok=True)
        if self._bridge_document is not None:
            self._bridge_document.save(target_path)
            return target_path

        workspace = self._ensure_bridge_workspace()
        source_hwp = workspace / "export_source.hwp"
        self._binary_document.save_copy(source_hwp)
        self._converter(source_hwp, target_path, "HWPX")
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

        workspace = self._ensure_bridge_workspace()
        temp_hwpx = workspace / "bridge_save.hwpx"
        self._bridge_document.save(temp_hwpx)
        self._converter(temp_hwpx, target_path, "HWP")
        self._binary_document = HwpBinaryDocument.open(target_path)
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
        )

    def append_picture(self, image_bytes: bytes | None = None, *, extension: str | None = None) -> None:
        self._invalidate_bridge()
        if self._pure_profile is None:
            self._pure_profile = HwpPureProfile.load_bundled()
        self._binary_document.append_picture(
            image_bytes,
            extension=extension,
            profile_root=self._pure_profile.root,
        )

    def append_hyperlink(
        self,
        url: str | None = None,
        *,
        text: str | None = None,
        metadata_fields: Sequence[str | int] | None = None,
    ) -> None:
        self._invalidate_bridge()
        if self._pure_profile is None:
            self._pure_profile = HwpPureProfile.load_bundled()
        self._binary_document.append_hyperlink(
            url,
            text=text,
            metadata_fields=metadata_fields,
            profile_root=self._pure_profile.root,
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
