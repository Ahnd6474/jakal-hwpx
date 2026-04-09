from __future__ import annotations

import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from ._hancom import convert_document
from .document import HwpxDocument
from .hwp_binary import HwpBinaryDocument, HwpBinaryFileHeader, HwpDocumentProperties, HwpParagraph
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

    def list_stream_paths(self) -> list[str]:
        return self._binary_document.list_stream_paths()

    def preview_text(self) -> str:
        return self._binary_document.preview_text()

    def set_preview_text(self, value: str) -> None:
        self._binary_document.set_preview_text(value)

    def sections(self) -> list["HwpSection"]:
        return [HwpSection(self, index) for index, _path in enumerate(self._binary_document.section_stream_paths())]

    def section(self, index: int) -> "HwpSection":
        return HwpSection(self, index)

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

    def append_table_pure(self) -> None:
        self._append_feature_pure("table")

    def append_picture_pure(self) -> None:
        self._append_feature_pure("picture")

    def append_hyperlink_pure(self) -> None:
        self._append_feature_pure("hyperlink")

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
        if self._pure_profile is None:
            raise RuntimeError("No pure profile is loaded. Use HwpDocument.blank_from_profile(...) or load_pure_profile(...).")
        append_feature_from_profile(self._binary_document, self._pure_profile, feature)


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

    def replace_text_same_length(self, old: str, new: str, *, count: int = -1) -> int:
        return self.document.replace_text_same_length(old, new, section_index=self.index, count=count)


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
