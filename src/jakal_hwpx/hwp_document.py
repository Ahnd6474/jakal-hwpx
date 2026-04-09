from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .hwp_binary import HwpBinaryDocument, HwpBinaryFileHeader, HwpDocumentProperties, HwpParagraph


class HwpDocument:
    def __init__(self, binary_document: HwpBinaryDocument):
        self._binary_document = binary_document

    @classmethod
    def open(cls, path: str | Path) -> "HwpDocument":
        return cls(HwpBinaryDocument.open(path))

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

    def save(self, path: str | Path | None = None) -> Path:
        return self._binary_document.save(path)


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
