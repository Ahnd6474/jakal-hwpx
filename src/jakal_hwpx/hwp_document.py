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
    ParagraphTextRecord,
    RecordNode,
    SectionModel,
    SectionParagraphModel,
    TAG_CTRL_HEADER,
    TAG_EQEDIT,
    TAG_LIST_HEADER,
    TAG_PARA_HEADER,
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

    def controls(self) -> list["HwpControlObject"]:
        controls: list[HwpControlObject] = []
        for paragraph in self.model().paragraphs():
            for control_node in paragraph.header.children:
                if control_node.tag_id != TAG_CTRL_HEADER:
                    continue
                wrapper = _build_control_wrapper(self.document, paragraph, control_node)
                if wrapper is not None:
                    controls.append(wrapper)
        return controls

    def tables(self) -> list["HwpTableObject"]:
        return [control for control in self.controls() if isinstance(control, HwpTableObject)]

    def pictures(self) -> list["HwpPictureObject"]:
        return [control for control in self.controls() if isinstance(control, HwpPictureObject)]

    def hyperlinks(self) -> list["HwpHyperlinkObject"]:
        return [control for control in self.controls() if isinstance(control, HwpHyperlinkObject)]

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


class HwpControlObject:
    def __init__(self, document: HwpDocument, paragraph: SectionParagraphModel, control_node: RecordNode):
        self._document = document
        self._paragraph = paragraph
        self._control_node = control_node

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
        return self._control_node

    @property
    def control_id(self) -> str:
        payload = self._control_node.payload
        if len(payload) < 4:
            return ""
        return payload[:4][::-1].decode("latin1", errors="replace")

    def descendant_nodes(self) -> list[RecordNode]:
        return list(self._control_node.iter_descendants())


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
        children = self._control_node.children
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

    def _table_record(self) -> RecordNode:
        for node in self._control_node.iter_descendants():
            if node.tag_id == TAG_TABLE:
                return node
        raise ValueError("HWP table control does not contain a table record.")


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

    def binary_data(self) -> bytes:
        return self.document.binary_document().read_stream(self.bindata_path(), decompress=False)

    def _shape_picture_record(self) -> RecordNode:
        for node in self._control_node.iter_descendants():
            if node.tag_id == TAG_SHAPE_COMPONENT_PICTURE:
                return node
        raise ValueError("HWP picture control does not contain a shape picture record.")


class HwpFieldObject(HwpControlObject):
    @property
    def field_type(self) -> str:
        return self.control_id


class HwpHyperlinkObject(HwpFieldObject):
    @property
    def command(self) -> str:
        payload = self._control_node.payload
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

class HwpEquationObject(HwpControlObject):
    @property
    def script(self) -> str:
        eqedit_record = self._eqedit_record()
        if eqedit_record is None:
            return ""
        decoded = eqedit_record.payload[4:].decode("utf-16-le", errors="ignore")
        script = decoded.split("Equation Version", 1)[0].replace("\x00", "").strip()
        if script.startswith("* "):
            script = script[2:]
        return script.strip()

    def _eqedit_record(self) -> RecordNode | None:
        for node in self._control_node.iter_descendants():
            if node.tag_id == TAG_EQEDIT:
                return node
        return None


class HwpShapeObject(HwpControlObject):
    @property
    def kind(self) -> str:
        return _shape_kind_from_control_node(self._control_node)

    @property
    def text(self) -> str:
        return self._paragraph.text.replace("\r", "")

    def descendant_tag_ids(self) -> list[int]:
        return [node.tag_id for node in self._control_node.iter_descendants()]

    def size(self) -> dict[str, int]:
        record = self._shape_component_record()
        if record is None or len(record.payload) < 28:
            return {"width": 0, "height": 0}
        return {
            "width": int.from_bytes(record.payload[20:24], "little", signed=False),
            "height": int.from_bytes(record.payload[24:28], "little", signed=False),
        }

    def _shape_component_record(self) -> RecordNode | None:
        for node in self._control_node.iter_descendants():
            if node.tag_id == TAG_SHAPE_COMPONENT:
                return node
        return None


class HwpOleObject(HwpShapeObject):
    @property
    def kind(self) -> str:
        return "ole"


def _build_control_wrapper(
    document: HwpDocument,
    paragraph: SectionParagraphModel,
    control_node: RecordNode,
) -> HwpControlObject | None:
    control_id = _control_id(control_node)
    if control_id == "tbl ":
        return HwpTableObject(document, paragraph, control_node)
    if control_id == "%hlk":
        return HwpHyperlinkObject(document, paragraph, control_node)
    if control_id.startswith("%"):
        return HwpFieldObject(document, paragraph, control_node)
    if control_id == "eqed":
        return HwpEquationObject(document, paragraph, control_node)
    if control_id == "gso " and _control_has_descendant(control_node, TAG_SHAPE_COMPONENT_OLE):
        return HwpOleObject(document, paragraph, control_node)
    if control_id == "gso " and any(node.tag_id == TAG_SHAPE_COMPONENT_PICTURE for node in control_node.iter_descendants()):
        return HwpPictureObject(document, paragraph, control_node)
    if control_id == "gso ":
        return HwpShapeObject(document, paragraph, control_node)
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
