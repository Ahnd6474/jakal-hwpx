from __future__ import annotations

import struct
from collections.abc import Iterator
from dataclasses import dataclass, field

# Imported after hwp_binary defines the raw and typed record classes.
from .hwp_binary import (
    TAG_BIN_DATA,
    TAG_BORDER_FILL,
    TAG_BULLET,
    TAG_CHAR_SHAPE,
    TAG_CHART_DATA,
    TAG_COMPATIBLE_DOCUMENT,
    TAG_CTRL_DATA,
    TAG_CTRL_HEADER,
    TAG_DOC_DATA,
    TAG_DOCUMENT_PROPERTIES,
    TAG_EQEDIT,
    TAG_FACE_NAME,
    TAG_FOOTNOTE_SHAPE,
    TAG_FORM_OBJECT,
    TAG_ID_MAPPINGS,
    TAG_LAYOUT_COMPATIBILITY,
    TAG_LIST_HEADER,
    TAG_MEMO_LIST,
    TAG_MEMO_SHAPE,
    TAG_NAMES,
    TAG_NUMBERING,
    TAG_PAGE_BORDER_FILL,
    TAG_PAGE_DEF,
    TAG_PARA_HEADER,
    TAG_PARA_SHAPE,
    TAG_PARA_TEXT,
    TAG_SHAPE_COMPONENT,
    TAG_SHAPE_COMPONENT_ARC,
    TAG_SHAPE_COMPONENT_CONTAINER,
    TAG_SHAPE_COMPONENT_CURVE,
    TAG_SHAPE_COMPONENT_ELLIPSE,
    TAG_SHAPE_COMPONENT_LINE,
    TAG_SHAPE_COMPONENT_OLE,
    TAG_SHAPE_COMPONENT_PICTURE,
    TAG_SHAPE_COMPONENT_POLYGON,
    TAG_SHAPE_COMPONENT_RECTANGLE,
    TAG_SHAPE_COMPONENT_TEXTART,
    TAG_STYLE,
    TAG_TAB_DEF,
    TAG_TRACK_CHANGE,
    TAG_TRACK_CHANGE_AUTHOR,
    ArcShapeComponentRecord,
    AutoNumberControlRecord,
    BinDataRecord,
    BookmarkControlRecord,
    BorderFillRecord,
    BulletRecord,
    CharShapeRecord,
    ChartDataRecord,
    CompatibleDocumentRecord,
    ContainerShapeComponentRecord,
    ControlDataRecord,
    ControlHeaderRecord,
    CurveShapeComponentRecord,
    DocDataRecord,
    DocumentPropertiesRecord,
    EllipseShapeComponentRecord,
    EqEditRecord,
    EquationControlRecord,
    FaceNameRecord,
    FieldControlRecord,
    FootnoteShapeRecord,
    FormObjectRecord,
    GraphicControlRecord,
    HeaderFooterControlRecord,
    HwpDocumentProperties,
    HwpBinaryEditError,
    HwpRecord,
    HyperlinkControlRecord,
    IdMappingsRecord,
    InvalidHwpFileError,
    LayoutCompatibilityRecord,
    LineShapeComponentRecord,
    ListHeaderRecord,
    MemoListRecord,
    MemoShapeRecord,
    NoteControlRecord,
    NumberingRecord,
    OleControlRecord,
    OleShapeComponentRecord,
    PageBorderFillRecord,
    PageDefRecord,
    PageHidingControlRecord,
    PageNumberControlRecord,
    PageStartsOnControlRecord,
    ParagraphHeaderRecord,
    ParagraphTextRecord,
    ParaShapeRecord,
    PictureControlRecord,
    PictureShapeComponentRecord,
    PolygonShapeComponentRecord,
    RectangleShapeComponentRecord,
    RecordNode,
    SectionDefinitionControlRecord,
    ShapeComponentRecord,
    ShapeControlRecord,
    StyleRecord,
    TabDefRecord,
    TableControlRecord,
    TextArtShapeComponentRecord,
    TrackChangeAuthorRecord,
    TrackChangeRecord,
    TypedRecord,
    UnknownRecord,
    _decode_control_id_payload,
    _find_paragraph_control_node,
    _find_section_definition_control_node,
)


def record_node_from_record(record: HwpRecord) -> RecordNode:
    if record.tag_id == TAG_DOCUMENT_PROPERTIES:
        return DocumentPropertiesRecord.from_record(record)
    if record.tag_id == TAG_ID_MAPPINGS:
        return IdMappingsRecord.from_record(record)
    if record.tag_id == TAG_BIN_DATA:
        return BinDataRecord.from_record(record)
    if record.tag_id == TAG_FACE_NAME:
        return FaceNameRecord.from_record(record)
    if record.tag_id == TAG_BORDER_FILL:
        return BorderFillRecord.from_record(record)
    if record.tag_id == TAG_CHAR_SHAPE:
        return CharShapeRecord.from_record(record)
    if record.tag_id == TAG_TAB_DEF:
        return TabDefRecord.from_record(record)
    if record.tag_id == TAG_NUMBERING:
        return NumberingRecord.from_record(record)
    if record.tag_id == TAG_BULLET:
        return BulletRecord.from_record(record)
    if record.tag_id == TAG_PARA_SHAPE:
        return ParaShapeRecord.from_record(record)
    if record.tag_id == TAG_STYLE:
        return StyleRecord.from_record(record)
    if record.tag_id == TAG_DOC_DATA:
        return DocDataRecord.from_record(record)
    if record.tag_id == TAG_COMPATIBLE_DOCUMENT:
        return CompatibleDocumentRecord.from_record(record)
    if record.tag_id == TAG_LAYOUT_COMPATIBILITY:
        return LayoutCompatibilityRecord.from_record(record)
    if record.tag_id == TAG_MEMO_SHAPE:
        return MemoShapeRecord.from_record(record)
    if record.tag_id == TAG_FORM_OBJECT:
        return FormObjectRecord.from_record(record)
    if record.tag_id == TAG_MEMO_LIST:
        return MemoListRecord.from_record(record)
    if record.tag_id == TAG_CHART_DATA:
        return ChartDataRecord.from_record(record)
    if record.tag_id == TAG_TRACK_CHANGE:
        return TrackChangeRecord.from_record(record)
    if record.tag_id == TAG_TRACK_CHANGE_AUTHOR:
        return TrackChangeAuthorRecord.from_record(record)
    if record.tag_id == TAG_CTRL_DATA:
        return ControlDataRecord.from_record(record)
    if record.tag_id == TAG_CTRL_HEADER:
        control_id = _decode_control_id_payload(record.payload)
        if control_id == "secd":
            return SectionDefinitionControlRecord.from_record(record)  # type: ignore[attr-defined]
        if control_id == "pghd":
            return PageHidingControlRecord.from_record(record)  # type: ignore[attr-defined]
        if control_id == "pgct":
            return PageStartsOnControlRecord.from_record(record)  # type: ignore[attr-defined]
        if control_id == "pgnp":
            return PageNumberControlRecord.from_record(record)  # type: ignore[attr-defined]
        if control_id == "bokm":
            return BookmarkControlRecord.from_record(record)  # type: ignore[attr-defined]
        if control_id in {"nwno", "atno"}:
            return AutoNumberControlRecord.from_record(record)  # type: ignore[attr-defined]
        if control_id in {"head", "foot"}:
            return HeaderFooterControlRecord.from_record(record)  # type: ignore[attr-defined]
        if control_id in {"fn  ", "en  "}:
            return NoteControlRecord.from_record(record)  # type: ignore[attr-defined]
        if control_id == "%hlk":
            return HyperlinkControlRecord.from_record(record)  # type: ignore[attr-defined]
        if control_id.startswith("%"):
            return FieldControlRecord.from_record(record)  # type: ignore[attr-defined]
        if control_id == "eqed":
            return EquationControlRecord.from_record(record)  # type: ignore[attr-defined]
        if control_id == "gso ":
            return GraphicControlRecord.from_record(record)  # type: ignore[attr-defined]
        if control_id == "tbl ":
            return TableControlRecord.from_record(record)  # type: ignore[attr-defined]
        return ControlHeaderRecord.from_record(record)  # type: ignore[attr-defined]
    if record.tag_id == TAG_PARA_HEADER:
        return ParagraphHeaderRecord.from_record(record)
    if record.tag_id == TAG_PARA_TEXT:
        return ParagraphTextRecord.from_record(record)
    if record.tag_id == TAG_LIST_HEADER:
        return ListHeaderRecord.from_record(record)
    if record.tag_id == TAG_PAGE_DEF:
        return PageDefRecord.from_record(record)
    if record.tag_id == TAG_FOOTNOTE_SHAPE:
        return FootnoteShapeRecord.from_record(record)
    if record.tag_id == TAG_PAGE_BORDER_FILL:
        return PageBorderFillRecord.from_record(record)
    if record.tag_id == TAG_EQEDIT:
        return EqEditRecord.from_record(record)
    if record.tag_id == TAG_SHAPE_COMPONENT:
        return ShapeComponentRecord.from_record(record)
    if record.tag_id == TAG_SHAPE_COMPONENT_LINE:
        return LineShapeComponentRecord.from_record(record)
    if record.tag_id == TAG_SHAPE_COMPONENT_RECTANGLE:
        return RectangleShapeComponentRecord.from_record(record)
    if record.tag_id == TAG_SHAPE_COMPONENT_ELLIPSE:
        return EllipseShapeComponentRecord.from_record(record)
    if record.tag_id == TAG_SHAPE_COMPONENT_ARC:
        return ArcShapeComponentRecord.from_record(record)
    if record.tag_id == TAG_SHAPE_COMPONENT_POLYGON:
        return PolygonShapeComponentRecord.from_record(record)
    if record.tag_id == TAG_SHAPE_COMPONENT_CURVE:
        return CurveShapeComponentRecord.from_record(record)
    if record.tag_id == TAG_SHAPE_COMPONENT_CONTAINER:
        return ContainerShapeComponentRecord.from_record(record)
    if record.tag_id == TAG_SHAPE_COMPONENT_OLE:
        return OleShapeComponentRecord.from_record(record)
    if record.tag_id == TAG_SHAPE_COMPONENT_PICTURE:
        return PictureShapeComponentRecord.from_record(record)
    if record.tag_id == TAG_SHAPE_COMPONENT_TEXTART:
        return TextArtShapeComponentRecord.from_record(record)
    if record.tag_id in TAG_NAMES:
        return TypedRecord(
            tag_id=record.tag_id,
            level=record.level,
            payload=record.payload,
            header_size=record.header_size,
            offset=record.offset,
        )
    return UnknownRecord(
        tag_id=record.tag_id,
        level=record.level,
        payload=record.payload,
        header_size=record.header_size,
        offset=record.offset,
    )


def build_record_tree(records: list[HwpRecord]) -> list[RecordNode]:
    roots: list[RecordNode] = []
    stack: list[RecordNode] = []
    for record in records:
        node = record_node_from_record(record)
        while stack and stack[-1].level >= node.level:
            stack.pop()
        if stack:
            stack[-1].children.append(node)
        else:
            roots.append(node)
        stack.append(node)
    _specialize_record_tree(roots)
    return roots


def _specialize_record_tree(nodes: list[RecordNode]) -> None:
    for index, node in enumerate(list(nodes)):
        if node.children:
            _specialize_record_tree(node.children)
        nodes[index] = _specialize_record_node(node)


def _specialize_record_node(node: RecordNode) -> RecordNode:
    if isinstance(node, GraphicControlRecord) and type(node) is GraphicControlRecord:
        if any(isinstance(descendant, OleShapeComponentRecord) for descendant in node.iter_descendants()):
            return OleControlRecord(
                level=node.level,
                payload=node.payload,
                header_size=node.header_size,
                offset=node.offset,
                children=node.children,
            )
        if any(isinstance(descendant, PictureShapeComponentRecord) for descendant in node.iter_descendants()):
            return PictureControlRecord(
                level=node.level,
                payload=node.payload,
                header_size=node.header_size,
                offset=node.offset,
                children=node.children,
            )
        return ShapeControlRecord(
            level=node.level,
            payload=node.payload,
            header_size=node.header_size,
            offset=node.offset,
            children=node.children,
        )
    return node


def flatten_record_tree(roots: list[RecordNode]) -> list[HwpRecord]:
    records: list[HwpRecord] = []
    for root in roots:
        records.extend(root.to_records())
    return records


@dataclass
class DocInfoModel:
    roots: list[RecordNode]

    @classmethod
    def from_records(cls, records: list[HwpRecord]) -> "DocInfoModel":
        return cls(roots=build_record_tree(records))

    def document_properties_record(self) -> DocumentPropertiesRecord:
        for root in self.roots:
            if isinstance(root, DocumentPropertiesRecord):
                return root
        raise InvalidHwpFileError("DocInfo does not start with a document properties record.")

    def document_properties(self) -> HwpDocumentProperties:
        return self.document_properties_record().to_properties()

    def id_mappings_record(self) -> IdMappingsRecord:
        for root in self.roots:
            if isinstance(root, IdMappingsRecord):
                return root
        raise InvalidHwpFileError("DocInfo does not contain an id mappings record.")

    def bin_data_records(self) -> list[BinDataRecord]:
        return [node for node in self.iter_nodes() if isinstance(node, BinDataRecord)]

    def face_name_records(self) -> list[FaceNameRecord]:
        return self.records_by_tag_id(TAG_FACE_NAME)

    def border_fill_records(self) -> list[BorderFillRecord]:
        return self.records_by_tag_id(TAG_BORDER_FILL)

    def char_shape_records(self) -> list[CharShapeRecord]:
        return self.records_by_tag_id(TAG_CHAR_SHAPE)

    def tab_def_records(self) -> list[TabDefRecord]:
        return self.records_by_tag_id(TAG_TAB_DEF)

    def numbering_records(self) -> list[NumberingRecord]:
        return self.records_by_tag_id(TAG_NUMBERING)

    def bullet_records(self) -> list[BulletRecord]:
        return self.records_by_tag_id(TAG_BULLET)

    def para_shape_records(self) -> list[ParaShapeRecord]:
        return self.records_by_tag_id(TAG_PARA_SHAPE)

    def style_records(self) -> list[StyleRecord]:
        return self.records_by_tag_id(TAG_STYLE)

    def face_name_fields(self) -> list[dict[str, object]]:
        return [record.fields() for record in self.face_name_records()]

    def border_fill_fields(self) -> list[dict[str, object]]:
        return [record.fields() for record in self.border_fill_records()]

    def char_shape_fields(self) -> list[dict[str, object]]:
        return [record.fields() for record in self.char_shape_records()]

    def tab_def_fields(self) -> list[dict[str, object]]:
        return [record.fields() for record in self.tab_def_records()]

    def numbering_fields(self) -> list[dict[str, object]]:
        return [record.fields() for record in self.numbering_records()]

    def bullet_fields(self) -> list[dict[str, object]]:
        return [record.fields() for record in self.bullet_records()]

    def para_shape_fields(self) -> list[dict[str, object]]:
        return [record.fields() for record in self.para_shape_records()]

    def style_fields(self) -> list[dict[str, object]]:
        return [record.fields(style_id=index) for index, record in enumerate(self.style_records())]

    def doc_data_records(self) -> list[DocDataRecord]:
        return self.records_by_tag_id(TAG_DOC_DATA)

    def compatible_document_records(self) -> list[CompatibleDocumentRecord]:
        return self.records_by_tag_id(TAG_COMPATIBLE_DOCUMENT)

    def layout_compatibility_records(self) -> list[LayoutCompatibilityRecord]:
        return self.records_by_tag_id(TAG_LAYOUT_COMPATIBILITY)

    def memo_shape_records(self) -> list[MemoShapeRecord]:
        return self.records_by_tag_id(TAG_MEMO_SHAPE)

    def track_change_records(self) -> list[TrackChangeRecord]:
        return self.records_by_tag_id(TAG_TRACK_CHANGE)

    def track_change_author_records(self) -> list[TrackChangeAuthorRecord]:
        return self.records_by_tag_id(TAG_TRACK_CHANGE_AUTHOR)

    def next_bin_data_id(self) -> int:
        id_record = self.id_mappings_record()
        return id_record.bin_data_count + 1

    def add_embedded_bindata(self, extension: str, *, storage_id: int | None = None, flags: int = 0x0001) -> BinDataRecord:
        id_record = self.id_mappings_record()
        if storage_id is None:
            storage_id = self.next_bin_data_id()
        bindata = BinDataRecord.embedded(
            level=id_record.level + 1,
            storage_id=storage_id,
            extension=extension,
            flags=flags,
        )
        insert_at = 0
        for index, child in enumerate(id_record.children):
            if isinstance(child, BinDataRecord):
                insert_at = index + 1
            else:
                break
        id_record.children.insert(insert_at, bindata)
        id_record.bin_data_count = max(id_record.bin_data_count, storage_id)
        return bindata

    def remove_bindata(self, storage_id: int) -> bool:
        id_record = self.id_mappings_record()
        for index, child in enumerate(id_record.children):
            if isinstance(child, BinDataRecord) and child.storage_id == storage_id:
                del id_record.children[index]
                remaining_ids = [record.storage_id or 0 for record in self.bin_data_records()]
                id_record.bin_data_count = max(remaining_ids, default=0)
                return True
        # Accept older malformed models long enough to remove and repair them.
        for index, root in enumerate(self.roots):
            if isinstance(root, BinDataRecord) and root.storage_id == storage_id:
                del self.roots[index]
                remaining_ids = [record.storage_id or 0 for record in self.bin_data_records()]
                id_record.bin_data_count = max(remaining_ids, default=0)
                return True
        return False

    def tag_counts(self) -> dict[str, int]:
        counts: dict[str, int] = {}
        for node in self.iter_nodes():
            counts[node.tag_name] = counts.get(node.tag_name, 0) + 1
        return counts

    def iter_nodes(self) -> Iterator[RecordNode]:
        for root in self.roots:
            yield from root.iter_preorder()

    def records_by_tag_id(self, tag_id: int) -> list[RecordNode]:
        return [node for node in self.iter_nodes() if node.tag_id == tag_id]

    def to_records(self) -> list[HwpRecord]:
        return flatten_record_tree(self.roots)

    def to_bytes(self) -> bytes:
        return b"".join(record.to_bytes() for record in self.to_records())


@dataclass
class SectionParagraphModel:
    section_index: int
    index: int
    header: ParagraphHeaderRecord

    def text_record(self) -> ParagraphTextRecord | None:
        for node in self.header.iter_descendants():
            if isinstance(node, ParagraphTextRecord):
                return node
        return None

    @property
    def raw_text(self) -> str:
        text_record = self.text_record()
        return text_record.raw_text if text_record is not None else ""

    @property
    def text(self) -> str:
        text_record = self.text_record()
        return text_record.text if text_record is not None else ""

    def control_nodes(self) -> list[RecordNode]:
        return [node for node in self.header.iter_descendants() if node.tag_id == TAG_CTRL_HEADER]

    def descendant_nodes(self) -> list[RecordNode]:
        return list(self.header.iter_descendants())

    def set_text_same_length(self, value: str) -> None:
        text_record = self.text_record()
        if text_record is None:
            raise HwpBinaryEditError("Paragraph does not contain a para text record.")
        if text_record.raw_text != text_record.text:
            raise HwpBinaryEditError(
                "SectionParagraphModel.set_text_same_length is only supported for paragraphs without hidden control characters."
            )
        text_record.set_text_same_length(value)
        self.header.char_count = len(value)

    def replace_text_same_length(self, old: str, new: str, *, count: int = -1) -> int:
        text_record = self.text_record()
        if text_record is None:
            return 0
        replaced = text_record.replace_text_same_length(old, new, count=count)
        self.header.char_count = len(text_record.raw_text)
        return replaced


@dataclass
class SectionModel:
    section_index: int
    roots: list[RecordNode]
    _paragraphs_cache: list[SectionParagraphModel] | None = field(default=None, init=False, repr=False)

    @classmethod
    def from_records(cls, section_index: int, records: list[HwpRecord]) -> "SectionModel":
        return cls(section_index=section_index, roots=build_record_tree(records))

    def iter_nodes(self) -> Iterator[RecordNode]:
        for root in self.roots:
            yield from root.iter_preorder()

    def paragraphs(self) -> list[SectionParagraphModel]:
        if self._paragraphs_cache is not None:
            return self._paragraphs_cache
        paragraphs: list[SectionParagraphModel] = []
        paragraph_index = 0
        for root in self.iter_nodes():
            if isinstance(root, ParagraphHeaderRecord):
                paragraphs.append(
                    SectionParagraphModel(
                        section_index=self.section_index,
                        index=paragraph_index,
                        header=root,
                    )
                )
                paragraph_index += 1
        self._paragraphs_cache = paragraphs
        return paragraphs

    def control_nodes(self) -> list[ControlHeaderRecord]:
        return [node for node in self.iter_nodes() if isinstance(node, ControlHeaderRecord)]

    def controls(self) -> list[ControlHeaderRecord]:
        return self.control_nodes()

    def controls_by_id(self, control_id: str) -> list[ControlHeaderRecord]:
        return [node for node in self.control_nodes() if node.control_id == control_id]

    def section_definition_control(self) -> SectionDefinitionControlRecord | None:
        control = _find_section_definition_control_node(self)
        return control if isinstance(control, SectionDefinitionControlRecord) else None

    def page_hiding_control(self) -> PageHidingControlRecord | None:
        control = _find_paragraph_control_node(self, "pghd")
        return control if isinstance(control, PageHidingControlRecord) else None

    def page_starts_on_control(self) -> PageStartsOnControlRecord | None:
        control = _find_paragraph_control_node(self, "pgct")
        return control if isinstance(control, PageStartsOnControlRecord) else None

    def page_number_controls(self) -> list[PageNumberControlRecord]:
        return [node for node in self.control_nodes() if isinstance(node, PageNumberControlRecord)]

    def form_object_records(self) -> list[FormObjectRecord]:
        return [node for node in self.iter_nodes() if isinstance(node, FormObjectRecord)]

    def memo_list_records(self) -> list[MemoListRecord]:
        return [node for node in self.iter_nodes() if isinstance(node, MemoListRecord)]

    def chart_data_records(self) -> list[ChartDataRecord]:
        return [node for node in self.iter_nodes() if isinstance(node, ChartDataRecord)]

    def append_paragraph(
        self,
        text: str,
        *,
        para_shape_id: int = 0,
        style_id: int = 0,
        split_flags: int = 0,
        control_mask: int = 0,
    ) -> SectionParagraphModel:
        header = ParagraphHeaderRecord(
            level=0,
            char_count=len(text),
            control_mask=control_mask,
            para_shape_id=para_shape_id,
            style_id=style_id,
            split_flags=split_flags,
        )
        header.add_child(ParagraphTextRecord(level=1, raw_text=text))
        paragraph = SectionParagraphModel(
            section_index=self.section_index,
            index=len(self.paragraphs()),
            header=header,
        )
        self.roots.append(header)
        self.paragraphs().append(paragraph)
        return paragraph

    def insert_paragraph(
        self,
        paragraph_index: int,
        text: str,
        *,
        para_shape_id: int = 0,
        style_id: int = 0,
        split_flags: int = 0,
        control_mask: int = 0,
    ) -> SectionParagraphModel:
        current_paragraphs = self.paragraphs()
        if paragraph_index < 0 or paragraph_index > len(current_paragraphs):
            raise IndexError(f"paragraph_index {paragraph_index} is out of range for section {self.section_index}.")
        header = ParagraphHeaderRecord(
            level=0,
            char_count=len(text),
            control_mask=control_mask,
            para_shape_id=para_shape_id,
            style_id=style_id,
            split_flags=split_flags,
        )
        header.add_child(ParagraphTextRecord(level=1, raw_text=text))
        if paragraph_index == len(current_paragraphs):
            self.roots.append(header)
        else:
            root_index = self._root_insert_index_for_paragraph(paragraph_index)
            self.roots.insert(root_index, header)
        self._paragraphs_cache = None
        return SectionParagraphModel(
            section_index=self.section_index,
            index=paragraph_index,
            header=header,
        )

    def _root_insert_index_for_paragraph(self, paragraph_index: int) -> int:
        seen = 0
        for root_index, root in enumerate(self.roots):
            paragraph_count = sum(1 for node in root.iter_preorder() if isinstance(node, ParagraphHeaderRecord))
            if seen + paragraph_count > paragraph_index:
                return root_index
            seen += paragraph_count
        return len(self.roots)

    def to_records(self) -> list[HwpRecord]:
        return flatten_record_tree(self.roots)

    def to_bytes(self) -> bytes:
        return b"".join(record.to_bytes() for record in self.to_records())


def _iter_records(data: bytes) -> Iterator[HwpRecord]:
    offset = 0
    while offset + 4 <= len(data):
        header = int.from_bytes(data[offset : offset + 4], "little")
        tag_id = header & 0x3FF
        level = (header >> 10) & 0x3FF
        size = (header >> 20) & 0xFFF
        header_size = 4
        if size == 0xFFF:
            if offset + 8 > len(data):
                break
            size = int.from_bytes(data[offset + 4 : offset + 8], "little")
            header_size = 8
        payload_offset = offset + header_size
        payload_end = payload_offset + size
        if payload_end > len(data):
            break
        yield HwpRecord(
            tag_id=tag_id,
            level=level,
            size=size,
            header_size=header_size,
            offset=offset,
            payload=data[payload_offset:payload_end],
        )
        offset = payload_end


def _strip_para_text_controls(raw_text: str) -> str:
    units = list(struct.unpack("<" + "H" * (len(raw_text.encode("utf-16-le")) // 2), raw_text.encode("utf-16-le")))
    output: list[str] = []
    index = 0
    while index < len(units):
        value = units[index]
        if value >= 32:
            output.append(chr(value))
            index += 1
            continue

        if value in (9, 10, 13):
            output.append(chr(value))

        # HWP extended/inline controls occupy 8 UTF-16 code units total.
        if (
            index + 7 < len(units)
            and units[index + 3] == 0
            and units[index + 4] == 0
            and units[index + 5] == 0
            and units[index + 6] == 0
        ):
            index += 8
            continue

        index += 1
    return "".join(output)

for _public_type in (DocInfoModel, SectionParagraphModel, SectionModel):
    _public_type.__module__ = "jakal_hwpx.hwp_binary"
del _public_type

