"""Native HWP section, field, bookmark, note, and story control record types."""
from __future__ import annotations

from ._hwp_component_records import (
    FootnoteShapeRecord,
    PageBorderFillRecord,
    PageDefRecord,
)
from ._hwp_constants import (
    _AUTO_NUMBER_TYPE_VALUES,
    TAG_CTRL_DATA,
    TAG_CTRL_HEADER,
)
from ._hwp_control_payloads import (
    _decode_control_id_payload,
    _parse_bookmark_ctrl_data_payload,
    _parse_header_footer_apply_page_type,
    _parse_hyperlink_command,
    _parse_shape_native_metadata,
    _semantic_field_type_from_native,
)
from ._hwp_docinfo_records import (
    ListHeaderRecord,
)
from ._hwp_page_payloads import (
    _SECTION_DEF_WONGGOJI_FORMAT_FLAG,
    _SECTION_PAGE_BORDER_FILL_TYPES,
    _parse_page_hiding_payload,
    _parse_page_num_payload,
    _parse_page_starts_on_payload,
    _parse_section_definition_payload,
)
from ._hwp_paragraph_records import (
    ParagraphHeaderRecord,
)
from ._hwp_record_base import (
    HwpRecord,
    RecordNode,
    TypedRecord,
)
from ._hwp_section_helpers import (
    _collect_text_from_node,
)


class ControlDataRecord(TypedRecord):
    def __init__(
        self,
        *,
        level: int,
        payload: bytes,
        header_size: int = 4,
        offset: int = -1,
        children: list[RecordNode] | None = None,
    ) -> None:
        super().__init__(
            tag_id=TAG_CTRL_DATA,
            level=level,
            payload=payload,
            header_size=header_size,
            offset=offset,
            children=list(children or []),
        )

    @classmethod
    def from_record(cls, record: HwpRecord) -> "ControlDataRecord":
        return cls(
            level=record.level,
            payload=record.payload,
            header_size=record.header_size,
            offset=record.offset,
        )

    def decoded_text(self) -> str:
        return self.payload.decode("utf-16-le", errors="ignore")

    def bookmark_name(self) -> str:
        return _parse_bookmark_ctrl_data_payload(self.payload)

    def native_metadata(self) -> dict[str, str]:
        return _parse_shape_native_metadata(self.payload)


class ControlHeaderRecord(TypedRecord):
    def __init__(
        self,
        *,
        level: int,
        payload: bytes,
        header_size: int = 4,
        offset: int = -1,
        children: list[RecordNode] | None = None,
    ) -> None:
        super().__init__(
            tag_id=TAG_CTRL_HEADER,
            level=level,
            payload=payload,
            header_size=header_size,
            offset=offset,
            children=list(children or []),
        )

    @classmethod
    def from_record(cls, record: HwpRecord) -> "ControlHeaderRecord":
        return cls(
            level=record.level,
            payload=record.payload,
            header_size=record.header_size,
            offset=record.offset,
        )

    @property
    def control_id(self) -> str:
        return _decode_control_id_payload(self.payload)

    def control_data_records(self) -> list[ControlDataRecord]:
        return [child for child in self.children if isinstance(child, ControlDataRecord)]

    def first_control_data(self) -> ControlDataRecord | None:
        return next(iter(self.control_data_records()), None)

    def list_header_records(self) -> list["ListHeaderRecord"]:
        return [child for child in self.children if isinstance(child, ListHeaderRecord)]

    def first_list_header(self) -> "ListHeaderRecord | None":
        return next(iter(self.list_header_records()), None)

    def nested_paragraphs(self) -> list["ParagraphHeaderRecord"]:
        return [node for node in self.iter_descendants() if isinstance(node, ParagraphHeaderRecord)]

    def text(self) -> str:
        return _collect_text_from_node(self).replace("\r", "").strip()


class SectionDefinitionControlRecord(ControlHeaderRecord):
    def settings(self) -> dict[str, object]:
        values = _parse_section_definition_payload(self.payload)
        attributes = int(values.get("attributes", 0))
        return {
            "attributes": attributes,
            "grid": {
                "space_columns": int(values.get("space_columns", 0)),
                "line_grid": int(values.get("line_grid", 0)),
                "char_grid": int(values.get("char_grid", 0)),
                "tab_stop": int(values.get("tab_stop", 0)),
                "wonggoji_format": bool(attributes & _SECTION_DEF_WONGGOJI_FORMAT_FLAG),
            },
            "start_numbers": {
                "page": int(values.get("page", 0)),
                "pic": int(values.get("pic", 0)),
                "tbl": int(values.get("tbl", 0)),
                "equation": int(values.get("equation", 0)),
            },
            "numbering_shape_id": int(values.get("numbering_shape_id", 0)),
            "language": int(values.get("language", 0)),
        }

    def page_def_record(self) -> "PageDefRecord | None":
        return next((child for child in self.children if isinstance(child, PageDefRecord)), None)

    def page_border_fill_records(self) -> list["PageBorderFillRecord"]:
        return [child for child in self.children if isinstance(child, PageBorderFillRecord)]

    def note_shape_records(self) -> list["FootnoteShapeRecord"]:
        return [child for child in self.children if isinstance(child, FootnoteShapeRecord)]

    def page_border_fills(self) -> list[dict[str, str | int]]:
        fills: list[dict[str, str | int]] = []
        for index, record in enumerate(self.page_border_fill_records()):
            border_type = _SECTION_PAGE_BORDER_FILL_TYPES[index] if index < len(_SECTION_PAGE_BORDER_FILL_TYPES) else f"EXTRA_{index}"
            fills.append(record.page_border_fill(border_type=border_type))
        return fills

    def note_settings(self) -> dict[str, dict[str, object]]:
        records = self.note_shape_records()
        settings: dict[str, dict[str, object]] = {}
        if len(records) >= 1:
            settings["footNotePr"] = records[0].note_shape(kind="footNotePr")
        if len(records) >= 2:
            settings["endNotePr"] = records[1].note_shape(kind="endNotePr")
        return settings


class PageHidingControlRecord(ControlHeaderRecord):
    def visibility(self) -> dict[str, str]:
        return _parse_page_hiding_payload(self.payload)


class PageStartsOnControlRecord(ControlHeaderRecord):
    def page_starts_on(self) -> str:
        return _parse_page_starts_on_payload(self.payload)


class PageNumberControlRecord(ControlHeaderRecord):
    def pos(self) -> str:
        return self.page_number().get("pos", "BOTTOM_CENTER")

    def format_type(self) -> str:
        return self.page_number().get("formatType", "DIGIT")

    def side_char(self) -> str:
        return self.page_number().get("sideChar", "-")

    def page_number(self) -> dict[str, str]:
        return _parse_page_num_payload(self.payload)

    def fields(self) -> dict[str, object]:
        values: dict[str, object] = dict(self.page_number())
        values.update(
            {
                "pos": self.pos(),
                "format_type": self.format_type(),
                "side_char": self.side_char(),
            }
        )
        return values


class BookmarkControlRecord(ControlHeaderRecord):
    def name(self) -> str:
        data = self.first_control_data()
        return "" if data is None else data.bookmark_name()


class FieldControlRecord(ControlHeaderRecord):
    def native_field_type(self) -> str:
        return self.control_id

    def name(self) -> str | None:
        for key in ("FieldName", "Name", "name"):
            value = self.get_parameter(key)
            if value:
                return value
        return None

    def parameters(self) -> dict[str, str]:
        parameters: dict[str, str] = {}
        for entry in self.control_data_text().split(";"):
            key, separator, value = entry.partition("=")
            if separator and key:
                parameters[key] = value
        return parameters

    def get_parameter(self, name: str) -> str | None:
        return self.parameters().get(name)

    def field_type(self) -> str:
        return _semantic_field_type_from_native(self.native_field_type(), self.parameters())

    def is_mail_merge(self) -> bool:
        return self.field_type() == "MAILMERGE"

    def is_calculation(self) -> bool:
        return self.field_type() == "FORMULA"

    def is_cross_reference(self) -> bool:
        return self.field_type() == "CROSSREF"

    def is_doc_property(self) -> bool:
        return self.field_type() == "DOCPROPERTY"

    def is_date(self) -> bool:
        return self.field_type() == "DATE"

    def merge_field_name(self) -> str | None:
        return self.get_parameter("MergeField") or self.get_parameter("FieldName") or self.name()

    def expression(self) -> str | None:
        return self.get_parameter("Expression") or self.get_parameter("Command")

    def bookmark_name(self) -> str | None:
        return self.get_parameter("BookmarkName") or self.get_parameter("Path") or self.name()

    def document_property_name(self) -> str | None:
        if not self.is_doc_property():
            return None
        return self.get_parameter("FieldName") or self.name()

    def command(self) -> str:
        return _parse_hyperlink_command(self.payload) if len(self.payload) >= 9 else ""

    def control_data_text(self) -> str:
        data = self.first_control_data()
        return "" if data is None else data.decoded_text()

    def fields(self) -> dict[str, object]:
        return {
            "field_type": self.field_type(),
            "native_field_type": self.native_field_type(),
            "name": self.name(),
            "parameters": self.parameters(),
            "is_mail_merge": self.is_mail_merge(),
            "is_calculation": self.is_calculation(),
            "is_cross_reference": self.is_cross_reference(),
            "is_doc_property": self.is_doc_property(),
            "is_date": self.is_date(),
            "merge_field_name": self.merge_field_name(),
            "expression": self.expression(),
            "bookmark_name": self.bookmark_name(),
            "document_property_name": self.document_property_name(),
            "command": self.command(),
            "display_text": self.text(),
            "control_data_text": self.control_data_text(),
        }


class HyperlinkControlRecord(FieldControlRecord):
    def command(self) -> str:
        return _parse_hyperlink_command(self.payload)

    def url(self) -> str:
        return self.command().split(";", 1)[0].replace("\\:", ":")

    def metadata_fields(self) -> list[str]:
        parts = self.command().split(";")
        if parts and parts[-1] == "":
            parts = parts[:-1]
        return parts[1:]

    def display_text(self) -> str:
        return self.text()

    def fields(self) -> dict[str, object]:
        values = super().fields()
        values.update(
            {
                "url": self.url(),
                "metadata_fields": self.metadata_fields(),
            }
        )
        return values


class AutoNumberControlRecord(ControlHeaderRecord):
    def kind(self) -> str:
        return "newNum" if self.control_id == "nwno" else "autoNum"

    def number(self) -> str | None:
        if self.control_id == "nwno" and len(self.payload) >= 10:
            return str(int.from_bytes(self.payload[8:10], "little", signed=False))
        if self.control_id == "atno":
            return "1"
        return None

    def number_type(self) -> str:
        if self.control_id == "atno" and len(self.payload) >= 12:
            code = int.from_bytes(self.payload[8:12], "little", signed=False)
            return _AUTO_NUMBER_TYPE_VALUES.get(code, "PAGE")
        return "PAGE"

    def fields(self) -> dict[str, object]:
        return {
            "kind": self.kind(),
            "number": self.number(),
            "number_type": self.number_type(),
        }


class HeaderFooterControlRecord(ControlHeaderRecord):
    def kind(self) -> str:
        return "header" if self.control_id == "head" else "footer"

    def apply_page_type(self) -> str:
        return _parse_header_footer_apply_page_type(self.payload)

    def paragraph_texts(self) -> list[str]:
        return [_collect_text_from_node(paragraph).replace("\r", "").strip() for paragraph in self.nested_paragraphs()]

    def list_header(self) -> dict[str, object]:
        header = self.first_list_header()
        return {} if header is None else header.fields()

    def fields(self) -> dict[str, object]:
        return {
            "kind": self.kind(),
            "apply_page_type": self.apply_page_type(),
            "text": self.text(),
            "paragraph_texts": self.paragraph_texts(),
            "list_header": self.list_header(),
        }


class NoteControlRecord(ControlHeaderRecord):
    def kind(self) -> str:
        return "footNote" if self.control_id == "fn  " else "endNote"

    def paragraph_texts(self) -> list[str]:
        return [_collect_text_from_node(paragraph).replace("\r", "").strip() for paragraph in self.nested_paragraphs()]

    def list_header(self) -> dict[str, object]:
        header = self.first_list_header()
        return {} if header is None else header.fields()

    def fields(self) -> dict[str, object]:
        return {
            "kind": self.kind(),
            "text": self.text(),
            "paragraph_texts": self.paragraph_texts(),
            "list_header": self.list_header(),
        }


# Preserve historical import and pickle identities.
for _public_type in (ControlDataRecord, ControlHeaderRecord, SectionDefinitionControlRecord, PageHidingControlRecord, PageStartsOnControlRecord, PageNumberControlRecord, BookmarkControlRecord, FieldControlRecord, HyperlinkControlRecord, AutoNumberControlRecord, HeaderFooterControlRecord, NoteControlRecord):
    _public_type.__module__ = "jakal_hwpx.hwp_binary"
del _public_type
