"""Native HWP section count, page layout, numbering, borders, and notes."""
from __future__ import annotations

import zlib
from typing import TYPE_CHECKING, cast

from ._hwp_constants import (
    TAG_CTRL_HEADER,
    TAG_FOOTNOTE_SHAPE,
    TAG_PAGE_BORDER_FILL,
    TAG_PAGE_DEF,
)
from ._hwp_nested_controls import (
    _apply_section_show_line_numbers,
    _ensure_paragraph_line_seg,
    _section_uses_synthetic_line_number_layout,
)
from ._hwp_page_payloads import (
    _SECTION_DEF_BORDER_FIRST_FLAG,
    _SECTION_DEF_FILL_FIRST_FLAG,
    _SECTION_DEF_HIDE_EMPTY_LINE_FLAG,
    _SECTION_DEF_WONGGOJI_FORMAT_FLAG,
    _SECTION_DEFINITION_STARTS_ON_MASK,
    _SECTION_DEFINITION_STARTS_ON_SHIFT,
    _SECTION_DEFINITION_STARTS_ON_VALUES,
    _SECTION_PAGE_BORDER_FILL_TYPES,
    _build_note_shape_payload,
    _build_page_border_fill_payload,
    _build_page_def_payload,
    _build_page_hiding_payload,
    _build_page_num_payload,
    _build_section_definition_payload,
    _parse_note_shape_payload,
    _parse_page_border_fill_payload,
    _parse_page_def_payload,
    _parse_page_hiding_payload,
    _parse_page_num_payload,
    _parse_page_starts_on_payload,
    _parse_section_definition_payload,
)
from ._hwp_record_tree import (
    DocInfoModel,
    SectionModel,
)
from .exceptions import (
    HwpBinaryEditError,
    InvalidHwpFileError,
)


class HwpBinarySectionMixin:
    """Native HWP section count, page layout, numbering, borders, and notes."""

    if TYPE_CHECKING:
        def _stream_is_compressed(self, path: str) -> bool: ...

        def add_stream(self, path: str | tuple[str, ...] | list[str], data: bytes) -> None: ...

        def replace_section_model(self, section_index: int, model: SectionModel) -> None: ...

        def section_model(self, section_index: int = 0) -> SectionModel: ...

        def section_stream_paths(self) -> list[str]: ...

        def set_preview_text(self, value: str) -> None: ...

        def docinfo_model(self) -> DocInfoModel: ...

        def replace_docinfo_model(self, model: DocInfoModel) -> None: ...

    def ensure_section_count(self, section_count: int) -> None:
        from ._hwp_section_helpers import (
            _build_blank_section_model,
            _normalize_section_definition_paragraph_for_multi_section,
        )

        if section_count < 1:
            raise ValueError("section_count must be at least 1.")

        current_count = len(self.section_stream_paths())
        if current_count == 0:
            raise InvalidHwpFileError("HWP document does not contain any BodyText/Section streams.")
        if current_count >= section_count:
            return

        if section_count > 1:
            for existing_section_index in range(current_count):
                existing_model = self.section_model(existing_section_index)
                _normalize_section_definition_paragraph_for_multi_section(existing_model)
                self.replace_section_model(existing_section_index, existing_model)

        while current_count < section_count:
            template_index = 1 if current_count > 1 else current_count - 1
            template_model = _build_blank_section_model(self.section_model(template_index), section_index=current_count)
            _normalize_section_definition_paragraph_for_multi_section(template_model)
            raw = template_model.to_bytes()
            if self._stream_is_compressed("BodyText/Section0"):
                compressor = zlib.compressobj(level=9, wbits=-15)
                payload = compressor.compress(raw) + compressor.flush()
            else:
                payload = raw
            self.add_stream(f"BodyText/Section{current_count}", payload)
            current_count += 1

        self._set_document_properties_section_count(current_count)

    def reset_body_sections_to_blank(self) -> None:
        from ._hwp_section_helpers import (
            _build_blank_section_model,
        )

        section_paths = self.section_stream_paths()
        for section_index, _path in enumerate(section_paths):
            self.replace_section_model(section_index, _build_blank_section_model(self.section_model(section_index)))
        self._set_document_properties_section_count(len(section_paths))
        self.set_preview_text("")

    def section_page_settings(self, section_index: int) -> dict[str, int]:
        from ._hwp_section_helpers import (
            _find_child_record,
            _find_section_definition_control_node,
        )

        self.ensure_section_count(section_index + 1)
        section_definition = _find_section_definition_control_node(self.section_model(section_index))
        if section_definition is None:
            return {}
        page_def = _find_child_record(section_definition, TAG_PAGE_DEF)
        if page_def is None:
            return {}
        return _parse_page_def_payload(page_def.payload)

    def set_section_page_settings(
        self,
        section_index: int,
        *,
        page_width: int | None = None,
        page_height: int | None = None,
        landscape: str | None = None,
        margins: dict[str, int] | None = None,
    ) -> None:
        from ._hwp_section_helpers import (
            _find_child_record,
            _find_section_definition_control_node,
        )

        self.ensure_section_count(section_index + 1)
        model = self.section_model(section_index)
        section_definition = _find_section_definition_control_node(model)
        if section_definition is None:
            raise HwpBinaryEditError(f"Section {section_index} does not contain a writable secd control.")
        page_def = _find_child_record(section_definition, TAG_PAGE_DEF)
        if page_def is None:
            raise HwpBinaryEditError(f"Section {section_index} does not contain a writable page_def record.")
        page_def.payload = _build_page_def_payload(
            page_def.payload,
            page_width=page_width,
            page_height=page_height,
            landscape=landscape,
            margins=margins,
        )
        self.replace_section_model(section_index, model)

    def section_page_border_fills(self, section_index: int) -> list[dict[str, str | int]]:
        from ._hwp_section_helpers import (
            _find_section_definition_control_node,
        )

        self.ensure_section_count(section_index + 1)
        section_definition = _find_section_definition_control_node(self.section_model(section_index))
        if section_definition is None:
            return []
        records = [child for child in section_definition.children if child.tag_id == TAG_PAGE_BORDER_FILL]
        fills: list[dict[str, str | int]] = []
        for index, record in enumerate(records):
            border_type = _SECTION_PAGE_BORDER_FILL_TYPES[index] if index < len(_SECTION_PAGE_BORDER_FILL_TYPES) else f"EXTRA_{index}"
            fills.append(_parse_page_border_fill_payload(record.payload, border_type=border_type))
        return fills

    def set_section_page_border_fills(
        self,
        section_index: int,
        page_border_fills: list[dict[str, str | int]],
    ) -> None:
        from ._hwp_section_helpers import (
            _find_section_definition_control_node,
        )

        self.ensure_section_count(section_index + 1)
        model = self.section_model(section_index)
        section_definition = _find_section_definition_control_node(model)
        if section_definition is None:
            raise HwpBinaryEditError(f"Section {section_index} does not contain a writable secd control.")
        records = [child for child in section_definition.children if child.tag_id == TAG_PAGE_BORDER_FILL]
        if not records:
            raise HwpBinaryEditError(f"Section {section_index} does not contain writable page_border_fill records.")
        record_by_type = {
            _SECTION_PAGE_BORDER_FILL_TYPES[index]: record
            for index, record in enumerate(records[: len(_SECTION_PAGE_BORDER_FILL_TYPES)])
        }
        for border_fill in page_border_fills:
            border_type = str(border_fill.get("type", "BOTH")).upper()
            record = record_by_type.get(border_type)
            if record is None:
                continue
            record.payload = _build_page_border_fill_payload(record.payload, border_fill=border_fill)
        self.replace_section_model(section_index, model)

    def section_definition_settings(self, section_index: int) -> dict[str, object]:
        from ._hwp_control_payloads import (
            _control_native_metadata,
        )
        from ._hwp_section_helpers import (
            _find_paragraph_control_node,
            _find_section_definition_control_node,
        )

        self.ensure_section_count(section_index + 1)
        model = self.section_model(section_index)
        section_definition = _find_section_definition_control_node(model)
        if section_definition is None:
            return {}
        values = _parse_section_definition_payload(section_definition.payload)
        metadata = _control_native_metadata(section_definition)
        page_hiding_control = _find_paragraph_control_node(model, "pghd")
        page_hiding = _parse_page_hiding_payload(page_hiding_control.payload) if page_hiding_control is not None else {}
        attributes = int(cast(int, values.get("attributes", 0)))
        page_starts_on_code = (
            attributes & _SECTION_DEFINITION_STARTS_ON_MASK
        ) >> _SECTION_DEFINITION_STARTS_ON_SHIFT
        page_starts_on = _SECTION_DEFINITION_STARTS_ON_VALUES.get(page_starts_on_code, "BOTH")
        page_starts_on_control = _find_paragraph_control_node(model, "pgct")
        if page_starts_on_code == 0 and page_starts_on_control is not None:
            page_starts_on = _parse_page_starts_on_payload(page_starts_on_control.payload)
        border_hidden = page_hiding.get("hideBorder") == "1"
        fill_hidden = page_hiding.get("hideFill") == "1"
        border_first = bool(attributes & _SECTION_DEF_BORDER_FIRST_FLAG)
        fill_first = bool(attributes & _SECTION_DEF_FILL_FIRST_FLAG)
        return {
            "visibility": {
                "hideFirstHeader": page_hiding.get("hideFirstHeader", "0"),
                "hideFirstFooter": page_hiding.get("hideFirstFooter", "0"),
                "hideFirstMasterPage": page_hiding.get("hideFirstMasterPage", "0"),
                "hideFirstPageNum": page_hiding.get("hideFirstPageNum", "0"),
                "hideFirstEmptyLine": "1" if attributes & _SECTION_DEF_HIDE_EMPTY_LINE_FLAG else "0",
                "border": "HIDE_FIRST" if border_hidden and border_first else "HIDE_ALL" if border_hidden else "SHOW_FIRST" if border_first else "SHOW_ALL",
                "fill": "HIDE_FIRST" if fill_hidden and fill_first else "HIDE_ALL" if fill_hidden else "SHOW_FIRST" if fill_first else "SHOW_ALL",
                "showLineNumber": "1" if _section_uses_synthetic_line_number_layout(model) else "0",
            },
            "grid": {
                "lineGrid": int(cast(int, values.get("line_grid", 0))),
                "charGrid": int(cast(int, values.get("char_grid", 0))),
                "wonggojiFormat": 1 if attributes & _SECTION_DEF_WONGGOJI_FORMAT_FLAG else 0,
            },
            "start_numbers": {
                "pageStartsOn": page_starts_on,
                "page": str(values.get("page", 0)),
                "pic": str(values.get("pic", 0)),
                "tbl": str(values.get("tbl", 0)),
                "equation": str(values.get("equation", 0)),
            },
            "numbering_shape_id": str(values.get("numbering_shape_id", 0)),
            "memo_shape_id": metadata.get("section.memoShapeId") or None,
        }

    def set_section_definition_settings(
        self,
        section_index: int,
        *,
        visibility: dict[str, str] | None = None,
        grid: dict[str, int] | None = None,
        start_numbers: dict[str, str] | None = None,
        numbering_shape_id: int | None = None,
        memo_shape_id: int | str | None = None,
    ) -> None:
        from ._hwp_control_payloads import (
            _update_control_native_metadata,
        )
        from ._hwp_record_base import (
            RecordNode,
        )
        from ._hwp_section_helpers import (
            _find_paragraph_control_node,
            _find_section_definition_control_node,
            _insert_paragraph_control_node,
        )

        self.ensure_section_count(section_index + 1)
        model = self.section_model(section_index)
        section_definition = _find_section_definition_control_node(model)
        if section_definition is None:
            raise HwpBinaryEditError(f"Section {section_index} does not contain a writable secd control.")
        section_definition.payload = _build_section_definition_payload(
            section_definition.payload,
            visibility=visibility,
            grid=grid,
            start_numbers=start_numbers,
            numbering_shape_id=numbering_shape_id,
        )
        if memo_shape_id is not None:
            _update_control_native_metadata(
                section_definition,
                updates={"section.memoShapeId": memo_shape_id},
            )
        if visibility is not None:
            current = cast(dict[str, str], self.section_definition_settings(section_index)["visibility"])
            current.update({key: value for key, value in visibility.items() if value is not None})
            border_value = str(current.get("border", "SHOW_ALL")).upper()
            fill_value = str(current.get("fill", "SHOW_ALL")).upper()
            payload = bytearray(section_definition.payload)
            attributes = int.from_bytes(payload[4:8], "little")
            if border_value in {"SHOW_FIRST", "HIDE_FIRST"}:
                attributes |= _SECTION_DEF_BORDER_FIRST_FLAG
            else:
                attributes &= ~_SECTION_DEF_BORDER_FIRST_FLAG
            if fill_value in {"SHOW_FIRST", "HIDE_FIRST"}:
                attributes |= _SECTION_DEF_FILL_FIRST_FLAG
            else:
                attributes &= ~_SECTION_DEF_FILL_FIRST_FLAG
            payload[4:8] = attributes.to_bytes(4, "little", signed=False)
            section_definition.payload = bytes(payload)
            page_hiding_visibility = {
                "hideFirstHeader": current.get("hideFirstHeader", "0"),
                "hideFirstFooter": current.get("hideFirstFooter", "0"),
                "hideFirstMasterPage": current.get("hideFirstMasterPage", "0"),
                "hideFirstPageNum": current.get("hideFirstPageNum", "0"),
                "hideBorder": "1" if border_value in {"HIDE_FIRST", "HIDE_ALL"} else "0",
                "hideFill": "1" if fill_value in {"HIDE_FIRST", "HIDE_ALL"} else "0",
            }
            page_hiding_control = _find_paragraph_control_node(model, "pghd")
            if any(value == "1" for value in page_hiding_visibility.values()):
                if page_hiding_control is None:
                    page_hiding_control = RecordNode(tag_id=TAG_CTRL_HEADER, level=1, payload=_build_page_hiding_payload(page_hiding_visibility))
                    _insert_paragraph_control_node(model, page_hiding_control, after_control_id="cold")
                else:
                    page_hiding_control.payload = _build_page_hiding_payload(page_hiding_visibility)
            elif page_hiding_control is not None:
                for paragraph in model.paragraphs():
                    if page_hiding_control in paragraph.header.children:
                        paragraph.header.children.remove(page_hiding_control)
                        break
            if visibility.get("showLineNumber") is not None:
                _apply_section_show_line_numbers(model, str(current.get("showLineNumber", "0")) in {"1", "true", "True"})
        if start_numbers is not None and start_numbers.get("pageStartsOn") is not None:
            page_starts_on_control = _find_paragraph_control_node(model, "pgct")
            if page_starts_on_control is not None:
                model.paragraphs()[0].header.children.remove(page_starts_on_control)
        self.replace_section_model(section_index, model)

    def section_page_numbers(self, section_index: int) -> list[dict[str, str]]:
        from ._hwp_control_payloads import (
            _build_control_id_payload,
        )

        self.ensure_section_count(section_index + 1)
        model = self.section_model(section_index)
        page_numbers: list[dict[str, str]] = []
        for paragraph in model.paragraphs():
            for child in paragraph.header.children:
                if child.tag_id == TAG_CTRL_HEADER and child.payload[:4] == _build_control_id_payload("pgnp"):
                    page_numbers.append(_parse_page_num_payload(child.payload))
        return page_numbers

    def set_section_page_numbers(self, section_index: int, page_numbers: list[dict[str, str]]) -> None:
        from ._hwp_control_payloads import (
            _build_control_id_payload,
            _build_control_marker_text,
        )
        from ._hwp_record_base import (
            RecordNode,
        )

        self.ensure_section_count(section_index + 1)
        model = self.section_model(section_index)
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(model)
        for paragraph in model.paragraphs():
            paragraph.header.children = [
                child
                for child in paragraph.header.children
                if not (child.tag_id == TAG_CTRL_HEADER and child.payload[:4] == _build_control_id_payload("pgnp"))
            ]
        for page_num in page_numbers:
            raw_text = _build_control_marker_text("pgnp") + "\r"
            paragraph = model.append_paragraph(raw_text)
            if use_synthetic_line_numbers:
                _ensure_paragraph_line_seg(paragraph)
            paragraph.header.add_child(
                RecordNode(tag_id=TAG_CTRL_HEADER, level=1, payload=_build_page_num_payload(page_num))
            )
        self.replace_section_model(section_index, model)

    def section_note_settings(self, section_index: int) -> dict[str, dict[str, object]]:
        from ._hwp_section_helpers import (
            _find_section_definition_control_node,
        )

        self.ensure_section_count(section_index + 1)
        section_definition = _find_section_definition_control_node(self.section_model(section_index))
        if section_definition is None:
            return {}
        records = [child for child in section_definition.children if child.tag_id == TAG_FOOTNOTE_SHAPE]
        settings: dict[str, dict[str, object]] = {}
        if len(records) >= 1:
            settings["footNotePr"] = _parse_note_shape_payload(records[0].payload, kind="footNotePr")
        if len(records) >= 2:
            settings["endNotePr"] = _parse_note_shape_payload(records[1].payload, kind="endNotePr")
        return settings

    def set_section_note_settings(
        self,
        section_index: int,
        *,
        footnote_pr: dict[str, object] | None = None,
        endnote_pr: dict[str, object] | None = None,
    ) -> None:
        from ._hwp_section_helpers import (
            _find_section_definition_control_node,
        )

        self.ensure_section_count(section_index + 1)
        model = self.section_model(section_index)
        section_definition = _find_section_definition_control_node(model)
        if section_definition is None:
            raise HwpBinaryEditError(f"Section {section_index} does not contain a writable secd control.")
        records = [child for child in section_definition.children if child.tag_id == TAG_FOOTNOTE_SHAPE]
        if len(records) >= 1 and footnote_pr is not None:
            records[0].payload = _build_note_shape_payload(records[0].payload, note_pr=footnote_pr, kind="footNotePr")
        if len(records) >= 2 and endnote_pr is not None:
            records[1].payload = _build_note_shape_payload(records[1].payload, note_pr=endnote_pr, kind="endNotePr")
        self.replace_section_model(section_index, model)

    def _set_document_properties_section_count(self, section_count: int) -> None:
        from ._hwp_record_base import (
            HwpDocumentProperties,
        )

        model = self.docinfo_model()
        properties_record = model.document_properties_record()
        properties = properties_record.to_properties()
        properties_record.properties = HwpDocumentProperties(
            section_count=section_count,
            page_start_number=properties.page_start_number,
            footnote_start_number=properties.footnote_start_number,
            endnote_start_number=properties.endnote_start_number,
            picture_start_number=properties.picture_start_number,
            table_start_number=properties.table_start_number,
            equation_start_number=properties.equation_start_number,
            list_id=properties.list_id,
            paragraph_id=properties.paragraph_id,
            character_unit_position=properties.character_unit_position,
        )
        self.replace_docinfo_model(model)
