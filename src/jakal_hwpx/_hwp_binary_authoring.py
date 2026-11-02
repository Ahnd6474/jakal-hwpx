"""Native HWP paragraph and object creation, profiles, and embedded media."""
from __future__ import annotations

import json
from collections.abc import (
    Sequence,
)
from pathlib import (
    Path,
)
from typing import TYPE_CHECKING

from ._hwp_constants import (
    _DEFAULT_EQUATION_CTRL_HEADER,
    _DEFAULT_OLE_LINE_WIDTH,
    _HYPERLINK_CHAR_SHAPE,
    _HYPERLINK_LINE_SEG,
    _HYPERLINK_PARA_HEADER_TAIL,
    _PICTURE_CTRL_HEADER,
    _PICTURE_SHAPE_COMPONENT,
    _PICTURE_SHAPE_PICTURE,
    _PICTURE_TOP_CHAR_SHAPE,
    _PICTURE_TOP_LINE_SEG,
    _PICTURE_TOP_PARA_HEADER_TAIL,
    _PICTURE_TOP_PARA_TEXT,
    _TABLE_CELL_CHAR_SHAPE,
    _TABLE_CELL_LINE_SEG,
    _TABLE_CELL_PARA_HEADER_TAIL,
    _TABLE_CTRL_HEADER,
    _TABLE_TOP_CHAR_SHAPE,
    _TABLE_TOP_LINE_SEG,
    _TABLE_TOP_PARA_HEADER_TAIL,
    _TABLE_TOP_PARA_TEXT,
    DEFAULT_HWP_EQUATION_FONT,
    DEFAULT_HWP_EQUATION_SCRIPT,
    DEFAULT_HWP_HYPERLINK_TEXT,
    DEFAULT_HWP_HYPERLINK_URL,
    DEFAULT_HWP_SHAPE_HEIGHT,
    DEFAULT_HWP_SHAPE_TEXT,
    DEFAULT_HWP_SHAPE_WIDTH,
    DEFAULT_HWP_TABLE_BORDER_FILL_ID,
    DEFAULT_HWP_TABLE_CELL_HEIGHT,
    DEFAULT_HWP_TABLE_CELL_WIDTH,
    DEFAULT_HWP_TABLE_COLS,
    DEFAULT_HWP_TABLE_ROWS,
    TAG_CHART_DATA,
    TAG_CTRL_DATA,
    TAG_CTRL_HEADER,
    TAG_EQEDIT,
    TAG_FORM_OBJECT,
    TAG_LIST_HEADER,
    TAG_MEMO_LIST,
    TAG_PARA_CHAR_SHAPE,
    TAG_PARA_HEADER,
    TAG_PARA_LINE_SEG,
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
    TAG_TABLE,
)
from ._hwp_graphic_payloads import (
    _build_control_native_metadata_payload,
    _build_graphic_control_header_payload,
    _build_ole_specific_payload,
    _build_shape_component_payload,
    _build_shape_native_metadata_payload,
    _build_shape_specific_payload,
    _shape_control_header_template,
)
from ._hwp_nested_controls import (
    _append_control_marker_to_paragraph,
    _apply_section_show_line_numbers,
    _build_header_footer_control_node,
    _build_note_control_node,
    _ensure_paragraph_line_seg,
    _section_uses_synthetic_line_number_layout,
    _set_header_footer_control_text,
)
from ._hwp_record_tree import (
    DocInfoModel,
    SectionModel,
    SectionParagraphModel,
)
from .exceptions import (
    HwpBinaryEditError,
)

if TYPE_CHECKING:
    from ._hwp_record_base import (
        HwpRecord,
    )


class HwpBinaryAuthoringMixin:
    """Native HWP paragraph and object creation, profiles, and embedded media."""

    if TYPE_CHECKING:
        _streams: dict[str, bytes]

        def ensure_section_count(self, section_count: int) -> None: ...

        def replace_section_model(self, section_index: int, model: SectionModel) -> None: ...

        def section_model(self, section_index: int = 0) -> SectionModel: ...

        def add_stream(self, path: str | tuple[str, ...] | list[str], data: bytes) -> None: ...

        def docinfo_model(self) -> DocInfoModel: ...

        def replace_docinfo_model(self, model: DocInfoModel) -> None: ...

        def remove_stream(self, path: str | tuple[str, ...] | list[str]) -> None: ...

        def bindata_stream_paths(self) -> list[str]: ...

        def replace_section_records(self, section_index: int, records: list[HwpRecord]) -> None: ...

        def section_records(self, section_index: int) -> list[HwpRecord]: ...

    def append_paragraph(
        self,
        text: str,
        *,
        section_index: int = 0,
        para_shape_id: int = 0,
        style_id: int = 0,
        split_flags: int = 0,
        control_mask: int = 0,
    ) -> SectionParagraphModel:
        self.ensure_section_count(section_index + 1)
        model = self.section_model(section_index)
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(model)
        paragraph = model.append_paragraph(
            text,
            para_shape_id=para_shape_id,
            style_id=style_id,
            split_flags=split_flags,
            control_mask=control_mask,
        )
        if use_synthetic_line_numbers:
            _ensure_paragraph_line_seg(paragraph)
        self.replace_section_model(section_index, model)
        return paragraph

    def append_table(
        self,
        cell_text: str | None = None,
        *,
        rows: int = DEFAULT_HWP_TABLE_ROWS,
        cols: int = DEFAULT_HWP_TABLE_COLS,
        cell_texts: Sequence[str] | Sequence[Sequence[str]] | None = None,
        row_heights: Sequence[int] | None = None,
        col_widths: Sequence[int] | None = None,
        cell_spans: dict[tuple[int, int], tuple[int, int]] | None = None,
        cell_border_fill_ids: dict[tuple[int, int], int] | None = None,
        table_border_fill_id: int = DEFAULT_HWP_TABLE_BORDER_FILL_ID,
        profile_root: str | Path | None = None,
        section_index: int | None = None,
    ) -> None:
        self._append_table_with_text(
            cell_text,
            rows=rows,
            cols=cols,
            cell_texts=cell_texts,
            row_heights=row_heights,
            col_widths=col_widths,
            cell_spans=cell_spans,
            cell_border_fill_ids=cell_border_fill_ids,
            table_border_fill_id=table_border_fill_id,
            profile_root=profile_root,
            section_index=section_index,
        )

    def append_picture(
        self,
        image_bytes: bytes | None = None,
        extension: str | None = None,
        *,
        width: int = DEFAULT_HWP_SHAPE_WIDTH,
        height: int = DEFAULT_HWP_SHAPE_HEIGHT,
        profile_root: str | Path | None = None,
        section_index: int | None = None,
    ) -> None:
        if image_bytes is None:
            self._append_picture_with_storage_id(
                self._resolve_default_picture_storage_id(),
                width=width,
                height=height,
                profile_root=profile_root,
                section_index=section_index,
            )
            return
        self._append_picture_with_bindata(
            image_bytes,
            extension=extension or "png",
            width=width,
            height=height,
            profile_root=profile_root,
            section_index=section_index,
        )

    def append_hyperlink(
        self,
        url: str | None = None,
        *,
        text: str | None = None,
        metadata_fields: Sequence[str | int] | None = None,
        profile_root: str | Path | None = None,
        section_index: int | None = None,
    ) -> None:
        self._append_hyperlink_with_text(
            url or DEFAULT_HWP_HYPERLINK_URL,
            text=text or DEFAULT_HWP_HYPERLINK_TEXT,
            metadata_fields=metadata_fields,
            profile_root=profile_root,
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
        section_index: int | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        from ._hwp_control_payloads import (
            _build_control_id_payload,
            _normalize_field_control_id,
        )
        from ._hwp_record_base import (
            RecordNode,
        )

        normalized_field_type = field_type.strip()
        if normalized_field_type.upper() in {"HYPERLINK", "%HLK"}:
            target = None
            if parameters is not None:
                target = (
                    parameters.get("Path")
                    or parameters.get("Command")
                    or parameters.get("url")
                    or parameters.get("URL")
                )
            self.append_hyperlink(
                str(target or DEFAULT_HWP_HYPERLINK_URL),
                text=display_text,
                section_index=section_index,
            )
            return

        target_section_index = self._resolve_append_section_index(None, section_index)
        display_value = display_text or name or normalized_field_type
        control_payload = _build_control_id_payload(_normalize_field_control_id(normalized_field_type))
        control_node = RecordNode(tag_id=TAG_CTRL_HEADER, level=1, payload=control_payload)
        native_parameters = dict(parameters or {})
        if name and not any(key in native_parameters for key in ("FieldName", "Name", "name")):
            native_parameters["FieldName"] = name
        if native_parameters:
            parameter_text = ";".join(f"{key}={value}" for key, value in native_parameters.items())
            control_node.add_child(RecordNode(tag_id=TAG_CTRL_DATA, level=2, payload=parameter_text.encode("utf-16-le")))

        model = self.section_model(target_section_index)
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(model)
        if paragraph_index is None or paragraph_index < 0:
            paragraph = model.append_paragraph(
                f"{display_value}\r",
                control_mask=1 if editable or dirty else 0,
            )
        else:
            paragraph = model.insert_paragraph(
                paragraph_index,
                f"{display_value}\r",
                control_mask=1 if editable or dirty else 0,
            )
        if use_synthetic_line_numbers:
            _ensure_paragraph_line_seg(paragraph)
        paragraph.header.add_child(control_node)
        self.replace_section_model(target_section_index, model)

    def append_auto_number(
        self,
        *,
        number: int | str = 1,
        number_type: str = "PAGE",
        kind: str = "newNum",
        section_index: int | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        from ._hwp_control_payloads import (
            _build_auto_number_payload,
            _build_control_marker_text,
        )
        from ._hwp_record_base import (
            RecordNode,
        )

        normalized_kind = kind.strip()
        if normalized_kind not in {"newNum", "autoNum"}:
            raise ValueError("append_auto_number() kind must be 'newNum' or 'autoNum'.")

        target_section_index = self._resolve_append_section_index(None, section_index)
        model = self.section_model(target_section_index)
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(model)
        control_id = "nwno" if normalized_kind == "newNum" else "atno"
        raw_text = _build_control_marker_text(control_id) + "\r"
        if paragraph_index is None or paragraph_index < 0:
            paragraph = model.append_paragraph(raw_text)
        else:
            paragraph = model.insert_paragraph(paragraph_index, raw_text)
        if use_synthetic_line_numbers:
            _ensure_paragraph_line_seg(paragraph)
        control_node = RecordNode(
            tag_id=TAG_CTRL_HEADER,
            level=1,
            payload=_build_auto_number_payload(normalized_kind, number=number, number_type=number_type),
        )
        paragraph.header.add_child(control_node)
        self.replace_section_model(target_section_index, model)

    def append_header(
        self,
        text: str,
        *,
        apply_page_type: str = "BOTH",
        section_index: int | None = None,
    ) -> None:
        self._upsert_header_footer_control("head", text=text, apply_page_type=apply_page_type, section_index=section_index)

    def append_footer(
        self,
        text: str,
        *,
        apply_page_type: str = "BOTH",
        section_index: int | None = None,
    ) -> None:
        self._upsert_header_footer_control("foot", text=text, apply_page_type=apply_page_type, section_index=section_index)

    def append_bookmark(
        self,
        name: str,
        *,
        section_index: int | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        from ._hwp_control_payloads import (
            _build_bookmark_ctrl_data_payload,
            _build_control_id_payload,
            _build_control_marker_text,
        )
        from ._hwp_record_base import (
            RecordNode,
        )

        target_section_index = self._resolve_append_section_index(None, section_index)
        model = self.section_model(target_section_index)
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(model)
        raw_text = _build_control_marker_text("bokm") + "\r"
        if paragraph_index is None or paragraph_index < 0:
            paragraph = model.append_paragraph(raw_text)
        else:
            paragraph = model.insert_paragraph(paragraph_index, raw_text)
        if use_synthetic_line_numbers:
            _ensure_paragraph_line_seg(paragraph)
        control_node = RecordNode(tag_id=TAG_CTRL_HEADER, level=1, payload=_build_control_id_payload("bokm"))
        control_node.add_child(RecordNode(tag_id=TAG_CTRL_DATA, level=2, payload=_build_bookmark_ctrl_data_payload(name)))
        paragraph.header.add_child(control_node)
        self.replace_section_model(target_section_index, model)

    def append_note(
        self,
        text: str,
        *,
        kind: str = "footNote",
        number: int | str | None = None,
        section_index: int | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        from ._hwp_control_payloads import (
            _build_control_marker_text,
        )

        normalized_kind = kind.strip()
        if normalized_kind not in {"footNote", "endNote"}:
            raise ValueError("append_note() kind must be 'footNote' or 'endNote'.")

        target_section_index = self._resolve_append_section_index(None, section_index)
        model = self.section_model(target_section_index)
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(model)
        control_id = "fn  " if normalized_kind == "footNote" else "en  "
        raw_text = _build_control_marker_text(control_id) + "\r"
        if paragraph_index is None or paragraph_index < 0:
            paragraph = model.append_paragraph(raw_text)
        else:
            paragraph = model.insert_paragraph(paragraph_index, raw_text)
        if use_synthetic_line_numbers:
            _ensure_paragraph_line_seg(paragraph)
        paragraph.header.add_child(_build_note_control_node(control_id, text))
        model._paragraphs_cache = None
        self.replace_section_model(target_section_index, model)

    def append_footnote(
        self,
        text: str,
        *,
        number: int | str | None = None,
        section_index: int | None = None,
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
        section_index: int | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        self.append_note(
            text,
            kind="endNote",
            number=number,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )

    def append_form(
        self,
        label: str = "",
        *,
        form_type: str = "INPUT",
        name: str | None = None,
        value: str | None = None,
        checked: bool = False,
        items: Sequence[str] | None = None,
        editable: bool = True,
        locked: bool = False,
        placeholder: str | None = None,
        section_index: int | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        from ._hwp_control_payloads import (
            _build_control_id_payload,
        )
        from ._hwp_record_base import (
            RecordNode,
        )

        target_section_index = self._resolve_append_section_index(None, section_index)
        model = self.section_model(target_section_index)
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(model)
        if paragraph_index is None or paragraph_index < 0:
            paragraph = model.append_paragraph("\r")
        else:
            paragraph = model.insert_paragraph(paragraph_index, "\r")
        if use_synthetic_line_numbers:
            _ensure_paragraph_line_seg(paragraph)
        control_node = RecordNode(tag_id=TAG_CTRL_HEADER, level=1, payload=_build_control_id_payload("mrof"))
        semantic_metadata = {
            "form.type": str(form_type).upper(),
            "form.name": name,
            "form.value": value,
            "form.checked": checked,
            "form.editable": editable,
            "form.locked": locked,
            "form.placeholder": placeholder,
            "form.label": label,
        }
        if items:
            semantic_metadata["form.items"] = json.dumps(list(items), ensure_ascii=True, separators=(",", ":"))
        if any(field is not None for field in semantic_metadata.values()):
            control_node.add_child(
                RecordNode(
                    tag_id=TAG_CTRL_DATA,
                    level=2,
                    payload=_build_control_native_metadata_payload(**semantic_metadata),
                )
            )
        control_node.add_child(RecordNode(tag_id=TAG_FORM_OBJECT, level=2, payload=str(label).encode("utf-16-le")))
        paragraph.header.add_child(control_node)
        self.replace_section_model(target_section_index, model)

    def append_memo(
        self,
        text: str,
        *,
        author: str | None = None,
        memo_id: str | None = None,
        anchor_id: str | None = None,
        order: int | None = None,
        visible: bool = True,
        section_index: int | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        from ._hwp_control_payloads import (
            _build_control_id_payload,
        )
        from ._hwp_record_base import (
            RecordNode,
        )

        target_section_index = self._resolve_append_section_index(None, section_index)
        model = self.section_model(target_section_index)
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(model)
        if paragraph_index is None or paragraph_index < 0:
            paragraph = model.append_paragraph("\r")
        else:
            paragraph = model.insert_paragraph(paragraph_index, "\r")
        if use_synthetic_line_numbers:
            _ensure_paragraph_line_seg(paragraph)
        control_node = RecordNode(tag_id=TAG_CTRL_HEADER, level=1, payload=_build_control_id_payload("omem"))
        semantic_metadata = {
            "memo.author": author,
            "memo.memoId": memo_id,
            "memo.anchorId": anchor_id,
            "memo.order": order,
            "memo.visible": visible,
        }
        if any(field is not None for field in semantic_metadata.values()):
            control_node.add_child(
                RecordNode(
                    tag_id=TAG_CTRL_DATA,
                    level=2,
                    payload=_build_control_native_metadata_payload(**semantic_metadata),
                )
            )
        control_node.add_child(RecordNode(tag_id=TAG_MEMO_LIST, level=2, payload=str(text).encode("utf-16-le")))
        paragraph.header.add_child(control_node)
        self.replace_section_model(target_section_index, model)

    def append_chart(
        self,
        title: str = "",
        *,
        chart_type: str = "BAR",
        categories: Sequence[str] | None = None,
        series: Sequence[dict[str, object]] | None = None,
        data_ref: str | None = None,
        legend_visible: bool = True,
        width: int = DEFAULT_HWP_SHAPE_WIDTH,
        height: int = DEFAULT_HWP_SHAPE_HEIGHT,
        shape_comment: str | None = None,
        section_index: int | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        from ._hwp_record_base import (
            RecordNode,
        )
        from ._hwp_section_helpers import (
            _insert_blank_like_paragraph,
        )

        target_section_index = self._resolve_append_section_index(None, section_index)
        model = self.section_model(target_section_index)
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(model)
        paragraph = _insert_blank_like_paragraph(model, paragraph_index)
        if use_synthetic_line_numbers:
            _ensure_paragraph_line_seg(paragraph)
        _append_control_marker_to_paragraph(paragraph, "gso ")
        control_node = RecordNode(
            tag_id=TAG_CTRL_HEADER,
            level=1,
            payload=_build_graphic_control_header_payload(
                _shape_control_header_template("rect"),
                "gso ",
                width=width,
                height=height,
                description=shape_comment,
            ),
        )
        semantic_metadata = {
            "chart.title": title,
            "chart.chartType": str(chart_type).upper(),
            "chart.dataRef": data_ref,
            "chart.legendVisible": legend_visible,
        }
        if categories:
            semantic_metadata["chart.categories"] = json.dumps(list(categories), ensure_ascii=True, separators=(",", ":"))
        if series:
            semantic_metadata["chart.series"] = json.dumps(list(series), ensure_ascii=True, separators=(",", ":"))
        control_node.add_child(
            RecordNode(
                tag_id=TAG_CTRL_DATA,
                level=2,
                payload=_build_shape_native_metadata_payload(**semantic_metadata),
            )
        )
        shape_component = RecordNode(
            tag_id=TAG_SHAPE_COMPONENT,
            level=2,
            payload=_build_shape_component_payload("rect", width, height),
        )
        shape_component.add_child(RecordNode(tag_id=TAG_CHART_DATA, level=3, payload=str(title).encode("utf-16-le")))
        control_node.add_child(shape_component)
        paragraph.header.add_child(control_node)
        self.replace_section_model(target_section_index, model)

    def append_equation(
        self,
        script: str = DEFAULT_HWP_EQUATION_SCRIPT,
        *,
        width: int = 4800,
        height: int = 2300,
        font: str = DEFAULT_HWP_EQUATION_FONT,
        shape_comment: str | None = None,
        section_index: int | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        from ._hwp_control_payloads import (
            _build_equation_payload,
        )
        from ._hwp_record_base import (
            RecordNode,
        )

        target_section_index = self._resolve_append_section_index(None, section_index)
        model = self.section_model(target_section_index)
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(model)
        if paragraph_index is None or paragraph_index < 0:
            paragraph = model.append_paragraph("\r")
        else:
            paragraph = model.insert_paragraph(paragraph_index, "\r")
        if use_synthetic_line_numbers:
            _ensure_paragraph_line_seg(paragraph)
        control_node = RecordNode(
            tag_id=TAG_CTRL_HEADER,
            level=1,
            payload=_build_graphic_control_header_payload(
                _DEFAULT_EQUATION_CTRL_HEADER,
                "eqed",
                width=width,
                height=height,
                description=shape_comment,
            ),
        )
        control_node.add_child(RecordNode(tag_id=TAG_EQEDIT, level=2, payload=_build_equation_payload(script, font=font)))
        paragraph.header.add_child(control_node)
        self.replace_section_model(target_section_index, model)

    def append_shape(
        self,
        *,
        kind: str = "rect",
        text: str = DEFAULT_HWP_SHAPE_TEXT,
        width: int = DEFAULT_HWP_SHAPE_WIDTH,
        height: int = DEFAULT_HWP_SHAPE_HEIGHT,
        fill_color: str = "#FFFFFF",
        line_color: str = "#000000",
        shape_comment: str | None = None,
        section_index: int | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        from ._hwp_record_base import (
            RecordNode,
        )
        from ._hwp_section_helpers import (
            _build_shape_text_records,
            _insert_blank_like_paragraph,
        )

        kind_to_tag = {
            "line": TAG_SHAPE_COMPONENT_LINE,
            "connectLine": TAG_SHAPE_COMPONENT_LINE,
            "rect": TAG_SHAPE_COMPONENT_RECTANGLE,
            "ellipse": TAG_SHAPE_COMPONENT_ELLIPSE,
            "arc": TAG_SHAPE_COMPONENT_ARC,
            "polygon": TAG_SHAPE_COMPONENT_POLYGON,
            "curve": TAG_SHAPE_COMPONENT_CURVE,
            "container": TAG_SHAPE_COMPONENT_CONTAINER,
            "textart": TAG_SHAPE_COMPONENT_TEXTART,
        }
        if kind not in kind_to_tag:
            supported = ", ".join(sorted(kind_to_tag))
            raise ValueError(f"append_shape() supports kind values: {supported}")

        target_section_index = self._resolve_append_section_index(None, section_index)
        model = self.section_model(target_section_index)
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(model)
        paragraph = _insert_blank_like_paragraph(model, paragraph_index)
        if use_synthetic_line_numbers:
            _ensure_paragraph_line_seg(paragraph)
        _append_control_marker_to_paragraph(paragraph, "gso ")
        control_node = RecordNode(
            tag_id=TAG_CTRL_HEADER,
            level=1,
            payload=_build_graphic_control_header_payload(
                _shape_control_header_template(kind),
                "gso ",
                width=width,
                height=height,
                description=shape_comment,
            ),
        )
        metadata_fields: dict[str, str] = {}
        if kind == "connectLine":
            metadata_fields["kind"] = "connectLine"
        normalized_fill_color = str(fill_color).strip().upper()
        if kind in {"ellipse", "arc", "polygon"} and normalized_fill_color and normalized_fill_color != "#FFFFFF":
            metadata_fields["fill_color"] = fill_color
        if kind in {"ellipse", "arc", "polygon"} and text:
            metadata_fields["text"] = text
        if metadata_fields:
            control_node.add_child(
                RecordNode(
                    tag_id=TAG_CTRL_DATA,
                    level=2,
                    payload=_build_shape_native_metadata_payload(**metadata_fields),
                )
            )
        shape_component = RecordNode(
            tag_id=TAG_SHAPE_COMPONENT,
            level=2,
            payload=_build_shape_component_payload(
                kind,
                width,
                height,
                fill_color=fill_color,
                line_color=line_color,
            ),
        )
        if text and kind not in {"line", "ellipse", "arc", "polygon", "textart"}:
            list_header, paragraph_header = _build_shape_text_records(text)
            shape_component.add_child(list_header)
            shape_component.add_child(paragraph_header)
            model._paragraphs_cache = None
        shape_component.add_child(
            RecordNode(
                tag_id=kind_to_tag[kind],
                level=3,
                payload=_build_shape_specific_payload(
                    kind,
                    text=text,
                    width=width,
                    height=height,
                    fill_color=fill_color,
                    line_color=line_color,
                ),
            )
        )
        control_node.add_child(shape_component)
        paragraph.header.add_child(control_node)
        self.replace_section_model(target_section_index, model)

    def append_ole(
        self,
        name: str,
        data: bytes,
        *,
        width: int = DEFAULT_HWP_SHAPE_WIDTH,
        height: int = DEFAULT_HWP_SHAPE_HEIGHT,
        shape_comment: str | None = None,
        object_type: str = "EMBEDDED",
        draw_aspect: str = "CONTENT",
        has_moniker: bool = False,
        eq_baseline: int = 0,
        line_color: str = "#000000",
        line_width: int = _DEFAULT_OLE_LINE_WIDTH,
        section_index: int | None = None,
        paragraph_index: int | None = None,
    ) -> None:
        from ._hwp_record_base import (
            RecordNode,
        )
        from ._hwp_section_helpers import (
            _build_shape_text_records,
        )

        target_section_index = self._resolve_append_section_index(None, section_index)
        storage_id, _ = self.add_embedded_bindata(data, extension="ole")
        model = self.section_model(target_section_index)
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(model)
        paragraph_text = f"{name}\r" if name else "\r"
        if paragraph_index is None or paragraph_index < 0:
            paragraph = model.append_paragraph(paragraph_text)
        else:
            paragraph = model.insert_paragraph(paragraph_index, paragraph_text)
        if use_synthetic_line_numbers:
            _ensure_paragraph_line_seg(paragraph)
        control_node = RecordNode(
            tag_id=TAG_CTRL_HEADER,
            level=1,
            payload=_build_graphic_control_header_payload(
                _shape_control_header_template("ole"),
                "gso ",
                width=width,
                height=height,
                description=shape_comment,
            ),
        )
        shape_component = RecordNode(
            tag_id=TAG_SHAPE_COMPONENT,
            level=2,
            payload=_build_shape_component_payload("ole", width, height),
        )
        if name:
            list_header, paragraph_header = _build_shape_text_records(name)
            shape_component.add_child(list_header)
            shape_component.add_child(paragraph_header)
        shape_component.add_child(
            RecordNode(
                tag_id=TAG_SHAPE_COMPONENT_OLE,
                level=3,
                payload=_build_ole_specific_payload(
                    storage_id=storage_id,
                    width=width,
                    height=height,
                    object_type=object_type,
                    draw_aspect=draw_aspect,
                    has_moniker=has_moniker,
                    eq_baseline=eq_baseline,
                    line_color=line_color,
                    line_width=line_width,
                ),
            )
        )
        control_node.add_child(shape_component)
        paragraph.header.add_child(control_node)
        self.replace_section_model(target_section_index, model)

    def _append_feature_via_profile(
        self,
        feature: str,
        *,
        profile_root: str | Path | None,
        section_index: int | None,
    ) -> None:
        from .hwp_pure_profile import HwpPureProfile, append_feature_from_profile

        profile = HwpPureProfile.load_bundled() if profile_root is None else HwpPureProfile.load(profile_root)
        if section_index is not None and section_index != profile.target_section_index:
            raise HwpBinaryEditError(
                f"Pure profile feature templates currently target section {profile.target_section_index}, "
                f"but section {section_index} was requested."
        )
        append_feature_from_profile(self, profile, feature)

    def _resolve_append_section_index(self, profile_root: str | Path | None, section_index: int | None) -> int:
        if section_index is not None:
            self.ensure_section_count(section_index + 1)
            return section_index
        if profile_root is None:
            from .hwp_pure_profile import HwpPureProfile

            target_section_index = HwpPureProfile.load_bundled().target_section_index
            self.ensure_section_count(target_section_index + 1)
            return target_section_index
        profile_path = Path(profile_root).expanduser().resolve()
        if not profile_path.exists():
            self.ensure_section_count(1)
            return 0
        metadata_path = profile_path / "profile.json"
        if not metadata_path.exists():
            self.ensure_section_count(1)
            return 0
        from .hwp_pure_profile import HwpPureProfile

        target_section_index = HwpPureProfile.load(profile_path).target_section_index
        self.ensure_section_count(target_section_index + 1)
        return target_section_index

    def _upsert_header_footer_control(
        self,
        control_id: str,
        *,
        text: str,
        apply_page_type: str,
        section_index: int | None,
    ) -> None:
        from ._hwp_control_payloads import (
            _build_header_footer_control_payload,
            _decode_control_id_payload,
        )

        target_section_index = self._resolve_append_section_index(None, section_index)
        model = self.section_model(target_section_index)
        paragraphs = model.paragraphs()
        if not paragraphs:
            model.append_paragraph("\r")
            paragraphs = model.paragraphs()
        target_paragraph = paragraphs[0]
        control_node = next(
            (
                child
                for child in target_paragraph.header.children
                if child.tag_id == TAG_CTRL_HEADER and _decode_control_id_payload(child.payload) == control_id
            ),
            None,
        )
        if control_node is None:
            instance_id = 1 + sum(
                1
                for child in target_paragraph.header.children
                if child.tag_id == TAG_CTRL_HEADER
                and _decode_control_id_payload(child.payload) in {"head", "foot"}
            )
            control_node = _build_header_footer_control_node(
                control_id,
                apply_page_type=apply_page_type,
                instance_id=instance_id,
            )
            target_paragraph.header.add_child(control_node)
            _append_control_marker_to_paragraph(target_paragraph, control_id)
        else:
            instance_id = int.from_bytes(control_node.payload[8:12].ljust(4, b"\x00"), "little") or 1
            control_node.payload = _build_header_footer_control_payload(
                control_id,
                apply_page_type=apply_page_type,
                instance_id=instance_id,
            )
        _set_header_footer_control_text(control_node, text)
        self.replace_section_model(target_section_index, model)

    def add_embedded_bindata(
        self,
        data: bytes,
        *,
        extension: str,
        storage_id: int | None = None,
        flags: int = 0x0001,
    ) -> tuple[int, str]:
        model = self.docinfo_model()
        bindata = model.add_embedded_bindata(extension, storage_id=storage_id, flags=flags)
        self.replace_docinfo_model(model)
        stream_path = f"BinData/BIN{int(bindata.storage_id or 0):04d}.{(bindata.extension or extension).lstrip('.').lower()}"
        self.add_stream(stream_path, data)
        return int(bindata.storage_id or 0), stream_path

    def remove_embedded_bindata(self, storage_id: int) -> bool:
        model = self.docinfo_model()
        removed = model.remove_bindata(storage_id)
        if not removed:
            return False
        self.replace_docinfo_model(model)
        prefix = f"BinData/BIN{storage_id:04d}."
        for stream_path in list(self._streams):
            if stream_path.startswith(prefix):
                self.remove_stream(stream_path)
        return True

    def _resolve_default_picture_storage_id(self) -> int:
        for stream_path in self.bindata_stream_paths():
            name = Path(stream_path).name
            if not name.startswith("BIN") or "." not in name:
                continue
            storage_token = name[3:].split(".", 1)[0]
            if storage_token.isdigit():
                return int(storage_token)
        raise HwpBinaryEditError(
            "append_picture() without image_bytes requires at least one existing BinData stream in the document."
        )

    def _build_picture_records(self, storage_id: int, *, width: int, height: int) -> list[HwpRecord]:
        from ._hwp_record_base import (
            HwpRecord,
        )

        top_char_count = len(_PICTURE_TOP_PARA_TEXT) // 2
        picture_payload = bytearray(_PICTURE_SHAPE_PICTURE)
        if len(picture_payload) <= 72:
            raise HwpBinaryEditError("Built picture payload is shorter than expected.")
        picture_payload[71:73] = int(storage_id).to_bytes(2, "little", signed=False)
        shape_component_payload = bytearray(_PICTURE_SHAPE_COMPONENT)
        shape_component_payload[20:24] = int(width).to_bytes(4, "little", signed=False)
        shape_component_payload[24:28] = int(height).to_bytes(4, "little", signed=False)
        return [
            HwpRecord(
                tag_id=TAG_PARA_HEADER,
                level=0,
                size=24,
                header_size=4,
                offset=-1,
                payload=(0x80000000 | top_char_count).to_bytes(4, "little") + _PICTURE_TOP_PARA_HEADER_TAIL,
            ),
            HwpRecord(tag_id=TAG_PARA_TEXT, level=1, size=len(_PICTURE_TOP_PARA_TEXT), header_size=4, offset=-1, payload=_PICTURE_TOP_PARA_TEXT),
            HwpRecord(tag_id=TAG_PARA_CHAR_SHAPE, level=1, size=len(_PICTURE_TOP_CHAR_SHAPE), header_size=4, offset=-1, payload=_PICTURE_TOP_CHAR_SHAPE),
            HwpRecord(tag_id=TAG_PARA_LINE_SEG, level=1, size=len(_PICTURE_TOP_LINE_SEG), header_size=4, offset=-1, payload=_PICTURE_TOP_LINE_SEG),
            HwpRecord(tag_id=TAG_CTRL_HEADER, level=1, size=len(_PICTURE_CTRL_HEADER), header_size=4, offset=-1, payload=_PICTURE_CTRL_HEADER),
            HwpRecord(
                tag_id=TAG_SHAPE_COMPONENT,
                level=2,
                size=len(_PICTURE_SHAPE_COMPONENT),
                header_size=4,
                offset=-1,
                payload=bytes(shape_component_payload),
            ),
            HwpRecord(
                tag_id=TAG_SHAPE_COMPONENT_PICTURE,
                level=3,
                size=len(picture_payload),
                header_size=4,
                offset=-1,
                payload=bytes(picture_payload),
            ),
        ]

    def _append_picture_with_storage_id(
        self,
        storage_id: int,
        *,
        width: int,
        height: int,
        profile_root: str | Path | None,
        section_index: int | None,
    ) -> None:
        target_section_index = self._resolve_append_section_index(profile_root, section_index)
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(self.section_model(target_section_index))
        section_records = self.section_records(target_section_index)
        section_records.extend(self._build_picture_records(storage_id, width=width, height=height))
        self.replace_section_records(target_section_index, section_records)
        if use_synthetic_line_numbers:
            model = self.section_model(target_section_index)
            _apply_section_show_line_numbers(model, True)
            self.replace_section_model(target_section_index, model)

    def _append_picture_with_bindata(
        self,
        image_bytes: bytes,
        *,
        extension: str,
        width: int,
        height: int,
        profile_root: str | Path | None,
        section_index: int | None,
    ) -> None:
        storage_id, _stream_path = self.add_embedded_bindata(image_bytes, extension=extension)
        self._append_picture_with_storage_id(
            storage_id,
            width=width,
            height=height,
            profile_root=profile_root,
            section_index=section_index,
        )

    def _append_table_with_text(
        self,
        cell_text: str | None,
        *,
        rows: int,
        cols: int,
        cell_texts: Sequence[str] | Sequence[Sequence[str]] | None,
        row_heights: Sequence[int] | None,
        col_widths: Sequence[int] | None,
        cell_spans: dict[tuple[int, int], tuple[int, int]] | None,
        cell_border_fill_ids: dict[tuple[int, int], int] | None,
        table_border_fill_id: int,
        profile_root: str | Path | None,
        section_index: int | None,
    ) -> None:
        from ._hwp_record_base import (
            HwpRecord,
        )
        from ._hwp_table_payloads import (
            _build_table_cell_list_header_payload,
            _build_table_cell_specs,
            _build_table_record_payload,
            _normalize_table_cell_texts,
            _normalize_table_measurements,
        )

        target_section_index = self._resolve_append_section_index(profile_root, section_index)
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(self.section_model(target_section_index))
        matrix = _normalize_table_cell_texts(rows, cols, cell_text, cell_texts)
        normalized_row_heights = _normalize_table_measurements(
            rows,
            row_heights,
            default=DEFAULT_HWP_TABLE_CELL_HEIGHT,
            name="row_heights",
        )
        normalized_col_widths = _normalize_table_measurements(
            cols,
            col_widths,
            default=DEFAULT_HWP_TABLE_CELL_WIDTH,
            name="col_widths",
        )
        cell_specs = _build_table_cell_specs(
            matrix,
            row_heights=normalized_row_heights,
            col_widths=normalized_col_widths,
            cell_spans=cell_spans,
            cell_border_fill_ids=cell_border_fill_ids,
            default_border_fill_id=table_border_fill_id,
            cell_margins=None,
        )
        top_char_count = len(_TABLE_TOP_PARA_TEXT) // 2
        records_to_append = [
            HwpRecord(
                tag_id=TAG_PARA_HEADER,
                level=0,
                size=24,
                header_size=4,
                offset=-1,
                payload=top_char_count.to_bytes(4, "little") + _TABLE_TOP_PARA_HEADER_TAIL,
            ),
            HwpRecord(tag_id=TAG_PARA_TEXT, level=1, size=len(_TABLE_TOP_PARA_TEXT), header_size=4, offset=-1, payload=_TABLE_TOP_PARA_TEXT),
            HwpRecord(tag_id=TAG_PARA_CHAR_SHAPE, level=1, size=len(_TABLE_TOP_CHAR_SHAPE), header_size=4, offset=-1, payload=_TABLE_TOP_CHAR_SHAPE),
            HwpRecord(tag_id=TAG_PARA_LINE_SEG, level=1, size=len(_TABLE_TOP_LINE_SEG), header_size=4, offset=-1, payload=_TABLE_TOP_LINE_SEG),
            HwpRecord(tag_id=TAG_CTRL_HEADER, level=1, size=len(_TABLE_CTRL_HEADER), header_size=4, offset=-1, payload=_TABLE_CTRL_HEADER),
        ]
        table_payload = _build_table_record_payload(
            rows,
            cols,
            default_row_height=1,
            border_fill_id=table_border_fill_id,
        )
        table_payload = bytearray(table_payload)
        row_sizes_offset = 18
        for row_index, row_height in enumerate(normalized_row_heights):
            offset = row_sizes_offset + row_index * 2
            table_payload[offset : offset + 2] = int(row_height).to_bytes(2, "little", signed=False)
        records_to_append.append(HwpRecord(tag_id=TAG_TABLE, level=2, size=len(table_payload), header_size=4, offset=-1, payload=bytes(table_payload)))
        for cell in cell_specs:
            cell_payload = _build_table_cell_list_header_payload(cell)
            records_to_append.append(
                HwpRecord(tag_id=TAG_LIST_HEADER, level=2, size=len(cell_payload), header_size=4, offset=-1, payload=cell_payload)
            )
            cell_raw_text = str(cell.text) + "\r"
            cell_char_count = len(cell_raw_text)
            records_to_append.append(
                HwpRecord(
                    tag_id=TAG_PARA_HEADER,
                    level=2,
                    size=24,
                    header_size=4,
                    offset=-1,
                    payload=(0x80000000 | cell_char_count).to_bytes(4, "little") + _TABLE_CELL_PARA_HEADER_TAIL,
                )
            )
            records_to_append.append(
                HwpRecord(
                    tag_id=TAG_PARA_TEXT,
                    level=3,
                    size=len(cell_raw_text.encode("utf-16-le")),
                    header_size=4,
                    offset=-1,
                    payload=cell_raw_text.encode("utf-16-le"),
                )
            )
            records_to_append.append(
                HwpRecord(
                    tag_id=TAG_PARA_CHAR_SHAPE,
                    level=3,
                    size=len(_TABLE_CELL_CHAR_SHAPE),
                    header_size=4,
                    offset=-1,
                    payload=_TABLE_CELL_CHAR_SHAPE,
                )
            )
            records_to_append.append(
                HwpRecord(
                    tag_id=TAG_PARA_LINE_SEG,
                    level=3,
                    size=len(_TABLE_CELL_LINE_SEG),
                    header_size=4,
                    offset=-1,
                    payload=_TABLE_CELL_LINE_SEG,
                )
            )
        section_records = self.section_records(target_section_index)
        section_records.extend(records_to_append)
        self.replace_section_records(target_section_index, section_records)
        if use_synthetic_line_numbers:
            model = self.section_model(target_section_index)
            _apply_section_show_line_numbers(model, True)
            self.replace_section_model(target_section_index, model)

    def _append_hyperlink_with_text(
        self,
        url: str,
        *,
        text: str | None,
        metadata_fields: Sequence[str | int] | None,
        profile_root: str | Path | None,
        section_index: int | None,
    ) -> None:
        from ._hwp_control_payloads import (
            _build_hyperlink_ctrl_payload,
            _build_hyperlink_para_text_payload,
        )
        from ._hwp_record_base import (
            HwpRecord,
        )

        target_section_index = self._resolve_append_section_index(profile_root, section_index)
        use_synthetic_line_numbers = _section_uses_synthetic_line_number_layout(self.section_model(target_section_index))
        display_text = text or url
        raw_payload = _build_hyperlink_para_text_payload(display_text)
        char_count = len(raw_payload) // 2
        control_payload = _build_hyperlink_ctrl_payload(url, metadata_fields=metadata_fields)

        records_to_append = [
            HwpRecord(
                tag_id=TAG_PARA_HEADER,
                level=0,
                size=22,
                header_size=4,
                offset=-1,
                payload=(0x80000000 | char_count).to_bytes(4, "little") + _HYPERLINK_PARA_HEADER_TAIL,
            ),
            HwpRecord(tag_id=TAG_PARA_TEXT, level=1, size=len(raw_payload), header_size=4, offset=-1, payload=raw_payload),
            HwpRecord(tag_id=TAG_PARA_CHAR_SHAPE, level=1, size=len(_HYPERLINK_CHAR_SHAPE), header_size=4, offset=-1, payload=_HYPERLINK_CHAR_SHAPE),
            HwpRecord(tag_id=TAG_PARA_LINE_SEG, level=1, size=len(_HYPERLINK_LINE_SEG), header_size=4, offset=-1, payload=_HYPERLINK_LINE_SEG),
            HwpRecord(tag_id=TAG_CTRL_HEADER, level=1, size=len(control_payload), header_size=4, offset=-1, payload=control_payload),
        ]
        section_records = self.section_records(target_section_index)
        section_records.extend(records_to_append)
        self.replace_section_records(target_section_index, section_records)
        if use_synthetic_line_numbers:
            model = self.section_model(target_section_index)
            _apply_section_show_line_numbers(model, True)
            self.replace_section_model(target_section_index, model)
