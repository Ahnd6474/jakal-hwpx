"""Strict native HWP validation of DocInfo, binary references, and controls."""
from __future__ import annotations

from pathlib import (
    Path,
)
from typing import TYPE_CHECKING

from ._hwp_object_payloads import (
    _collect_control_native_metadata,
    _collect_shape_native_metadata,
    _metadata_bool_literal_is_valid,
    _metadata_json_value,
)
from .exceptions import (
    ValidationIssue,
)
from .hwp_binary import (
    TAG_CHART_DATA,
    TAG_CTRL_DATA,
    TAG_FORM_OBJECT,
    TAG_MEMO_LIST,
    TAG_SHAPE_COMPONENT,
    TAG_SHAPE_COMPONENT_OLE,
    TAG_SHAPE_COMPONENT_PICTURE,
    TAG_TABLE,
    ParagraphHeaderRecord,
)

if TYPE_CHECKING:
    from .hwp_document import HwpDocument


_HWP_STRICT_LINT_HINTS = {
    "binary_closure": "Keep DocInfo BinData records and BinData/BINxxxx streams in sync. Re-add the picture/OLE data or remove the broken control.",
    "binary_media_type": "Use replace_binary(..., extension=...) or rebuild the BinData entry so the DocInfo extension matches the stream name.",
    "binary_reference": "Re-add the missing embedded binary or update the picture/OLE control to point at an existing BinData id.",
    "control_subtree": "Recreate the object with the public append/setter API, or restore the missing child record before saving.",
    "docinfo_mapping": "Repair DocInfo id mappings before saving. This usually means recreating the referenced memo/BinData entry through the HWP API.",
    "docinfo_reference": "Repair the paragraph header or restore the missing DocInfo record before saving.",
    "metadata_schema": "Use the semantic setter API instead of editing JAKAL metadata text directly; values must keep their documented type.",
    "payload_sanity": "The native HWP payload is truncated or malformed. Recreate that control or replace it from a known-good source document.",
}


def _issue(
    kind: str,
    message: str,
    *,
    section_index: int | None = None,
    paragraph_index: int | None = None,
    context: str | None = None,
    hint: str | None = None,
) -> ValidationIssue:
    return ValidationIssue(
        kind=kind,
        message=message,
        section_index=section_index,
        paragraph_index=paragraph_index,
        context=context,
        hint=hint or _HWP_STRICT_LINT_HINTS.get(kind),
    )


def collect_hwp_validation_issues(document: HwpDocument) -> list[ValidationIssue]:
    from ._hwp_objects import _control_has_descendant, _control_id

    errors: list[ValidationIssue] = []
    docinfo = document.docinfo_model()
    docinfo_bindata: dict[int, object] = {}
    for record in docinfo.bin_data_records():
        storage_id = int(record.storage_id or 0)
        if storage_id <= 0:
            continue
        if storage_id in docinfo_bindata:
            errors.append(_issue("docinfo_mapping", f"Duplicate BinData docinfo storage id: {storage_id}", context="DocInfo"))
            continue
        docinfo_bindata[storage_id] = record

    stream_paths = document.bindata_stream_paths()
    bindata_streams: dict[int, str] = {}
    for stream_path in stream_paths:
        name = Path(stream_path).name
        if not name.startswith("BIN") or "." not in name:
            continue
        storage_token = name[3:].split(".", 1)[0]
        if not storage_token.isdigit():
            continue
        storage_id = int(storage_token)
        if storage_id in bindata_streams:
            errors.append(_issue("binary_closure", f"Duplicate BinData stream for storage id {storage_id}: {stream_path}", context="BinData"))
            continue
        bindata_streams[storage_id] = stream_path

    for storage_id, record in docinfo_bindata.items():
        stream_path = bindata_streams.get(storage_id)
        if stream_path is None:
            errors.append(_issue("binary_closure", f"DocInfo BinData id {storage_id} is missing a BinData stream.", context="DocInfo"))
            continue
        expected_extension = str(getattr(record, "extension", "") or "").lstrip(".").lower()
        actual_extension = Path(stream_path).suffix.lstrip(".").lower()
        if expected_extension and actual_extension and expected_extension != actual_extension:
            errors.append(
                _issue(
                    "binary_media_type",
                    f"DocInfo BinData id {storage_id} expects .{expected_extension} but stream path is {stream_path}.",
                    context="BinData",
                )
            )

    for storage_id, stream_path in bindata_streams.items():
        if storage_id not in docinfo_bindata:
            errors.append(
                _issue(
                    "binary_closure",
                    f"BinData stream is not declared in DocInfo: {stream_path}",
                    context="BinData",
                )
            )

    id_mappings = docinfo.id_mappings_record()
    count_specs = [
        (8, "border_fills", len(docinfo.border_fill_records())),
        (9, "char_shapes", len(docinfo.char_shape_records())),
        (10, "tab_defs", len(docinfo.tab_def_records())),
        (11, "numberings", len(docinfo.numbering_records())),
        (12, "bullets", len(docinfo.bullet_records())),
        (13, "para_shapes", len(docinfo.para_shape_records())),
        (14, "styles", len(docinfo.style_records())),
        (15, "memo_shapes", len(docinfo.memo_shape_records())),
        (16, "track_changes", len(docinfo.track_change_records())),
        (17, "track_change_authors", len(docinfo.track_change_author_records())),
    ]
    for count_index, count_name, actual_count in count_specs:
        declared_count = id_mappings.get_count(count_index)
        if declared_count < actual_count:
            errors.append(
                _issue(
                    "docinfo_mapping",
                    f"DocInfo id_mappings.{count_name}={declared_count} is smaller than actual {count_name} record count {actual_count}.",
                    context="DocInfo",
                )
            )

    max_bindata_id = max([0, *docinfo_bindata.keys(), *bindata_streams.keys()])
    if id_mappings.bin_data_count < max_bindata_id:
        errors.append(
            _issue(
                "docinfo_mapping",
                f"DocInfo id_mappings.bin_data_count={id_mappings.bin_data_count} is smaller than max BinData id {max_bindata_id}.",
                context="DocInfo",
            )
        )

    memo_shape_records = docinfo.memo_shape_records()
    memo_shape_count = id_mappings.get_count(15)
    if memo_shape_count < len(memo_shape_records):
        errors.append(
            _issue(
                "docinfo_mapping",
                f"DocInfo id_mappings.memo_shapes={memo_shape_count} is smaller than actual memo shape record count {len(memo_shape_records)}.",
                context="DocInfo",
            )
        )

    para_shape_count = len(docinfo.para_shape_records())
    style_count = len(docinfo.style_records())
    for section_index, section in enumerate(document.sections()):
        section_model = document.section_model(section_index)
        memo_shape_id = document.binary_document().section_definition_settings(section_index).get("memo_shape_id")
        if memo_shape_id not in (None, "", "0"):
            if not str(memo_shape_id).isdigit():
                errors.append(
                    _issue(
                        "docinfo_mapping",
                        f"Section memo_shape_id must be numeric, got {memo_shape_id!r}.",
                        section_index=section_index,
                        context="SectionSettings",
                    )
                )
            elif int(str(memo_shape_id)) >= len(memo_shape_records):
                errors.append(
                    _issue(
                        "docinfo_mapping",
                        f"Section memo_shape_id {memo_shape_id} points to missing TAG_MEMO_SHAPE record.",
                        section_index=section_index,
                        context="SectionSettings",
                    )
                )
        for paragraph in section_model.paragraphs():
            if paragraph.header.para_shape_id >= para_shape_count:
                errors.append(
                    _issue(
                        "docinfo_reference",
                        f"Paragraph references missing para shape id {paragraph.header.para_shape_id}.",
                        section_index=section_index,
                        paragraph_index=paragraph.index,
                        context="ParagraphHeader",
                    )
                )
            if paragraph.header.style_id >= style_count:
                errors.append(
                    _issue(
                        "docinfo_reference",
                        f"Paragraph references missing style id {paragraph.header.style_id}.",
                        section_index=section_index,
                        paragraph_index=paragraph.index,
                        context="ParagraphHeader",
                    )
                )
            for control_node in paragraph.control_nodes():
                control_id = _control_id(control_node)
                context = f"control[{control_id or 'UNKNOWN'}]"
                if control_id == "gso ":
                    has_shape_component = _control_has_descendant(control_node, TAG_SHAPE_COMPONENT)
                    has_picture = _control_has_descendant(control_node, TAG_SHAPE_COMPONENT_PICTURE)
                    has_ole = _control_has_descendant(control_node, TAG_SHAPE_COMPONENT_OLE)
                    has_chart = _control_has_descendant(control_node, TAG_CHART_DATA)
                    shape_metadata = _collect_shape_native_metadata(control_node)
                    expects_chart_data = any(key.startswith("chart.") for key in shape_metadata)
                    if not (has_shape_component or has_chart):
                        errors.append(
                            _issue(
                                "control_subtree",
                                "Graphic control does not contain a shape component or chart data record.",
                                section_index=section_index,
                                paragraph_index=paragraph.index,
                                context=context,
                            )
                        )
                    if expects_chart_data and not has_chart:
                        errors.append(
                            _issue(
                                "control_subtree",
                                "Chart control metadata is present but TAG_CHART_DATA is missing.",
                                section_index=section_index,
                                paragraph_index=paragraph.index,
                                context=context,
                                )
                            )
                    if not _metadata_bool_literal_is_valid(shape_metadata.get("chart.legendVisible")):
                        errors.append(
                            _issue(
                                "metadata_schema",
                                "chart.legendVisible metadata must be a bool-like literal.",
                                section_index=section_index,
                                paragraph_index=paragraph.index,
                                context=context,
                            )
                        )
                    chart_categories, chart_categories_error = _metadata_json_value(shape_metadata, "chart.categories")
                    if chart_categories_error is not None:
                        errors.append(
                            _issue(
                                "metadata_schema",
                                f"chart.categories metadata must be valid JSON: {chart_categories_error}",
                                section_index=section_index,
                                paragraph_index=paragraph.index,
                                context=context,
                            )
                        )
                    elif chart_categories is not None and not isinstance(chart_categories, list):
                        errors.append(
                            _issue(
                                "metadata_schema",
                                "chart.categories metadata must decode to a JSON list.",
                                section_index=section_index,
                                paragraph_index=paragraph.index,
                                context=context,
                            )
                        )
                    chart_series, chart_series_error = _metadata_json_value(shape_metadata, "chart.series")
                    if chart_series_error is not None:
                        errors.append(
                            _issue(
                                "metadata_schema",
                                f"chart.series metadata must be valid JSON: {chart_series_error}",
                                section_index=section_index,
                                paragraph_index=paragraph.index,
                                context=context,
                            )
                        )
                    elif chart_series is not None:
                        if not isinstance(chart_series, list):
                            errors.append(
                                _issue(
                                    "metadata_schema",
                                    "chart.series metadata must decode to a JSON list.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                        elif not all(isinstance(item, dict) for item in chart_series):
                            errors.append(
                                _issue(
                                    "metadata_schema",
                                    "chart.series metadata must decode to a JSON list of objects.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                    if has_chart:
                        chart_nodes = [node for node in control_node.iter_descendants() if node.tag_id == TAG_CHART_DATA]
                        for chart_node in chart_nodes:
                            if len(chart_node.payload) % 2:
                                errors.append(
                                    _issue(
                                        "payload_sanity",
                                        "Chart payload is not UTF-16 aligned.",
                                        section_index=section_index,
                                        paragraph_index=paragraph.index,
                                        context=context,
                                    )
                                )
                    if has_picture:
                        picture_nodes = [node for node in control_node.iter_descendants() if node.tag_id == TAG_SHAPE_COMPONENT_PICTURE]
                        for picture_node in picture_nodes:
                            if len(picture_node.payload) < 73:
                                errors.append(
                                    _issue(
                                        "payload_sanity",
                                        "Picture payload is shorter than the expected 73-byte minimum.",
                                        section_index=section_index,
                                        paragraph_index=paragraph.index,
                                        context=context,
                                    )
                                )
                            else:
                                storage_id = int.from_bytes(picture_node.payload[71:73], "little", signed=False)
                                if storage_id and storage_id not in docinfo_bindata:
                                    errors.append(
                                        _issue(
                                            "binary_reference",
                                            f"Picture control references missing BinData id {storage_id}.",
                                            section_index=section_index,
                                            paragraph_index=paragraph.index,
                                            context=context,
                                        )
                                    )
                    if has_ole:
                        ole_nodes = [node for node in control_node.iter_descendants() if node.tag_id == TAG_SHAPE_COMPONENT_OLE]
                        for ole_node in ole_nodes:
                            if len(ole_node.payload) < 20:
                                errors.append(
                                    _issue(
                                        "payload_sanity",
                                        "OLE payload is shorter than the expected 20-byte minimum.",
                                        section_index=section_index,
                                        paragraph_index=paragraph.index,
                                        context=context,
                                    )
                                )
                            storage_id = int.from_bytes(ole_node.payload[12:14].ljust(2, b"\x00"), "little", signed=False)
                            if storage_id and storage_id not in docinfo_bindata:
                                errors.append(
                                    _issue(
                                        "binary_reference",
                                        f"OLE control references missing BinData id {storage_id}.",
                                        section_index=section_index,
                                        paragraph_index=paragraph.index,
                                        context=context,
                                        )
                                    )
                elif control_id == "mrof":
                    metadata = _collect_control_native_metadata(control_node)
                    form_nodes = [node for node in control_node.iter_descendants() if node.tag_id == TAG_FORM_OBJECT]
                    if not form_nodes:
                        errors.append(
                            _issue(
                                "control_subtree",
                                "Form control does not contain a TAG_FORM_OBJECT record.",
                                section_index=section_index,
                                paragraph_index=paragraph.index,
                                context=context,
                            )
                        )
                    for form_node in form_nodes:
                        if len(form_node.payload) % 2:
                            errors.append(
                                _issue(
                                    "payload_sanity",
                                    "Form payload is not UTF-16 aligned.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                    if not _metadata_bool_literal_is_valid(metadata.get("form.checked")):
                        errors.append(
                            _issue(
                                "metadata_schema",
                                "form.checked metadata must be a bool-like literal.",
                                section_index=section_index,
                                paragraph_index=paragraph.index,
                                context=context,
                            )
                        )
                    if not _metadata_bool_literal_is_valid(metadata.get("form.editable")):
                        errors.append(
                            _issue(
                                "metadata_schema",
                                "form.editable metadata must be a bool-like literal.",
                                section_index=section_index,
                                paragraph_index=paragraph.index,
                                context=context,
                            )
                        )
                    if not _metadata_bool_literal_is_valid(metadata.get("form.locked")):
                        errors.append(
                            _issue(
                                "metadata_schema",
                                "form.locked metadata must be a bool-like literal.",
                                section_index=section_index,
                                paragraph_index=paragraph.index,
                                context=context,
                            )
                        )
                    form_items, form_items_error = _metadata_json_value(metadata, "form.items")
                    if form_items_error is not None:
                        errors.append(
                            _issue(
                                "metadata_schema",
                                f"form.items metadata must be valid JSON: {form_items_error}",
                                section_index=section_index,
                                paragraph_index=paragraph.index,
                                context=context,
                            )
                        )
                    elif form_items is not None and not isinstance(form_items, list):
                        errors.append(
                            _issue(
                                "metadata_schema",
                                "form.items metadata must decode to a JSON list.",
                                section_index=section_index,
                                paragraph_index=paragraph.index,
                                context=context,
                            )
                        )
                elif control_id == "omem":
                    metadata = _collect_control_native_metadata(control_node)
                    memo_nodes = [node for node in control_node.iter_descendants() if node.tag_id == TAG_MEMO_LIST]
                    if not memo_nodes:
                        errors.append(
                            _issue(
                                "control_subtree",
                                "Memo control does not contain a TAG_MEMO_LIST record.",
                                section_index=section_index,
                                paragraph_index=paragraph.index,
                                context=context,
                            )
                        )
                    for memo_node in memo_nodes:
                        if len(memo_node.payload) % 2:
                            errors.append(
                                _issue(
                                    "payload_sanity",
                                    "Memo payload is not UTF-16 aligned.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                    memo_order = metadata.get("memo.order")
                    if memo_order is not None and memo_order != "":
                        try:
                            int(memo_order)
                        except ValueError:
                            errors.append(
                                _issue(
                                    "metadata_schema",
                                    "memo.order metadata must be an integer.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                    if not _metadata_bool_literal_is_valid(metadata.get("memo.visible")):
                        errors.append(
                            _issue(
                                "metadata_schema",
                                "memo.visible metadata must be a bool-like literal.",
                                section_index=section_index,
                                paragraph_index=paragraph.index,
                                context=context,
                            )
                        )
                elif control_id == "tbl ":
                    table_nodes = [node for node in control_node.iter_descendants() if node.tag_id == TAG_TABLE]
                    if not table_nodes:
                        errors.append(
                            _issue(
                                "control_subtree",
                                "Table control does not contain a TAG_TABLE record.",
                                section_index=section_index,
                                paragraph_index=paragraph.index,
                                context=context,
                            )
                        )
                    for table_node in table_nodes:
                        if len(table_node.payload) < 20:
                            errors.append(
                                _issue(
                                    "payload_sanity",
                                    "Table payload is shorter than the expected minimum.",
                                    section_index=section_index,
                                    paragraph_index=paragraph.index,
                                    context=context,
                                )
                            )
                elif control_id in {"head", "foot", "fn  ", "en  "}:
                    if not any(isinstance(node, ParagraphHeaderRecord) for node in control_node.children):
                        errors.append(
                            _issue(
                                "control_subtree",
                                "Nested control is missing a writable paragraph header.",
                                section_index=section_index,
                                paragraph_index=paragraph.index,
                                context=context,
                            )
                        )
                    if control_id in {"head", "foot"} and len(control_node.payload) < 8:
                        errors.append(
                            _issue(
                                "payload_sanity",
                                "Header/footer payload is shorter than the expected 8-byte minimum.",
                                section_index=section_index,
                                paragraph_index=paragraph.index,
                                context=context,
                            )
                        )
                elif control_id == "pgnp" and len(control_node.payload) < 8:
                    errors.append(
                        _issue(
                            "payload_sanity",
                            "Page number payload is shorter than the expected 8-byte minimum.",
                            section_index=section_index,
                            paragraph_index=paragraph.index,
                            context=context,
                        )
                    )
                elif control_id.startswith("%") and control_id != "%hlk":
                    if not _control_has_descendant(control_node, TAG_CTRL_DATA):
                        errors.append(
                            _issue(
                                "control_subtree",
                                "Field control is missing TAG_CTRL_DATA parameters.",
                                section_index=section_index,
                                paragraph_index=paragraph.index,
                                context=context,
                            )
                        )

    return errors
