from __future__ import annotations

from typing import TYPE_CHECKING

from ._hancom_models import HancomBlock
from ._hancom_styles import _build_hwp_colorref
from .hwp_binary import TAG_BORDER_FILL, TypedRecord

if TYPE_CHECKING:
    from .hwp_document import HwpDocument

# Imported after hancom_document finishes defining shared conversion helpers.
from .hancom_document import (
    AutoNumber,
    Bookmark,
    Chart,
    Equation,
    Field,
    Form,
    Hyperlink,
    Memo,
    Note,
    Ole,
    Paragraph,
    Picture,
    Shape,
    Table,
    _apply_hwp_shape_specific_fields,
    _block_owned_text_for_write,
    _hwpx_layout_kwargs,
    _hwpx_line_style_kwargs,
    _hwpx_rotation_kwargs,
    _normalize_int_map,
    _resolve_hwp_paragraph_attributes,
)


def _ensure_hwp_table_border_fill(
    document: HwpDocument,
    fill_color: str | None,
    borders: dict[str, dict[str, str | int]],
) -> int:
    key = (fill_color, tuple((side, tuple(sorted(spec.items()))) for side, spec in sorted(borders.items())))
    cache = getattr(document, "_docx_table_border_fill_cache", None)
    if cache is None:
        cache = {}
        setattr(document, "_docx_table_border_fill_cache", cache)
    if key in cache:
        return cache[key]
    docinfo = document.docinfo_model()
    records = docinfo.records_by_tag_id(TAG_BORDER_FILL)
    if not records:
        raise ValueError("The HWP profile has no border-fill template.")
    template = records[1] if fill_color and len(records) > 1 and len(records[1].payload) >= 53 else records[0]
    payload = bytearray(template.payload)
    if fill_color:
        if len(payload) < 53:
            payload = bytearray(records[0].payload[:32]) + bytearray(21)
            payload[32:36] = (1).to_bytes(4, "little")
        payload[36:40] = _build_hwp_colorref(fill_color)
    style_codes = {"single": 0, "dashed": 1, "dotted": 2, "double": 7}
    widths = (0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0)
    for side, offset in (("left", 2), ("right", 8), ("top", 14), ("bottom", 20)):
        spec = borders.get(side)
        if not spec or spec.get("style") in {"none", "nil"}:
            continue
        payload[offset] = style_codes.get(str(spec.get("style", "single")), 0)
        width_mm = int(spec.get("size", 4)) * 0.3527778 / 8
        payload[offset + 1] = min(range(len(widths)), key=lambda index: abs(widths[index] - width_mm))
        payload[offset + 2:offset + 6] = _build_hwp_colorref(str(spec.get("color", "000000")))
    record = TypedRecord(tag_id=TAG_BORDER_FILL, level=records[-1].level, payload=bytes(payload))
    parent = docinfo.id_mappings_record()
    parent.children.insert(parent.children.index(records[-1]) + 1, record)
    parent.set_count(8, len(records) + 1)
    document.binary_document().replace_docinfo_model(docinfo)
    style_id = len(records)
    cache[key] = style_id
    return style_id


def _append_hancom_block_to_hwp(
    document: HwpDocument,
    block: HancomBlock,
    *,
    section_index: int,
    paragraph_attribute_limits: tuple[int, int] | None = None,
) -> None:
    unsupported = getattr(block, "_unsupported_word_math", None)
    if unsupported:
        raise ValueError(unsupported)
    if isinstance(block, Paragraph):
        if block.text:
            para_shape_id, style_id, split_flags, control_mask = _resolve_hwp_paragraph_attributes(
                document,
                block,
                paragraph_attribute_limits=paragraph_attribute_limits,
            )
            document.append_paragraph(
                block.text,
                section_index=section_index,
                para_shape_id=para_shape_id,
                style_id=style_id,
                split_flags=split_flags,
                control_mask=control_mask,
            )
        return
    if isinstance(block, Table):
        cell_border_fill_ids = dict(block.cell_border_fill_ids or {})
        for row, column in set(block.cell_fill_colors or {}) | set(block.cell_border_styles or {}):
            color = (block.cell_fill_colors or {}).get((row, column))
            borders = (block.cell_border_styles or {}).get((row, column), {})
            cell_border_fill_ids[(row, column)] = _ensure_hwp_table_border_fill(document, color, borders)
        table = document.append_table(
            rows=block.rows,
            cols=block.cols,
            cell_texts=block.cell_texts,
            row_heights=block.row_heights,
            col_widths=block.col_widths,
            cell_spans=block.cell_spans,
            cell_border_fill_ids=cell_border_fill_ids or None,
            table_border_fill_id=block.table_border_fill_id,
            section_index=section_index,
        )
        if table is None:
            table = document.section(section_index).tables()[-1]
        if block.col_widths:
            table.set_column_widths(block.col_widths)
        if block.cell_spacing is not None:
            table.set_cell_spacing(block.cell_spacing)
        if block.table_margins:
            table.set_table_margins(**block.table_margins)
        for (row, column), border_fill_id in cell_border_fill_ids.items():
            table.set_cell_border_fill_id(row, column, border_fill_id)
        for (row, column), margins in (block.cell_margins or {}).items():
            table.set_cell_margins(row, column, **_normalize_int_map(margins))
        return
    if isinstance(block, Picture):
        picture = document.append_picture(
            block.data,
            extension=block.extension,
            width=block.width,
            height=block.height,
            section_index=section_index,
        )
        if picture is None:
            picture = document.section(section_index).pictures()[-1]
        if block.shape_comment:
            picture.set_shape_comment(block.shape_comment)
        if block.layout:
            picture.set_layout(**_hwpx_layout_kwargs(block.layout))
        if block.out_margins:
            picture.set_out_margins(**block.out_margins)
        if block.rotation:
            picture.set_rotation(**_hwpx_rotation_kwargs(block.rotation))
        if block.image_adjustment:
            picture.set_image_adjustment(**block.image_adjustment)
        if block.crop:
            picture.set_crop(**block.crop)
        line_style_kwargs = _hwpx_line_style_kwargs(block.line_style)
        if line_style_kwargs:
            picture.set_line_style(**line_style_kwargs)
        elif block.line_color or block.line_width:
            picture.set_line_style(color=block.line_color, width=block.line_width)
        return
    if isinstance(block, Hyperlink):
        document.append_hyperlink(
            block.target,
            text=_block_owned_text_for_write(block) or None,
            metadata_fields=block.metadata_fields or None,
            section_index=section_index,
        )
        return
    if isinstance(block, Bookmark):
        if block.name:
            document.append_bookmark(block.name, section_index=section_index)
        return
    if isinstance(block, Field):
        document.append_field(
            field_type=block.effective_native_field_type,
            display_text=_block_owned_text_for_write(block),
            name=block.name,
            parameters=block.parameters,
            editable=block.editable,
            dirty=block.dirty,
            section_index=section_index,
        )
        return
    if isinstance(block, Form):
        document.append_form(
            block.label,
            form_type=block.form_type,
            name=block.name,
            value=block.value,
            checked=block.checked,
            items=block.items,
            editable=block.editable,
            locked=block.locked,
            placeholder=block.placeholder,
            section_index=section_index,
        )
        return
    if isinstance(block, Memo):
        document.append_memo(
            block.text,
            author=block.author,
            memo_id=block.memo_id,
            anchor_id=block.anchor_id,
            order=block.order,
            visible=block.visible,
            section_index=section_index,
        )
        return
    if isinstance(block, AutoNumber):
        document.append_auto_number(number=block.number, number_type=block.number_type, kind=block.kind, section_index=section_index)
        return
    if isinstance(block, Note):
        document.append_note(block.text, kind=block.kind, number=block.number, section_index=section_index)
        return
    if isinstance(block, Equation):
        equation = document.append_equation(
            block.script,
            width=block.width,
            height=block.height,
            font=block.font,
            shape_comment=block.shape_comment,
            section_index=section_index,
        )
        if equation is None:
            equation = document.section(section_index).equations()[-1]
        if block.layout:
            equation.set_layout(**_hwpx_layout_kwargs(block.layout))
        if block.out_margins:
            equation.set_out_margins(**block.out_margins)
        if block.rotation:
            equation.set_rotation(**_hwpx_rotation_kwargs(block.rotation))
        return
    if isinstance(block, Shape):
        kind = block.kind if block.kind in {"line", "connectLine", "rect", "ellipse", "arc", "polygon", "curve", "container", "textart"} else "rect"
        shape = document.append_shape(
            kind=kind,
            text=block.text,
            width=block.width,
            height=block.height,
            fill_color=block.fill_color,
            line_color=block.line_color,
            shape_comment=block.shape_comment,
            section_index=section_index,
        )
        if shape is None:
            shape = document.section(section_index).shapes()[-1]
        _apply_hwp_shape_specific_fields(shape, block)
        if block.layout:
            shape.set_layout(**_hwpx_layout_kwargs(block.layout))
        if block.out_margins:
            shape.set_out_margins(**block.out_margins)
        if block.rotation:
            shape.set_rotation(**_hwpx_rotation_kwargs(block.rotation))
        line_style_kwargs = _hwpx_line_style_kwargs(block.line_style)
        if line_style_kwargs:
            shape.set_line_style(**line_style_kwargs)
        else:
            shape.set_line_style(color=block.line_color, width=block.line_width, style="SOLID" if block.line_width else "NONE")
        if block.text_margins and hasattr(shape, "set_text_margins"):
            shape.set_text_margins(**block.text_margins)
        return
    if isinstance(block, Chart):
        chart = document.append_chart(
            block.title,
            chart_type=block.chart_type,
            categories=block.categories,
            series=block.series,
            data_ref=block.data_ref,
            legend_visible=block.legend_visible,
            width=block.width,
            height=block.height,
            shape_comment=block.shape_comment,
            section_index=section_index,
        )
        if chart is None:
            chart = document.section(section_index).charts()[-1]
        if block.layout:
            chart.set_layout(**_hwpx_layout_kwargs(block.layout))
        if block.out_margins:
            chart.set_out_margins(**block.out_margins)
        if block.rotation:
            chart.set_rotation(**_hwpx_rotation_kwargs(block.rotation))
        return
    if isinstance(block, Ole):
        ole = document.append_ole(
            block.name,
            block.data,
            width=block.width,
            height=block.height,
            shape_comment=block.shape_comment,
            object_type=block.object_type,
            draw_aspect=block.draw_aspect,
            has_moniker=block.has_moniker,
            eq_baseline=block.eq_baseline,
            line_color=block.line_color,
            line_width=block.line_width,
            section_index=section_index,
        )
        if ole is None:
            ole = document.section(section_index).oles()[-1]
        if block.layout:
            ole.set_layout(**_hwpx_layout_kwargs(block.layout))
        if block.out_margins:
            ole.set_out_margins(**block.out_margins)
        if block.rotation:
            ole.set_rotation(**_hwpx_rotation_kwargs(block.rotation))
        if block.extent:
            ole.set_extent(**block.extent)
        line_style_kwargs = _hwpx_line_style_kwargs(block.line_style)
        if line_style_kwargs:
            ole.set_line_style(**line_style_kwargs)
        elif block.line_color or block.line_width:
            ole.set_line_style(color=block.line_color, width=block.line_width)
        return

_append_hancom_block_to_hwp.__module__ = "jakal_hwpx.hancom_document"

