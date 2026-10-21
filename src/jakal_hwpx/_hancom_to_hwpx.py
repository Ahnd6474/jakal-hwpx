from __future__ import annotations

from typing import TYPE_CHECKING

from ._hancom_models import HancomBlock, HancomSection

if TYPE_CHECKING:
    from .document import HwpxDocument

# Imported after hancom_document finishes defining shared conversion helpers.
from .hancom_document import (
    _HWPX_MEMO_FIELD_TYPE,
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
    _append_control_host_paragraph,
    _apply_hwpx_equation_block,
    _apply_hwpx_nested_text_blocks,
    _apply_hwpx_ole_block,
    _apply_hwpx_paragraph_style_block,
    _apply_hwpx_picture_block,
    _apply_hwpx_shape_block,
    _apply_hwpx_table_block,
    _block_owned_text_for_write,
    _build_hwpx_memo_parameters,
    _has_nested_blocks,
    _hwpx_layout_kwargs,
    _hwpx_rotation_kwargs,
    _media_type_for_extension,
    _memo_requires_hwpx_carrier,
    _resolve_hwpx_char_pr_id,
    _resolve_hwpx_para_pr_id,
    _resolve_hwpx_style_id,
    _section_uses_source_paragraphs,
    _write_grouped_section_to_hwpx,
)


def _write_section_to_hwpx(document: HwpxDocument, section_index: int, section: HancomSection) -> None:
    if _section_uses_source_paragraphs(section):
        _write_grouped_section_to_hwpx(document, section_index, section)
        return

    blocks_written = 0
    first_paragraph_consumed = False
    for block in section.blocks:
        if isinstance(block, Paragraph):
            if not first_paragraph_consumed and blocks_written == 0:
                document.set_paragraph_text(section_index, 0, block.text, char_pr_id=_resolve_hwpx_char_pr_id(document, block))
                _apply_hwpx_paragraph_style_block(document, section_index, 0, block)
                first_paragraph_consumed = True
            else:
                document.append_paragraph(
                    block.text,
                    section_index=section_index,
                    para_pr_id=_resolve_hwpx_para_pr_id(document, block),
                    style_id=_resolve_hwpx_style_id(document, block),
                    char_pr_id=_resolve_hwpx_char_pr_id(document, block),
                )
            blocks_written += 1
            continue

        paragraph_index = 0 if blocks_written == 0 else _append_control_host_paragraph(document, section_index)

        if isinstance(block, Table):
            appended = document.append_table(
                block.rows,
                block.cols,
                section_index=section_index,
                paragraph_index=paragraph_index,
                cell_texts=block.cell_texts,
            )
            _apply_hwpx_table_block(appended, block)
        elif isinstance(block, Picture):
            appended = document.append_picture(
                block.name,
                block.data,
                section_index=section_index,
                paragraph_index=paragraph_index,
                width=block.width,
                height=block.height,
                original_width=block.original_width,
                original_height=block.original_height,
                current_width=block.current_width,
                current_height=block.current_height,
                media_type=_media_type_for_extension(block.extension),
            )
            _apply_hwpx_picture_block(appended, block)
        elif isinstance(block, Hyperlink):
            document.append_hyperlink(
                block.target,
                display_text=_block_owned_text_for_write(block),
                section_index=section_index,
                paragraph_index=paragraph_index,
            )
        elif isinstance(block, Bookmark):
            document.append_bookmark(
                block.name,
                section_index=section_index,
                paragraph_index=paragraph_index,
            )
        elif isinstance(block, Field):
            document.append_field(
                field_type=block.effective_native_field_type,
                display_text=_block_owned_text_for_write(block),
                name=block.name,
                parameters=block.parameters,
                editable=block.editable,
                dirty=block.dirty,
                section_index=section_index,
                paragraph_index=paragraph_index,
            )
        elif isinstance(block, Form):
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
                paragraph_index=paragraph_index,
            )
        elif isinstance(block, Memo):
            document.append_memo(
                block.text,
                section_index=section_index,
                paragraph_index=paragraph_index,
            )
            if _memo_requires_hwpx_carrier(block):
                document.append_field(
                    field_type=_HWPX_MEMO_FIELD_TYPE,
                    display_text=None,
                    name=block.memo_id,
                    parameters=_build_hwpx_memo_parameters(block),
                    section_index=section_index,
                    paragraph_index=paragraph_index,
                )
        elif isinstance(block, AutoNumber):
            document.append_auto_number(
                number=block.number,
                number_type=block.number_type,
                kind=block.kind,
                section_index=section_index,
                paragraph_index=paragraph_index,
            )
        elif isinstance(block, Note):
            appended = document.append_note(
                "" if _has_nested_blocks(block) else block.text,
                kind=block.kind,
                number=block.number,
                section_index=section_index,
                paragraph_index=paragraph_index,
            )
            _apply_hwpx_nested_text_blocks(document, appended, block, vertical_align="TOP")
        elif isinstance(block, Equation):
            appended = document.append_equation(
                block.script,
                section_index=section_index,
                paragraph_index=paragraph_index,
                width=block.width,
                height=block.height,
                shape_comment=block.shape_comment,
                text_color=block.text_color,
                base_unit=block.base_unit,
                font=block.font,
            )
            _apply_hwpx_equation_block(appended, block)
        elif isinstance(block, Shape):
            appended = document.append_shape(
                kind=block.kind,
                text="" if _has_nested_blocks(block) else block.text,
                width=block.width,
                height=block.height,
                fill_color=block.fill_color,
                line_color=block.line_color,
                line_width=block.line_width,
                original_width=block.original_width,
                original_height=block.original_height,
                current_width=block.current_width,
                current_height=block.current_height,
                shape_comment=block.shape_comment,
                section_index=section_index,
                paragraph_index=paragraph_index,
            )
            _apply_hwpx_shape_block(appended, block)
            _apply_hwpx_nested_text_blocks(document, appended, block)
        elif isinstance(block, Chart):
            appended = document.append_chart(
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
                paragraph_index=paragraph_index,
            )
            if block.layout:
                appended.set_layout(**_hwpx_layout_kwargs(block.layout))
            if block.out_margins:
                appended.set_out_margins(**block.out_margins)
            if block.rotation:
                appended.set_rotation(**_hwpx_rotation_kwargs(block.rotation))
        elif isinstance(block, Ole):
            appended = document.append_ole(
                block.name,
                block.data,
                width=block.width,
                height=block.height,
                shape_comment=block.shape_comment,
                object_type=block.object_type,
                draw_aspect=block.draw_aspect,
                has_moniker=block.has_moniker,
                eq_baseline=block.eq_baseline,
                original_width=block.original_width,
                original_height=block.original_height,
                current_width=block.current_width if block.current_width is not None else 0,
                current_height=block.current_height if block.current_height is not None else 0,
                section_index=section_index,
                paragraph_index=paragraph_index,
            )
            _apply_hwpx_ole_block(appended, block)
        blocks_written += 1


def _append_non_paragraph_block_to_hwpx(
    document: HwpxDocument,
    section_index: int,
    paragraph_index: int,
    block: HancomBlock,
) -> None:
    if isinstance(block, Table):
        appended = document.append_table(
            block.rows,
            block.cols,
            section_index=section_index,
            paragraph_index=paragraph_index,
            cell_texts=block.cell_texts,
        )
        _apply_hwpx_table_block(appended, block)
    elif isinstance(block, Picture):
        appended = document.append_picture(
            block.name,
            block.data,
            section_index=section_index,
            paragraph_index=paragraph_index,
            width=block.width,
            height=block.height,
            original_width=block.original_width,
            original_height=block.original_height,
            current_width=block.current_width,
            current_height=block.current_height,
            media_type=_media_type_for_extension(block.extension),
        )
        _apply_hwpx_picture_block(appended, block)
    elif isinstance(block, Hyperlink):
        document.append_hyperlink(
            block.target,
            display_text=_block_owned_text_for_write(block),
            section_index=section_index,
            paragraph_index=paragraph_index,
        )
    elif isinstance(block, Bookmark):
        document.append_bookmark(
            block.name,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )
    elif isinstance(block, Field):
        document.append_field(
            field_type=block.effective_native_field_type,
            display_text=_block_owned_text_for_write(block),
            name=block.name,
            parameters=block.parameters,
            editable=block.editable,
            dirty=block.dirty,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )
    elif isinstance(block, Form):
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
            paragraph_index=paragraph_index,
        )
    elif isinstance(block, Memo):
        document.append_memo(
            block.text,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )
        if _memo_requires_hwpx_carrier(block):
            document.append_field(
                field_type=_HWPX_MEMO_FIELD_TYPE,
                display_text=None,
                name=block.memo_id,
                parameters=_build_hwpx_memo_parameters(block),
                section_index=section_index,
                paragraph_index=paragraph_index,
            )
    elif isinstance(block, AutoNumber):
        document.append_auto_number(
            number=block.number,
            number_type=block.number_type,
            kind=block.kind,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )
    elif isinstance(block, Note):
        appended = document.append_note(
            "" if _has_nested_blocks(block) else block.text,
            kind=block.kind,
            number=block.number,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )
        _apply_hwpx_nested_text_blocks(document, appended, block, vertical_align="TOP")
    elif isinstance(block, Equation):
        appended = document.append_equation(
            block.script,
            section_index=section_index,
            paragraph_index=paragraph_index,
            width=block.width,
            height=block.height,
            shape_comment=block.shape_comment,
            text_color=block.text_color,
            base_unit=block.base_unit,
            font=block.font,
        )
        _apply_hwpx_equation_block(appended, block)
    elif isinstance(block, Shape):
        appended = document.append_shape(
            kind=block.kind,
            text="" if _has_nested_blocks(block) else block.text,
            width=block.width,
            height=block.height,
            fill_color=block.fill_color,
            line_color=block.line_color,
            line_width=block.line_width,
            original_width=block.original_width,
            original_height=block.original_height,
            current_width=block.current_width,
            current_height=block.current_height,
            shape_comment=block.shape_comment,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )
        _apply_hwpx_shape_block(appended, block)
        _apply_hwpx_nested_text_blocks(document, appended, block)
    elif isinstance(block, Chart):
        appended = document.append_chart(
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
            paragraph_index=paragraph_index,
        )
        if block.layout:
            appended.set_layout(**_hwpx_layout_kwargs(block.layout))
        if block.out_margins:
            appended.set_out_margins(**block.out_margins)
        if block.rotation:
            appended.set_rotation(**_hwpx_rotation_kwargs(block.rotation))
    elif isinstance(block, Ole):
        appended = document.append_ole(
            block.name,
            block.data,
            width=block.width,
            height=block.height,
            shape_comment=block.shape_comment,
            object_type=block.object_type,
            draw_aspect=block.draw_aspect,
            has_moniker=block.has_moniker,
            eq_baseline=block.eq_baseline,
            original_width=block.original_width,
            original_height=block.original_height,
            current_width=block.current_width if block.current_width is not None else 0,
            current_height=block.current_height if block.current_height is not None else 0,
            section_index=section_index,
            paragraph_index=paragraph_index,
        )
        _apply_hwpx_ole_block(appended, block)

_write_section_to_hwpx.__module__ = "jakal_hwpx.hancom_document"
_append_non_paragraph_block_to_hwpx.__module__ = "jakal_hwpx.hancom_document"

