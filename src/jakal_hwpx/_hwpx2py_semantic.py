from __future__ import annotations

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
)

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from .hwpx2py import (
        _assign_if_value,
        _emit_asset_data,
        _emit_nested_blocks,
        _emit_non_default_fields,
        _emit_source_paragraph_index,
        _kw,
        _kwargs,
        _line,
        _py,
    )

_SHARED_NAMES = ['_assign_if_value','_emit_asset_data','_emit_nested_blocks','_emit_non_default_fields','_emit_source_paragraph_index','_kw','_kwargs','_line','_py']


def _bind_shared_helpers() -> None:
    if _SHARED_NAMES[0] in globals():
        return
    from . import hwpx2py as shared

    globals().update({name: getattr(shared, name) for name in _SHARED_NAMES})

def _emit_block(lines: list[str], block: object, section_index: int, *, include_binary_assets: bool) -> None:
    _bind_shared_helpers()
    if isinstance(block, Paragraph):
        _line(lines, f"block = doc.append_paragraph({_py(block.text)}, section_index={section_index})", indent=1)
        _emit_non_default_fields(
            lines,
            "block",
            block,
            {
                "text": block.text,
                "style_id": None,
                "para_pr_id": None,
                "char_pr_id": None,
                "hwp_para_shape_id": None,
                "hwp_style_id": None,
                "hwp_split_flags": None,
                "hwp_control_mask": None,
            },
            skip={"text"},
        )
        _emit_source_paragraph_index(lines, block)
        return

    if isinstance(block, Table):
        kwargs = _kwargs(
            [
                f"rows={block.rows}",
                f"cols={block.cols}",
                _kw("cell_texts", block.cell_texts, default=[]),
                _kw("row_heights", block.row_heights),
                _kw("col_widths", block.col_widths),
                _kw("cell_spans", block.cell_spans),
                _kw("cell_border_fill_ids", block.cell_border_fill_ids),
                _kw("table_border_fill_id", block.table_border_fill_id, default=1),
                f"section_index={section_index}",
            ]
        )
        _line(lines, f"block = doc.append_table({kwargs})", indent=1)
        _emit_non_default_fields(
            lines,
            "block",
            block,
            {
                "rows": block.rows,
                "cols": block.cols,
                "cell_texts": block.cell_texts,
                "row_heights": block.row_heights,
                "col_widths": block.col_widths,
                "cell_spans": block.cell_spans,
                "cell_border_fill_ids": block.cell_border_fill_ids,
                "cell_margins": None,
                "cell_vertical_aligns": None,
                "table_border_fill_id": block.table_border_fill_id,
                "cell_spacing": None,
                "table_margins": None,
                "layout": {},
                "out_margins": {},
                "page_break": None,
                "repeat_header": None,
            },
            skip={
                "rows",
                "cols",
                "cell_texts",
                "row_heights",
                "col_widths",
                "cell_spans",
                "cell_border_fill_ids",
                "table_border_fill_id",
            },
        )
        _emit_source_paragraph_index(lines, block)
        return

    if isinstance(block, Picture):
        if not include_binary_assets:
            _line(lines, f"# Skipped Picture asset: name={_py(block.name)}, size={len(block.data)} bytes", indent=1)
            return
        _emit_asset_data(lines, block.data)
        kwargs = _kwargs(
            [
                _kw("extension", block.extension),
                _kw("width", block.width, default=7200),
                _kw("height", block.height, default=7200),
                _kw("original_width", block.original_width),
                _kw("original_height", block.original_height),
                _kw("current_width", block.current_width),
                _kw("current_height", block.current_height),
                f"section_index={section_index}",
            ]
        )
        _line(lines, f"block = doc.append_picture({_py(block.name)}, asset_data, {kwargs})", indent=1)
        _emit_source_paragraph_index(lines, block)
        _assign_if_value(lines, "block", "shape_comment", block.shape_comment, default=None)
        _assign_if_value(lines, "block", "layout", block.layout, default={})
        _assign_if_value(lines, "block", "out_margins", block.out_margins, default={})
        _assign_if_value(lines, "block", "rotation", block.rotation, default={})
        _assign_if_value(lines, "block", "image_adjustment", block.image_adjustment, default={})
        _assign_if_value(lines, "block", "crop", block.crop, default={})
        _assign_if_value(lines, "block", "line_color", block.line_color, default="#000000")
        _assign_if_value(lines, "block", "line_width", block.line_width, default=0)
        _assign_if_value(lines, "block", "has_size_node", block.has_size_node, default=True)
        _assign_if_value(lines, "block", "size_attributes", block.size_attributes, default={})
        _assign_if_value(lines, "block", "has_line_node", block.has_line_node, default=False)
        _assign_if_value(lines, "block", "has_position_node", block.has_position_node, default=True)
        _assign_if_value(lines, "block", "has_out_margin_node", block.has_out_margin_node, default=True)
        return

    if isinstance(block, Hyperlink):
        kwargs = _kwargs(
            [
                _kw("display_text", block.display_text),
                _kw("metadata_fields", block.metadata_fields, default=[]),
                f"section_index={section_index}",
            ]
        )
        _line(lines, f"block = doc.append_hyperlink({_py(block.target)}, {kwargs})", indent=1)
        _emit_source_paragraph_index(lines, block)
        return

    if isinstance(block, Bookmark):
        _line(lines, f"block = doc.append_bookmark({_py(block.name)}, section_index={section_index})", indent=1)
        _emit_source_paragraph_index(lines, block)
        return

    if isinstance(block, Field):
        kwargs = _kwargs(
            [
                f"field_type={_py(block.field_type)}",
                _kw("display_text", block.display_text),
                _kw("name", block.name),
                _kw("parameters", block.parameters, default={}),
                _kw("editable", block.editable, default=False),
                _kw("dirty", block.dirty, default=False),
                f"section_index={section_index}",
            ]
        )
        _line(lines, f"block = doc.append_field({kwargs})", indent=1)
        _emit_source_paragraph_index(lines, block)
        _assign_if_value(lines, "block", "native_field_type", block.native_field_type, default=None)
        return

    if isinstance(block, AutoNumber):
        kwargs = _kwargs(
            [
                _kw("number", block.number, default=1),
                _kw("number_type", block.number_type, default="PAGE"),
                _kw("kind", block.kind, default="newNum"),
                f"section_index={section_index}",
            ]
        )
        _line(lines, f"block = doc.append_auto_number({kwargs})", indent=1)
        _emit_source_paragraph_index(lines, block)
        return

    if isinstance(block, Note):
        kwargs = _kwargs(
            [
                _kw("kind", block.kind, default="footNote"),
                _kw("number", block.number),
                f"section_index={section_index}",
            ]
        )
        _line(lines, f"block = doc.append_note({_py(block.text)}, {kwargs})", indent=1)
        _emit_source_paragraph_index(lines, block)
        _emit_nested_blocks(lines, block)
        return

    if isinstance(block, Equation):
        kwargs = _kwargs(
            [
                _kw("width", block.width, default=4800),
                _kw("height", block.height, default=2300),
                _kw("shape_comment", block.shape_comment),
                _kw("text_color", block.text_color, default="#000000"),
                _kw("base_unit", block.base_unit, default=1100),
                _kw("font", block.font, default="HYhwpEQ"),
                f"section_index={section_index}",
            ]
        )
        _line(lines, f"block = doc.append_equation({_py(block.script)}, {kwargs})", indent=1)
        _emit_source_paragraph_index(lines, block)
        _assign_if_value(lines, "block", "layout", block.layout, default={})
        _assign_if_value(lines, "block", "out_margins", block.out_margins, default={})
        _assign_if_value(lines, "block", "rotation", block.rotation, default={})
        return

    if isinstance(block, Shape):
        kwargs = _kwargs(
            [
                _kw("kind", block.kind, default="rect"),
                _kw("text", block.text, default=""),
                _kw("width", block.width, default=12000),
                _kw("height", block.height, default=3200),
                _kw("fill_color", block.fill_color, default="#FFFFFF"),
                _kw("line_color", block.line_color, default="#000000"),
                _kw("line_width", block.line_width, default=33),
                _kw("shape_comment", block.shape_comment),
                _kw("original_width", block.original_width),
                _kw("original_height", block.original_height),
                _kw("current_width", block.current_width),
                _kw("current_height", block.current_height),
                f"section_index={section_index}",
            ]
        )
        _line(lines, f"block = doc.append_shape({kwargs})", indent=1)
        _emit_source_paragraph_index(lines, block)
        _emit_nested_blocks(lines, block)
        _assign_if_value(lines, "block", "layout", block.layout, default={})
        _assign_if_value(lines, "block", "out_margins", block.out_margins, default={})
        _assign_if_value(lines, "block", "rotation", block.rotation, default={})
        _assign_if_value(lines, "block", "text_margins", block.text_margins, default={})
        _assign_if_value(lines, "block", "specific_fields", block.specific_fields, default={})
        _assign_if_value(lines, "block", "line_style", block.line_style, default={})
        _assign_if_value(lines, "block", "has_line_node", block.has_line_node, default=True)
        _assign_if_value(lines, "block", "has_size_node", block.has_size_node, default=True)
        _assign_if_value(lines, "block", "size_attributes", block.size_attributes, default={})
        _assign_if_value(lines, "block", "has_position_node", block.has_position_node, default=True)
        _assign_if_value(lines, "block", "has_out_margin_node", block.has_out_margin_node, default=True)
        _assign_if_value(lines, "block", "has_fill_node", block.has_fill_node, default=True)
        _assign_if_value(lines, "block", "has_text_margin_node", block.has_text_margin_node, default=True)
        return

    if isinstance(block, Ole):
        if not include_binary_assets:
            _line(lines, f"# Skipped OLE asset: name={_py(block.name)}, size={len(block.data)} bytes", indent=1)
            return
        _emit_asset_data(lines, block.data)
        kwargs = _kwargs(
            [
                _kw("width", block.width, default=42001),
                _kw("height", block.height, default=13501),
                _kw("shape_comment", block.shape_comment),
                _kw("object_type", block.object_type, default="EMBEDDED"),
                _kw("draw_aspect", block.draw_aspect, default="CONTENT"),
                _kw("has_moniker", block.has_moniker, default=False),
                _kw("eq_baseline", block.eq_baseline, default=0),
                _kw("line_color", block.line_color, default="#000000"),
                _kw("line_width", block.line_width, default=0),
                _kw("original_width", block.original_width),
                _kw("original_height", block.original_height),
                _kw("current_width", block.current_width),
                _kw("current_height", block.current_height),
                f"section_index={section_index}",
            ]
        )
        _line(lines, f"block = doc.append_ole({_py(block.name)}, asset_data, {kwargs})", indent=1)
        _emit_source_paragraph_index(lines, block)
        _assign_if_value(lines, "block", "layout", block.layout, default={})
        _assign_if_value(lines, "block", "out_margins", block.out_margins, default={})
        _assign_if_value(lines, "block", "rotation", block.rotation, default={})
        _assign_if_value(lines, "block", "extent", block.extent, default={})
        _assign_if_value(lines, "block", "has_size_node", block.has_size_node, default=True)
        _assign_if_value(lines, "block", "size_attributes", block.size_attributes, default={})
        _assign_if_value(lines, "block", "has_position_node", block.has_position_node, default=True)
        _assign_if_value(lines, "block", "has_out_margin_node", block.has_out_margin_node, default=True)
        return

    if isinstance(block, Form):
        kwargs = _kwargs(
            [
                _kw("form_type", block.form_type, default="INPUT"),
                _kw("name", block.name),
                _kw("value", block.value),
                _kw("checked", block.checked, default=False),
                _kw("items", block.items, default=[]),
                _kw("editable", block.editable, default=True),
                _kw("locked", block.locked, default=False),
                _kw("placeholder", block.placeholder),
                f"section_index={section_index}",
            ]
        )
        _line(lines, f"block = doc.append_form({_py(block.label)}, {kwargs})", indent=1)
        _emit_source_paragraph_index(lines, block)
        return

    if isinstance(block, Memo):
        kwargs = _kwargs(
            [
                _kw("author", block.author),
                _kw("memo_id", block.memo_id),
                _kw("anchor_id", block.anchor_id),
                _kw("order", block.order),
                _kw("visible", block.visible, default=True),
                f"section_index={section_index}",
            ]
        )
        _line(lines, f"block = doc.append_memo({_py(block.text)}, {kwargs})", indent=1)
        _emit_source_paragraph_index(lines, block)
        return

    if isinstance(block, Chart):
        kwargs = _kwargs(
            [
                _kw("chart_type", block.chart_type, default="BAR"),
                _kw("categories", block.categories, default=[]),
                _kw("series", block.series, default=[]),
                _kw("data_ref", block.data_ref),
                _kw("legend_visible", block.legend_visible, default=True),
                _kw("width", block.width, default=12000),
                _kw("height", block.height, default=3200),
                _kw("shape_comment", block.shape_comment),
                f"section_index={section_index}",
            ]
        )
        _line(lines, f"block = doc.append_chart({_py(block.title)}, {kwargs})", indent=1)
        _emit_source_paragraph_index(lines, block)
        _assign_if_value(lines, "block", "layout", block.layout, default={})
        _assign_if_value(lines, "block", "out_margins", block.out_margins, default={})
        _assign_if_value(lines, "block", "rotation", block.rotation, default={})
        return

    _line(lines, f"# Skipped unsupported block: {type(block).__name__}", indent=1)

_emit_block.__module__ = "jakal_hwpx.hwpx2py"

