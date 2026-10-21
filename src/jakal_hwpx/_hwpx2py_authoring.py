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
        _authoring_cell_border_fill_ids,
        _emit_asset_data,
        _emit_nested_blocks,
        _emit_non_default_fields,
        _emit_source_paragraph_index,
        _kw,
        _kwargs,
        _line,
        _py,
        _raw_py_str,
        _table_size_matches_cells,
    )

_SHARED_NAMES = ['_assign_if_value','_authoring_cell_border_fill_ids','_emit_asset_data','_emit_nested_blocks','_emit_non_default_fields','_emit_source_paragraph_index','_kw','_kwargs','_line','_py','_raw_py_str','_table_size_matches_cells']


def _bind_shared_helpers() -> None:
    if _SHARED_NAMES[0] in globals():
        return
    from . import hwpx2py as shared

    globals().update({name: getattr(shared, name) for name in _SHARED_NAMES})

def _emit_authoring_block(lines: list[str], block: object, section_index: int, *, include_binary_assets: bool) -> None:
    _bind_shared_helpers()
    section_kw = _kw("section", section_index, default=0)
    if isinstance(block, Paragraph):
        kwargs = _kwargs(
            [
                section_kw,
                _kw("style", block.style_id, default="0"),
                _kw("para", block.para_pr_id, default="0"),
                _kw("char", block.char_pr_id, default="0"),
            ]
        )
        suffix = f", {kwargs}" if kwargs else ""
        _line(lines, f"block = p(doc, {_py(block.text)}{suffix})", indent=1)
        _emit_non_default_fields(
            lines,
            "block",
            block,
            {
                "text": block.text,
                "style_id": "0",
                "para_pr_id": "0",
                "char_pr_id": "0",
                "hwp_para_shape_id": None,
                "hwp_style_id": None,
                "hwp_split_flags": None,
                "hwp_control_mask": None,
            },
            skip={"text", "style_id", "para_pr_id", "char_pr_id"},
        )
        _emit_source_paragraph_index(lines, block)
        return

    if isinstance(block, Table):
        kwargs = _kwargs(
            [
                section_kw,
                _kw("rows", block.rows) if not _table_size_matches_cells(block) else None,
                _kw("cols", block.cols) if not _table_size_matches_cells(block) else None,
                _kw("row_heights", block.row_heights),
                _kw("col_widths", block.col_widths),
                _kw("cell_spans", block.cell_spans),
                _kw("cell_border_fill_ids", _authoring_cell_border_fill_ids(block)),
                _kw("table_border_fill_id", block.table_border_fill_id, default=1),
            ]
        )
        suffix = f", {kwargs}" if kwargs else ""
        _line(lines, f"block = table(doc, {_py(block.cell_texts)}{suffix})", indent=1)
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
            skip={"rows", "cols", "cell_texts", "row_heights", "col_widths", "cell_spans", "cell_border_fill_ids", "table_border_fill_id"},
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
                section_kw,
                _kw("extension", block.extension),
                _kw("width", block.width, default=7200),
                _kw("height", block.height, default=7200),
                _kw("original_width", block.original_width),
                _kw("original_height", block.original_height),
                _kw("current_width", block.current_width),
                _kw("current_height", block.current_height),
            ]
        )
        suffix = f", {kwargs}" if kwargs else ""
        _line(lines, f"block = image(doc, {_py(block.name)}, asset_data{suffix})", indent=1)
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
        kwargs = _kwargs([section_kw, _kw("display_text", block.display_text), _kw("metadata_fields", block.metadata_fields, default=[])])
        suffix = f", {kwargs}" if kwargs else ""
        _line(lines, f"block = link(doc, {_py(block.target)}{suffix})", indent=1)
        _emit_source_paragraph_index(lines, block)
        return

    if isinstance(block, Bookmark):
        kwargs = _kwargs([section_kw])
        suffix = f", {kwargs}" if kwargs else ""
        _line(lines, f"block = bookmark(doc, {_py(block.name)}{suffix})", indent=1)
        _emit_source_paragraph_index(lines, block)
        return

    if isinstance(block, Field):
        kwargs = _kwargs(
            [
                section_kw,
                _kw("display_text", block.display_text),
                _kw("name", block.name),
                _kw("parameters", block.parameters, default={}),
                _kw("editable", block.editable, default=False),
                _kw("dirty", block.dirty, default=False),
            ]
        )
        suffix = f", {kwargs}" if kwargs else ""
        _line(lines, f"block = field(doc, {_py(block.field_type)}{suffix})", indent=1)
        _emit_source_paragraph_index(lines, block)
        _assign_if_value(lines, "block", "native_field_type", block.native_field_type, default=None)
        return

    if isinstance(block, AutoNumber):
        kwargs = _kwargs(
            [
                section_kw,
                _kw("number", block.number, default=1),
                _kw("number_type", block.number_type, default="PAGE"),
                _kw("kind", block.kind, default="newNum"),
            ]
        )
        _line(lines, f"block = autonum(doc{', ' + kwargs if kwargs else ''})", indent=1)
        _emit_source_paragraph_index(lines, block)
        return

    if isinstance(block, Note):
        kwargs = _kwargs([section_kw, _kw("kind", block.kind, default="footNote"), _kw("number", block.number)])
        suffix = f", {kwargs}" if kwargs else ""
        _line(lines, f"block = note(doc, {_py(block.text)}{suffix})", indent=1)
        _emit_source_paragraph_index(lines, block)
        _emit_nested_blocks(lines, block)
        return

    if isinstance(block, Equation):
        kwargs = _kwargs(
            [
                section_kw,
                _kw("width", block.width, default=4800),
                _kw("height", block.height, default=2300),
                _kw("shape_comment", block.shape_comment),
                _kw("text_color", block.text_color, default="#000000"),
                _kw("base_unit", block.base_unit, default=1100),
                _kw("font", block.font, default="HYhwpEQ"),
            ]
        )
        suffix = f", {kwargs}" if kwargs else ""
        _line(lines, f"block = eq(doc, {_raw_py_str(block.script)}{suffix})", indent=1)
        _emit_source_paragraph_index(lines, block)
        _assign_if_value(lines, "block", "layout", block.layout, default={})
        _assign_if_value(lines, "block", "out_margins", block.out_margins, default={})
        _assign_if_value(lines, "block", "rotation", block.rotation, default={})
        return

    if isinstance(block, Shape):
        kwargs = _kwargs(
            [
                section_kw,
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
            ]
        )
        _line(lines, f"block = shape(doc{', ' + kwargs if kwargs else ''})", indent=1)
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
                section_kw,
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
            ]
        )
        suffix = f", {kwargs}" if kwargs else ""
        _line(lines, f"block = ole(doc, {_py(block.name)}, asset_data{suffix})", indent=1)
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
                section_kw,
                _kw("form_type", block.form_type, default="INPUT"),
                _kw("name", block.name),
                _kw("value", block.value),
                _kw("checked", block.checked, default=False),
                _kw("items", block.items, default=[]),
                _kw("editable", block.editable, default=True),
                _kw("locked", block.locked, default=False),
                _kw("placeholder", block.placeholder),
            ]
        )
        suffix = f", {kwargs}" if kwargs else ""
        _line(lines, f"block = form(doc, {_py(block.label)}{suffix})", indent=1)
        _emit_source_paragraph_index(lines, block)
        return

    if isinstance(block, Memo):
        kwargs = _kwargs(
            [
                section_kw,
                _kw("author", block.author),
                _kw("memo_id", block.memo_id),
                _kw("anchor_id", block.anchor_id),
                _kw("order", block.order),
                _kw("visible", block.visible, default=True),
            ]
        )
        suffix = f", {kwargs}" if kwargs else ""
        _line(lines, f"block = memo(doc, {_py(block.text)}{suffix})", indent=1)
        _emit_source_paragraph_index(lines, block)
        return

    if isinstance(block, Chart):
        kwargs = _kwargs(
            [
                section_kw,
                _kw("chart_type", block.chart_type, default="BAR"),
                _kw("categories", block.categories, default=[]),
                _kw("series", block.series, default=[]),
                _kw("data_ref", block.data_ref),
                _kw("legend_visible", block.legend_visible, default=True),
                _kw("width", block.width, default=12000),
                _kw("height", block.height, default=3200),
                _kw("shape_comment", block.shape_comment),
            ]
        )
        suffix = f", {kwargs}" if kwargs else ""
        _line(lines, f"block = chart(doc, {_py(block.title)}{suffix})", indent=1)
        _emit_source_paragraph_index(lines, block)
        _assign_if_value(lines, "block", "layout", block.layout, default={})
        _assign_if_value(lines, "block", "out_margins", block.out_margins, default={})
        _assign_if_value(lines, "block", "rotation", block.rotation, default={})
        return

    _line(lines, f"# Skipped unsupported block: {type(block).__name__}", indent=1)

_emit_authoring_block.__module__ = "jakal_hwpx.hwpx2py"

