from __future__ import annotations

from .hancom_document import (
    Bookmark,
    Equation,
    Hyperlink,
    Paragraph,
    Picture,
    Shape,
    Table,
)

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from .hwpx2py import (
        _authoring_cell_border_fill_ids,
        _emit_asset_data,
        _kw,
        _macro_attrs,
        _macro_attrs_kw,
        _macro_expr,
        _py,
        _raw_py_str,
        _table_size_matches_cells,
    )

_SHARED_NAMES = ['_authoring_cell_border_fill_ids','_emit_asset_data','_kw','_macro_attrs','_macro_attrs_kw','_macro_expr','_py','_raw_py_str','_table_size_matches_cells']


def _bind_shared_helpers() -> None:
    if _SHARED_NAMES[0] in globals():
        return
    from . import hwpx2py as shared

    globals().update({name: getattr(shared, name) for name in _SHARED_NAMES})

def _macro_expr_for_block(
    lines: list[str],
    block: object,
    section_index: int,
    *,
    include_binary_assets: bool,
    asset_counter: list[int],
) -> str | None:
    _bind_shared_helpers()
    section_kw = None
    if isinstance(block, Paragraph):
        attrs = _macro_attrs(
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
        return _macro_expr(
            "text",
            [_py(block.text)],
            [
                section_kw,
                _kw("style", block.style_id, default="0"),
                _kw("para", block.para_pr_id, default="0"),
                _kw("char", block.char_pr_id, default="0"),
                _macro_attrs_kw(attrs),
            ],
        )

    if isinstance(block, Equation):
        attrs = _macro_attrs(
            block,
            {
                "script": block.script,
                "width": block.width,
                "height": block.height,
                "shape_comment": block.shape_comment,
                "text_color": block.text_color,
                "base_unit": block.base_unit,
                "font": block.font,
                "layout": {},
                "out_margins": {},
                "rotation": {},
            },
            skip={"script", "width", "height", "shape_comment", "text_color", "base_unit", "font"},
        )
        return _macro_expr(
            "math",
            [_raw_py_str(block.script)],
            [
                section_kw,
                _kw("width", block.width, default=4800),
                _kw("height", block.height, default=2300),
                _kw("shape_comment", block.shape_comment),
                _kw("text_color", block.text_color, default="#000000"),
                _kw("base_unit", block.base_unit, default=1100),
                _kw("font", block.font, default="HYhwpEQ"),
                _macro_attrs_kw(attrs),
            ],
        )

    if isinstance(block, Table):
        attrs = _macro_attrs(
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
        return _macro_expr(
            "tabular",
            [_py(block.cell_texts)],
            [
                section_kw,
                _kw("rows", block.rows) if not _table_size_matches_cells(block) else None,
                _kw("cols", block.cols) if not _table_size_matches_cells(block) else None,
                _kw("row_heights", block.row_heights),
                _kw("col_widths", block.col_widths),
                _kw("cell_spans", block.cell_spans),
                _kw("cell_border_fill_ids", _authoring_cell_border_fill_ids(block)),
                _kw("table_border_fill_id", block.table_border_fill_id, default=1),
                _macro_attrs_kw(attrs),
            ],
        )

    if isinstance(block, Shape) and not getattr(block, "nested_blocks", []):
        attrs = _macro_attrs(
            block,
            {
                "kind": block.kind,
                "text": block.text,
                "width": block.width,
                "height": block.height,
                "fill_color": block.fill_color,
                "line_color": block.line_color,
                "shape_comment": block.shape_comment,
                "layout": {},
                "out_margins": {},
                "rotation": {},
                "text_margins": {},
                "specific_fields": {},
                "line_width": block.line_width,
                "line_style": {},
                "has_line_node": True,
                "original_width": block.original_width,
                "original_height": block.original_height,
                "current_width": block.current_width,
                "current_height": block.current_height,
                "has_size_node": True,
                "size_attributes": {},
                "has_position_node": True,
                "has_out_margin_node": True,
                "has_fill_node": True,
                "has_text_margin_node": True,
                "nested_blocks": [],
            },
            skip={
                "kind",
                "text",
                "width",
                "height",
                "fill_color",
                "line_color",
                "line_width",
                "shape_comment",
                "original_width",
                "original_height",
                "current_width",
                "current_height",
            },
        )
        return _macro_expr(
            "draw",
            [],
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
                _macro_attrs_kw(attrs),
            ],
        )

    if isinstance(block, Picture):
        if not include_binary_assets:
            return None
        asset_counter[0] += 1
        asset_name = f"asset_data_{asset_counter[0]}"
        _emit_asset_data(lines, block.data, name=asset_name)
        attrs = _macro_attrs(
            block,
            {
                "name": block.name,
                "data": block.data,
                "extension": block.extension,
                "width": block.width,
                "height": block.height,
                "shape_comment": None,
                "layout": {},
                "out_margins": {},
                "rotation": {},
                "image_adjustment": {},
                "crop": {},
                "line_color": "#000000",
                "line_width": 0,
                "original_width": block.original_width,
                "original_height": block.original_height,
                "current_width": block.current_width,
                "current_height": block.current_height,
                "has_size_node": True,
                "size_attributes": {},
                "has_line_node": False,
                "has_position_node": True,
                "has_out_margin_node": True,
            },
            skip={"name", "data", "extension", "width", "height", "original_width", "original_height", "current_width", "current_height"},
        )
        return _macro_expr(
            "figure",
            [_py(block.name), asset_name],
            [
                section_kw,
                _kw("extension", block.extension),
                _kw("width", block.width, default=7200),
                _kw("height", block.height, default=7200),
                _kw("original_width", block.original_width),
                _kw("original_height", block.original_height),
                _kw("current_width", block.current_width),
                _kw("current_height", block.current_height),
                _macro_attrs_kw(attrs),
            ],
        )

    if isinstance(block, Hyperlink):
        attrs = _macro_attrs(block, {"target": block.target, "display_text": block.display_text, "metadata_fields": []}, skip={"target", "display_text", "metadata_fields"})
        return _macro_expr(
            "href",
            [_py(block.target)],
            [section_kw, _kw("display_text", block.display_text), _kw("metadata_fields", block.metadata_fields, default=[]), _macro_attrs_kw(attrs)],
        )

    if isinstance(block, Bookmark):
        attrs = _macro_attrs(block, {"name": block.name}, skip={"name"})
        return _macro_expr("label", [_py(block.name)], [_macro_attrs_kw(attrs)])

    return None

_macro_expr_for_block.__module__ = "jakal_hwpx.hwpx2py"

