from __future__ import annotations

from ._hancom_models import HancomBlock

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
    Path,
    Picture,
    Shape,
    Table,
    _extract_hwp_table_column_widths,
    _normalize_hwp_text,
)


def _extract_hwp_control(control) -> HancomBlock | None:
    from .hwp_document import (
        HwpAutoNumberObject,
        HwpBookmarkObject,
        HwpChartObject,
        HwpEquationObject,
        HwpFieldObject,
        HwpFormObject,
        HwpHyperlinkObject,
        HwpMemoObject,
        HwpNoteObject,
        HwpOleObject,
        HwpPictureObject,
        HwpShapeObject,
        HwpTableObject,
    )

    if isinstance(control, HwpTableObject):
        cells = control.cells()
        cell_texts = control.cell_text_matrix()
        cell_spans = {
            (cell.row, cell.column): (cell.row_span, cell.col_span)
            for cell in cells
            if cell.row_span != 1 or cell.col_span != 1
        }
        cell_border_fill_ids = {
            (cell.row, cell.column): cell.border_fill_id
            for cell in cells
        }
        cell_margins = {
            (cell.row, cell.column): {
                "left": cell.margins[0],
                "right": cell.margins[1],
                "top": cell.margins[2],
                "bottom": cell.margins[3],
            }
            for cell in cells
        }
        return Table(
            rows=control.row_count,
            cols=control.column_count,
            cell_texts=cell_texts,
            row_heights=list(control.row_heights),
            col_widths=_extract_hwp_table_column_widths(control),
            cell_spans=cell_spans or None,
            cell_border_fill_ids=cell_border_fill_ids or None,
            cell_margins=cell_margins or None,
            table_border_fill_id=control.table_border_fill_id,
            cell_spacing=control.cell_spacing,
            table_margins={
                "left": control.table_margins[0],
                "right": control.table_margins[1],
                "top": control.table_margins[2],
                "bottom": control.table_margins[3],
            },
        )
    if isinstance(control, HwpPictureObject):
        size = control.size() if hasattr(control, "size") else {"width": 7200, "height": 7200}
        path = control.bindata_path()
        line_style = dict(control.line_style()) if hasattr(control, "line_style") else {}
        return Picture(
            name=Path(path).name,
            data=control.binary_data(),
            extension=control.extension,
            width=size.get("width", 7200),
            height=size.get("height", 7200),
            shape_comment=getattr(control, "shape_comment", "") or None,
            layout=dict(control.layout()) if hasattr(control, "layout") else {},
            out_margins=dict(control.out_margins()) if hasattr(control, "out_margins") else {},
            rotation=dict(control.rotation()) if hasattr(control, "rotation") else {},
            image_adjustment=dict(control.image_adjustment()) if hasattr(control, "image_adjustment") else {},
            crop=dict(control.crop()) if hasattr(control, "crop") else {},
            line_color=getattr(control, "line_color", "#000000"),
            line_width=int(getattr(control, "line_width", 0)),
            line_style=line_style,
        )
    if isinstance(control, HwpHyperlinkObject):
        return Hyperlink(
            target=control.url,
            display_text=control.display_text or None,
            metadata_fields=list(control.metadata_fields),
        )
    if isinstance(control, HwpBookmarkObject):
        return Bookmark(name=control.name or "")
    if isinstance(control, HwpAutoNumberObject):
        return AutoNumber(
            kind=control.kind,
            number=control.number or 1,
            number_type=control.number_type or "PAGE",
        )
    if isinstance(control, HwpFormObject):
        return Form(
            label=control.label,
            form_type=control.form_type,
            name=control.name,
            value=control.value,
            checked=control.checked,
            items=list(control.items),
            editable=control.editable,
            locked=control.locked,
            placeholder=control.placeholder,
        )
    if isinstance(control, HwpMemoObject):
        return Memo(
            text=control.text,
            author=control.author,
            memo_id=control.memo_id,
            anchor_id=control.anchor_id,
            order=control.order,
            visible=control.visible,
        )
    if isinstance(control, HwpNoteObject):
        number = int(control.number) if control.number and str(control.number).isdigit() else None
        return Note(kind=control.kind, text=control.text, number=number)
    if isinstance(control, HwpOleObject):
        size = control.size()
        line_style = dict(control.line_style()) if hasattr(control, "line_style") else {}
        return Ole(
            name=control.text or "embedded.ole",
            data=control.binary_data(),
            width=size.get("width", 42001),
            height=size.get("height", 13501),
            shape_comment=getattr(control, "shape_comment", "") or None,
            object_type=getattr(control, "object_type", "EMBEDDED"),
            draw_aspect=getattr(control, "draw_aspect", "CONTENT"),
            has_moniker=bool(getattr(control, "has_moniker", False)),
            eq_baseline=int(getattr(control, "eq_baseline", 0)),
            line_color=getattr(control, "line_color", "#000000"),
            line_width=int(getattr(control, "line_width", 0)),
            line_style=line_style,
            layout=dict(control.layout()) if hasattr(control, "layout") else {},
            out_margins=dict(control.out_margins()) if hasattr(control, "out_margins") else {},
            rotation=dict(control.rotation()) if hasattr(control, "rotation") else {},
            extent=dict(control.extent()) if hasattr(control, "extent") else {},
        )
    if isinstance(control, HwpChartObject):
        size = control.size()
        return Chart(
            title=control.title,
            chart_type=control.chart_type,
            categories=list(control.categories),
            series=[dict(value) for value in control.series],
            data_ref=control.data_ref,
            legend_visible=control.legend_visible,
            width=size.get("width", 12000),
            height=size.get("height", 3200),
            shape_comment=getattr(control, "shape_comment", "") or None,
            layout=dict(control.layout()) if hasattr(control, "layout") else {},
            out_margins=dict(control.out_margins()) if hasattr(control, "out_margins") else {},
            rotation=dict(control.rotation()) if hasattr(control, "rotation") else {},
        )
    if isinstance(control, HwpShapeObject):
        size = control.size()
        line_style = dict(control.line_style()) if hasattr(control, "line_style") else {}
        return Shape(
            kind=control.kind,
            text=control.text,
            width=size.get("width", 12000),
            height=size.get("height", 3200),
            fill_color=getattr(control, "fill_color", "#FFFFFF"),
            line_color=getattr(control, "line_color", "#000000"),
            shape_comment=getattr(control, "shape_comment", "") or None,
            layout=dict(control.layout()) if hasattr(control, "layout") else {},
            out_margins=dict(control.out_margins()) if hasattr(control, "out_margins") else {},
            rotation=dict(control.rotation()) if hasattr(control, "rotation") else {},
            text_margins=dict(control.text_margins()) if hasattr(control, "text_margins") else {},
            specific_fields=dict(control.specific_fields()) if hasattr(control, "specific_fields") else {},
            line_width=int(getattr(control, "line_width", 33)),
            line_style=line_style,
        )
    if isinstance(control, HwpEquationObject):
        size = control.size() if hasattr(control, "size") else {"width": 4800, "height": 2300}
        return Equation(
            script=control.script,
            width=size.get("width", 4800),
            height=size.get("height", 2300),
            shape_comment=getattr(control, "shape_comment", "") or None,
            font=getattr(control, "font", "HYhwpEQ"),
            layout=dict(control.layout()) if hasattr(control, "layout") else {},
            out_margins=dict(control.out_margins()) if hasattr(control, "out_margins") else {},
            rotation=dict(control.rotation()) if hasattr(control, "rotation") else {},
        )
    if isinstance(control, HwpFieldObject):
        return Field(
            field_type=control.field_type,
            display_text=_normalize_hwp_text(control.display_text) or None,
            name=control.name,
            parameters=dict(control.parameters),
            editable=control.editable,
            dirty=control.dirty,
            native_field_type=control.native_field_type,
        )
    return None

_extract_hwp_control.__module__ = "jakal_hwpx.hancom_document"

