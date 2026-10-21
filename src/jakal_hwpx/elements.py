from __future__ import annotations

from dataclasses import dataclass

from lxml import etree

from .exceptions import ValidationIssue
from .namespaces import NS, qname


_FORM_TAG_BY_TYPE = {
    "CHECKBOX": "checkBtn",
    "RADIOBUTTON": "radioBtn",
    "BUTTON": "btn",
    "INPUT": "edit",
    "EDIT": "edit",
    "COMBOBOX": "comboBox",
    "LISTBOX": "listBox",
    "SCROLLBAR": "scrollBar",
}
_FORM_TYPE_BY_TAG = {
    "checkBtn": "CHECKBOX",
    "radioBtn": "RADIOBUTTON",
    "btn": "BUTTON",
    "edit": "INPUT",
    "comboBox": "COMBOBOX",
    "listBox": "LISTBOX",
    "scrollBar": "SCROLLBAR",
}
_BUTTON_FORM_TAGS = {"checkBtn", "radioBtn", "btn"}
_LIST_FORM_TAGS = {"comboBox", "listBox"}
_FORM_METADATA_COMMAND_PREFIX = "JAKAL_FORM_META:"
from ._hwpx_chart_xml import (  # noqa: E402, F401 - compatibility re-exports
    _CHART_METADATA_URI,
    _CHART_NS,
    _CHART_TYPE_ELEMENT_BY_NAME,
    _CHART_TYPE_NAME_BY_ELEMENT,
    _HANCOM_CHART_STYLE_URI,
    _build_chart_part_root,
    _build_chart_series_node,
    _build_chart_text_properties,
    _chart_categories_from_root,
    _chart_data_ref_from_root,
    _chart_formula_column_name,
    _chart_formula_range,
    _chart_formula_sheet_name,
    _chart_metadata_from_root,
    _chart_part_root,
    _chart_qname,
    _chart_series_from_root,
    _chart_sheet_anchor_from_formula,
    _chart_title_from_root,
    _chart_type_from_root,
    _ensure_chart_hancom_extension,
    _ensure_chart_plot_area_style,
    _ensure_chart_root_text_properties,
    _ensure_chart_style_block,
    _set_chart_legend_visible,
    _set_chart_metadata,
    _set_chart_title,
)


from ._hwpx_element_helpers import (  # noqa: E402, F401 - compatibility re-exports
    _text_nodes,
    _first_node,
    _bool_attr,
    _set_optional_attributes,
    _optional_int_attr,
    _graphic_layout,
    _set_graphic_layout,
    _margin_values,
    _set_margin_values,
    _graphic_size,
    _set_graphic_size,
    _graphic_rotation,
    _set_graphic_rotation,
    _sync_rotation_center_for_size,
    _sync_picture_size_geometry,
    _sync_shape_point_geometry,
    _sync_extent_for_size,
    _line_style,
    _set_line_style,
    _fill_style,
    _set_fill_style,
    _text_margin,
    _set_text_margin,
    _image_adjustment,
    _set_image_adjustment,
    _crop_values,
    _set_crop_values,
    _extract_text,
    _replace_text,
    _paragraphs_affected_by_text_edit,
    _invalidate_paragraph_layout,
    _paragraph_nodes,
    _ensure_first_paragraph,
    _default_char_pr_id,
    _ensure_run_with_text,
    _reset_paragraph_text,
    _replace_paragraph_text_preserving_controls,
    _preserved_structure_signature,
    _capture_protected_paragraph_signatures,
    _missing_preserved_tokens,
    _clone_paragraph,
    _split_score,
    _snap_split_index,
    _distribute_text_across_paragraphs,
    _set_text,
    _form_command_metadata,
    _write_form_command_metadata,
    _form_list_item_values,
    _set_form_list_items,
    _form_caption_text,
    _set_form_caption_text,
    _first_form_text_node,
)


@dataclass
class HeaderFooterXml:
    document: object
    section: object
    element: etree._Element

    @property
    def kind(self) -> str:
        return etree.QName(self.element).localname

    @property
    def apply_page_type(self) -> str | None:
        return self.element.get("applyPageType")

    def set_apply_page_type(self, value: str) -> None:
        self.element.set("applyPageType", value)
        self.section.mark_modified()

    @property
    def text(self) -> str:
        return _extract_text(self.element)

    def replace_text(self, old: str, new: str, count: int = -1) -> int:
        self.section.mark_modified()
        return _replace_text(self.element, old, new, count=count)

    def set_text(self, text: str) -> None:
        paragraphs, signatures = _capture_protected_paragraph_signatures(self.element)
        self.section.mark_modified()
        _set_text(self.element, text)
        self.document._track_control_preservation_targets(
            paragraphs,
            signatures,
            issues=[
                ValidationIssue(
                    kind="control_preservation",
                    message="",
                    part_path=self.section.path,
                    section_index=self.section.section_index(),
                    paragraph_index=index,
                    context=f"{self.kind}(applyPageType={self.apply_page_type or 'UNKNOWN'})",
                )
                for index, _ in enumerate(paragraphs)
            ],
        )


@dataclass
class TableCellXml:
    table: "TableXml"
    element: etree._Element

    @property
    def row(self) -> int:
        addr = self.element.xpath("./hp:cellAddr/@rowAddr", namespaces=NS)
        return int(addr[0]) if addr else 0

    @property
    def column(self) -> int:
        addr = self.element.xpath("./hp:cellAddr/@colAddr", namespaces=NS)
        return int(addr[0]) if addr else 0

    @property
    def text(self) -> str:
        return _extract_text(self.element)

    @property
    def row_span(self) -> int:
        span = self.element.xpath("./hp:cellSpan/@rowSpan", namespaces=NS)
        return int(span[0]) if span else 1

    @property
    def col_span(self) -> int:
        span = self.element.xpath("./hp:cellSpan/@colSpan", namespaces=NS)
        return int(span[0]) if span else 1

    @property
    def border_fill_id(self) -> int:
        return int(self.element.get("borderFillIDRef", "0") or 0)

    @property
    def margins(self) -> dict[str, int]:
        return _margin_values(self.element, "./hp:cellMargin")

    @property
    def vertical_align(self) -> str:
        values = self.element.xpath("./hp:subList/@vertAlign", namespaces=NS)
        return values[0] if values else "CENTER"

    def set_text(self, text: str) -> None:
        paragraphs, signatures = _capture_protected_paragraph_signatures(self.element)
        self.table.section.mark_modified()
        _set_text(self.element, text)
        self.table.document._track_control_preservation_targets(
            paragraphs,
            signatures,
            issues=[
                ValidationIssue(
                    kind="control_preservation",
                    message="",
                    part_path=self.table.section.path,
                    section_index=self.table.section.section_index(),
                    paragraph_index=index,
                    cell_row=self.row,
                    cell_column=self.column,
                    context=f"table cell(row={self.row}, column={self.column})",
                )
                for index, _ in enumerate(paragraphs)
            ],
        )

    def set_border_fill_id(self, value: int | str) -> None:
        self.element.set("borderFillIDRef", str(value))
        self.table.section.mark_modified()

    def set_margins(
        self,
        *,
        left: int | str | None = None,
        right: int | str | None = None,
        top: int | str | None = None,
        bottom: int | str | None = None,
    ) -> None:
        _set_margin_values(self.element, "./hp:cellMargin", left=left, right=right, top=top, bottom=bottom)
        self.table.section.mark_modified()

    def set_vertical_align(self, value: str) -> None:
        sub_list = _first_node(self.element, "./hp:subList")
        if sub_list is None:
            raise ValueError("Table cell does not contain hp:subList.")
        sub_list.set("vertAlign", value)
        self.table.section.mark_modified()


@dataclass
class TableXml:
    document: object
    section: object
    element: etree._Element

    @property
    def row_count(self) -> int:
        value = self.element.get("rowCnt")
        return int(value) if value else len(self.element.xpath("./hp:tr", namespaces=NS))

    @property
    def column_count(self) -> int:
        value = self.element.get("colCnt")
        return int(value) if value else 0

    def cells(self) -> list[TableCellXml]:
        return [TableCellXml(self, cell) for cell in self.element.xpath("./hp:tr/hp:tc", namespaces=NS)]

    def cell(self, row: int, column: int) -> TableCellXml:
        for cell in self.cells():
            if cell.row == row and cell.column == column:
                return cell
        raise IndexError(f"Cell ({row}, {column}) was not found.")

    def set_cell_text(self, row: int, column: int, text: str) -> None:
        self.cell(row, column).set_text(text)

    def rows(self) -> list[list[TableCellXml]]:
        grouped: dict[int, list[TableCellXml]] = {}
        for cell in self.cells():
            grouped.setdefault(cell.row, []).append(cell)
        return [sorted(grouped[index], key=lambda item: item.column) for index in sorted(grouped)]

    def layout(self) -> dict[str, str]:
        return _graphic_layout(self.element)

    def out_margins(self) -> dict[str, int]:
        return _margin_values(self.element, "./hp:outMargin")

    def in_margins(self) -> dict[str, int]:
        return _margin_values(self.element, "./hp:inMargin")

    @property
    def cell_spacing(self) -> int:
        return int(self.element.get("cellSpacing", "0") or 0)

    @property
    def table_border_fill_id(self) -> int:
        return int(self.element.get("borderFillIDRef", "0") or 0)

    @property
    def page_break(self) -> str:
        return self.element.get("pageBreak", "CELL")

    @property
    def repeat_header(self) -> bool:
        return self.element.get("repeatHeader", "0") == "1"

    def set_layout(
        self,
        *,
        text_wrap: str | None = None,
        text_flow: str | None = None,
        treat_as_char: bool | None = None,
        affect_line_spacing: bool | None = None,
        flow_with_text: bool | None = None,
        allow_overlap: bool | None = None,
        hold_anchor_and_so: bool | None = None,
        vert_rel_to: str | None = None,
        horz_rel_to: str | None = None,
        vert_align: str | None = None,
        horz_align: str | None = None,
        vert_offset: int | str | None = None,
        horz_offset: int | str | None = None,
    ) -> None:
        _set_graphic_layout(
            self.element,
            text_wrap=text_wrap,
            text_flow=text_flow,
            treat_as_char=treat_as_char,
            affect_line_spacing=affect_line_spacing,
            flow_with_text=flow_with_text,
            allow_overlap=allow_overlap,
            hold_anchor_and_so=hold_anchor_and_so,
            vert_rel_to=vert_rel_to,
            horz_rel_to=horz_rel_to,
            vert_align=vert_align,
            horz_align=horz_align,
            vert_offset=vert_offset,
            horz_offset=horz_offset,
        )
        self.section.mark_modified()

    def set_out_margins(
        self,
        *,
        left: int | str | None = None,
        right: int | str | None = None,
        top: int | str | None = None,
        bottom: int | str | None = None,
    ) -> None:
        _set_margin_values(self.element, "./hp:outMargin", left=left, right=right, top=top, bottom=bottom)
        self.section.mark_modified()

    def set_in_margins(
        self,
        *,
        left: int | str | None = None,
        right: int | str | None = None,
        top: int | str | None = None,
        bottom: int | str | None = None,
    ) -> None:
        _set_margin_values(self.element, "./hp:inMargin", left=left, right=right, top=top, bottom=bottom)
        self.section.mark_modified()

    def set_cell_spacing(self, value: int | str) -> None:
        self.element.set("cellSpacing", str(value))
        self.section.mark_modified()

    def set_table_border_fill_id(self, value: int | str) -> None:
        self.element.set("borderFillIDRef", str(value))
        self.section.mark_modified()

    def set_page_break(self, value: str) -> None:
        self.element.set("pageBreak", value)
        self.section.mark_modified()

    def set_repeat_header(self, value: bool) -> None:
        self.element.set("repeatHeader", "1" if value else "0")
        self.section.mark_modified()

    def append_row(self) -> list[TableCellXml]:
        rows = self.element.xpath("./hp:tr", namespaces=NS)
        if not rows:
            raise ValueError("Table does not contain rows.")
        template = rows[-1]
        protected = [
            token
            for paragraph in template.xpath(".//hp:p", namespaces=NS)
            for token in _preserved_structure_signature(paragraph).elements()
        ]
        if protected:
            raise ValueError(
                "append_row() cannot clone a template row containing preserved controls. "
                f"Unsupported template row nodes: {', '.join(protected)}"
            )
        new_row = etree.fromstring(etree.tostring(template))
        next_row_index = max((cell.row for cell in self.cells()), default=-1) + 1
        for column_index, cell in enumerate(new_row.xpath("./hp:tc", namespaces=NS)):
            for paragraph in cell.xpath(".//hp:p", namespaces=NS):
                for child in list(paragraph):
                    paragraph.remove(child)
                _ensure_run_with_text(paragraph, "")
            cell_addr = cell.xpath("./hp:cellAddr", namespaces=NS)
            if cell_addr:
                cell_addr[0].set("rowAddr", str(next_row_index))
                cell_addr[0].set("colAddr", str(column_index))
        self.element.append(new_row)
        self.element.set("rowCnt", str(self.row_count + 1))
        self.section.mark_modified()
        return [TableCellXml(self, cell) for cell in new_row.xpath("./hp:tc", namespaces=NS)]

    def merge_cells(self, start_row: int, start_column: int, end_row: int, end_column: int) -> TableCellXml:
        if end_row < start_row or end_column < start_column:
            raise ValueError("Invalid merge range.")
        anchor = self.cell(start_row, start_column)
        anchor_element = anchor.element
        anchor_span = anchor.element.xpath("./hp:cellSpan", namespaces=NS)
        if anchor_span:
            span = anchor_span[0]
        else:
            span = etree.SubElement(anchor.element, qname("hp", "cellSpan"))
        span.set("rowSpan", str(end_row - start_row + 1))
        span.set("colSpan", str(end_column - start_column + 1))

        for cell in list(self.cells()):
            if cell.element is anchor_element:
                continue
            if start_row <= cell.row <= end_row and start_column <= cell.column <= end_column:
                parent = cell.element.getparent()
                if parent is not None:
                    parent.remove(cell.element)
        self.section.mark_modified()
        return anchor


@dataclass
class PictureXml:
    document: object
    section: object
    element: etree._Element

    @property
    def binary_item_id(self) -> str | None:
        values = self.element.xpath("./hc:img/@binaryItemIDRef", namespaces=NS)
        return values[0] if values else None

    @property
    def shape_comment(self) -> str:
        values = self.element.xpath("./hp:shapeComment", namespaces=NS)
        if not values:
            return ""
        return values[0].text or ""

    @shape_comment.setter
    def shape_comment(self, value: str) -> None:
        comment = self.element.xpath("./hp:shapeComment", namespaces=NS)
        if comment:
            comment[0].text = value
        else:
            node = etree.SubElement(self.element, qname("hp", "shapeComment"))
            node.text = value
        self.section.mark_modified()

    def binary_part_path(self) -> str:
        item_id = self.binary_item_id
        if item_id is None:
            raise ValueError("Picture is not bound to a binary manifest item.")
        for item in self.document.content_hpf.manifest_items():
            if item.get("id") == item_id:
                href = item.get("href")
                if href:
                    return href
        raise KeyError(f"Manifest item {item_id} was not found.")

    def binary_data(self) -> bytes:
        part = self.document.get_part(self.binary_part_path())
        return part.data

    def replace_binary(self, data: bytes) -> None:
        part = self.document.get_part(self.binary_part_path())
        part.data = data

    def bind_binary_item(self, item_id: str) -> None:
        image_nodes = self.element.xpath("./hc:img", namespaces=NS)
        if not image_nodes:
            raise ValueError("Picture does not contain an hc:img node.")
        image_nodes[0].set("binaryItemIDRef", item_id)
        self.section.mark_modified()

    def layout(self) -> dict[str, str]:
        return _graphic_layout(self.element)

    def size(self) -> dict[str, int]:
        return _graphic_size(self.element)

    def rotation(self) -> dict[str, str]:
        return _graphic_rotation(self.element)

    def out_margins(self) -> dict[str, int]:
        return _margin_values(self.element, "./hp:outMargin")

    def image_adjustment(self) -> dict[str, str]:
        return _image_adjustment(self.element)

    def crop(self) -> dict[str, int]:
        return _crop_values(self.element)

    def line_style(self) -> dict[str, str]:
        return _line_style(self.element)

    def set_layout(
        self,
        *,
        text_wrap: str | None = None,
        text_flow: str | None = None,
        treat_as_char: bool | None = None,
        affect_line_spacing: bool | None = None,
        flow_with_text: bool | None = None,
        allow_overlap: bool | None = None,
        hold_anchor_and_so: bool | None = None,
        vert_rel_to: str | None = None,
        horz_rel_to: str | None = None,
        vert_align: str | None = None,
        horz_align: str | None = None,
        vert_offset: int | str | None = None,
        horz_offset: int | str | None = None,
    ) -> None:
        _set_graphic_layout(
            self.element,
            text_wrap=text_wrap,
            text_flow=text_flow,
            treat_as_char=treat_as_char,
            affect_line_spacing=affect_line_spacing,
            flow_with_text=flow_with_text,
            allow_overlap=allow_overlap,
            hold_anchor_and_so=hold_anchor_and_so,
            vert_rel_to=vert_rel_to,
            horz_rel_to=horz_rel_to,
            vert_align=vert_align,
            horz_align=horz_align,
            vert_offset=vert_offset,
            horz_offset=horz_offset,
        )
        self.section.mark_modified()

    def set_out_margins(
        self,
        *,
        left: int | str | None = None,
        right: int | str | None = None,
        top: int | str | None = None,
        bottom: int | str | None = None,
    ) -> None:
        _set_margin_values(self.element, "./hp:outMargin", left=left, right=right, top=top, bottom=bottom)
        self.section.mark_modified()

    def set_size(
        self,
        *,
        width: int | str | None = None,
        height: int | str | None = None,
        original_width: int | str | None = None,
        original_height: int | str | None = None,
        current_width: int | str | None = None,
        current_height: int | str | None = None,
    ) -> None:
        _set_graphic_size(
            self.element,
            width=width,
            height=height,
            original_width=original_width,
            original_height=original_height,
            current_width=current_width,
            current_height=current_height,
        )
        self.section.mark_modified()

    def set_rotation(
        self,
        *,
        angle: int | str | None = None,
        center_x: int | str | None = None,
        center_y: int | str | None = None,
        rotate_image: bool | None = None,
    ) -> None:
        _set_graphic_rotation(
            self.element,
            angle=angle,
            center_x=center_x,
            center_y=center_y,
            rotate_image=rotate_image,
        )
        self.section.mark_modified()

    def set_image_adjustment(
        self,
        *,
        bright: int | str | None = None,
        contrast: int | str | None = None,
        effect: str | None = None,
        alpha: int | str | None = None,
    ) -> None:
        _set_image_adjustment(self.element, bright=bright, contrast=contrast, effect=effect, alpha=alpha)
        self.section.mark_modified()

    def set_crop(
        self,
        *,
        left: int | str | None = None,
        right: int | str | None = None,
        top: int | str | None = None,
        bottom: int | str | None = None,
    ) -> None:
        _set_crop_values(self.element, left=left, right=right, top=top, bottom=bottom)
        self.section.mark_modified()

    def set_line_style(
        self,
        *,
        color: str | None = None,
        width: int | str | None = None,
        style: str | None = None,
        end_cap: str | None = None,
        head_style: str | None = None,
        tail_style: str | None = None,
        head_fill: bool | None = None,
        tail_fill: bool | None = None,
        head_size: str | None = None,
        tail_size: str | None = None,
        outline_style: str | None = None,
        alpha: int | str | None = None,
    ) -> None:
        _set_line_style(
            self.element,
            color=color,
            width=width,
            style=style,
            end_cap=end_cap,
            head_style=head_style,
            tail_style=tail_style,
            head_fill=head_fill,
            tail_fill=tail_fill,
            head_size=head_size,
            tail_size=tail_size,
            outline_style=outline_style,
            alpha=alpha,
        )
        self.section.mark_modified()


@dataclass
class StyleDefinitionXml:
    header_part: object
    element: etree._Element

    @property
    def style_id(self) -> str | None:
        return self.element.get("id")

    @property
    def name(self) -> str | None:
        return self.element.get("name")

    @property
    def english_name(self) -> str | None:
        return self.element.get("engName")

    @property
    def para_pr_id(self) -> str | None:
        return self.element.get("paraPrIDRef")

    @property
    def char_pr_id(self) -> str | None:
        return self.element.get("charPrIDRef")

    def set_name(self, value: str) -> None:
        self.element.set("name", value)
        self.header_part.mark_modified()

    def set_english_name(self, value: str) -> None:
        self.element.set("engName", value)
        self.header_part.mark_modified()

    def bind_refs(self, *, para_pr_id: str | None = None, char_pr_id: str | None = None) -> None:
        if para_pr_id is not None:
            self.element.set("paraPrIDRef", para_pr_id)
        if char_pr_id is not None:
            self.element.set("charPrIDRef", char_pr_id)
        self.header_part.mark_modified()

    def configure(
        self,
        *,
        style_type: str | None = None,
        para_pr_id: str | None = None,
        char_pr_id: str | None = None,
        next_style_id: str | None = None,
        lang_id: str | None = None,
        lock_form: bool | None = None,
    ) -> None:
        _set_optional_attributes(
            self.element,
            type=style_type,
            paraPrIDRef=para_pr_id,
            charPrIDRef=char_pr_id,
            nextStyleIDRef=next_style_id,
            langID=lang_id,
            lockForm=lock_form,
        )
        self.header_part.mark_modified()


@dataclass
class ParagraphStyleXml:
    header_part: object
    element: etree._Element

    @property
    def style_id(self) -> str | None:
        return self.element.get("id")

    @property
    def alignment_horizontal(self) -> str | None:
        nodes = self.element.xpath("./hh:align/@horizontal", namespaces=NS)
        return nodes[0] if nodes else None

    @property
    def line_spacing(self) -> str | None:
        nodes = self.element.xpath(".//hh:lineSpacing/@value", namespaces=NS)
        return nodes[0] if nodes else None

    def set_alignment(self, *, horizontal: str | None = None, vertical: str | None = None) -> None:
        nodes = self.element.xpath("./hh:align", namespaces=NS)
        align = nodes[0] if nodes else etree.SubElement(self.element, qname("hh", "align"))
        if horizontal is not None:
            align.set("horizontal", horizontal)
        if vertical is not None:
            align.set("vertical", vertical)
        self.header_part.mark_modified()

    def set_line_spacing(self, value: int | str, spacing_type: str | None = None) -> None:
        nodes = self.element.xpath(".//hh:lineSpacing", namespaces=NS)
        if not nodes:
            return
        for node in nodes:
            node.set("value", str(value))
            if spacing_type is not None:
                node.set("type", spacing_type)
        self.header_part.mark_modified()

    def set_margin(
        self,
        *,
        intent: int | str | None = None,
        left: int | str | None = None,
        right: int | str | None = None,
        prev: int | str | None = None,
        next: int | str | None = None,
        unit: str | None = None,
    ) -> None:
        margin = _first_node(self.element, "./hh:margin")
        if margin is None:
            return
        for key, value in {
            "intent": intent,
            "left": left,
            "right": right,
            "prev": prev,
            "next": next,
        }.items():
            if value is None:
                continue
            node = _first_node(margin, f"./hc:{key}")
            if node is None:
                node = etree.SubElement(margin, qname("hc", key))
            node.set("value", str(value))
            if unit is not None:
                node.set("unit", unit)
        self.header_part.mark_modified()

    def set_break_setting(
        self,
        *,
        break_latin_word: str | None = None,
        break_non_latin_word: str | None = None,
        widow_orphan: bool | None = None,
        keep_with_next: bool | None = None,
        keep_lines: bool | None = None,
        page_break_before: bool | None = None,
        line_wrap: str | None = None,
    ) -> None:
        node = _first_node(self.element, "./hh:breakSetting")
        _set_optional_attributes(
            node,
            breakLatinWord=break_latin_word,
            breakNonLatinWord=break_non_latin_word,
            widowOrphan=widow_orphan,
            keepWithNext=keep_with_next,
            keepLines=keep_lines,
            pageBreakBefore=page_break_before,
            lineWrap=line_wrap,
        )
        self.header_part.mark_modified()

    def set_auto_spacing(self, *, e_asian_eng: bool | None = None, e_asian_num: bool | None = None) -> None:
        _set_optional_attributes(
            _first_node(self.element, "./hh:autoSpacing"),
            eAsianEng=e_asian_eng,
            eAsianNum=e_asian_num,
        )
        self.header_part.mark_modified()


@dataclass
class CharacterStyleXml:
    header_part: object
    element: etree._Element

    @property
    def style_id(self) -> str | None:
        return self.element.get("id")

    @property
    def text_color(self) -> str | None:
        return self.element.get("textColor")

    @property
    def height(self) -> str | None:
        return self.element.get("height")

    def set_text_color(self, color: str) -> None:
        self.element.set("textColor", color)
        self.header_part.mark_modified()

    def set_height(self, value: int | str) -> None:
        self.element.set("height", str(value))
        self.header_part.mark_modified()

    def set_font_refs(
        self,
        *,
        hangul: str | None = None,
        latin: str | None = None,
        hanja: str | None = None,
        japanese: str | None = None,
        other: str | None = None,
        symbol: str | None = None,
        user: str | None = None,
    ) -> None:
        font_ref = _first_node(self.element, "./hh:fontRef")
        _set_optional_attributes(
            font_ref,
            hangul=hangul,
            latin=latin,
            hanja=hanja,
            japanese=japanese,
            other=other,
            symbol=symbol,
            user=user,
        )
        self.header_part.mark_modified()

    def set_relative_shape(
        self,
        tag_name: str,
        *,
        hangul: int | str | None = None,
        latin: int | str | None = None,
        hanja: int | str | None = None,
        japanese: int | str | None = None,
        other: int | str | None = None,
        symbol: int | str | None = None,
        user: int | str | None = None,
    ) -> None:
        node = _first_node(self.element, f"./hh:{tag_name}")
        _set_optional_attributes(
            node,
            hangul=hangul,
            latin=latin,
            hanja=hanja,
            japanese=japanese,
            other=other,
            symbol=symbol,
            user=user,
        )
        self.header_part.mark_modified()

    def set_underline(
        self,
        *,
        underline_type: str | None = None,
        shape: str | None = None,
        color: str | None = None,
    ) -> None:
        node = _first_node(self.element, "./hh:underline")
        _set_optional_attributes(node, type=underline_type, shape=shape, color=color)
        self.header_part.mark_modified()

    def set_effects(
        self,
        *,
        shade_color: str | None = None,
        use_font_space: bool | None = None,
        use_kerning: bool | None = None,
        sym_mark: str | None = None,
    ) -> None:
        _set_optional_attributes(
            self.element,
            shadeColor=shade_color,
            useFontSpace=use_font_space,
            useKerning=use_kerning,
            symMark=sym_mark,
        )
        self.header_part.mark_modified()


def _optional_int_attribute(element: etree._Element, name: str) -> int | None:
    value = element.get(name)
    if value in (None, ""):
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


@dataclass
class MemoShapeXml:
    header_part: object
    element: etree._Element

    @property
    def memo_shape_id(self) -> str | None:
        return self.element.get("id")

    @property
    def width(self) -> int | None:
        return _optional_int_attribute(self.element, "width")

    @property
    def line_width(self) -> int | None:
        return _optional_int_attribute(self.element, "lineWidth")

    @property
    def line_type(self) -> str | None:
        return self.element.get("lineType")

    @property
    def line_color(self) -> str | None:
        return self.element.get("lineColor")

    @property
    def fill_color(self) -> str | None:
        return self.element.get("fillColor")

    @property
    def active_color(self) -> str | None:
        return self.element.get("activeColor")

    @property
    def memo_type(self) -> str | None:
        return self.element.get("memoType")

    def configure(
        self,
        *,
        memo_shape_id: int | str | None = None,
        width: int | str | None = None,
        line_width: int | str | None = None,
        line_type: str | None = None,
        line_color: str | None = None,
        fill_color: str | None = None,
        active_color: str | None = None,
        memo_type: str | None = None,
    ) -> None:
        _set_optional_attributes(
            self.element,
            id=memo_shape_id,
            width=width,
            lineWidth=line_width,
            lineType=line_type,
            lineColor=line_color,
            fillColor=fill_color,
            activeColor=active_color,
            memoType=memo_type,
        )
        self.header_part.mark_modified()

    def set_width(self, value: int | str) -> None:
        self.element.set("width", str(value))
        self.header_part.mark_modified()

    def set_line_width(self, value: int | str) -> None:
        self.element.set("lineWidth", str(value))
        self.header_part.mark_modified()

    def set_line_style(
        self,
        *,
        line_type: str | None = None,
        line_color: str | None = None,
        fill_color: str | None = None,
        active_color: str | None = None,
    ) -> None:
        _set_optional_attributes(
            self.element,
            lineType=line_type,
            lineColor=line_color,
            fillColor=fill_color,
            activeColor=active_color,
        )
        self.header_part.mark_modified()

    def set_memo_type(self, value: str) -> None:
        self.element.set("memoType", value)
        self.header_part.mark_modified()


@dataclass
class SectionSettingsXml:
    section: object
    element: etree._Element

    @property
    def page_width(self) -> int | None:
        nodes = self.element.xpath("./hp:pagePr/@width", namespaces=NS)
        return int(nodes[0]) if nodes else None

    @property
    def page_height(self) -> int | None:
        nodes = self.element.xpath("./hp:pagePr/@height", namespaces=NS)
        return int(nodes[0]) if nodes else None

    @property
    def landscape(self) -> str | None:
        nodes = self.element.xpath("./hp:pagePr/@landscape", namespaces=NS)
        return nodes[0] if nodes else None

    @property
    def memo_shape_id(self) -> str | None:
        return self.element.get("memoShapeIDRef")

    def set_page_size(self, *, width: int | None = None, height: int | None = None, landscape: str | None = None) -> None:
        page_pr_nodes = self.element.xpath("./hp:pagePr", namespaces=NS)
        if not page_pr_nodes:
            raise ValueError("Section does not contain hp:pagePr.")
        page_pr = page_pr_nodes[0]
        if width is not None:
            page_pr.set("width", str(width))
        if height is not None:
            page_pr.set("height", str(height))
        if landscape is not None:
            page_pr.set("landscape", landscape)
        self.section.mark_modified()

    def margins(self) -> dict[str, int]:
        margin_nodes = self.element.xpath("./hp:pagePr/hp:margin", namespaces=NS)
        if not margin_nodes:
            return {}
        margin = margin_nodes[0]
        keys = ["header", "footer", "gutter", "left", "right", "top", "bottom"]
        return {key: int(margin.get(key, "0")) for key in keys}

    def set_margins(self, **values: int) -> None:
        margin_nodes = self.element.xpath("./hp:pagePr/hp:margin", namespaces=NS)
        if not margin_nodes:
            raise ValueError("Section does not contain hp:margin.")
        margin = margin_nodes[0]
        for key, value in values.items():
            margin.set(key, str(value))
        self.section.mark_modified()

    def visibility(self) -> dict[str, str]:
        node = _first_node(self.element, "./hp:visibility")
        if node is None:
            return {}
        return {
            key: node.get(key, "")
            for key in (
                "hideFirstHeader",
                "hideFirstFooter",
                "hideFirstMasterPage",
                "border",
                "fill",
                "hideFirstPageNum",
                "hideFirstEmptyLine",
                "showLineNumber",
            )
        }

    def set_visibility(
        self,
        *,
        hide_first_header: bool | None = None,
        hide_first_footer: bool | None = None,
        hide_first_master_page: bool | None = None,
        border: str | None = None,
        fill: str | None = None,
        hide_first_page_num: bool | None = None,
        hide_first_empty_line: bool | None = None,
        show_line_number: bool | None = None,
    ) -> None:
        node = _first_node(self.element, "./hp:visibility")
        _set_optional_attributes(
            node,
            hideFirstHeader=hide_first_header,
            hideFirstFooter=hide_first_footer,
            hideFirstMasterPage=hide_first_master_page,
            border=border,
            fill=fill,
            hideFirstPageNum=hide_first_page_num,
            hideFirstEmptyLine=hide_first_empty_line,
            showLineNumber=show_line_number,
        )
        self.section.mark_modified()

    def grid(self) -> dict[str, int]:
        node = _first_node(self.element, "./hp:grid")
        if node is None:
            return {}
        return {
            key: int(node.get(key, "0"))
            for key in ("lineGrid", "charGrid", "wonggojiFormat")
        }

    def set_grid(
        self,
        *,
        line_grid: int | str | None = None,
        char_grid: int | str | None = None,
        wonggoji_format: bool | None = None,
    ) -> None:
        _set_optional_attributes(
            _first_node(self.element, "./hp:grid"),
            lineGrid=line_grid,
            charGrid=char_grid,
            wonggojiFormat=wonggoji_format,
        )
        self.section.mark_modified()

    def start_numbers(self) -> dict[str, str]:
        node = _first_node(self.element, "./hp:startNum")
        if node is None:
            return {}
        return {
            key: node.get(key, "")
            for key in ("pageStartsOn", "page", "pic", "tbl", "equation")
        }

    def set_start_numbers(
        self,
        *,
        page_starts_on: str | None = None,
        page: int | str | None = None,
        pic: int | str | None = None,
        tbl: int | str | None = None,
        equation: int | str | None = None,
    ) -> None:
        _set_optional_attributes(
            _first_node(self.element, "./hp:startNum"),
            pageStartsOn=page_starts_on,
            page=page,
            pic=pic,
            tbl=tbl,
            equation=equation,
        )
        self.section.mark_modified()

    def set_memo_shape_id(self, value: int | str | None) -> None:
        self.element.set("memoShapeIDRef", "0" if value is None else str(value))
        self.section.mark_modified()


@dataclass
class NoteXml:
    document: object
    section: object
    element: etree._Element

    @property
    def kind(self) -> str:
        return etree.QName(self.element).localname

    @property
    def number(self) -> str | None:
        return self.element.get("number")

    @property
    def text(self) -> str:
        return _extract_text(self.element)

    def set_text(self, text: str) -> None:
        paragraphs, signatures = _capture_protected_paragraph_signatures(self.element)
        self.section.mark_modified()
        _set_text(self.element, text)
        self.document._track_control_preservation_targets(
            paragraphs,
            signatures,
            issues=[
                ValidationIssue(
                    kind="control_preservation",
                    message="",
                    part_path=self.section.path,
                    section_index=self.section.section_index(),
                    paragraph_index=index,
                    context=f"{self.kind}(number={self.number or 'UNKNOWN'})",
                )
                for index, _ in enumerate(paragraphs)
            ],
        )

    def set_number(self, value: int | str) -> None:
        self.element.set("number", str(value))
        self.section.mark_modified()


@dataclass
class MemoXml:
    document: object
    section: object
    element: etree._Element

    @property
    def text(self) -> str:
        return _extract_text(self.element)

    def set_text(self, text: str) -> None:
        paragraphs, signatures = _capture_protected_paragraph_signatures(self.element)
        self.section.mark_modified()
        _set_text(self.element, text)
        self.document._track_control_preservation_targets(
            paragraphs,
            signatures,
            issues=[
                ValidationIssue(
                    kind="control_preservation",
                    message="",
                    part_path=self.section.path,
                    section_index=self.section.section_index(),
                    paragraph_index=index,
                    context="hiddenComment",
                )
                for index, _ in enumerate(paragraphs)
            ],
        )


@dataclass
class FormXml:
    document: object
    section: object
    element: etree._Element

    @property
    def kind(self) -> str:
        return etree.QName(self.element).localname

    @property
    def form_type(self) -> str:
        return _FORM_TYPE_BY_TAG.get(self.kind, self.kind.upper())

    @property
    def name(self) -> str | None:
        return self.element.get("name")

    @property
    def label(self) -> str:
        if self.kind in _BUTTON_FORM_TAGS:
            return self.element.get("caption", "")
        caption_label = _form_caption_text(self.element)
        if caption_label:
            return caption_label
        metadata = _form_command_metadata(self.element)
        label = metadata.get("label")
        if isinstance(label, str):
            return label
        if self.kind == "edit":
            text_node = _first_form_text_node(self.element)
            return text_node.text or "" if text_node is not None else ""
        if self.kind in _LIST_FORM_TAGS:
            selected = self.element.get("selectedValue", "")
            if selected:
                return selected
            values = _form_list_item_values(self.element)
            return values[0] if values else ""
        return ""

    @property
    def value(self) -> str | None:
        if self.kind in {"checkBtn", "radioBtn", "scrollBar"}:
            return self.element.get("value")
        if self.kind == "edit":
            text_node = _first_form_text_node(self.element)
            return text_node.text if text_node is not None else None
        if self.kind in _LIST_FORM_TAGS:
            selected = self.element.get("selectedValue")
            if selected not in (None, ""):
                return selected
            values = _form_list_item_values(self.element)
            return values[0] if values else None
        return None

    @property
    def checked(self) -> bool:
        value = (self.value or "").strip().upper()
        return self.kind in {"checkBtn", "radioBtn"} and value not in {"", "0", "FALSE", "NO", "OFF", "UNCHECKED"}

    @property
    def items(self) -> list[str]:
        metadata = _form_command_metadata(self.element)
        if "items" in metadata and isinstance(metadata["items"], list):
            return [str(value) for value in metadata["items"]]
        return _form_list_item_values(self.element) if self.kind in _LIST_FORM_TAGS else []

    @property
    def editable(self) -> bool:
        return self.element.get("editable", "0") == "1"

    @property
    def locked(self) -> bool:
        return self.element.get("enabled", "1") != "1"

    @property
    def placeholder(self) -> str | None:
        metadata = _form_command_metadata(self.element)
        value = metadata.get("placeholder")
        return value if isinstance(value, str) and value else None

    def layout(self) -> dict[str, str]:
        return _graphic_layout(self.element)

    def size(self) -> dict[str, int]:
        return _graphic_size(self.element)

    def out_margins(self) -> dict[str, int]:
        return _margin_values(self.element, "./hp:outMargin")

    def set_name(self, value: str | None) -> None:
        self.element.set("name", value or "")
        self.section.mark_modified()

    def set_label(self, value: str) -> None:
        if self.kind in _BUTTON_FORM_TAGS:
            self.element.set("caption", value)
        else:
            form_char_pr = _first_node(self.element, "./hp:formCharPr")
            _set_form_caption_text(self.element, value or None, char_pr_id=None if form_char_pr is None else form_char_pr.get("charPrIDRef"))
            _write_form_command_metadata(self.element, updates={"label": None})
        self.section.mark_modified()

    def set_value(self, value: str | None) -> None:
        if self.kind in {"checkBtn", "radioBtn", "scrollBar"}:
            self.element.set("value", value or "")
        elif self.kind == "edit":
            text_node = _first_form_text_node(self.element)
            if text_node is None:
                text_node = etree.SubElement(self.element, qname("hp", "text"))
            text_node.text = value or ""
        elif self.kind in _LIST_FORM_TAGS:
            self.element.set("selectedValue", value or "")
            if value and not self.items:
                _set_form_list_items(self.element, [value])
        self.section.mark_modified()

    def set_checked(self, value: bool) -> None:
        if self.kind in {"checkBtn", "radioBtn"}:
            self.element.set("value", "CHECKED" if value else "UNCHECKED")
            self.section.mark_modified()

    def set_items(self, values: list[str]) -> None:
        if self.kind in _LIST_FORM_TAGS:
            _set_form_list_items(self.element, list(values))
            if values and not self.element.get("selectedValue"):
                self.element.set("selectedValue", values[0])
            self.section.mark_modified()

    def set_editable(self, value: bool) -> None:
        self.element.set("editable", "1" if value else "0")
        self.section.mark_modified()

    def set_locked(self, value: bool) -> None:
        self.element.set("enabled", "0" if value else "1")
        self.section.mark_modified()

    def set_placeholder(self, value: str | None) -> None:
        _write_form_command_metadata(self.element, updates={"placeholder": value})
        self.section.mark_modified()

    def set_size(
        self,
        *,
        width: int | str | None = None,
        height: int | str | None = None,
    ) -> None:
        _set_graphic_size(self.element, width=width, height=height)
        self.section.mark_modified()

    def set_layout(
        self,
        *,
        text_wrap: str | None = None,
        text_flow: str | None = None,
        treat_as_char: bool | None = None,
        affect_line_spacing: bool | None = None,
        flow_with_text: bool | None = None,
        allow_overlap: bool | None = None,
        hold_anchor_and_so: bool | None = None,
        vert_rel_to: str | None = None,
        horz_rel_to: str | None = None,
        vert_align: str | None = None,
        horz_align: str | None = None,
        vert_offset: int | str | None = None,
        horz_offset: int | str | None = None,
    ) -> None:
        _set_graphic_layout(
            self.element,
            text_wrap=text_wrap,
            text_flow=text_flow,
            treat_as_char=treat_as_char,
            affect_line_spacing=affect_line_spacing,
            flow_with_text=flow_with_text,
            allow_overlap=allow_overlap,
            hold_anchor_and_so=hold_anchor_and_so,
            vert_rel_to=vert_rel_to,
            horz_rel_to=horz_rel_to,
            vert_align=vert_align,
            horz_align=horz_align,
            vert_offset=vert_offset,
            horz_offset=horz_offset,
        )
        self.section.mark_modified()

    def set_out_margins(
        self,
        *,
        left: int | str | None = None,
        right: int | str | None = None,
        top: int | str | None = None,
        bottom: int | str | None = None,
    ) -> None:
        _set_margin_values(self.element, "./hp:outMargin", left=left, right=right, top=top, bottom=bottom)
        self.section.mark_modified()


@dataclass
class BookmarkXml:
    document: object
    section: object
    element: etree._Element

    @property
    def name(self) -> str | None:
        return self.element.get("name")

    def rename(self, value: str) -> None:
        self.element.set("name", value)
        self.section.mark_modified()


@dataclass
class FieldXml:
    document: object
    section: object
    element: etree._Element

    @property
    def field_type(self) -> str | None:
        return self.element.get("type")

    @property
    def normalized_field_type(self) -> str:
        return (self.field_type or "").strip().upper()

    @property
    def field_id(self) -> str | None:
        return self.element.get("fieldid")

    @property
    def control_id(self) -> str | None:
        return self.element.get("id")

    @property
    def name(self) -> str | None:
        return self.element.get("name")

    def parameter_map(self) -> dict[str, str]:
        mapping: dict[str, str] = {}
        for parameter in self.element.xpath("./hp:parameters/*", namespaces=NS):
            name = parameter.get("name")
            if name:
                mapping[name] = parameter.text or ""
        return mapping

    def get_parameter(self, name: str) -> str | None:
        return self.parameter_map().get(name)

    def set_name(self, value: str) -> None:
        self.element.set("name", value)
        self.section.mark_modified()

    def set_field_type(self, value: str) -> None:
        self.element.set("type", value)
        self.section.mark_modified()

    def set_parameter(self, name: str, value: str) -> None:
        parameters = self.element.xpath("./hp:parameters", namespaces=NS)
        if not parameters:
            params = etree.SubElement(self.element, qname("hp", "parameters"))
            params.set("cnt", "0")
            params.set("name", "")
        else:
            params = parameters[0]

        for parameter in params:
            if parameter.get("name") == name:
                parameter.text = value
                self.section.mark_modified()
                return

        new_param = etree.SubElement(params, qname("hp", "stringParam"))
        new_param.set("name", name)
        new_param.text = value
        params.set("cnt", str(len(params)))
        self.section.mark_modified()

    @property
    def is_hyperlink(self) -> bool:
        return self.normalized_field_type == "HYPERLINK"

    @property
    def is_mail_merge(self) -> bool:
        return self.normalized_field_type in {"%MAI", "MAILMERGE", "MAIL_MERGE", "MERGEFIELD"}

    @property
    def is_calculation(self) -> bool:
        return self.normalized_field_type in {"%FOR", "%CAL", "CALCULATE", "CALC", "FORMULA"}

    @property
    def is_cross_reference(self) -> bool:
        return self.normalized_field_type in {"%REF", "%PAG", "%CRO", "REF", "PAGEREF", "BOOKMARKREF", "CROSSREF", "CROSS_REF"}

    @property
    def hyperlink_target(self) -> str | None:
        return self.get_parameter("Path") or self.get_parameter("Command")

    def set_hyperlink_target(self, value: str) -> None:
        self.set_parameter("Path", value)
        self.set_parameter("Command", value)

    @property
    def display_text(self) -> str:
        paragraph = self.element.xpath("ancestor::hp:p[1]", namespaces=NS)
        begin_run = self.element.xpath("ancestor::hp:run[1]", namespaces=NS)
        matching_end = self.element.xpath(
            "ancestor::hp:p[1]//hp:fieldEnd[@beginIDRef=$begin_id][1]",
            namespaces=NS,
            begin_id=self.control_id,
        )
        if not paragraph or not begin_run or not matching_end:
            return ""
        end_run = matching_end[0].xpath("ancestor::hp:run[1]", namespaces=NS)
        if not end_run:
            return ""
        runs = paragraph[0].xpath("./hp:run", namespaces=NS)
        try:
            begin_index = runs.index(begin_run[0])
            end_index = runs.index(end_run[0])
        except ValueError:
            return ""
        if end_index <= begin_index:
            return ""
        return "".join(_extract_text(run) for run in runs[begin_index + 1 : end_index])

    def set_display_text(self, value: str) -> None:
        paragraph = self.element.xpath("ancestor::hp:p[1]", namespaces=NS)
        begin_run = self.element.xpath("ancestor::hp:run[1]", namespaces=NS)
        matching_end = self.element.xpath(
            "ancestor::hp:p[1]//hp:fieldEnd[@beginIDRef=$begin_id][1]",
            namespaces=NS,
            begin_id=self.control_id,
        )
        if not paragraph or not begin_run or not matching_end:
            raise ValueError("Field display text can only be edited when a matching fieldEnd exists in the same paragraph.")
        end_run = matching_end[0].xpath("ancestor::hp:run[1]", namespaces=NS)
        if not end_run:
            raise ValueError("Matching fieldEnd is not contained in a run.")
        runs = paragraph[0].xpath("./hp:run", namespaces=NS)
        try:
            begin_index = runs.index(begin_run[0])
            end_index = runs.index(end_run[0])
        except ValueError as exc:
            raise ValueError("Field runs could not be located in the paragraph.") from exc
        middle_runs = runs[begin_index + 1 : end_index]
        if middle_runs:
            _set_text(middle_runs[0], value)
            for extra in middle_runs[1:]:
                _set_text(extra, "")
        else:
            new_run = etree.Element(qname("hp", "run"))
            new_run.set("charPrIDRef", begin_run[0].get("charPrIDRef", "0"))
            text_node = etree.SubElement(new_run, qname("hp", "t"))
            text_node.text = value
            end_run[0].addprevious(new_run)
        _invalidate_paragraph_layout(paragraph[0])
        self.section.mark_modified()

    def configure_mail_merge(self, field_name: str, *, display_text: str | None = None) -> None:
        self.set_field_type("MAILMERGE")
        self.set_name(field_name)
        self.set_parameter("FieldName", field_name)
        self.set_parameter("MergeField", field_name)
        if display_text is not None:
            self.set_display_text(display_text)

    def configure_calculation(self, expression: str, *, display_text: str | None = None) -> None:
        self.set_field_type("FORMULA")
        self.set_parameter("Expression", expression)
        self.set_parameter("Command", expression)
        if display_text is not None:
            self.set_display_text(display_text)

    def configure_cross_reference(self, bookmark_name: str, *, display_text: str | None = None) -> None:
        self.set_field_type("CROSSREF")
        self.set_name(bookmark_name)
        self.set_parameter("BookmarkName", bookmark_name)
        self.set_parameter("Path", bookmark_name)
        if display_text is not None:
            self.set_display_text(display_text)

    def configure_doc_property(self, property_name: str, *, display_text: str | None = None) -> None:
        self.set_field_type("DOCPROPERTY")
        self.set_name(property_name)
        self.set_parameter("FieldName", property_name)
        if display_text is not None:
            self.set_display_text(display_text)

    def configure_date(self, *, display_text: str | None = None) -> None:
        self.set_field_type("DATE")
        if display_text is not None:
            self.set_display_text(display_text)


@dataclass
class AutoNumberXml:
    document: object
    section: object
    element: etree._Element

    @property
    def kind(self) -> str:
        return etree.QName(self.element).localname

    @property
    def number(self) -> str | None:
        return self.element.get("num")

    @property
    def number_type(self) -> str | None:
        return self.element.get("numType")

    def set_number(self, value: int | str) -> None:
        self.element.set("num", str(value))
        self.section.mark_modified()

    def set_number_type(self, value: str) -> None:
        self.element.set("numType", value)
        self.section.mark_modified()


@dataclass
class EquationXml:
    document: object
    section: object
    element: etree._Element

    @property
    def script(self) -> str:
        nodes = self.element.xpath("./hp:script", namespaces=NS)
        return nodes[0].text or "" if nodes else ""

    @script.setter
    def script(self, value: str) -> None:
        nodes = self.element.xpath("./hp:script", namespaces=NS)
        if nodes:
            nodes[0].text = value
        else:
            node = etree.SubElement(self.element, qname("hp", "script"))
            node.text = value
        self.section.mark_modified()

    @property
    def shape_comment(self) -> str:
        nodes = self.element.xpath("./hp:shapeComment", namespaces=NS)
        return nodes[0].text or "" if nodes else ""

    def layout(self) -> dict[str, str]:
        return _graphic_layout(self.element)

    def size(self) -> dict[str, int]:
        return _graphic_size(self.element)

    def rotation(self) -> dict[str, str]:
        return _graphic_rotation(self.element)

    def out_margins(self) -> dict[str, int]:
        return _margin_values(self.element, "./hp:outMargin")

    def set_layout(
        self,
        *,
        text_wrap: str | None = None,
        text_flow: str | None = None,
        treat_as_char: bool | None = None,
        affect_line_spacing: bool | None = None,
        flow_with_text: bool | None = None,
        allow_overlap: bool | None = None,
        hold_anchor_and_so: bool | None = None,
        vert_rel_to: str | None = None,
        horz_rel_to: str | None = None,
        vert_align: str | None = None,
        horz_align: str | None = None,
        vert_offset: int | str | None = None,
        horz_offset: int | str | None = None,
    ) -> None:
        _set_graphic_layout(
            self.element,
            text_wrap=text_wrap,
            text_flow=text_flow,
            treat_as_char=treat_as_char,
            affect_line_spacing=affect_line_spacing,
            flow_with_text=flow_with_text,
            allow_overlap=allow_overlap,
            hold_anchor_and_so=hold_anchor_and_so,
            vert_rel_to=vert_rel_to,
            horz_rel_to=horz_rel_to,
            vert_align=vert_align,
            horz_align=horz_align,
            vert_offset=vert_offset,
            horz_offset=horz_offset,
        )
        self.section.mark_modified()

    def set_out_margins(
        self,
        *,
        left: int | str | None = None,
        right: int | str | None = None,
        top: int | str | None = None,
        bottom: int | str | None = None,
    ) -> None:
        _set_margin_values(self.element, "./hp:outMargin", left=left, right=right, top=top, bottom=bottom)
        self.section.mark_modified()

    def set_size(
        self,
        *,
        width: int | str | None = None,
        height: int | str | None = None,
    ) -> None:
        _set_graphic_size(self.element, width=width, height=height)
        self.section.mark_modified()

    def set_rotation(
        self,
        *,
        angle: int | str | None = None,
        center_x: int | str | None = None,
        center_y: int | str | None = None,
        rotate_image: bool | None = None,
    ) -> None:
        _set_graphic_rotation(
            self.element,
            angle=angle,
            center_x=center_x,
            center_y=center_y,
            rotate_image=rotate_image,
        )
        self.section.mark_modified()


@dataclass
class ShapeXml:
    document: object
    section: object
    element: etree._Element

    @property
    def kind(self) -> str:
        return etree.QName(self.element).localname

    @property
    def shape_comment(self) -> str:
        nodes = self.element.xpath("./hp:shapeComment", namespaces=NS)
        return nodes[0].text or "" if nodes else ""

    @shape_comment.setter
    def shape_comment(self, value: str) -> None:
        nodes = self.element.xpath("./hp:shapeComment", namespaces=NS)
        if nodes:
            nodes[0].text = value
        else:
            node = etree.SubElement(self.element, qname("hp", "shapeComment"))
            node.text = value
        self.section.mark_modified()

    @property
    def text(self) -> str:
        direct_text = self.element.get("text")
        if direct_text is not None:
            return direct_text
        return _extract_text(self.element)

    def set_text(self, value: str) -> None:
        if self.element.get("text") is not None:
            self.element.set("text", value)
            self.section.mark_modified()
            return

        draw_text_nodes = self.element.xpath("./hp:drawText", namespaces=NS)
        if draw_text_nodes:
            paragraphs, signatures = _capture_protected_paragraph_signatures(draw_text_nodes[0])
            _set_text(draw_text_nodes[0], value)
            self.section.mark_modified()
            self.document._track_control_preservation_targets(
                paragraphs,
                signatures,
                issues=[
                    ValidationIssue(
                        kind="control_preservation",
                        message="",
                        part_path=self.section.path,
                        section_index=self.section.section_index(),
                        paragraph_index=index,
                        context=f"{self.kind} drawText",
                    )
                    for index, _ in enumerate(paragraphs)
                ],
            )
            return

        paragraphs, signatures = _capture_protected_paragraph_signatures(self.element)
        _set_text(self.element, value)
        self.section.mark_modified()
        self.document._track_control_preservation_targets(
            paragraphs,
            signatures,
            issues=[
                ValidationIssue(
                    kind="control_preservation",
                    message="",
                    part_path=self.section.path,
                    section_index=self.section.section_index(),
                    paragraph_index=index,
                    context=self.kind,
                )
                for index, _ in enumerate(paragraphs)
            ],
        )

    def layout(self) -> dict[str, str]:
        return _graphic_layout(self.element)

    def size(self) -> dict[str, int]:
        return _graphic_size(self.element)

    def rotation(self) -> dict[str, str]:
        return _graphic_rotation(self.element)

    def out_margins(self) -> dict[str, int]:
        return _margin_values(self.element, "./hp:outMargin")

    def line_style(self) -> dict[str, str]:
        return _line_style(self.element)

    def fill_style(self) -> dict[str, str]:
        return _fill_style(self.element)

    def text_margins(self) -> dict[str, int]:
        return _text_margin(self.element)

    def set_layout(
        self,
        *,
        text_wrap: str | None = None,
        text_flow: str | None = None,
        treat_as_char: bool | None = None,
        affect_line_spacing: bool | None = None,
        flow_with_text: bool | None = None,
        allow_overlap: bool | None = None,
        hold_anchor_and_so: bool | None = None,
        vert_rel_to: str | None = None,
        horz_rel_to: str | None = None,
        vert_align: str | None = None,
        horz_align: str | None = None,
        vert_offset: int | str | None = None,
        horz_offset: int | str | None = None,
    ) -> None:
        _set_graphic_layout(
            self.element,
            text_wrap=text_wrap,
            text_flow=text_flow,
            treat_as_char=treat_as_char,
            affect_line_spacing=affect_line_spacing,
            flow_with_text=flow_with_text,
            allow_overlap=allow_overlap,
            hold_anchor_and_so=hold_anchor_and_so,
            vert_rel_to=vert_rel_to,
            horz_rel_to=horz_rel_to,
            vert_align=vert_align,
            horz_align=horz_align,
            vert_offset=vert_offset,
            horz_offset=horz_offset,
        )
        self.section.mark_modified()

    def set_out_margins(
        self,
        *,
        left: int | str | None = None,
        right: int | str | None = None,
        top: int | str | None = None,
        bottom: int | str | None = None,
    ) -> None:
        _set_margin_values(self.element, "./hp:outMargin", left=left, right=right, top=top, bottom=bottom)
        self.section.mark_modified()

    def set_size(
        self,
        *,
        width: int | str | None = None,
        height: int | str | None = None,
        original_width: int | str | None = None,
        original_height: int | str | None = None,
        current_width: int | str | None = None,
        current_height: int | str | None = None,
    ) -> None:
        _set_graphic_size(
            self.element,
            width=width,
            height=height,
            original_width=original_width,
            original_height=original_height,
            current_width=current_width,
            current_height=current_height,
        )
        self.section.mark_modified()

    def set_rotation(
        self,
        *,
        angle: int | str | None = None,
        center_x: int | str | None = None,
        center_y: int | str | None = None,
        rotate_image: bool | None = None,
    ) -> None:
        _set_graphic_rotation(
            self.element,
            angle=angle,
            center_x=center_x,
            center_y=center_y,
            rotate_image=rotate_image,
        )
        self.section.mark_modified()

    def set_line_style(
        self,
        *,
        color: str | None = None,
        width: int | str | None = None,
        style: str | None = None,
        end_cap: str | None = None,
        head_style: str | None = None,
        tail_style: str | None = None,
        head_fill: bool | None = None,
        tail_fill: bool | None = None,
        head_size: str | None = None,
        tail_size: str | None = None,
        outline_style: str | None = None,
        alpha: int | str | None = None,
    ) -> None:
        _set_line_style(
            self.element,
            color=color,
            width=width,
            style=style,
            end_cap=end_cap,
            head_style=head_style,
            tail_style=tail_style,
            head_fill=head_fill,
            tail_fill=tail_fill,
            head_size=head_size,
            tail_size=tail_size,
            outline_style=outline_style,
            alpha=alpha,
        )
        self.section.mark_modified()

    def set_fill_style(
        self,
        *,
        face_color: str | None = None,
        hatch_color: str | None = None,
        alpha: int | str | None = None,
    ) -> None:
        _set_fill_style(self.element, face_color=face_color, hatch_color=hatch_color, alpha=alpha)
        self.section.mark_modified()

    def set_text_margins(
        self,
        *,
        left: int | str | None = None,
        right: int | str | None = None,
        top: int | str | None = None,
        bottom: int | str | None = None,
    ) -> None:
        _set_text_margin(self.element, left=left, right=right, top=top, bottom=bottom)
        self.section.mark_modified()


@dataclass
class ChartXml:
    document: object
    section: object
    element: etree._Element

    @property
    def chart_part_path(self) -> str:
        return self.element.get("chartIDRef", "")

    def _chart_part(self):
        return self.document.get_part(self.chart_part_path, expected_type=object)

    def _chart_root(self) -> etree._Element:
        return _chart_part_root(self.document.get_part(self.chart_part_path))

    def _fallback_ole_element(self) -> etree._Element | None:
        case = self.element.getparent()
        if case is None or etree.QName(case).localname != "case":
            return None
        switch = case.getparent()
        if switch is None or etree.QName(switch).localname != "switch":
            return None
        return _first_node(switch, "./hp:default/hp:ole")

    def _chart_metadata(self) -> dict[str, object]:
        if not self.chart_part_path:
            return {}
        return _chart_metadata_from_root(self._chart_root())

    def _rebuild_chart_part(self, **overrides: object) -> None:
        root = self._chart_root()
        metadata = _chart_metadata_from_root(root)
        state = {
            "title": self.title,
            "chart_type": self.chart_type,
            "categories": self.categories,
            "series": self.series,
            "data_ref": self.data_ref,
            "legend_visible": self.legend_visible,
            "metadata": metadata,
        }
        state.update(overrides)
        part = self.document.get_part(self.chart_part_path)
        part._root = _build_chart_part_root(
            title=str(state["title"]),
            chart_type=str(state["chart_type"]),
            categories=[str(value) for value in state["categories"]],
            series=[dict(value) for value in state["series"]],
            data_ref=str(state["data_ref"]) if state["data_ref"] not in (None, "") else None,
            legend_visible=bool(state["legend_visible"]),
            metadata=dict(state["metadata"]),
        )
        part.mark_modified()

    @property
    def title(self) -> str:
        return _chart_title_from_root(self._chart_root())

    @property
    def chart_type(self) -> str:
        return _chart_type_from_root(self._chart_root())

    @property
    def categories(self) -> list[str]:
        return _chart_categories_from_root(self._chart_root())

    @property
    def series(self) -> list[dict[str, object]]:
        return _chart_series_from_root(self._chart_root())

    @property
    def data_ref(self) -> str | None:
        native_value = _chart_data_ref_from_root(self._chart_root())
        if native_value not in (None, ""):
            return native_value
        value = self._chart_metadata().get("dataRef")
        return str(value) if value not in (None, "") else None

    @property
    def shape_comment(self) -> str:
        nodes = self.element.xpath("./hp:shapeComment", namespaces=NS)
        if nodes:
            return nodes[0].text or ""
        value = self._chart_metadata().get("shapeComment")
        return str(value) if value not in (None, "") else ""

    @property
    def legend_visible(self) -> bool:
        return bool(self._chart_root().xpath("boolean(.//c:chart/c:legend)", namespaces=_CHART_NS))

    def layout(self) -> dict[str, str]:
        return _graphic_layout(self.element)

    def size(self) -> dict[str, int]:
        return _graphic_size(self.element)

    def out_margins(self) -> dict[str, int]:
        return _margin_values(self.element, "./hp:outMargin")

    def rotation(self) -> dict[str, str]:
        metadata = self._chart_metadata().get("rotation")
        metadata_rotation: dict[str, str] = {}
        if isinstance(metadata, dict):
            metadata_rotation.update({str(key): str(value) for key, value in metadata.items() if value not in (None, "")})
        fallback_ole = self._fallback_ole_element()
        native_rotation: dict[str, str] = {}
        if fallback_ole is not None:
            native_rotation.update({key: value for key, value in _graphic_rotation(fallback_ole).items() if value not in (None, "")})
        if metadata_rotation and native_rotation:
            differs = any(
                metadata_rotation.get(key) not in (None, "") and metadata_rotation.get(key) != native_rotation.get(key)
                for key in set(metadata_rotation) | set(native_rotation)
            )
            metadata_is_non_default = any(
                metadata_rotation.get(key) not in (None, "", "0")
                for key in ("angle", "centerX", "centerY", "rotateimage")
            )
            if differs and native_rotation.get("angle", "0") == "0" and metadata_is_non_default:
                return metadata_rotation
        rotation: dict[str, str] = {}
        rotation.update(metadata_rotation)
        rotation.update(native_rotation)
        return rotation

    def set_title(self, value: str) -> None:
        self._rebuild_chart_part(title=value)

    def set_chart_type(self, value: str) -> None:
        metadata = self._chart_metadata()
        metadata["chartType"] = value
        self._rebuild_chart_part(chart_type=value, metadata=metadata)

    def set_categories(self, values: list[str]) -> None:
        self._rebuild_chart_part(categories=[str(value) for value in values])

    def set_series(self, values: list[dict[str, object]]) -> None:
        self._rebuild_chart_part(series=[dict(value) for value in values])

    def set_data_ref(self, value: str | None) -> None:
        metadata = self._chart_metadata()
        metadata.pop("dataRef", None)
        self._rebuild_chart_part(
            data_ref=None if value in (None, "") else str(value),
            metadata=metadata,
        )

    def set_shape_comment(self, value: str | None) -> None:
        nodes = self.element.xpath("./hp:shapeComment", namespaces=NS)
        if value in (None, ""):
            for node in nodes:
                self.element.remove(node)
        elif nodes:
            nodes[0].text = str(value)
        else:
            node = etree.SubElement(self.element, qname("hp", "shapeComment"))
            node.text = str(value)
        metadata = self._chart_metadata()
        metadata.pop("shapeComment", None)
        self._rebuild_chart_part(metadata=metadata)
        self.section.mark_modified()

    def set_legend_visible(self, value: bool) -> None:
        self._rebuild_chart_part(legend_visible=value)

    def set_layout(
        self,
        *,
        text_wrap: str | None = None,
        text_flow: str | None = None,
        treat_as_char: bool | None = None,
        affect_line_spacing: bool | None = None,
        flow_with_text: bool | None = None,
        allow_overlap: bool | None = None,
        hold_anchor_and_so: bool | None = None,
        vert_rel_to: str | None = None,
        horz_rel_to: str | None = None,
        vert_align: str | None = None,
        horz_align: str | None = None,
        vert_offset: int | str | None = None,
        horz_offset: int | str | None = None,
    ) -> None:
        _set_graphic_layout(
            self.element,
            text_wrap=text_wrap,
            text_flow=text_flow,
            treat_as_char=treat_as_char,
            affect_line_spacing=affect_line_spacing,
            flow_with_text=flow_with_text,
            allow_overlap=allow_overlap,
            hold_anchor_and_so=hold_anchor_and_so,
            vert_rel_to=vert_rel_to,
            horz_rel_to=horz_rel_to,
            vert_align=vert_align,
            horz_align=horz_align,
            vert_offset=vert_offset,
            horz_offset=horz_offset,
        )
        self.section.mark_modified()

    def set_size(
        self,
        *,
        width: int | str | None = None,
        height: int | str | None = None,
    ) -> None:
        _set_graphic_size(self.element, width=width, height=height)
        fallback_ole = self._fallback_ole_element()
        if fallback_ole is not None:
            _set_graphic_size(
                fallback_ole,
                width=width,
                height=height,
                original_width=width,
                original_height=height,
                extent_x=width,
                extent_y=height,
            )
        self.section.mark_modified()

    def set_out_margins(
        self,
        *,
        left: int | str | None = None,
        right: int | str | None = None,
        top: int | str | None = None,
        bottom: int | str | None = None,
    ) -> None:
        _set_margin_values(self.element, "./hp:outMargin", left=left, right=right, top=top, bottom=bottom)
        self.section.mark_modified()

    def set_rotation(
        self,
        *,
        angle: int | str | None = None,
        center_x: int | str | None = None,
        center_y: int | str | None = None,
        rotate_image: bool | None = None,
    ) -> None:
        fallback_ole = self._fallback_ole_element()
        if fallback_ole is not None:
            _set_graphic_rotation(
                fallback_ole,
                angle=angle,
                center_x=center_x,
                center_y=center_y,
                rotate_image=rotate_image,
            )
            metadata = self._chart_metadata()
            if "rotation" in metadata:
                metadata.pop("rotation", None)
                self._rebuild_chart_part(metadata=metadata)
            self.section.mark_modified()
            return

        metadata = self._chart_metadata()
        current = metadata.get("rotation")
        rotation = dict(current) if isinstance(current, dict) else {}
        updates = {
            "angle": angle,
            "centerX": center_x,
            "centerY": center_y,
            "rotateimage": "1" if rotate_image else ("0" if rotate_image is False else None),
        }
        for key, value in updates.items():
            if value is None:
                continue
            rotation[key] = str(value)
        metadata["rotation"] = rotation
        self._rebuild_chart_part(metadata=metadata)
        self.section.mark_modified()


@dataclass
class OleXml(ShapeXml):
    @property
    def binary_item_id(self) -> str | None:
        return self.element.get("binaryItemIDRef")

    @property
    def object_type(self) -> str | None:
        return self.element.get("objectType")

    @property
    def draw_aspect(self) -> str | None:
        return self.element.get("drawAspect")

    @property
    def has_moniker(self) -> bool:
        return self.element.get("hasMoniker") == "1"

    def extent(self) -> dict[str, int]:
        node = _first_node(self.element, "./hc:extent")
        if node is None:
            return {}
        return {
            "x": int(node.get("x", "0")),
            "y": int(node.get("y", "0")),
        }

    def binary_part_path(self) -> str:
        item_id = self.binary_item_id
        if item_id is None:
            raise ValueError("OLE object is not bound to a binary manifest item.")
        for item in self.document.content_hpf.manifest_items():
            if item.get("id") == item_id:
                href = item.get("href")
                if href:
                    return href
        raise KeyError(f"Manifest item {item_id} was not found.")

    def binary_data(self) -> bytes:
        part = self.document.get_part(self.binary_part_path())
        return part.data

    def replace_binary(self, data: bytes) -> None:
        part = self.document.get_part(self.binary_part_path())
        part.data = data

    def bind_binary_item(self, item_id: str) -> None:
        self.element.set("binaryItemIDRef", item_id)
        self.section.mark_modified()

    def set_object_metadata(
        self,
        *,
        object_type: str | None = None,
        draw_aspect: str | None = None,
        has_moniker: bool | None = None,
        eq_baseline: int | str | None = None,
    ) -> None:
        _set_optional_attributes(
            self.element,
            objectType=object_type,
            drawAspect=draw_aspect,
            hasMoniker=has_moniker,
            eqBaseLine=eq_baseline,
        )
        self.section.mark_modified()

    def set_extent(self, *, x: int | str | None = None, y: int | str | None = None) -> None:
        _set_graphic_size(self.element, extent_x=x, extent_y=y)
        self.section.mark_modified()


HeaderFooterBlock = HeaderFooterXml
TableCell = TableCellXml
Table = TableXml
Picture = PictureXml
StyleDefinition = StyleDefinitionXml
ParagraphStyle = ParagraphStyleXml
CharacterStyle = CharacterStyleXml
MemoShape = MemoShapeXml
SectionSettings = SectionSettingsXml
Note = NoteXml
Memo = MemoXml
Bookmark = BookmarkXml
Field = FieldXml
AutoNumber = AutoNumberXml
Equation = EquationXml
ShapeObject = ShapeXml
ChartObject = ChartXml
OleObject = OleXml
