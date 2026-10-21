from __future__ import annotations

import json
from pathlib import PurePosixPath

from lxml import etree

from ._hwpx_common import (
    _BUTTON_FORM_TAGS,
    _FORM_METADATA_COMMAND_PREFIX,
    _FORM_TAG_BY_TYPE,
    _graphic_attribute_xpath,
)
from ._hwpx_document_fields import HwpxDocumentFieldMixin
from ._hwpx_document_graphics import HwpxDocumentGraphicMixin
from .elements import (
    AutoNumberXml,
    CharacterStyleXml,
    ChartXml,
    EquationXml,
    FormXml,
    HeaderFooterXml,
    MemoShapeXml,
    MemoXml,
    NoteXml,
    ParagraphStyleXml,
    StyleDefinitionXml,
    _build_chart_part_root,
    _invalidate_paragraph_layout,
    _set_form_caption_text,
    _set_graphic_layout,
    _set_margin_values,
)
from .namespaces import NS, qname
from .xmlnode import HwpxXmlNode


_APPEND_BLOCK_TYPE_ALIASES = {
    "text": "paragraph",
    "para": "paragraph",
    "paragraph": "paragraph",
    "eq": "equation",
    "equ": "equation",
    "equation": "equation",
    "image": "picture",
    "img": "picture",
    "pic": "picture",
    "picture": "picture",
    "shape": "shape",
    "ole": "ole",
    "table": "table",
    "tbl": "table",
    "bookmark": "bookmark",
    "field": "field",
    "link": "hyperlink",
    "hyperlink": "hyperlink",
    "note": "note",
    "footnote": "note",
    "endnote": "note",
    "form": "form",
    "memo": "memo",
    "comment": "memo",
    "chart": "chart",
    "autonum": "auto_number",
    "auto_number": "auto_number",
    "auto-number": "auto_number",
    "newnum": "auto_number",
    "new_number": "auto_number",
    "new-number": "auto_number",
    "header": "header",
    "footer": "footer",
}


def _canonical_append_block_type(value: str) -> str:
    normalized = value.strip().lower().replace(" ", "_")
    resolved = _APPEND_BLOCK_TYPE_ALIASES.get(normalized)
    if resolved is None:
        raise ValueError(f"Unsupported block type: {value!r}")
    return resolved


class HwpxDocumentControlMixin(HwpxDocumentGraphicMixin, HwpxDocumentFieldMixin):
    def append_header(
        self,
        text: str = "",
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
        apply_page_type: str = "BOTH",
        hide_first: bool | None = None,
    ) -> HeaderFooterXml:
        self._ensure_editable_sections()
        block = self._build_header_footer_block(
            "header",
            text=text,
            char_pr_id=char_pr_id,
            apply_page_type=apply_page_type,
        )
        section = self._append_control(section_index, paragraph_index, block, char_pr_id=char_pr_id)
        if hide_first is not None:
            self.section_settings(section_index).set_visibility(hide_first_header=hide_first)
        return HeaderFooterXml(self, section, block)

    def append_footer(
        self,
        text: str = "",
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
        apply_page_type: str = "BOTH",
        hide_first: bool | None = None,
    ) -> HeaderFooterXml:
        self._ensure_editable_sections()
        block = self._build_header_footer_block(
            "footer",
            text=text,
            char_pr_id=char_pr_id,
            apply_page_type=apply_page_type,
        )
        section = self._append_control(section_index, paragraph_index, block, char_pr_id=char_pr_id)
        if hide_first is not None:
            self.section_settings(section_index).set_visibility(hide_first_footer=hide_first)
        return HeaderFooterXml(self, section, block)

    def append_note(
        self,
        text: str = "",
        *,
        kind: str = "footNote",
        number: int | str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> NoteXml:
        self._ensure_editable_sections()
        if kind not in {"footNote", "endNote"}:
            raise ValueError("append_note() kind must be 'footNote' or 'endNote'.")
        note = etree.Element(qname("hp", kind))
        note.set("id", self._next_control_number(".//hp:footNote/@id | .//hp:endNote/@id"))
        note.set("number", str(number if number is not None else self._next_control_number(".//hp:footNote/@number | .//hp:endNote/@number")))
        note.append(self._build_sublist_with_text(text, char_pr_id=char_pr_id, vertical_align="TOP"))
        section = self._append_control(section_index, paragraph_index, note, char_pr_id=char_pr_id)
        return NoteXml(self, section, note)

    def append_footnote(
        self,
        text: str = "",
        *,
        number: int | str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> NoteXml:
        return self.append_note(
            text,
            kind="footNote",
            number=number,
            section_index=section_index,
            paragraph_index=paragraph_index,
            char_pr_id=char_pr_id,
        )

    def append_endnote(
        self,
        text: str = "",
        *,
        number: int | str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> NoteXml:
        return self.append_note(
            text,
            kind="endNote",
            number=number,
            section_index=section_index,
            paragraph_index=paragraph_index,
            char_pr_id=char_pr_id,
        )

    def append_form(
        self,
        label: str = "",
        *,
        form_type: str = "INPUT",
        name: str | None = None,
        value: str | None = None,
        checked: bool = False,
        items: list[str] | None = None,
        editable: bool = True,
        locked: bool = False,
        placeholder: str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> FormXml:
        self._ensure_editable_sections()
        normalized_form_type = form_type.strip().upper()
        tag = _FORM_TAG_BY_TYPE.get(normalized_form_type)
        if tag is None:
            raise ValueError(f"Unsupported HWPX form_type: {form_type}")

        form = etree.Element(qname("hp", tag))
        form.set("name", name or "")
        form.set("foreColor", "#000000")
        form.set("backColor", "#7FFFFFFF")
        form.set("groupName", "")
        form.set("tabStop", "1")
        form.set("editable", "1" if editable else "0")
        form.set("tabOrder", "1")
        form.set("enabled", "0" if locked else "1")
        form.set("borderTypeIDRef", "0")
        form.set("drawFrame", "1")
        form.set("printable", "1")
        form.set("command", "")

        extra_metadata: dict[str, object] = {}
        if tag == "checkBtn":
            form.set("caption", label)
            form.set("value", value or ("CHECKED" if checked else "UNCHECKED"))
            form.set("radioGroupName", "")
            form.set("triState", "0")
            form.set("backStyle", "TRANSPARENT")
        elif tag == "radioBtn":
            form.set("caption", label)
            form.set("value", value or ("CHECKED" if checked else "UNCHECKED"))
            form.set("radioGroupName", name or "")
            form.set("triState", "0")
            form.set("backStyle", "TRANSPARENT")
        elif tag == "btn":
            form.set("caption", label)
            form.set("radioGroupName", "")
            form.set("triState", "0")
        elif tag == "edit":
            form.set("multiLine", "0")
            form.set("passwordChar", "")
            form.set("maxLength", "32")
            form.set("scrollBars", "NONE")
            form.set("tabKeyBehavior", "NEXT_OBJECT")
            form.set("numOnly", "0")
            form.set("readOnly", "0")
            form.set("alignText", "LEFT")
            text_node = etree.SubElement(form, qname("hp", "text"))
            text_node.text = value or ""
        elif tag == "comboBox":
            form.set("listBoxRows", "4")
            form.set("listBoxWidth", "2400")
            form.set("editEnable", "1" if editable else "0")
            form.set("selectedValue", value or "")
            for item_value in items or ([value] if value else []):
                item = etree.SubElement(form, qname("hp", "listItem"))
                item.set("displayText", "")
                item.set("value", item_value or "")
        elif tag == "listBox":
            form.set("itemHeight", "300")
            form.set("topIdx", "4294967295")
            form.set("selectedValue", value or "")
            for item_value in items or ([value] if value else []):
                item = etree.SubElement(form, qname("hp", "listItem"))
                item.set("displayText", "")
                item.set("value", item_value or "")
        elif tag == "scrollBar":
            form.set("delay", "0")
            form.set("largeChange", "10")
            form.set("smallChange", "1")
            form.set("min", "0")
            form.set("max", "100")
            form.set("page", "10")
            form.set("value", value or "0")
            form.set("type", "HORIZONTAL")

        if tag != "scrollBar":
            form_char_pr = etree.SubElement(form, qname("hp", "formCharPr"))
            form_char_pr.set("charPrIDRef", char_pr_id or "0")
            form_char_pr.set("followContext", "0")
            form_char_pr.set("autoSz", "0")
            form_char_pr.set("wordWrap", "0")
        elif label:
            _set_form_caption_text(form, label, char_pr_id=char_pr_id or "0")

        size = etree.SubElement(form, qname("hp", "sz"))
        if tag == "btn":
            size.set("width", "3200")
            size.set("height", "1400")
        elif tag == "edit":
            size.set("width", "4200")
            size.set("height", "1400")
        elif tag == "listBox":
            size.set("width", "4200")
            size.set("height", "1800")
        elif tag == "scrollBar":
            size.set("width", "4200")
            size.set("height", "900")
        else:
            size.set("width", "2400" if tag in {"checkBtn", "radioBtn"} else "4200")
            size.set("height", "1200" if tag in {"checkBtn", "radioBtn"} else "1400")
        size.set("widthRelTo", "ABSOLUTE")
        size.set("heightRelTo", "ABSOLUTE")
        size.set("protect", "0")

        pos = etree.SubElement(form, qname("hp", "pos"))
        pos.set("treatAsChar", "1")
        pos.set("affectLSpacing", "0")
        pos.set("flowWithText", "1")
        pos.set("allowOverlap", "0")
        pos.set("holdAnchorAndSO", "0")
        pos.set("vertRelTo", "PARA")
        pos.set("horzRelTo", "COLUMN")
        pos.set("vertAlign", "TOP")
        pos.set("horzAlign", "LEFT")
        pos.set("vertOffset", "0")
        pos.set("horzOffset", "0")

        out_margin = etree.SubElement(form, qname("hp", "outMargin"))
        out_margin.set("left", "0")
        out_margin.set("right", "0")
        out_margin.set("top", "0")
        out_margin.set("bottom", "0")

        if label and tag not in _BUTTON_FORM_TAGS and tag != "scrollBar":
            _set_form_caption_text(form, label, char_pr_id=char_pr_id or "0")

        if placeholder:
            extra_metadata["placeholder"] = placeholder
        if items and tag not in {"comboBox", "listBox"}:
            extra_metadata["items"] = list(items)
        if extra_metadata:
            form.set(
                "command",
                _FORM_METADATA_COMMAND_PREFIX + json.dumps(extra_metadata, ensure_ascii=False, separators=(",", ":"), sort_keys=True),
            )

        section = self._append_run_child(section_index, paragraph_index, form, char_pr_id=char_pr_id)
        return FormXml(self, section, form)

    def append_memo(
        self,
        text: str = "",
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> MemoXml:
        self._ensure_editable_sections()
        memo = etree.Element(qname("hp", "hiddenComment"))
        memo.append(self._build_sublist_with_text(text, char_pr_id=char_pr_id, vertical_align="TOP"))
        section = self._append_control(section_index, paragraph_index, memo, char_pr_id=char_pr_id)
        return MemoXml(self, section, memo)

    def append_memo_shape(
        self,
        *,
        memo_shape_id: int | str | None = None,
        width: int | str = 12000,
        line_width: int | str = 12,
        line_type: str = "SOLID",
        line_color: str = "#808080",
        fill_color: str = "#FFF2CC",
        active_color: str = "#FFC000",
        memo_type: str = "NOMAL",
    ) -> MemoShapeXml:
        resolved_id = str(memo_shape_id) if memo_shape_id is not None else self._next_header_numeric_id(".//hh:memoPr/@id")
        memo_shape = etree.Element(qname("hh", "memoPr"))
        memo_shape.set("id", resolved_id)
        memo_shape.set("width", str(width))
        memo_shape.set("lineWidth", str(line_width))
        memo_shape.set("lineType", line_type)
        memo_shape.set("lineColor", line_color)
        memo_shape.set("fillColor", fill_color)
        memo_shape.set("activeColor", active_color)
        memo_shape.set("memoType", memo_type)

        parent = self._ensure_header_collection(".//hh:memoProperties", "memoProperties")
        parent.append(memo_shape)
        parent.set("itemCnt", str(len(parent.xpath("./hh:memoPr", namespaces=NS))))
        self.header.mark_modified()
        return MemoShapeXml(self.header, memo_shape)

    def append_chart(
        self,
        title: str = "",
        *,
        chart_type: str = "BAR",
        categories: list[str] | None = None,
        series: list[dict[str, object]] | None = None,
        data_ref: str | None = None,
        legend_visible: bool = True,
        width: int = 12000,
        height: int = 3200,
        shape_comment: str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
        text_wrap: str = "TOP_AND_BOTTOM",
        text_flow: str = "BOTH_SIDES",
        treat_as_char: bool = True,
        affect_line_spacing: bool = False,
        flow_with_text: bool = True,
        allow_overlap: bool = False,
        hold_anchor_and_so: bool = False,
        vert_rel_to: str = "PARA",
        horz_rel_to: str = "COLUMN",
        vert_align: str = "TOP",
        horz_align: str = "LEFT",
        vert_offset: int = 0,
        horz_offset: int = 0,
        out_margin_left: int = 0,
        out_margin_right: int = 0,
        out_margin_top: int = 0,
        out_margin_bottom: int = 0,
    ) -> ChartXml:
        self._ensure_editable_sections()

        chart_path = self._next_chart_part_path()
        chart_control_id = self._next_control_number(".//hp:chart/@id | .//hp:ole/@id")
        chart_z_order = self._next_control_number(".//hp:chart/@zOrder | .//hp:ole/@zOrder")
        metadata: dict[str, object] = {"chartType": str(chart_type or "BAR").upper()}
        chart_part = self.add_part(
            chart_path,
            etree.tostring(
                _build_chart_part_root(
                    title=title,
                    chart_type=str(chart_type or "BAR").upper(),
                    categories=[str(value) for value in (categories or [])],
                    series=[dict(value) for value in (series or [])],
                    data_ref=None if data_ref in (None, "") else str(data_ref),
                    legend_visible=legend_visible,
                    metadata=metadata,
                ),
                encoding="UTF-8",
                xml_declaration=True,
                standalone=True,
                pretty_print=False,
            ),
        )
        chart_manifest_id = self.content_hpf.next_manifest_id(PurePosixPath(chart_path).stem or "chart")
        self.content_hpf.ensure_manifest_item(chart_manifest_id, chart_path, "application/xml")

        chart = etree.Element(qname("hp", "chart"))
        chart.set("id", chart_control_id)
        chart.set("zOrder", chart_z_order)
        chart.set("numberingType", "PICTURE")
        chart.set("textWrap", text_wrap)
        chart.set("textFlow", text_flow)
        chart.set("lock", "0")
        chart.set("dropcapstyle", "None")
        chart.set("chartIDRef", chart_path)
        if shape_comment:
            comment = etree.SubElement(chart, qname("hp", "shapeComment"))
            comment.text = shape_comment

        size = etree.SubElement(chart, qname("hp", "sz"))
        size.set("width", str(width))
        size.set("widthRelTo", "ABSOLUTE")
        size.set("height", str(height))
        size.set("heightRelTo", "ABSOLUTE")
        size.set("protect", "0")

        position = etree.SubElement(chart, qname("hp", "pos"))
        position.set("treatAsChar", "1" if treat_as_char else "0")
        position.set("affectLSpacing", "0")
        position.set("flowWithText", "1")
        position.set("allowOverlap", "0")
        position.set("holdAnchorAndSO", "0")
        position.set("vertRelTo", "PARA")
        position.set("horzRelTo", "COLUMN")
        position.set("vertAlign", "TOP")
        position.set("horzAlign", "LEFT")
        position.set("vertOffset", "0")
        position.set("horzOffset", "0")
        _set_graphic_layout(
            chart,
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

        out_margin = etree.SubElement(chart, qname("hp", "outMargin"))
        out_margin.set("left", "0")
        out_margin.set("right", "0")
        out_margin.set("top", "0")
        out_margin.set("bottom", "0")
        _set_margin_values(
            chart,
            "./hp:outMargin",
            left=out_margin_left,
            right=out_margin_right,
            top=out_margin_top,
            bottom=out_margin_bottom,
        )

        switch = etree.Element(qname("hp", "switch"))
        case = etree.SubElement(switch, qname("hp", "case"))
        case.set(qname("hp", "required-namespace"), NS["ooxmlchart"])
        case.append(chart)
        default = etree.SubElement(switch, qname("hp", "default"))
        default.append(
            self._build_chart_fallback_ole(
                control_id=chart_control_id,
                z_order=chart_z_order,
                width=width,
                height=height,
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
                out_margin_left=out_margin_left,
                out_margin_right=out_margin_right,
                out_margin_top=out_margin_top,
                out_margin_bottom=out_margin_bottom,
            )
        )

        paragraph = self._resolve_paragraph_for_insert(section_index=section_index, paragraph_index=paragraph_index)
        run = self._append_run(paragraph, char_pr_id=char_pr_id)
        run.append(switch)
        etree.SubElement(run, qname("hp", "t"))

        section = self.sections[section_index]
        section.mark_modified()
        chart_part.mark_modified()
        return ChartXml(self, section, chart)

    def append_auto_number(
        self,
        *,
        number: int | str = 1,
        number_type: str = "PAGE",
        kind: str = "newNum",
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> AutoNumberXml:
        self._ensure_editable_sections()
        if kind not in {"newNum", "autoNum"}:
            raise ValueError("append_auto_number() kind must be 'newNum' or 'autoNum'.")
        node = etree.Element(qname("hp", kind))
        node.set("num", str(number))
        node.set("numType", number_type)
        section = self._append_control(section_index, paragraph_index, node, char_pr_id=char_pr_id)
        return AutoNumberXml(self, section, node)

    def append_new_number(
        self,
        *,
        number: int | str = 1,
        number_type: str = "PAGE",
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> AutoNumberXml:
        return self.append_auto_number(
            number=number,
            number_type=number_type,
            kind="newNum",
            section_index=section_index,
            paragraph_index=paragraph_index,
            char_pr_id=char_pr_id,
        )

    def append_equation(
        self,
        script: str,
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
        width: int = 4800,
        height: int = 2300,
        treat_as_char: bool = True,
        shape_comment: str | None = None,
        text_color: str = "#000000",
        base_unit: int = 1100,
        font: str = "HYhwpEQ",
        text_wrap: str = "TOP_AND_BOTTOM",
        text_flow: str = "BOTH_SIDES",
        affect_line_spacing: bool = False,
        flow_with_text: bool = True,
        allow_overlap: bool = False,
        hold_anchor_and_so: bool = False,
        vert_rel_to: str = "PARA",
        horz_rel_to: str = "COLUMN",
        vert_align: str = "TOP",
        horz_align: str = "LEFT",
        vert_offset: int = 0,
        horz_offset: int = 0,
        out_margin_left: int = 0,
        out_margin_right: int = 0,
        out_margin_top: int = 0,
        out_margin_bottom: int = 0,
    ) -> EquationXml:
        self._ensure_editable_sections()
        if width < 1 or height < 1:
            raise ValueError("width and height must be positive.")

        equation = etree.Element(qname("hp", "equation"))
        equation.set("id", self._next_control_number(_graphic_attribute_xpath("id")))
        equation.set("zOrder", self._next_control_number(_graphic_attribute_xpath("zOrder")))
        equation.set("numberingType", "EQUATION")
        equation.set("textWrap", "TOP_AND_BOTTOM")
        equation.set("textFlow", "BOTH_SIDES")
        equation.set("lock", "0")
        equation.set("dropcapstyle", "None")
        equation.set("version", "Equation Version 60")
        equation.set("baseLine", "93")
        equation.set("textColor", text_color)
        equation.set("baseUnit", str(base_unit))
        equation.set("lineMode", "CHAR")
        equation.set("font", font)

        size = etree.SubElement(equation, qname("hp", "sz"))
        size.set("width", str(width))
        size.set("widthRelTo", "ABSOLUTE")
        size.set("height", str(height))
        size.set("heightRelTo", "ABSOLUTE")
        size.set("protect", "0")

        position = etree.SubElement(equation, qname("hp", "pos"))
        position.set("treatAsChar", "1" if treat_as_char else "0")
        position.set("affectLSpacing", "0")
        position.set("flowWithText", "1")
        position.set("allowOverlap", "0")
        position.set("holdAnchorAndSO", "0")
        position.set("vertRelTo", "PARA")
        position.set("horzRelTo", "COLUMN")
        position.set("vertAlign", "TOP")
        position.set("horzAlign", "LEFT")
        position.set("vertOffset", "0")
        position.set("horzOffset", "0")

        out_margin = etree.SubElement(equation, qname("hp", "outMargin"))
        for key in ("left", "right", "top", "bottom"):
            out_margin.set(key, "0")
        _set_graphic_layout(
            equation,
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
        _set_margin_values(
            equation,
            "./hp:outMargin",
            left=out_margin_left,
            right=out_margin_right,
            top=out_margin_top,
            bottom=out_margin_bottom,
        )

        if shape_comment:
            comment = etree.SubElement(equation, qname("hp", "shapeComment"))
            comment.text = shape_comment

        script_node = etree.SubElement(equation, qname("hp", "script"))
        script_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        script_node.text = script

        section = self._append_run_child(section_index, paragraph_index, equation, char_pr_id=char_pr_id)
        return EquationXml(self, section, equation)

    def append_inline_equation(
        self,
        script: str,
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
        width: int = 4800,
        height: int = 2300,
        treat_as_char: bool = True,
        shape_comment: str | None = None,
        text_color: str = "#000000",
        base_unit: int = 1100,
        font: str = "HYhwpEQ",
        text_wrap: str = "TOP_AND_BOTTOM",
        text_flow: str = "BOTH_SIDES",
        affect_line_spacing: bool = False,
        flow_with_text: bool = True,
        allow_overlap: bool = False,
        hold_anchor_and_so: bool = False,
        vert_rel_to: str = "PARA",
        horz_rel_to: str = "COLUMN",
        vert_align: str = "TOP",
        horz_align: str = "LEFT",
        vert_offset: int = 0,
        horz_offset: int = 0,
        out_margin_left: int = 0,
        out_margin_right: int = 0,
        out_margin_top: int = 0,
        out_margin_bottom: int = 0,
    ) -> EquationXml:
        self._ensure_editable_sections()
        resolved_paragraph_index = paragraph_index
        if resolved_paragraph_index is None:
            paragraph_count = self.paragraph_count(section_index)
            if paragraph_count <= 0:
                self.append_paragraph("", section_index=section_index, char_pr_id=char_pr_id)
                paragraph_count = self.paragraph_count(section_index)
            resolved_paragraph_index = paragraph_count - 1
        return self.append_equation(
            script,
            section_index=section_index,
            paragraph_index=resolved_paragraph_index,
            char_pr_id=char_pr_id,
            width=width,
            height=height,
            treat_as_char=treat_as_char,
            shape_comment=shape_comment,
            text_color=text_color,
            base_unit=base_unit,
            font=font,
            text_wrap=text_wrap,
            text_flow=text_flow,
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
            out_margin_left=out_margin_left,
            out_margin_right=out_margin_right,
            out_margin_top=out_margin_top,
            out_margin_bottom=out_margin_bottom,
        )

    def _dispatch_append_block(
        self,
        type: str,
        content: object | None,
        *,
        section_index: int,
        paragraph_index: int | None,
        char_pr_id: str | None,
        inline: bool,
        extra_kwargs: dict[str, object],
    ):
        requested_type = type.strip().lower().replace(" ", "_")
        normalized_type = _canonical_append_block_type(type)
        kwargs = dict(extra_kwargs)
        kwargs.setdefault("section_index", section_index)
        if char_pr_id is not None:
            kwargs.setdefault("char_pr_id", char_pr_id)

        if inline:
            if normalized_type in {"header", "footer"}:
                raise ValueError(f"append_inline() does not support block type {type!r}.")
            kwargs["paragraph_index"] = self._resolve_inline_paragraph_index(
                section_index=section_index,
                paragraph_index=paragraph_index,
                char_pr_id=char_pr_id,
            )
        elif paragraph_index is not None:
            kwargs["paragraph_index"] = paragraph_index

        if normalized_type == "paragraph":
            if inline:
                return self._append_inline_text(
                    str(content or ""),
                    section_index=section_index,
                    paragraph_index=int(kwargs["paragraph_index"]),
                    char_pr_id=char_pr_id,
                )
            return self.append_paragraph(str(content or ""), **kwargs)

        if normalized_type == "equation":
            script = str(content or "")
            if inline:
                return self.append_inline_equation(script, **kwargs)
            return self.append_equation(script, **kwargs)

        if normalized_type == "picture":
            data = kwargs.pop("data", None)
            if not isinstance(data, bytes):
                raise ValueError("append_block(type='picture') requires bytes data=...")
            return self.append_picture(str(content or ""), data, **kwargs)

        if normalized_type == "shape":
            if content is not None and "text" not in kwargs:
                kwargs["text"] = str(content)
            return self.append_shape(**kwargs)

        if normalized_type == "ole":
            data = kwargs.pop("data", None)
            if not isinstance(data, bytes):
                raise ValueError("append_block(type='ole') requires bytes data=...")
            return self.append_ole(str(content or ""), data, **kwargs)

        if normalized_type == "table":
            if isinstance(content, list):
                cell_texts = [[str(value) for value in row] for row in content]
                kwargs.setdefault("cell_texts", cell_texts)
                kwargs.setdefault("rows", len(cell_texts))
                kwargs.setdefault("columns", max((len(row) for row in cell_texts), default=0))
            rows = kwargs.pop("rows", None)
            columns = kwargs.pop("columns", None)
            if not isinstance(rows, int) or not isinstance(columns, int):
                raise ValueError("append_block(type='table') requires rows=... and columns=..., or a 2D content list.")
            return self.append_table(rows, columns, **kwargs)

        if normalized_type == "bookmark":
            return self.append_bookmark(str(content or ""), **kwargs)

        if normalized_type == "field":
            if content is not None and "display_text" not in kwargs:
                kwargs["display_text"] = str(content)
            return self.append_field(**kwargs)

        if normalized_type == "hyperlink":
            kwargs.setdefault("display_text", None)
            return self.append_hyperlink(str(content or ""), **kwargs)

        if normalized_type == "note":
            if requested_type == "endnote":
                kwargs.setdefault("kind", "endNote")
            elif requested_type in {"note", "footnote"}:
                kwargs.setdefault("kind", "footNote")
            return self.append_note(str(content or ""), **kwargs)

        if normalized_type == "form":
            return self.append_form(str(content or ""), **kwargs)

        if normalized_type == "memo":
            return self.append_memo(str(content or ""), **kwargs)

        if normalized_type == "chart":
            return self.append_chart(str(content or ""), **kwargs)

        if normalized_type == "auto_number":
            if requested_type in {"newnum", "new_number", "new-number"}:
                kwargs.setdefault("kind", "newNum")
            if content is not None and "number" not in kwargs:
                kwargs["number"] = content
            return self.append_auto_number(**kwargs)

        if normalized_type == "header":
            return self.append_header(str(content or ""), **kwargs)

        if normalized_type == "footer":
            return self.append_footer(str(content or ""), **kwargs)

        raise ValueError(f"Unsupported block type: {type!r}")

    def _resolve_inline_paragraph_index(
        self,
        *,
        section_index: int,
        paragraph_index: int | None,
        char_pr_id: str | None = None,
    ) -> int:
        if paragraph_index is not None:
            return paragraph_index
        paragraph_count = self.paragraph_count(section_index)
        if paragraph_count <= 0:
            self.append_paragraph("", section_index=section_index, char_pr_id=char_pr_id)
            paragraph_count = self.paragraph_count(section_index)
        return paragraph_count - 1

    def _append_inline_text(
        self,
        text: str,
        *,
        section_index: int,
        paragraph_index: int,
        char_pr_id: str | None = None,
    ) -> HwpxXmlNode:
        paragraph = self._paragraph_element(section_index, paragraph_index)
        run = self._append_run(paragraph, char_pr_id=char_pr_id)
        text_node = etree.SubElement(run, qname("hp", "t"))
        text_node.text = text
        _invalidate_paragraph_layout(paragraph)
        section = self.sections[section_index]
        section.mark_modified()
        return HwpxXmlNode(text_node, section)

    def append_paragraph_style(
        self,
        *,
        style_id: str | None = None,
        template_id: str | None = None,
        alignment_horizontal: str | None = None,
        alignment_vertical: str | None = None,
        line_spacing: int | str | None = None,
    ) -> ParagraphStyleXml:
        template = self._clone_header_template(".//hh:paraPr", template_id=template_id)
        resolved_id = style_id or self._next_header_numeric_id(".//hh:paraPr/@id")
        template.set("id", resolved_id)
        parent = self._ensure_header_collection(".//hh:paraProperties", "paraProperties")
        parent.append(template)
        parent.set("itemCnt", str(len(parent.xpath("./hh:paraPr", namespaces=NS))))
        self.header.mark_modified()

        style = ParagraphStyleXml(self.header, template)
        if alignment_horizontal is not None or alignment_vertical is not None:
            style.set_alignment(horizontal=alignment_horizontal, vertical=alignment_vertical)
        if line_spacing is not None:
            style.set_line_spacing(line_spacing)
        return style

    def append_character_style(
        self,
        *,
        style_id: str | None = None,
        template_id: str | None = None,
        text_color: str | None = None,
        height: int | str | None = None,
    ) -> CharacterStyleXml:
        template = self._clone_header_template(".//hh:charPr", template_id=template_id)
        resolved_id = style_id or self._next_header_numeric_id(".//hh:charPr/@id")
        template.set("id", resolved_id)
        parent = self._ensure_header_collection(".//hh:charProperties", "charProperties")
        parent.append(template)
        parent.set("itemCnt", str(len(parent.xpath("./hh:charPr", namespaces=NS))))
        self.header.mark_modified()

        style = CharacterStyleXml(self.header, template)
        if text_color is not None:
            style.set_text_color(text_color)
        if height is not None:
            style.set_height(height)
        return style

    def append_style(
        self,
        name: str,
        *,
        english_name: str | None = None,
        style_id: str | None = None,
        style_type: str = "PARA",
        para_pr_id: str | None = None,
        char_pr_id: str | None = None,
        next_style_id: str | None = None,
        lang_id: str = "1042",
        lock_form: bool = False,
        template_id: str | None = None,
    ) -> StyleDefinitionXml:
        template = self._clone_header_template(".//hh:style", template_id=template_id)
        resolved_id = style_id or self._next_header_numeric_id(".//hh:style/@id")
        template.set("id", resolved_id)
        template.set("type", style_type)
        template.set("name", name)
        template.set("engName", english_name or name)
        template.set("paraPrIDRef", para_pr_id or template.get("paraPrIDRef", "0"))
        template.set("charPrIDRef", char_pr_id or template.get("charPrIDRef", "0"))
        template.set("nextStyleIDRef", next_style_id or template.get("nextStyleIDRef", resolved_id))
        template.set("langID", lang_id)
        template.set("lockForm", "1" if lock_form else "0")

        parent = self._ensure_header_collection(".//hh:styles", "styles")
        parent.append(template)
        parent.set("itemCnt", str(len(parent.xpath("./hh:style", namespaces=NS))))
        self.header.mark_modified()
        return StyleDefinitionXml(self.header, template)



for _compat_method in HwpxDocumentControlMixin.__dict__.values():
    if callable(_compat_method):
        _compat_method.__module__ = "jakal_hwpx.document"

