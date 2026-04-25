from __future__ import annotations

from collections import Counter
from copy import deepcopy
from dataclasses import dataclass
import json

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
_CHART_METADATA_URI = "urn:jakal-hwpx:chart-metadata"
_HANCOM_CHART_STYLE_URI = "CC8EB2C9-7E31-499d-B8F2-F6CE61031016"
_CHART_NS = {
    "c": NS["c"],
    "a": NS["a"],
    "r": NS["r"],
    "c14": NS["c14"],
    "ho": NS["ho"],
    "mc": NS["mc"],
    "jakalchart": NS["jakalchart"],
}
_CHART_TYPE_ELEMENT_BY_NAME = {
    "BAR": "barChart",
    "COLUMN": "barChart",
    "LINE": "lineChart",
    "PIE": "pieChart",
    "DOUGHNUT": "doughnutChart",
    "AREA": "areaChart",
}
_CHART_TYPE_NAME_BY_ELEMENT = {
    "barChart": "COLUMN",
    "lineChart": "LINE",
    "pieChart": "PIE",
    "doughnutChart": "DOUGHNUT",
    "areaChart": "AREA",
}


def _chart_qname(prefix: str, local_name: str) -> str:
    return f"{{{_CHART_NS[prefix]}}}{local_name}"


def _build_chart_text_properties(*, font_size: int | None = None) -> etree._Element:
    tx_pr = etree.Element(_chart_qname("c", "txPr"))
    body_pr = etree.SubElement(tx_pr, _chart_qname("a", "bodyPr"))
    body_pr.set("rot", "0")
    body_pr.set("vert", "horz")
    body_pr.set("wrap", "none")
    body_pr.set("lIns", "0")
    body_pr.set("tIns", "0")
    body_pr.set("rIns", "0")
    body_pr.set("bIns", "0")
    body_pr.set("anchor", "ctr")
    body_pr.set("anchorCtr", "1")
    paragraph = etree.SubElement(tx_pr, _chart_qname("a", "p"))
    paragraph_pr = etree.SubElement(paragraph, _chart_qname("a", "pPr"))
    paragraph_pr.set("algn", "l")
    default_run_pr = etree.SubElement(paragraph_pr, _chart_qname("a", "defRPr"))
    if font_size is not None:
        default_run_pr.set("sz", str(font_size))
    default_run_pr.set("b", "0")
    default_run_pr.set("i", "0")
    default_run_pr.set("u", "none")
    etree.SubElement(paragraph, _chart_qname("a", "endParaRPr"))
    return tx_pr


def _ensure_chart_style_block(root: etree._Element) -> None:
    if root.find(_chart_qname("mc", "AlternateContent")) is not None:
        return
    alternate_content = etree.Element(_chart_qname("mc", "AlternateContent"))
    choice = etree.SubElement(alternate_content, _chart_qname("mc", "Choice"))
    choice.set("Requires", "c14")
    style_2010 = etree.SubElement(choice, _chart_qname("c14", "style"))
    style_2010.set("val", "102")
    fallback = etree.SubElement(alternate_content, _chart_qname("mc", "Fallback"))
    style = etree.SubElement(fallback, _chart_qname("c", "style"))
    style.set("val", "2")
    insert_index = 2 if len(root) >= 2 else len(root)
    root.insert(insert_index, alternate_content)


def _ensure_chart_root_text_properties(root: etree._Element) -> None:
    if root.find(_chart_qname("c", "txPr")) is None:
        root.append(_build_chart_text_properties(font_size=1000))


def _ensure_chart_hancom_extension(root: etree._Element) -> None:
    ext_list = root.find(_chart_qname("c", "extLst"))
    if ext_list is None:
        ext_list = etree.SubElement(root, _chart_qname("c", "extLst"))
    for candidate in ext_list.findall(_chart_qname("c", "ext")):
        if candidate.get("uri") == _HANCOM_CHART_STYLE_URI:
            return
    ext = etree.Element(_chart_qname("c", "ext"))
    ext.set("uri", _HANCOM_CHART_STYLE_URI)
    style = etree.SubElement(ext, _chart_qname("ho", "hncChartStyle"))
    style.set("layoutIndex", "-1")
    style.set("colorIndex", "0")
    style.set("styleIndex", "0")
    ext_list.insert(0, ext)


def _ensure_chart_plot_area_style(plot_area: etree._Element) -> None:
    if plot_area.find(_chart_qname("c", "spPr")) is not None:
        return
    sp_pr = etree.SubElement(plot_area, _chart_qname("c", "spPr"))
    etree.SubElement(sp_pr, _chart_qname("a", "noFill"))
    line = etree.SubElement(sp_pr, _chart_qname("a", "ln"))
    line.set("w", "9525")
    line.set("cap", "flat")
    line.set("cmpd", "sng")
    line.set("algn", "ctr")
    etree.SubElement(line, _chart_qname("a", "noFill"))
    dash = etree.SubElement(line, _chart_qname("a", "prstDash"))
    dash.set("val", "solid")
    etree.SubElement(line, _chart_qname("a", "round"))
    head_end = etree.SubElement(line, _chart_qname("a", "headEnd"))
    head_end.set("w", "med")
    head_end.set("len", "med")
    tail_end = etree.SubElement(line, _chart_qname("a", "tailEnd"))
    tail_end.set("w", "med")
    tail_end.set("len", "med")


def _text_nodes(element: etree._Element) -> list[etree._Element]:
    return list(element.xpath(".//hp:t", namespaces=NS))


def _first_node(element: etree._Element, expression: str) -> etree._Element | None:
    nodes = element.xpath(expression, namespaces=NS)
    return nodes[0] if nodes else None


def _bool_attr(value: bool) -> str:
    return "1" if value else "0"


def _set_optional_attributes(element: etree._Element | None, **attrs: object) -> None:
    if element is None:
        return
    for key, value in attrs.items():
        if value is None:
            continue
        if isinstance(value, bool):
            element.set(key, _bool_attr(value))
        else:
            element.set(key, str(value))


def _graphic_layout(element: etree._Element) -> dict[str, str]:
    pos = _first_node(element, "./hp:pos")
    layout = {
        "textWrap": element.get("textWrap", ""),
        "textFlow": element.get("textFlow", ""),
    }
    if pos is None:
        return layout
    for key in (
        "treatAsChar",
        "affectLSpacing",
        "flowWithText",
        "allowOverlap",
        "holdAnchorAndSO",
        "vertRelTo",
        "horzRelTo",
        "vertAlign",
        "horzAlign",
        "vertOffset",
        "horzOffset",
    ):
        layout[key] = pos.get(key, "")
    return layout


def _set_graphic_layout(
    element: etree._Element,
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
    _set_optional_attributes(
        element,
        textWrap=text_wrap,
        textFlow=text_flow,
    )
    pos = _first_node(element, "./hp:pos")
    _set_optional_attributes(
        pos,
        treatAsChar=treat_as_char,
        affectLSpacing=affect_line_spacing,
        flowWithText=flow_with_text,
        allowOverlap=allow_overlap,
        holdAnchorAndSO=hold_anchor_and_so,
        vertRelTo=vert_rel_to,
        horzRelTo=horz_rel_to,
        vertAlign=vert_align,
        horzAlign=horz_align,
        vertOffset=vert_offset,
        horzOffset=horz_offset,
    )


def _margin_values(element: etree._Element, expression: str) -> dict[str, int]:
    margin = _first_node(element, expression)
    if margin is None:
        return {}
    return {key: int(margin.get(key, "0")) for key in ("left", "right", "top", "bottom")}


def _set_margin_values(
    element: etree._Element,
    expression: str,
    *,
    left: int | str | None = None,
    right: int | str | None = None,
    top: int | str | None = None,
    bottom: int | str | None = None,
) -> None:
    margin = _first_node(element, expression)
    _set_optional_attributes(margin, left=left, right=right, top=top, bottom=bottom)


def _graphic_size(element: etree._Element) -> dict[str, int]:
    size = _first_node(element, "./hp:sz")
    if size is None:
        return {}
    return {
        "width": int(size.get("width", "0")),
        "height": int(size.get("height", "0")),
    }


def _set_graphic_size(
    element: etree._Element,
    *,
    width: int | str | None = None,
    height: int | str | None = None,
    original_width: int | str | None = None,
    original_height: int | str | None = None,
    current_width: int | str | None = None,
    current_height: int | str | None = None,
    extent_x: int | str | None = None,
    extent_y: int | str | None = None,
) -> None:
    _set_optional_attributes(_first_node(element, "./hp:sz"), width=width, height=height)
    _set_optional_attributes(_first_node(element, "./hp:orgSz"), width=original_width, height=original_height)
    _set_optional_attributes(_first_node(element, "./hp:curSz"), width=current_width, height=current_height)
    _set_optional_attributes(_first_node(element, "./hc:extent"), x=extent_x, y=extent_y)


def _graphic_rotation(element: etree._Element) -> dict[str, str]:
    node = _first_node(element, "./hp:rotationInfo")
    if node is None:
        return {}
    return {
        key: node.get(key, "")
        for key in ("angle", "centerX", "centerY", "rotateimage")
    }


def _set_graphic_rotation(
    element: etree._Element,
    *,
    angle: int | str | None = None,
    center_x: int | str | None = None,
    center_y: int | str | None = None,
    rotate_image: bool | None = None,
) -> None:
    node = _first_node(element, "./hp:rotationInfo")
    if node is None:
        node = etree.SubElement(element, qname("hp", "rotationInfo"))
    _set_optional_attributes(
        node,
        angle=angle,
        centerX=center_x,
        centerY=center_y,
        rotateimage=rotate_image,
    )


def _line_style(element: etree._Element) -> dict[str, str]:
    node = _first_node(element, "./hp:lineShape")
    if node is None:
        return {}
    return {
        key: node.get(key, "")
        for key in (
            "color",
            "width",
            "style",
            "endCap",
            "headStyle",
            "tailStyle",
            "headfill",
            "tailfill",
            "headSz",
            "tailSz",
            "outlineStyle",
            "alpha",
        )
    }


def _set_line_style(
    element: etree._Element,
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
    node = _first_node(element, "./hp:lineShape")
    if node is None:
        node = etree.SubElement(element, qname("hp", "lineShape"))
    _set_optional_attributes(
        node,
        color=color,
        width=width,
        style=style,
        endCap=end_cap,
        headStyle=head_style,
        tailStyle=tail_style,
        headfill=head_fill,
        tailfill=tail_fill,
        headSz=head_size,
        tailSz=tail_size,
        outlineStyle=outline_style,
        alpha=alpha,
    )


def _chart_part_root(chart_part) -> etree._Element:
    return chart_part._root


def _chart_metadata_from_root(root: etree._Element) -> dict[str, object]:
    values = root.xpath(
        "./c:extLst/c:ext[@uri=$uri]/jakalchart:metadata/text()",
        namespaces=_CHART_NS,
        uri=_CHART_METADATA_URI,
    )
    if not values:
        return {}
    try:
        decoded = json.loads(values[0])
    except json.JSONDecodeError:
        return {}
    return decoded if isinstance(decoded, dict) else {}


def _set_chart_metadata(root: etree._Element, metadata: dict[str, object]) -> None:
    ext_list = root.find(_chart_qname("c", "extLst"))
    ext = None
    if ext_list is not None:
        for candidate in ext_list.findall(_chart_qname("c", "ext")):
            if candidate.get("uri") == _CHART_METADATA_URI:
                ext = candidate
                break
    if not metadata:
        if ext is not None and ext_list is not None:
            ext_list.remove(ext)
            if len(ext_list) == 0:
                root.remove(ext_list)
        return
    if ext_list is None:
        ext_list = etree.SubElement(root, _chart_qname("c", "extLst"))
    if ext is None:
        ext = etree.SubElement(ext_list, _chart_qname("c", "ext"))
        ext.set("uri", _CHART_METADATA_URI)
    nodes = ext.findall(_chart_qname("jakalchart", "metadata"))
    metadata_node = nodes[0] if nodes else etree.SubElement(ext, _chart_qname("jakalchart", "metadata"))
    metadata_node.text = json.dumps(metadata, ensure_ascii=False, separators=(",", ":"), sort_keys=True)


def _chart_title_from_root(root: etree._Element) -> str:
    values = root.xpath(
        ".//c:chart/c:title//a:t/text() | .//c:chart/c:title//c:v/text()",
        namespaces=_CHART_NS,
    )
    return "".join(values)


def _set_chart_title(root: etree._Element, title: str) -> None:
    chart = root.find(_chart_qname("c", "chart"))
    if chart is None:
        chart = etree.SubElement(root, _chart_qname("c", "chart"))
    title_node = chart.find(_chart_qname("c", "title"))
    if title_node is not None:
        chart.remove(title_node)
    auto_title = chart.find(_chart_qname("c", "autoTitleDeleted"))
    if title:
        title_node = etree.Element(_chart_qname("c", "title"))
        tx_node = etree.SubElement(title_node, _chart_qname("c", "tx"))
        rich = etree.SubElement(tx_node, _chart_qname("c", "rich"))
        etree.SubElement(rich, _chart_qname("a", "bodyPr"))
        etree.SubElement(rich, _chart_qname("a", "lstStyle"))
        paragraph = etree.SubElement(rich, _chart_qname("a", "p"))
        run = etree.SubElement(paragraph, _chart_qname("a", "r"))
        etree.SubElement(run, _chart_qname("a", "t")).text = title
        etree.SubElement(title_node, _chart_qname("c", "layout"))
        overlay = etree.SubElement(title_node, _chart_qname("c", "overlay"))
        overlay.set("val", "0")
        title_node.append(_build_chart_text_properties())
        insert_index = list(chart).index(auto_title) if auto_title is not None else 0
        chart.insert(insert_index, title_node)
    if auto_title is None:
        auto_title = etree.SubElement(chart, _chart_qname("c", "autoTitleDeleted"))
    auto_title.set("val", "0" if title else "1")


def _chart_type_from_root(root: etree._Element) -> str:
    metadata = _chart_metadata_from_root(root)
    metadata_type = metadata.get("chartType")
    if isinstance(metadata_type, str) and metadata_type:
        return metadata_type
    plot_area = root.find(".//" + _chart_qname("c", "plotArea"))
    if plot_area is None:
        return "BAR"
    for child in plot_area:
        local_name = etree.QName(child).localname
        if local_name not in _CHART_TYPE_NAME_BY_ELEMENT:
            continue
        if local_name == "barChart":
            bar_dir = child.find(_chart_qname("c", "barDir"))
            if bar_dir is not None and bar_dir.get("val") == "bar":
                return "BAR"
        return _CHART_TYPE_NAME_BY_ELEMENT[local_name]
    return "BAR"


def _chart_data_ref_from_root(root: etree._Element) -> str | None:
    formulas = root.xpath(
        ".//c:plotArea//c:ser[1]/c:cat/c:strRef/c:f/text()"
        " | .//c:plotArea//c:ser[1]/c:val/c:numRef/c:f/text()"
        " | .//c:plotArea//c:ser[1]/c:xVal/c:numRef/c:f/text()"
        " | .//c:plotArea//c:ser[1]/c:yVal/c:numRef/c:f/text()",
        namespaces=_CHART_NS,
    )
    for formula in formulas:
        anchor = _chart_sheet_anchor_from_formula(str(formula))
        if anchor:
            return anchor
    return None


def _chart_categories_from_root(root: etree._Element) -> list[str]:
    values = root.xpath(
        ".//c:plotArea//c:ser[1]/c:cat//c:pt/c:v/text()",
        namespaces=_CHART_NS,
    )
    return [str(value) for value in values]


def _chart_series_from_root(root: etree._Element) -> list[dict[str, object]]:
    series_values: list[dict[str, object]] = []
    for node in root.xpath(".//c:plotArea//c:ser", namespaces=_CHART_NS):
        name_values = node.xpath("./c:tx//c:v/text()", namespaces=_CHART_NS)
        point_values: list[object] = []
        for value in node.xpath("./c:val//c:pt/c:v/text()", namespaces=_CHART_NS):
            try:
                numeric = float(value)
                point_values.append(int(numeric) if numeric.is_integer() else numeric)
            except ValueError:
                point_values.append(value)
        series_values.append(
            {
                "name": name_values[0] if name_values else "",
                "values": point_values,
            }
        )
    return series_values


def _set_chart_legend_visible(root: etree._Element, visible: bool) -> None:
    chart = root.find(_chart_qname("c", "chart"))
    if chart is None:
        chart = etree.SubElement(root, _chart_qname("c", "chart"))
    legend = chart.find(_chart_qname("c", "legend"))
    if visible:
        if legend is None:
            legend = etree.SubElement(chart, _chart_qname("c", "legend"))
            legend_pos = etree.SubElement(legend, _chart_qname("c", "legendPos"))
            legend_pos.set("val", "r")
            etree.SubElement(legend, _chart_qname("c", "layout"))
            overlay = etree.SubElement(legend, _chart_qname("c", "overlay"))
            overlay.set("val", "0")
    elif legend is not None:
        chart.remove(legend)


def _chart_sheet_anchor_from_formula(formula: str) -> str | None:
    normalized = formula.strip()
    if "!" not in normalized:
        return None
    sheet_name, _ = normalized.split("!", 1)
    sheet_name = sheet_name.strip()
    if sheet_name.startswith("'") and sheet_name.endswith("'") and len(sheet_name) >= 2:
        sheet_name = sheet_name[1:-1].replace("''", "'")
    return sheet_name or None


def _chart_formula_sheet_name(data_ref: str) -> str:
    return "'" + data_ref.replace("'", "''") + "'"


def _chart_formula_column_name(index: int) -> str:
    if index < 0:
        raise ValueError("chart formula column index must be non-negative")
    value = index + 1
    column_name = ""
    while value > 0:
        value, remainder = divmod(value - 1, 26)
        column_name = chr(ord("A") + remainder) + column_name
    return column_name


def _chart_formula_range(data_ref: str, *, column_index: int, point_count: int) -> str:
    end_row = max(point_count, 1) + 1
    column_name = _chart_formula_column_name(column_index)
    return f"{_chart_formula_sheet_name(data_ref)}!${column_name}$2:${column_name}${end_row}"


def _build_chart_series_node(
    *,
    chart_type: str,
    categories: list[str],
    data_ref: str | None,
    series_index: int,
    series: dict[str, object],
) -> etree._Element:
    series_node = etree.Element(_chart_qname("c", "ser"))
    idx_node = etree.SubElement(series_node, _chart_qname("c", "idx"))
    idx_node.set("val", str(series_index))
    order_node = etree.SubElement(series_node, _chart_qname("c", "order"))
    order_node.set("val", str(series_index))

    name = str(series.get("name", ""))
    if name:
        tx_node = etree.SubElement(series_node, _chart_qname("c", "tx"))
        etree.SubElement(tx_node, _chart_qname("c", "v")).text = name

    if chart_type not in {"PIE", "DOUGHNUT"}:
        invert = etree.SubElement(series_node, _chart_qname("c", "invertIfNegative"))
        invert.set("val", "0")

    if categories:
        cat_node = etree.SubElement(series_node, _chart_qname("c", "cat"))
        if data_ref:
            category_ref = etree.SubElement(cat_node, _chart_qname("c", "strRef"))
            etree.SubElement(category_ref, _chart_qname("c", "f")).text = _chart_formula_range(
                data_ref,
                column_index=0,
                point_count=len(categories),
            )
            literal = etree.SubElement(category_ref, _chart_qname("c", "strCache"))
        else:
            literal = etree.SubElement(cat_node, _chart_qname("c", "strLit"))
        point_count = etree.SubElement(literal, _chart_qname("c", "ptCount"))
        point_count.set("val", str(len(categories)))
        for index, category in enumerate(categories):
            point = etree.SubElement(literal, _chart_qname("c", "pt"))
            point.set("idx", str(index))
            etree.SubElement(point, _chart_qname("c", "v")).text = str(category)

    values = series.get("values", [])
    if isinstance(values, list):
        val_node = etree.SubElement(series_node, _chart_qname("c", "val"))
        if data_ref:
            value_ref = etree.SubElement(val_node, _chart_qname("c", "numRef"))
            etree.SubElement(value_ref, _chart_qname("c", "f")).text = _chart_formula_range(
                data_ref,
                column_index=series_index + 1,
                point_count=len(values),
            )
            literal = etree.SubElement(value_ref, _chart_qname("c", "numCache"))
        else:
            literal = etree.SubElement(val_node, _chart_qname("c", "numLit"))
        format_code = etree.SubElement(literal, _chart_qname("c", "formatCode"))
        format_code.text = "General"
        point_count = etree.SubElement(literal, _chart_qname("c", "ptCount"))
        point_count.set("val", str(len(values)))
        for index, value in enumerate(values):
            point = etree.SubElement(literal, _chart_qname("c", "pt"))
            point.set("idx", str(index))
            etree.SubElement(point, _chart_qname("c", "v")).text = str(value)
    return series_node


def _build_chart_part_root(
    *,
    title: str,
    chart_type: str,
    categories: list[str],
    series: list[dict[str, object]],
    data_ref: str | None = None,
    legend_visible: bool,
    metadata: dict[str, object] | None = None,
) -> etree._Element:
    normalized_chart_type = str(chart_type or "BAR").upper()
    element_name = _CHART_TYPE_ELEMENT_BY_NAME.get(normalized_chart_type, "barChart")
    root = etree.Element(_chart_qname("c", "chartSpace"), nsmap=_CHART_NS)

    date1904 = etree.SubElement(root, _chart_qname("c", "date1904"))
    date1904.set("val", "0")
    rounded = etree.SubElement(root, _chart_qname("c", "roundedCorners"))
    rounded.set("val", "0")
    _ensure_chart_style_block(root)

    chart = etree.SubElement(root, _chart_qname("c", "chart"))
    _set_chart_title(root, title)
    plot_area = etree.SubElement(chart, _chart_qname("c", "plotArea"))
    etree.SubElement(plot_area, _chart_qname("c", "layout"))

    chart_node = etree.SubElement(plot_area, _chart_qname("c", element_name))
    if element_name == "barChart":
        bar_dir = etree.SubElement(chart_node, _chart_qname("c", "barDir"))
        bar_dir.set("val", "bar" if normalized_chart_type == "BAR" else "col")
        grouping = etree.SubElement(chart_node, _chart_qname("c", "grouping"))
        grouping.set("val", "clustered")
    elif element_name == "lineChart":
        grouping = etree.SubElement(chart_node, _chart_qname("c", "grouping"))
        grouping.set("val", "standard")
    elif element_name == "areaChart":
        grouping = etree.SubElement(chart_node, _chart_qname("c", "grouping"))
        grouping.set("val", "standard")

    vary_colors = etree.SubElement(chart_node, _chart_qname("c", "varyColors"))
    vary_colors.set("val", "1" if normalized_chart_type in {"PIE", "DOUGHNUT"} else "0")
    for index, value in enumerate(series):
        chart_node.append(
            _build_chart_series_node(
                chart_type=normalized_chart_type,
                categories=list(categories),
                data_ref=data_ref,
                series_index=index,
                series=value,
            )
        )

    if element_name == "barChart":
        gap_width = etree.SubElement(chart_node, _chart_qname("c", "gapWidth"))
        gap_width.set("val", "150")
        overlap = etree.SubElement(chart_node, _chart_qname("c", "overlap"))
        overlap.set("val", "0")

    if normalized_chart_type == "PIE":
        first_slice = etree.SubElement(chart_node, _chart_qname("c", "firstSliceAng"))
        first_slice.set("val", "0")
    elif normalized_chart_type == "DOUGHNUT":
        hole_size = etree.SubElement(chart_node, _chart_qname("c", "holeSize"))
        hole_size.set("val", "50")

    if normalized_chart_type not in {"PIE", "DOUGHNUT"}:
        for axis_id in ("1", "2"):
            axis = etree.SubElement(chart_node, _chart_qname("c", "axId"))
            axis.set("val", axis_id)

        cat_axis = etree.SubElement(plot_area, _chart_qname("c", "catAx"))
        cat_axis_id = etree.SubElement(cat_axis, _chart_qname("c", "axId"))
        cat_axis_id.set("val", "1")
        scaling = etree.SubElement(cat_axis, _chart_qname("c", "scaling"))
        orientation = etree.SubElement(scaling, _chart_qname("c", "orientation"))
        orientation.set("val", "minMax")
        axis_pos = etree.SubElement(cat_axis, _chart_qname("c", "axPos"))
        axis_pos.set("val", "b")
        cross_axis = etree.SubElement(cat_axis, _chart_qname("c", "crossAx"))
        cross_axis.set("val", "2")
        delete_axis = etree.SubElement(cat_axis, _chart_qname("c", "delete"))
        delete_axis.set("val", "0")
        major_tick = etree.SubElement(cat_axis, _chart_qname("c", "majorTickMark"))
        major_tick.set("val", "out")
        minor_tick = etree.SubElement(cat_axis, _chart_qname("c", "minorTickMark"))
        minor_tick.set("val", "none")
        tick_label = etree.SubElement(cat_axis, _chart_qname("c", "tickLblPos"))
        tick_label.set("val", "nextTo")
        crosses = etree.SubElement(cat_axis, _chart_qname("c", "crosses"))
        crosses.set("val", "autoZero")
        auto = etree.SubElement(cat_axis, _chart_qname("c", "auto"))
        auto.set("val", "1")
        label_align = etree.SubElement(cat_axis, _chart_qname("c", "lblAlgn"))
        label_align.set("val", "ctr")
        label_offset = etree.SubElement(cat_axis, _chart_qname("c", "lblOffset"))
        label_offset.set("val", "100")
        tick_mark_skip = etree.SubElement(cat_axis, _chart_qname("c", "tickMarkSkip"))
        tick_mark_skip.set("val", "1")
        no_multi_level = etree.SubElement(cat_axis, _chart_qname("c", "noMultiLvlLbl"))
        no_multi_level.set("val", "0")

        value_axis = etree.SubElement(plot_area, _chart_qname("c", "valAx"))
        value_axis_id = etree.SubElement(value_axis, _chart_qname("c", "axId"))
        value_axis_id.set("val", "2")
        scaling = etree.SubElement(value_axis, _chart_qname("c", "scaling"))
        orientation = etree.SubElement(scaling, _chart_qname("c", "orientation"))
        orientation.set("val", "minMax")
        axis_pos = etree.SubElement(value_axis, _chart_qname("c", "axPos"))
        axis_pos.set("val", "l")
        cross_axis = etree.SubElement(value_axis, _chart_qname("c", "crossAx"))
        cross_axis.set("val", "1")
        delete_axis = etree.SubElement(value_axis, _chart_qname("c", "delete"))
        delete_axis.set("val", "0")
        etree.SubElement(value_axis, _chart_qname("c", "majorGridlines"))
        number_format = etree.SubElement(value_axis, _chart_qname("c", "numFmt"))
        number_format.set("formatCode", "General")
        number_format.set("sourceLinked", "1")
        major_tick = etree.SubElement(value_axis, _chart_qname("c", "majorTickMark"))
        major_tick.set("val", "out")
        minor_tick = etree.SubElement(value_axis, _chart_qname("c", "minorTickMark"))
        minor_tick.set("val", "none")
        tick_label = etree.SubElement(value_axis, _chart_qname("c", "tickLblPos"))
        tick_label.set("val", "nextTo")
        crosses = etree.SubElement(value_axis, _chart_qname("c", "crosses"))
        crosses.set("val", "autoZero")
        cross_between = etree.SubElement(value_axis, _chart_qname("c", "crossBetween"))
        cross_between.set("val", "between")

    _ensure_chart_plot_area_style(plot_area)

    _set_chart_legend_visible(root, legend_visible)
    plot_vis_only = etree.SubElement(chart, _chart_qname("c", "plotVisOnly"))
    plot_vis_only.set("val", "0")
    display_blanks = etree.SubElement(chart, _chart_qname("c", "dispBlanksAs"))
    display_blanks.set("val", "gap")
    _ensure_chart_root_text_properties(root)
    _ensure_chart_hancom_extension(root)

    stored_metadata = {key: value for key, value in (metadata or {}).items() if value not in (None, "", {}, [])}
    if stored_metadata:
        _set_chart_metadata(root, stored_metadata)
    return root


def _fill_style(element: etree._Element) -> dict[str, str]:
    node = _first_node(element, "./hc:fillBrush/hc:winBrush")
    if node is None:
        return {}
    return {
        key: node.get(key, "")
        for key in ("faceColor", "hatchColor", "alpha")
    }


def _set_fill_style(
    element: etree._Element,
    *,
    face_color: str | None = None,
    hatch_color: str | None = None,
    alpha: int | str | None = None,
) -> None:
    _set_optional_attributes(
        _first_node(element, "./hc:fillBrush/hc:winBrush"),
        faceColor=face_color,
        hatchColor=hatch_color,
        alpha=alpha,
    )


def _text_margin(element: etree._Element) -> dict[str, int]:
    return _margin_values(element, "./hp:drawText/hp:textMargin")


def _set_text_margin(
    element: etree._Element,
    *,
    left: int | str | None = None,
    right: int | str | None = None,
    top: int | str | None = None,
    bottom: int | str | None = None,
) -> None:
    _set_margin_values(element, "./hp:drawText/hp:textMargin", left=left, right=right, top=top, bottom=bottom)


def _image_adjustment(element: etree._Element) -> dict[str, str]:
    node = _first_node(element, "./hc:img")
    if node is None:
        return {}
    return {
        key: node.get(key, "")
        for key in ("bright", "contrast", "effect", "alpha")
    }


def _set_image_adjustment(
    element: etree._Element,
    *,
    bright: int | str | None = None,
    contrast: int | str | None = None,
    effect: str | None = None,
    alpha: int | str | None = None,
) -> None:
    _set_optional_attributes(_first_node(element, "./hc:img"), bright=bright, contrast=contrast, effect=effect, alpha=alpha)


def _crop_values(element: etree._Element) -> dict[str, int]:
    node = _first_node(element, "./hp:imgClip")
    if node is None:
        return {}
    return {key: int(node.get(key, "0")) for key in ("left", "right", "top", "bottom")}


def _set_crop_values(
    element: etree._Element,
    *,
    left: int | str | None = None,
    right: int | str | None = None,
    top: int | str | None = None,
    bottom: int | str | None = None,
) -> None:
    node = _first_node(element, "./hp:imgClip")
    if node is None:
        node = etree.SubElement(element, qname("hp", "imgClip"))
    _set_optional_attributes(node, left=left, right=right, top=top, bottom=bottom)


def _extract_text(element: etree._Element) -> str:
    return "".join(node.text or "" for node in _text_nodes(element))


def _replace_text(element: etree._Element, old: str, new: str, count: int = -1) -> int:
    if not old:
        raise ValueError("old must be non-empty.")
    remaining = count
    replaced = 0
    for node in _text_nodes(element):
        current = node.text or ""
        if old not in current:
            continue
        if remaining < 0:
            changed = current.count(old)
            node.text = current.replace(old, new)
        else:
            changed = 0
            updated = current
            while changed < remaining and old in updated:
                updated = updated.replace(old, new, 1)
                changed += 1
            node.text = updated
            remaining -= changed
        replaced += changed
        if changed and remaining == 0:
            break
    if replaced:
        _invalidate_paragraph_layout(element)
    return replaced


def _paragraphs_affected_by_text_edit(element: etree._Element) -> list[etree._Element]:
    paragraphs: list[etree._Element] = []
    seen: set[int] = set()

    def add(node: etree._Element) -> None:
        marker = id(node)
        if marker in seen:
            return
        seen.add(marker)
        paragraphs.append(node)

    if etree.QName(element).localname == "p":
        add(element)

    ancestors = element.xpath("ancestor::hp:p[1]", namespaces=NS)
    if ancestors:
        add(ancestors[0])

    for paragraph in _paragraph_nodes(element):
        add(paragraph)

    return paragraphs


def _invalidate_paragraph_layout(element: etree._Element) -> None:
    for paragraph in _paragraphs_affected_by_text_edit(element):
        for line_seg_array in paragraph.xpath("./hp:linesegarray", namespaces=NS):
            paragraph.remove(line_seg_array)


def _paragraph_nodes(element: etree._Element) -> list[etree._Element]:
    if etree.QName(element).localname == "p":
        return [element]

    paragraphs = list(element.xpath("./hp:p", namespaces=NS))
    for sublist in element.xpath("./hp:subList", namespaces=NS):
        paragraphs.extend(sublist.xpath("./hp:p", namespaces=NS))
    return paragraphs


def _ensure_first_paragraph(element: etree._Element) -> etree._Element:
    paragraph = element.xpath(".//hp:p[1]", namespaces=NS)
    if paragraph:
        return paragraph[0]

    sublists = element.xpath(".//hp:subList[1]", namespaces=NS)
    if sublists:
        sublist = sublists[0]
    else:
        sublist = etree.SubElement(element, qname("hp", "subList"))
        sublist.set("id", "")
        sublist.set("textDirection", "HORIZONTAL")
        sublist.set("lineWrap", "BREAK")
        sublist.set("vertAlign", "TOP")
        sublist.set("linkListIDRef", "0")
        sublist.set("linkListNextIDRef", "0")
        sublist.set("textWidth", "0")
        sublist.set("textHeight", "0")
        sublist.set("hasTextRef", "0")
        sublist.set("hasNumRef", "0")

    paragraph = etree.SubElement(sublist, qname("hp", "p"))
    paragraph.set("id", "0")
    paragraph.set("paraPrIDRef", "0")
    paragraph.set("styleIDRef", "0")
    paragraph.set("pageBreak", "0")
    paragraph.set("columnBreak", "0")
    paragraph.set("merged", "0")
    return paragraph


def _default_char_pr_id(paragraph: etree._Element) -> str:
    for run in paragraph.xpath("./hp:run", namespaces=NS):
        value = run.get("charPrIDRef")
        if value is not None:
            return value
    return "0"


def _ensure_run_with_text(paragraph: etree._Element, text: str) -> etree._Element:
    run = etree.SubElement(paragraph, qname("hp", "run"))
    run.set("charPrIDRef", _default_char_pr_id(paragraph))
    text_node = etree.SubElement(run, qname("hp", "t"))
    text_node.text = text
    return run


def _reset_paragraph_text(paragraph: etree._Element, text: str) -> None:
    for child in list(paragraph):
        paragraph.remove(child)
    _ensure_run_with_text(paragraph, text)


def _replace_paragraph_text_preserving_controls(
    paragraph: etree._Element,
    text: str,
    *,
    char_pr_id: str | None = None,
) -> None:
    resolved_char_pr_id = char_pr_id or _default_char_pr_id(paragraph)
    preserved_runs = [
        deepcopy(child)
        for child in paragraph.xpath("./hp:run[hp:secPr or hp:ctrl]", namespaces=NS)
    ]
    for preserved_run in preserved_runs:
        for text_node in preserved_run.xpath("./hp:t", namespaces=NS):
            preserved_run.remove(text_node)
    for child in list(paragraph):
        paragraph.remove(child)
    for preserved in preserved_runs:
        paragraph.append(preserved)
    run = etree.SubElement(paragraph, qname("hp", "run"))
    run.set("charPrIDRef", resolved_char_pr_id)
    text_node = etree.SubElement(run, qname("hp", "t"))
    text_node.text = text


def _preserved_structure_signature(paragraph: etree._Element) -> Counter[str]:
    signature: Counter[str] = Counter()
    for node in paragraph.xpath("./hp:run/hp:secPr | ./hp:run/hp:ctrl/*", namespaces=NS):
        local_name = etree.QName(node).localname
        parent_local_name = etree.QName(node.getparent()).localname if node.getparent() is not None else ""
        label = f"{parent_local_name}/{local_name}" if parent_local_name == "ctrl" else local_name
        attributes = []
        for key in ("id", "fieldid", "beginIDRef", "name", "instid", "num", "numType", "type"):
            value = node.get(key)
            if value:
                attributes.append(f"{key}={value}")
        if attributes:
            label = f"{label}[{', '.join(attributes)}]"
        signature[label] += 1
    return signature


def _capture_protected_paragraph_signatures(element: etree._Element) -> tuple[list[etree._Element], list[Counter[str]]]:
    paragraphs = _paragraph_nodes(element)
    return paragraphs, [_preserved_structure_signature(paragraph) for paragraph in paragraphs]


def _missing_preserved_tokens(expected: Counter[str], actual: Counter[str]) -> list[str]:
    missing: list[str] = []
    for token, count in expected.items():
        deficit = count - actual.get(token, 0)
        if deficit > 0:
            missing.extend([token] * deficit)
    return missing


def _clone_paragraph(paragraph: etree._Element) -> etree._Element:
    clone = deepcopy(paragraph)
    _reset_paragraph_text(clone, "")
    return clone


def _split_score(text: str, index: int, target: int) -> tuple[int, int]:
    penalty = abs(index - target) * 10
    prev_char = text[index - 1] if index > 0 else ""
    next_char = text[index] if index < len(text) else ""

    if prev_char in ")]":
        penalty -= 40
    elif prev_char == ",":
        penalty -= 30
    elif prev_char == " ":
        penalty -= 10

    if next_char == "[":
        penalty -= 25
    elif next_char == "(":
        penalty -= 10
    elif next_char == " ":
        penalty -= 5

    return penalty, index


def _snap_split_index(text: str, target: int, *, minimum: int, maximum: int) -> int:
    if minimum >= maximum:
        return minimum

    target = min(max(target, minimum), maximum)
    window = max(8, len(text) // 12)
    search_start = max(minimum, target - window)
    search_end = min(maximum, target + window)
    candidates = range(search_start, search_end + 1)
    return min(candidates, key=lambda index: _split_score(text, index, target))


def _distribute_text_across_paragraphs(text: str, paragraphs: list[etree._Element]) -> list[str]:
    if not paragraphs:
        return [text]

    if "\n" in text:
        return text.split("\n")

    if len(paragraphs) == 1:
        return [text]

    original_lengths = [len("".join(node.text or "" for node in paragraph.xpath(".//hp:t", namespaces=NS))) for paragraph in paragraphs]
    total_length = sum(original_lengths)
    if total_length <= 0:
        return [text] + [""] * (len(paragraphs) - 1)

    segments: list[str] = []
    start = 0

    for index in range(1, len(original_lengths)):
        remaining_paragraphs = len(original_lengths) - index
        minimum = start
        maximum = len(text) - remaining_paragraphs
        target = round(len(text) * (sum(original_lengths[:index]) / total_length))
        split_at = _snap_split_index(text, target, minimum=minimum, maximum=maximum)
        segments.append(text[start:split_at])
        start = split_at

    segments.append(text[start:])
    return segments


def _set_text(element: etree._Element, text: str) -> None:
    paragraphs = _paragraph_nodes(element)
    if paragraphs:
        parts = _distribute_text_across_paragraphs(text, paragraphs) if (len(paragraphs) > 1 or "\n" in text) else [text]
        while len(paragraphs) < len(parts):
            template = paragraphs[-1] if paragraphs else _ensure_first_paragraph(element)
            clone = _clone_paragraph(template)
            template.addnext(clone)
            paragraphs = _paragraph_nodes(element)
        for paragraph, value in zip(paragraphs, parts):
            _replace_paragraph_text_preserving_controls(paragraph, value)
        for paragraph in paragraphs[len(parts) :]:
            _replace_paragraph_text_preserving_controls(paragraph, "")
        _invalidate_paragraph_layout(element)
        return

    text_nodes = _text_nodes(element)
    if text_nodes:
        text_nodes[0].text = text
        for extra in text_nodes[1:]:
            extra.text = ""
        _invalidate_paragraph_layout(element)
        return

    paragraph = _ensure_first_paragraph(element)
    _replace_paragraph_text_preserving_controls(paragraph, text)
    _invalidate_paragraph_layout(element)


def _form_command_metadata(element: etree._Element) -> dict[str, object]:
    command = element.get("command", "")
    if not command.startswith(_FORM_METADATA_COMMAND_PREFIX):
        return {}
    try:
        payload = json.loads(command[len(_FORM_METADATA_COMMAND_PREFIX) :])
    except json.JSONDecodeError:
        return {}
    return payload if isinstance(payload, dict) else {}


def _write_form_command_metadata(
    element: etree._Element,
    *,
    updates: dict[str, object | None] | None = None,
) -> None:
    command = element.get("command", "")
    if command and not command.startswith(_FORM_METADATA_COMMAND_PREFIX):
        return
    metadata = dict(_form_command_metadata(element))
    for key, value in (updates or {}).items():
        if value is None or value == "":
            metadata.pop(key, None)
        else:
            metadata[str(key)] = value
    if metadata:
        element.set(
            "command",
            _FORM_METADATA_COMMAND_PREFIX + json.dumps(metadata, ensure_ascii=False, separators=(",", ":"), sort_keys=True),
        )
    else:
        element.set("command", "")


def _form_list_item_values(element: etree._Element) -> list[str]:
    values: list[str] = []
    for item in element.xpath("./hp:listItem", namespaces=NS):
        values.append(item.get("displayText") or item.get("value") or "")
    return values


def _set_form_list_items(element: etree._Element, values: list[str]) -> None:
    for item in list(element.xpath("./hp:listItem", namespaces=NS)):
        element.remove(item)
    for value in values:
        item = etree.SubElement(element, qname("hp", "listItem"))
        item.set("displayText", "")
        item.set("value", value)


def _form_caption_text(element: etree._Element) -> str | None:
    caption = _first_node(element, "./hp:caption")
    if caption is None:
        return None
    text = _extract_text(caption)
    return text if text else None


def _set_form_caption_text(
    element: etree._Element,
    value: str | None,
    *,
    char_pr_id: str | None = None,
) -> None:
    caption = _first_node(element, "./hp:caption")
    if not value:
        if caption is not None:
            element.remove(caption)
        return

    if caption is None:
        caption = etree.Element(qname("hp", "caption"))
        element.insert(0, caption)
    _set_optional_attributes(caption, side="LEFT", fullSz=False, width=-1, gap=850, lastWidth=0)

    sub_list = _first_node(caption, "./hp:subList")
    if sub_list is None:
        sub_list = etree.SubElement(caption, qname("hp", "subList"))
    _set_optional_attributes(
        sub_list,
        id="",
        textDirection="HORIZONTAL",
        lineWrap="BREAK",
        vertAlign="TOP",
        linkListIDRef="0",
        linkListNextIDRef="0",
        textWidth="0",
        textHeight="0",
        hasTextRef=False,
        hasNumRef=False,
    )

    paragraph = _first_node(sub_list, "./hp:p")
    if paragraph is None:
        paragraph = etree.SubElement(sub_list, qname("hp", "p"))
    _set_optional_attributes(
        paragraph,
        id="0",
        paraPrIDRef="0",
        styleIDRef="0",
        pageBreak="0",
        columnBreak="0",
        merged="0",
    )

    run = _first_node(paragraph, "./hp:run")
    if run is None:
        run = etree.SubElement(paragraph, qname("hp", "run"))
    run.set("charPrIDRef", char_pr_id or run.get("charPrIDRef") or "0")

    text_node = _first_form_text_node(run)
    if text_node is None:
        text_node = _first_node(run, "./hp:t")
    if text_node is None:
        text_node = etree.SubElement(run, qname("hp", "t"))
    text_node.text = value

    for extra_paragraph in list(sub_list.xpath("./hp:p[position()>1]", namespaces=NS)):
        sub_list.remove(extra_paragraph)
    for extra_run in list(paragraph.xpath("./hp:run[position()>1]", namespaces=NS)):
        paragraph.remove(extra_run)


def _first_form_text_node(element: etree._Element) -> etree._Element | None:
    nodes = element.xpath("./hp:text", namespaces=NS)
    return nodes[0] if nodes else None


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
        return self.field_type == "HYPERLINK"

    @property
    def is_mail_merge(self) -> bool:
        return self.field_type in {"MAILMERGE", "MAIL_MERGE", "MERGEFIELD"}

    @property
    def is_calculation(self) -> bool:
        return self.field_type in {"CALCULATE", "CALC", "FORMULA"}

    @property
    def is_cross_reference(self) -> bool:
        return self.field_type in {"REF", "PAGEREF", "BOOKMARKREF", "CROSSREF", "CROSS_REF"}

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
