from __future__ import annotations

import json

from lxml import etree

from .namespaces import NS


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

