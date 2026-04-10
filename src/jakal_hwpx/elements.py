from __future__ import annotations

from collections import Counter
from copy import deepcopy
from dataclasses import dataclass

from lxml import etree

from .exceptions import ValidationIssue
from .namespaces import NS, qname


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
    _set_optional_attributes(
        _first_node(element, "./hp:lineShape"),
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
SectionSettings = SectionSettingsXml
Note = NoteXml
Bookmark = BookmarkXml
Field = FieldXml
AutoNumber = AutoNumberXml
Equation = EquationXml
ShapeObject = ShapeXml
OleObject = OleXml
