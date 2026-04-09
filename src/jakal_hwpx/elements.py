from __future__ import annotations

from collections import Counter
from copy import deepcopy
from dataclasses import dataclass

from lxml import etree

from .exceptions import ValidationIssue
from .namespaces import NS, qname


def _text_nodes(element: etree._Element) -> list[etree._Element]:
    return list(element.xpath(".//hp:t", namespaces=NS))


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

    for paragraph in element.xpath(".//hp:p", namespaces=NS):
        add(paragraph)

    return paragraphs


def _invalidate_paragraph_layout(element: etree._Element) -> None:
    for paragraph in _paragraphs_affected_by_text_edit(element):
        for line_seg_array in paragraph.xpath("./hp:linesegarray", namespaces=NS):
            paragraph.remove(line_seg_array)


def _paragraph_nodes(element: etree._Element) -> list[etree._Element]:
    return list(element.xpath(".//hp:p", namespaces=NS))


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
    if paragraphs and (len(paragraphs) > 1 or "\n" in text):
        parts = _distribute_text_across_paragraphs(text, paragraphs)
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
class HeaderFooterBlock:
    document: object
    section: object
    element: etree._Element

    @property
    def kind(self) -> str:
        return etree.QName(self.element).localname

    @property
    def apply_page_type(self) -> str | None:
        return self.element.get("applyPageType")

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
class TableCell:
    table: "Table"
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
class Table:
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

    def cells(self) -> list[TableCell]:
        return [TableCell(self, cell) for cell in self.element.xpath("./hp:tr/hp:tc", namespaces=NS)]

    def cell(self, row: int, column: int) -> TableCell:
        for cell in self.cells():
            if cell.row == row and cell.column == column:
                return cell
        raise IndexError(f"Cell ({row}, {column}) was not found.")

    def set_cell_text(self, row: int, column: int, text: str) -> None:
        self.cell(row, column).set_text(text)

    def rows(self) -> list[list[TableCell]]:
        grouped: dict[int, list[TableCell]] = {}
        for cell in self.cells():
            grouped.setdefault(cell.row, []).append(cell)
        return [sorted(grouped[index], key=lambda item: item.column) for index in sorted(grouped)]

    def append_row(self) -> list[TableCell]:
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
        return [TableCell(self, cell) for cell in new_row.xpath("./hp:tc", namespaces=NS)]

    def merge_cells(self, start_row: int, start_column: int, end_row: int, end_column: int) -> TableCell:
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
class Picture:
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


@dataclass
class StyleDefinition:
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


@dataclass
class ParagraphStyle:
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


@dataclass
class CharacterStyle:
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


@dataclass
class SectionSettings:
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


@dataclass
class Note:
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
class Bookmark:
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
class Field:
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
class AutoNumber:
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
class Equation:
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


@dataclass
class ShapeObject:
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
