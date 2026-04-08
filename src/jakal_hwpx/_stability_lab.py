from __future__ import annotations

import json
from collections import Counter
from copy import deepcopy
from dataclasses import asdict, dataclass
from pathlib import Path

from lxml import etree

from .document import HwpxDocument
from .elements import _preserved_structure_signature
from .exceptions import HwpxValidationError, ValidationIssue
from .namespaces import NS, qname


def _new_paragraph(*, paragraph_id: str, para_pr_id: str = "0", style_id: str = "0") -> etree._Element:
    paragraph = etree.Element(qname("hp", "p"))
    paragraph.set("id", paragraph_id)
    paragraph.set("paraPrIDRef", para_pr_id)
    paragraph.set("styleIDRef", style_id)
    paragraph.set("pageBreak", "0")
    paragraph.set("columnBreak", "0")
    paragraph.set("merged", "0")
    return paragraph


def _new_sublist(*, vertical_align: str = "CENTER") -> etree._Element:
    sublist = etree.Element(qname("hp", "subList"))
    sublist.set("id", "")
    sublist.set("textDirection", "HORIZONTAL")
    sublist.set("lineWrap", "BREAK")
    sublist.set("vertAlign", vertical_align)
    sublist.set("linkListIDRef", "0")
    sublist.set("linkListNextIDRef", "0")
    sublist.set("textWidth", "0")
    sublist.set("textHeight", "0")
    sublist.set("hasTextRef", "0")
    sublist.set("hasNumRef", "0")
    return sublist


def _append_text_run(paragraph: etree._Element, text: str, *, char_pr_id: str = "0") -> None:
    run = etree.SubElement(paragraph, qname("hp", "run"))
    run.set("charPrIDRef", char_pr_id)
    etree.SubElement(run, qname("hp", "t")).text = text


def _append_bookmark_run(paragraph: etree._Element, *, name: str, char_pr_id: str = "0") -> None:
    run = etree.SubElement(paragraph, qname("hp", "run"))
    run.set("charPrIDRef", char_pr_id)
    ctrl = etree.SubElement(run, qname("hp", "ctrl"))
    bookmark = etree.SubElement(ctrl, qname("hp", "bookmark"))
    bookmark.set("name", name)


def _append_newnum_run(paragraph: etree._Element, *, num: str = "1", num_type: str = "PAGE", char_pr_id: str = "0") -> None:
    run = etree.SubElement(paragraph, qname("hp", "run"))
    run.set("charPrIDRef", char_pr_id)
    ctrl = etree.SubElement(run, qname("hp", "ctrl"))
    node = etree.SubElement(ctrl, qname("hp", "newNum"))
    node.set("num", num)
    node.set("numType", num_type)


def _append_field_runs(
    paragraph: etree._Element,
    *,
    case_name: str,
    display_text: str,
    char_pr_id: str = "0",
) -> None:
    field_seed = sum(ord(char) for char in case_name)
    begin_id = str(1000 + field_seed)
    field_id = str(2000 + field_seed)

    begin_run = etree.SubElement(paragraph, qname("hp", "run"))
    begin_run.set("charPrIDRef", char_pr_id)
    begin_ctrl = etree.SubElement(begin_run, qname("hp", "ctrl"))
    field_begin = etree.SubElement(begin_ctrl, qname("hp", "fieldBegin"))
    field_begin.set("id", begin_id)
    field_begin.set("type", "HYPERLINK")
    field_begin.set("name", "")
    field_begin.set("editable", "0")
    field_begin.set("dirty", "0")
    field_begin.set("zorder", "0")
    field_begin.set("fieldid", field_id)

    display_run = etree.SubElement(paragraph, qname("hp", "run"))
    display_run.set("charPrIDRef", char_pr_id)
    etree.SubElement(display_run, qname("hp", "t")).text = display_text

    end_run = etree.SubElement(paragraph, qname("hp", "run"))
    end_run.set("charPrIDRef", char_pr_id)
    end_ctrl = etree.SubElement(end_run, qname("hp", "ctrl"))
    field_end = etree.SubElement(end_ctrl, qname("hp", "fieldEnd"))
    field_end.set("beginIDRef", begin_id)
    field_end.set("fieldid", field_id)


def _append_tokens(paragraph: etree._Element, tokens: tuple[str, ...], *, case_name: str, paragraph_offset: int) -> None:
    for token in tokens:
        if token == "text":
            _append_text_run(paragraph, f"{case_name}-TEXT-{paragraph_offset}")
            continue
        if token == "bookmark":
            _append_bookmark_run(paragraph, name=f"{case_name}-bookmark-{paragraph_offset}")
            continue
        if token == "field_pair":
            _append_field_runs(
                paragraph,
                case_name=f"{case_name}-{paragraph_offset}",
                display_text=f"{case_name}-FIELD-{paragraph_offset}",
            )
            continue
        if token == "newnum":
            _append_newnum_run(paragraph)
            continue
        raise ValueError(f"Unsupported token: {token}")


def _build_content_paragraphs(
    tokens_by_paragraph: tuple[tuple[str, ...], ...],
    *,
    case_name: str,
    start_id: int,
) -> list[etree._Element]:
    paragraphs: list[etree._Element] = []
    for offset, tokens in enumerate(tokens_by_paragraph):
        paragraph = _new_paragraph(paragraph_id=str(start_id + offset))
        _append_tokens(paragraph, tokens, case_name=case_name, paragraph_offset=offset)
        paragraphs.append(paragraph)
    return paragraphs


def _replace_paragraph_children_preserving_secpr(paragraph: etree._Element) -> None:
    preserved_secpr_runs = [
        deepcopy(run)
        for run in paragraph.xpath("./hp:run[hp:secPr]", namespaces=NS)
    ]
    for child in list(paragraph):
        paragraph.remove(child)
    for run in preserved_secpr_runs:
        paragraph.append(run)


def _install_section_case(document: HwpxDocument, case_name: str, tokens_by_paragraph: tuple[tuple[str, ...], ...]) -> None:
    section = document.sections[0]
    root = section.root_element
    base_paragraphs = root.xpath("./hp:p", namespaces=NS)
    if not base_paragraphs:
        raise ValueError("Blank document section must contain at least one paragraph.")

    first_paragraph = base_paragraphs[0]
    _replace_paragraph_children_preserving_secpr(first_paragraph)
    _append_tokens(first_paragraph, tokens_by_paragraph[0], case_name=case_name, paragraph_offset=0)

    for extra in base_paragraphs[1:]:
        root.remove(extra)

    for paragraph in _build_content_paragraphs(tokens_by_paragraph[1:], case_name=case_name, start_id=100):
        root.append(paragraph)
    section.mark_modified()


def _attach_control_block(document: HwpxDocument, block: etree._Element) -> None:
    paragraph = document.sections[0].root_element.xpath("./hp:p", namespaces=NS)[0]
    run = etree.SubElement(paragraph, qname("hp", "run"))
    run.set("charPrIDRef", "0")
    ctrl = etree.SubElement(run, qname("hp", "ctrl"))
    ctrl.append(block)
    document.sections[0].mark_modified()


def _attach_shape_block(document: HwpxDocument, block: etree._Element) -> None:
    paragraph = document.sections[0].root_element.xpath("./hp:p", namespaces=NS)[0]
    run = etree.SubElement(paragraph, qname("hp", "run"))
    run.set("charPrIDRef", "0")
    run.append(block)
    document.sections[0].mark_modified()


def _build_header_footer(kind: str, case_name: str, tokens_by_paragraph: tuple[tuple[str, ...], ...]) -> etree._Element:
    block = etree.Element(qname("hp", kind))
    block.set("id", "0")
    block.set("applyPageType", "BOTH")
    sublist = _new_sublist(vertical_align="TOP")
    for paragraph in _build_content_paragraphs(tokens_by_paragraph, case_name=case_name, start_id=200):
        sublist.append(paragraph)
    block.append(sublist)
    return block


def _build_table(case_name: str, tokens_by_paragraph: tuple[tuple[str, ...], ...]) -> etree._Element:
    table = etree.Element(qname("hp", "tbl"))
    table.set("rowCnt", "1")
    table.set("colCnt", "1")
    row = etree.SubElement(table, qname("hp", "tr"))
    cell = etree.SubElement(row, qname("hp", "tc"))
    sublist = _new_sublist()
    for paragraph in _build_content_paragraphs(tokens_by_paragraph, case_name=case_name, start_id=300):
        sublist.append(paragraph)
    cell.append(sublist)
    cell_addr = etree.SubElement(cell, qname("hp", "cellAddr"))
    cell_addr.set("colAddr", "0")
    cell_addr.set("rowAddr", "0")
    cell_span = etree.SubElement(cell, qname("hp", "cellSpan"))
    cell_span.set("colSpan", "1")
    cell_span.set("rowSpan", "1")
    return table


def _build_shape(case_name: str, tokens_by_paragraph: tuple[tuple[str, ...], ...]) -> etree._Element:
    rect = etree.Element(qname("hp", "rect"))
    draw_text = etree.SubElement(rect, qname("hp", "drawText"))
    sublist = _new_sublist()
    for paragraph in _build_content_paragraphs(tokens_by_paragraph, case_name=case_name, start_id=400):
        sublist.append(paragraph)
    draw_text.append(sublist)
    return rect


def _build_note(kind: str, case_name: str, tokens_by_paragraph: tuple[tuple[str, ...], ...]) -> etree._Element:
    note = etree.Element(qname("hp", kind))
    note.set("id", "0")
    note.set("number", "1")
    sublist = _new_sublist(vertical_align="TOP")
    for paragraph in _build_content_paragraphs(tokens_by_paragraph, case_name=case_name, start_id=500):
        sublist.append(paragraph)
    note.append(sublist)
    return note


def _section_signatures(document: HwpxDocument, expected_count: int) -> list[Counter[str]]:
    paragraphs = document.sections[0].root_element.xpath("./hp:p", namespaces=NS)[:expected_count]
    return [_preserved_structure_signature(paragraph) for paragraph in paragraphs]


def _header_signatures(document: HwpxDocument, kind: str) -> list[Counter[str]]:
    blocks = document.headers() if kind == "header" else document.footers()
    block = blocks[0]
    return [_preserved_structure_signature(paragraph) for paragraph in block.element.xpath(".//hp:p", namespaces=NS)]


def _table_signatures(document: HwpxDocument) -> list[Counter[str]]:
    cell = document.tables()[0].cell(0, 0)
    return [_preserved_structure_signature(paragraph) for paragraph in cell.element.xpath(".//hp:p", namespaces=NS)]


def _shape_signatures(document: HwpxDocument) -> list[Counter[str]]:
    shape = document.shapes()[0]
    return [_preserved_structure_signature(paragraph) for paragraph in shape.element.xpath(".//hp:drawText//hp:p", namespaces=NS)]


def _note_signatures(document: HwpxDocument) -> list[Counter[str]]:
    note = document.notes()[0]
    return [_preserved_structure_signature(paragraph) for paragraph in note.element.xpath(".//hp:p", namespaces=NS)]


@dataclass(frozen=True)
class StabilityCase:
    name: str
    container: str
    description: str
    tokens_by_paragraph: tuple[tuple[str, ...], ...]
    edit_text: str
    edit_mode: str = "set_text"


def _replace_target_text(case_name: str, tokens: tuple[str, ...], paragraph_offset: int = 0) -> str:
    if "text" in tokens:
        return f"{case_name}-TEXT-{paragraph_offset}"
    if "field_pair" in tokens:
        return f"{case_name}-FIELD-{paragraph_offset}"
    raise ValueError(f"No replaceable text token was found for case {case_name}: {tokens!r}")


@dataclass
class StabilityCaseResult:
    name: str
    container: str
    description: str
    baseline_errors: list[ValidationIssue]
    edited_errors: list[ValidationIssue]
    reopened_errors: list[ValidationIssue]
    control_errors: list[ValidationIssue]
    signature_mismatch: list[ValidationIssue]
    outputs: list[str]

    @property
    def ok(self) -> bool:
        return not (
            self.baseline_errors
            or self.edited_errors
            or self.reopened_errors
            or self.control_errors
            or self.signature_mismatch
        )


def stability_cases() -> list[StabilityCase]:
    cases: list[StabilityCase] = []
    container_labels = {
        "section": "Section",
        "header": "Header",
        "footer": "Footer",
        "table_cell": "Table cell",
        "shape": "Shape drawText",
        "footnote": "Footnote",
        "endnote": "Endnote",
    }
    section_combos = [
        ("plain_text", ("text",)),
        ("bookmark_text", ("bookmark", "text")),
        ("field_pair", ("field_pair",)),
        ("bookmark_field_text", ("bookmark", "field_pair", "text")),
        ("newnum_text", ("newnum", "text")),
    ]
    for name_suffix, tokens in section_combos:
        cases.append(
            StabilityCase(
                f"section_{name_suffix}",
                "section",
                f"{container_labels['section']} paragraph using tokens {tokens!r}.",
                (tokens,),
                "EDITED-SECTION",
            )
        )
        if "text" in tokens or "field_pair" in tokens:
            cases.append(
                StabilityCase(
                    f"section_{name_suffix}_replace",
                    "section",
                    f"{container_labels['section']} paragraph using document.replace_text for tokens {tokens!r}.",
                    (tokens,),
                    "EDITED-SECTION-REPLACE",
                    edit_mode="replace_text",
                )
            )

    multi_container_combos = {
        "header": [
            ("plain_text", ("text",)),
            ("bookmark_text", ("bookmark", "text")),
            ("field_pair", ("field_pair",)),
            ("newnum_text", ("newnum", "text")),
            ("bookmark_field_text", ("bookmark", "field_pair", "text")),
        ],
        "footer": [
            ("plain_text", ("text",)),
            ("bookmark_text", ("bookmark", "text")),
            ("field_pair", ("field_pair",)),
            ("newnum_text", ("newnum", "text")),
            ("bookmark_field_text", ("bookmark", "field_pair", "text")),
        ],
        "table_cell": [
            ("plain_text", ("text",)),
            ("bookmark_text", ("bookmark", "text")),
            ("field_pair", ("field_pair",)),
            ("bookmark_field_text", ("bookmark", "field_pair", "text")),
        ],
        "shape": [
            ("plain_text", ("text",)),
            ("bookmark_text", ("bookmark", "text")),
            ("field_pair", ("field_pair",)),
            ("bookmark_field_text", ("bookmark", "field_pair", "text")),
        ],
        "footnote": [
            ("plain_text", ("text",)),
            ("bookmark_text", ("bookmark", "text")),
            ("field_pair", ("field_pair",)),
            ("bookmark_field_text", ("bookmark", "field_pair", "text")),
        ],
        "endnote": [
            ("plain_text", ("text",)),
            ("bookmark_text", ("bookmark", "text")),
            ("field_pair", ("field_pair",)),
            ("bookmark_field_text", ("bookmark", "field_pair", "text")),
        ],
    }
    for container, combos in multi_container_combos.items():
        for name_suffix, tokens in combos:
            cases.append(
                StabilityCase(
                    f"{container}_{name_suffix}_single",
                    container,
                    f"{container_labels[container]} single-paragraph case using tokens {tokens!r}.",
                    (tokens,),
                    f"{container.upper()}-EDIT",
                )
            )
            if container in {"header", "footer"} and ("text" in tokens or "field_pair" in tokens):
                cases.append(
                    StabilityCase(
                        f"{container}_{name_suffix}_replace",
                        container,
                        f"{container_labels[container]} single-paragraph case using document.replace_text for tokens {tokens!r}.",
                        (tokens,),
                        f"{container.upper()}-REPLACE",
                        edit_mode="replace_text",
                    )
                )
            cases.append(
                StabilityCase(
                    f"{container}_{name_suffix}_dual",
                    container,
                    f"{container_labels[container]} dual-paragraph case using tokens {tokens!r} followed by plain text.",
                    (tokens, ("text",)),
                    f"{container.upper()}-A\n{container.upper()}-B",
                )
            )
    return cases


def _build_case_document(case: StabilityCase) -> tuple[HwpxDocument, list[Counter[str]], callable[[HwpxDocument], list[Counter[str]]]]:
    document = HwpxDocument.blank()
    if case.container == "section":
        _install_section_case(document, case.name, case.tokens_by_paragraph)
        return document, _section_signatures(document, len(case.tokens_by_paragraph)), lambda doc: _section_signatures(
            doc, len(case.tokens_by_paragraph)
        )
    if case.container == "header":
        _attach_control_block(document, _build_header_footer("header", case.name, case.tokens_by_paragraph))
        return document, _header_signatures(document, "header"), lambda doc: _header_signatures(doc, "header")
    if case.container == "footer":
        _attach_control_block(document, _build_header_footer("footer", case.name, case.tokens_by_paragraph))
        return document, _header_signatures(document, "footer"), lambda doc: _header_signatures(doc, "footer")
    if case.container == "table_cell":
        _attach_control_block(document, _build_table(case.name, case.tokens_by_paragraph))
        return document, _table_signatures(document), _table_signatures
    if case.container == "shape":
        _attach_shape_block(document, _build_shape(case.name, case.tokens_by_paragraph))
        return document, _shape_signatures(document), _shape_signatures
    if case.container == "footnote":
        _attach_control_block(document, _build_note("footNote", case.name, case.tokens_by_paragraph))
        return document, _note_signatures(document), _note_signatures
    if case.container == "endnote":
        _attach_control_block(document, _build_note("endNote", case.name, case.tokens_by_paragraph))
        return document, _note_signatures(document), _note_signatures
    raise ValueError(f"Unsupported container: {case.container}")


def _apply_case_edit(document: HwpxDocument, case: StabilityCase) -> int:
    if case.edit_mode == "replace_text":
        old_text = _replace_target_text(case.name, case.tokens_by_paragraph[0])
        include_header = case.container in {"header", "footer"}
        return document.replace_text(old_text, case.edit_text, count=1, include_header=include_header)

    if case.container == "section":
        document.set_paragraph_text(0, 0, case.edit_text)
        return 1
    if case.container == "header":
        document.headers()[0].set_text(case.edit_text)
        return 1
    if case.container == "footer":
        document.footers()[0].set_text(case.edit_text)
        return 1
    if case.container == "table_cell":
        document.tables()[0].cell(0, 0).set_text(case.edit_text)
        return 1
    if case.container == "shape":
        document.shapes()[0].set_text(case.edit_text)
        return 1
    if case.container in {"footnote", "endnote"}:
        document.notes()[0].set_text(case.edit_text)
        return 1
    raise ValueError(f"Unsupported container: {case.container}")


def run_stability_case(case: StabilityCase, output_dir: Path) -> StabilityCaseResult:
    output_dir.mkdir(parents=True, exist_ok=True)
    document, baseline_signatures, signature_reader = _build_case_document(case)
    outputs: list[str] = []

    baseline_path = output_dir / f"{case.name}_baseline.hwpx"
    document.save(baseline_path, validate=True)
    outputs.append(str(baseline_path))
    reopened_baseline = HwpxDocument.open(baseline_path)
    baseline_errors = reopened_baseline.validation_errors()

    replacements = _apply_case_edit(document, case)
    control_errors = document.control_preservation_validation_errors()

    edited_path = output_dir / f"{case.name}_edited.hwpx"
    edited_errors: list[ValidationIssue] = []
    reopened_errors: list[ValidationIssue] = []
    signature_mismatch: list[ValidationIssue] = []

    if case.edit_mode == "replace_text" and replacements != 1:
        edited_errors.append(
            ValidationIssue(
                kind="stability_lab",
                message=f"replace_text expected 1 replacement but got {replacements}",
                context=case.name,
            )
        )

    try:
        document.save(edited_path, validate=True)
        outputs.append(str(edited_path))
    except HwpxValidationError as exc:
        edited_errors.extend(exc.errors)
        return StabilityCaseResult(
            name=case.name,
            container=case.container,
            description=case.description,
            baseline_errors=baseline_errors,
            edited_errors=edited_errors,
            reopened_errors=reopened_errors,
            control_errors=control_errors,
            signature_mismatch=signature_mismatch,
            outputs=outputs,
        )
    except Exception as exc:  # noqa: BLE001
        edited_errors.append(ValidationIssue(kind="stability_lab", message=str(exc), context=case.name))
        return StabilityCaseResult(
            name=case.name,
            container=case.container,
            description=case.description,
            baseline_errors=baseline_errors,
            edited_errors=edited_errors,
            reopened_errors=reopened_errors,
            control_errors=control_errors,
            signature_mismatch=signature_mismatch,
            outputs=outputs,
        )

    reopened = HwpxDocument.open(edited_path)
    reopened_errors = reopened.validation_errors()
    reopened_signatures = signature_reader(reopened)
    if reopened_signatures != baseline_signatures:
        signature_mismatch.append(
            ValidationIssue(
                kind="stability_lab",
                message=f"expected signatures {baseline_signatures!r}, reopened signatures {reopened_signatures!r}",
                context=case.name,
            )
        )

    return StabilityCaseResult(
        name=case.name,
        container=case.container,
        description=case.description,
        baseline_errors=baseline_errors,
        edited_errors=edited_errors,
        reopened_errors=reopened_errors,
        control_errors=control_errors,
        signature_mismatch=signature_mismatch,
        outputs=outputs,
    )


def run_stability_matrix(output_dir: Path) -> list[StabilityCaseResult]:
    return [run_stability_case(case, output_dir) for case in stability_cases()]


def write_stability_report(results: list[StabilityCaseResult], output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps([asdict(result) for result in results], indent=2), encoding="utf-8")
    return output_path
