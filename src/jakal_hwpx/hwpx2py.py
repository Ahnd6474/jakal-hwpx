from __future__ import annotations

import argparse
import base64
from dataclasses import fields, is_dataclass
from pathlib import Path
from typing import Literal, Sequence

from .document import HwpxDocument
from .exceptions import HwpxValidationError
from .hancom_document import (
    AutoNumber,
    Bookmark,
    BulletDefinition,
    Chart,
    CharacterStyle,
    Equation,
    Field,
    Form,
    HancomDocument,
    HeaderFooter,
    Hyperlink,
    Memo,
    MemoShapeDefinition,
    Note,
    NumberingDefinition,
    Ole,
    Paragraph,
    ParagraphStyle,
    Picture,
    SectionSettings,
    Shape,
    StyleDefinition,
    Table,
)


_INDENT = "    "
_BASE64_CHUNK_SIZE = 76
HancomScriptMode = Literal["semantic", "authoring", "macro"]
Hwpx2PyMode = Literal["semantic", "authoring", "macro", "raw"]


def _py(value: object) -> str:
    return repr(value)


def _kw(name: str, value: object, *, default: object = None) -> str | None:
    if value == default:
        return None
    return f"{name}={_py(value)}"


def _kwargs(items: Sequence[str | None]) -> str:
    return ", ".join(item for item in items if item is not None)


def _line(lines: list[str], text: str = "", *, indent: int = 0) -> None:
    lines.append(f"{_INDENT * indent}{text}")


def _base64_chunks(data: bytes) -> list[str]:
    encoded = base64.b64encode(data).decode("ascii")
    return [encoded[index : index + _BASE64_CHUNK_SIZE] for index in range(0, len(encoded), _BASE64_CHUNK_SIZE)]


def _emit_bytes_data(lines: list[str], data: bytes, *, name: str, indent: int = 1) -> None:
    chunks = _base64_chunks(data)
    if not chunks:
        _line(lines, f"{name} = b''", indent=indent)
        return
    _line(lines, f"{name} = _decode_asset(", indent=indent)
    for chunk in chunks:
        _line(lines, _py(chunk), indent=indent + 1)
    _line(lines, ")", indent=indent)


def _emit_asset_data(lines: list[str], data: bytes, *, name: str = "asset_data", indent: int = 1) -> None:
    _emit_bytes_data(lines, data, name=name, indent=indent)


def _normalize_hancom_mode(mode: str) -> HancomScriptMode:
    normalized = mode.strip().lower().replace("_", "-")
    if normalized in {"semantic", "api", "public-api"}:
        return "semantic"
    if normalized in {"authoring", "dsl"}:
        return "authoring"
    if normalized in {"macro", "latex", "latex-like"}:
        return "macro"
    raise ValueError("mode must be 'semantic', 'authoring', or 'macro'.")


def _normalize_mode(mode: str) -> Hwpx2PyMode:
    normalized = mode.strip().lower().replace("_", "-")
    if normalized in {"raw", "high-fidelity", "package"}:
        return "raw"
    try:
        return _normalize_hancom_mode(normalized)
    except ValueError:
        raise ValueError("mode must be 'semantic', 'authoring', 'macro', or 'raw'.") from None


def _raw_py_str(value: str) -> str:
    if value.endswith("\\") or "\n" in value or "\r" in value or any(ord(character) < 32 for character in value):
        return _py(value)
    if "'" not in value:
        return f"r'{value}'"
    if '"' not in value:
        return f'r"{value}"'
    return _py(value)


def _assign_if_value(
    lines: list[str],
    target: str,
    name: str,
    value: object,
    *,
    default: object = None,
    indent: int = 1,
) -> None:
    if value != default:
        _line(lines, f"{target}.{name} = {_py(value)}", indent=indent)


def _emit_non_default_fields(
    lines: list[str],
    target: str,
    value: object,
    defaults: dict[str, object],
    *,
    skip: set[str] | None = None,
    indent: int = 1,
) -> None:
    if not is_dataclass(value):
        return
    skip_names = skip or set()
    for field_info in fields(value):
        name = field_info.name
        if name in skip_names:
            continue
        current = getattr(value, name)
        if isinstance(current, bytes):
            continue
        default = defaults.get(name)
        _assign_if_value(lines, target, name, current, default=default, indent=indent)


def _emit_source_paragraph_index(lines: list[str], block: object, *, target: str = "block") -> None:
    value = getattr(block, "source_paragraph_index", None)
    if value is not None:
        _line(lines, f"{target}.source_paragraph_index = {_py(value)}", indent=1)
    order = getattr(block, "source_order", None)
    if order is not None:
        _line(lines, f"{target}.source_order = {_py(order)}", indent=1)
    if getattr(block, "source_text_owned_by_container", False):
        _line(lines, f"{target}.source_text_owned_by_container = True", indent=1)


def _emit_nested_block_position(lines: list[str], block: object, *, target: str = "nested") -> None:
    paragraph_index = getattr(block, "nested_paragraph_index", None)
    if paragraph_index is not None:
        _line(lines, f"{target}.nested_paragraph_index = {_py(paragraph_index)}", indent=1)
    source_order = getattr(block, "nested_source_order", None)
    if source_order is not None:
        _line(lines, f"{target}.nested_source_order = {_py(source_order)}", indent=1)


def _emit_nested_blocks(lines: list[str], block: object, *, target: str = "block") -> None:
    nested_blocks = [
        nested
        for nested in getattr(block, "nested_blocks", [])
        if isinstance(nested, (Paragraph, Bookmark, Hyperlink, Field, AutoNumber))
    ]
    if not nested_blocks:
        return
    _line(lines, f"{target}.nested_blocks = []", indent=1)
    for nested in nested_blocks:
        _emit_nested_block(lines, nested, owner=target)


def _emit_nested_block(lines: list[str], block: object, *, owner: str) -> None:
    if isinstance(block, Paragraph):
        kwargs = _kwargs(
            [
                _kw("style_id", block.style_id),
                _kw("para_pr_id", block.para_pr_id),
                _kw("char_pr_id", block.char_pr_id),
            ]
        )
        suffix = f", {kwargs}" if kwargs else ""
        _line(lines, f"nested = Paragraph({_py(block.text)}{suffix})", indent=1)
    elif isinstance(block, Bookmark):
        _line(lines, f"nested = Bookmark({_py(block.name)})", indent=1)
    elif isinstance(block, Hyperlink):
        kwargs = _kwargs(
            [
                _kw("display_text", block.display_text),
                _kw("metadata_fields", block.metadata_fields, default=[]),
            ]
        )
        suffix = f", {kwargs}" if kwargs else ""
        _line(lines, f"nested = Hyperlink({_py(block.target)}{suffix})", indent=1)
    elif isinstance(block, Field):
        kwargs = _kwargs(
            [
                _kw("display_text", block.display_text),
                _kw("name", block.name),
                _kw("parameters", block.parameters, default={}),
                _kw("editable", block.editable, default=False),
                _kw("dirty", block.dirty, default=False),
            ]
        )
        suffix = f", {kwargs}" if kwargs else ""
        _line(lines, f"nested = Field({_py(block.field_type)}{suffix})", indent=1)
        _assign_if_value(lines, "nested", "native_field_type", block.native_field_type, default=None)
    elif isinstance(block, AutoNumber):
        kwargs = _kwargs(
            [
                _kw("kind", block.kind, default="newNum"),
                _kw("number", block.number, default=1),
                _kw("number_type", block.number_type, default="PAGE"),
            ]
        )
        _line(lines, f"nested = AutoNumber({kwargs})", indent=1)
    else:
        return
    _emit_nested_block_position(lines, block)
    _line(lines, f"{owner}.nested_blocks.append(nested)", indent=1)


def _emit_metadata(lines: list[str], document: HancomDocument) -> None:
    _line(lines, "# Metadata", indent=1)
    metadata_defaults = {
        "title": None,
        "language": None,
        "creator": None,
        "subject": None,
        "description": None,
        "lastsaveby": None,
        "created": None,
        "modified": None,
        "date": None,
        "keyword": None,
        "extra": {},
    }
    wrote = False
    for name, default in metadata_defaults.items():
        value = getattr(document.metadata, name)
        if value != default:
            _line(lines, f"doc.metadata.{name} = {_py(value)}", indent=1)
            wrote = True
    if not wrote:
        _line(lines, "pass", indent=1)
    _line(lines)


def _emit_style_definitions(lines: list[str], document: HancomDocument) -> None:
    if not (
        document.style_definitions
        or document.paragraph_styles
        or document.character_styles
        or document.numbering_definitions
        or document.bullet_definitions
        or document.memo_shape_definitions
    ):
        return

    _line(lines, "# Styles and reusable definitions", indent=1)
    for style in document.style_definitions:
        kwargs = _kwargs(
            [
                _kw("style_id", style.style_id),
                _kw("english_name", style.english_name),
                _kw("style_type", style.style_type, default="PARA"),
                _kw("para_pr_id", style.para_pr_id),
                _kw("char_pr_id", style.char_pr_id),
                _kw("next_style_id", style.next_style_id),
                _kw("lang_id", style.lang_id, default="1042"),
                _kw("lock_form", style.lock_form, default=False),
            ]
        )
        suffix = f", {kwargs}" if kwargs else ""
        _line(lines, f"doc.append_style({_py(style.name)}{suffix})", indent=1)
        if style.native_hwp_payload:
            _line(lines, "# native_hwp_payload for a style was omitted.", indent=1)

    for style in document.paragraph_styles:
        kwargs = _kwargs(
            [
                _kw("style_id", style.style_id),
                _kw("alignment_horizontal", style.alignment_horizontal),
                _kw("alignment_vertical", style.alignment_vertical),
                _kw("line_spacing", style.line_spacing),
                _kw("line_spacing_type", style.line_spacing_type),
                _kw("margins", style.margins, default={}),
                _kw("tab_pr_id", style.tab_pr_id),
                _kw("snap_to_grid", style.snap_to_grid),
                _kw("condense", style.condense),
                _kw("font_line_height", style.font_line_height),
                _kw("suppress_line_numbers", style.suppress_line_numbers),
                _kw("checked", style.checked),
                _kw("heading", style.heading, default={}),
                _kw("break_setting", style.break_setting, default={}),
                _kw("auto_spacing", style.auto_spacing, default={}),
            ]
        )
        _line(lines, f"doc.append_paragraph_style({kwargs})", indent=1)
        if style.native_hwp_payload:
            _line(lines, "# native_hwp_payload for a paragraph style was omitted.", indent=1)

    for style in document.character_styles:
        kwargs = _kwargs(
            [
                _kw("style_id", style.style_id),
                _kw("text_color", style.text_color),
                _kw("shade_color", style.shade_color),
                _kw("height", style.height),
                _kw("use_font_space", style.use_font_space),
                _kw("use_kerning", style.use_kerning),
                _kw("sym_mark", style.sym_mark),
                _kw("font_refs", style.font_refs, default={}),
            ]
        )
        _line(lines, f"doc.append_character_style({kwargs})", indent=1)
        if style.native_hwp_payload:
            _line(lines, "# native_hwp_payload for a character style was omitted.", indent=1)

    for definition in document.numbering_definitions:
        kwargs = _kwargs(
            [
                _kw("style_id", definition.style_id),
                _kw("formats", definition.formats, default=[]),
                _kw("para_heads", definition.para_heads, default=[]),
                _kw("para_head_flags", definition.para_head_flags, default=[]),
                _kw("width_adjusts", definition.width_adjusts, default=[]),
                _kw("text_offsets", definition.text_offsets, default=[]),
                _kw("char_pr_ids", definition.char_pr_ids, default=[]),
                _kw("start_numbers", definition.start_numbers, default=[]),
                _kw("unknown_short", definition.unknown_short, default=0),
                _kw("unknown_short_bits", definition.unknown_short_bits, default={}),
            ]
        )
        _line(lines, f"doc.append_numbering_definition({kwargs})", indent=1)
        if definition.native_hwp_payload:
            _line(lines, "# native_hwp_payload for a numbering definition was omitted.", indent=1)

    for definition in document.bullet_definitions:
        kwargs = _kwargs(
            [
                _kw("style_id", definition.style_id),
                _kw("flags", definition.flags),
                _kw("flag_bits", definition.flag_bits, default={}),
                _kw("bullet_char", definition.bullet_char, default="\uf0a1"),
                _kw("width_adjust", definition.width_adjust, default=0),
                _kw("text_offset", definition.text_offset, default=50),
                _kw("char_pr_id", definition.char_pr_id),
            ]
        )
        _line(lines, f"doc.append_bullet_definition({kwargs})", indent=1)
        if definition.unknown_tail or definition.native_hwp_payload:
            _line(lines, "# unknown/native payload for a bullet definition was omitted.", indent=1)

    for definition in document.memo_shape_definitions:
        kwargs = _kwargs(
            [
                _kw("memo_shape_id", definition.memo_shape_id),
                _kw("width", definition.width),
                _kw("line_width", definition.line_width),
                _kw("line_type", definition.line_type),
                _kw("line_color", definition.line_color),
                _kw("fill_color", definition.fill_color),
                _kw("active_color", definition.active_color),
                _kw("memo_type", definition.memo_type),
            ]
        )
        _line(lines, f"doc.append_memo_shape_definition({kwargs})", indent=1)
        if definition.native_hwp_payload:
            _line(lines, "# native_hwp_payload for a memo shape definition was omitted.", indent=1)
    _line(lines)


def _document_has_binary_assets(document: HancomDocument) -> bool:
    for section in document.sections:
        for block in section.blocks:
            if isinstance(block, (Picture, Ole)):
                return True
    return False


def _block_display_text(block: object) -> str | None:
    if isinstance(block, Hyperlink):
        return block.display_text
    if isinstance(block, Field):
        return block.display_text
    return None


def _emit_sections(lines: list[str], document: HancomDocument, *, include_binary_assets: bool) -> None:
    section_count = max(len(document.sections), 1)
    _line(lines, "# Sections", indent=1)
    _line(lines, f"while len(doc.sections) < {section_count}:", indent=1)
    _line(lines, "doc.append_section()", indent=2)
    _line(lines)

    for section_index, section in enumerate(document.sections):
        _line(lines, f"section = doc.sections[{section_index}]", indent=1)
        _emit_section_settings(lines, section.settings)
        for header_footer in section.header_footer_blocks:
            _emit_header_footer(lines, header_footer, section_index)
        for block_index, block in enumerate(section.blocks):
            next_block = section.blocks[block_index + 1] if block_index + 1 < len(section.blocks) else None
            if (
                isinstance(block, Paragraph)
                and block.text
                and block.text == _block_display_text(next_block)
                and getattr(block, "source_paragraph_index", None)
                == getattr(next_block, "source_paragraph_index", None)
            ):
                _line(lines, f"# Paragraph text is owned by the following inline control: {_py(block.text)}", indent=1)
                continue
            _emit_block(lines, block, section_index, include_binary_assets=include_binary_assets)
        _line(lines)


def _emit_section_settings(lines: list[str], settings: SectionSettings) -> None:
    defaults = {
        "page_width": None,
        "page_height": None,
        "landscape": None,
        "margins": {},
        "page_border_fills": [],
        "visibility": {},
        "grid": {},
        "start_numbers": {},
        "page_numbers": [],
        "page_number_positions": [],
        "footnote_pr": {},
        "endnote_pr": {},
        "line_number_shape": {},
        "numbering_shape_id": None,
        "memo_shape_id": None,
    }
    _line(lines, "settings = section.settings", indent=1)
    _emit_non_default_fields(lines, "settings", settings, defaults)


def _emit_header_footer(lines: list[str], block: HeaderFooter, section_index: int) -> None:
    method = "append_header" if block.kind == "header" else "append_footer"
    kwargs = _kwargs([_kw("apply_page_type", block.apply_page_type, default="BOTH"), f"section_index={section_index}"])
    _line(lines, f"block = doc.{method}({_py(block.text)}, {kwargs})", indent=1)
    _emit_source_paragraph_index(lines, block)
    _emit_nested_blocks(lines, block)


def _emit_block(lines: list[str], block: object, section_index: int, *, include_binary_assets: bool) -> None:
    if isinstance(block, Paragraph):
        _line(lines, f"block = doc.append_paragraph({_py(block.text)}, section_index={section_index})", indent=1)
        _emit_non_default_fields(
            lines,
            "block",
            block,
            {
                "text": block.text,
                "style_id": None,
                "para_pr_id": None,
                "char_pr_id": None,
                "hwp_para_shape_id": None,
                "hwp_style_id": None,
                "hwp_split_flags": None,
                "hwp_control_mask": None,
            },
            skip={"text"},
        )
        _emit_source_paragraph_index(lines, block)
        return

    if isinstance(block, Table):
        kwargs = _kwargs(
            [
                f"rows={block.rows}",
                f"cols={block.cols}",
                _kw("cell_texts", block.cell_texts, default=[]),
                _kw("row_heights", block.row_heights),
                _kw("col_widths", block.col_widths),
                _kw("cell_spans", block.cell_spans),
                _kw("cell_border_fill_ids", block.cell_border_fill_ids),
                _kw("table_border_fill_id", block.table_border_fill_id, default=1),
                f"section_index={section_index}",
            ]
        )
        _line(lines, f"block = doc.append_table({kwargs})", indent=1)
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
            skip={
                "rows",
                "cols",
                "cell_texts",
                "row_heights",
                "col_widths",
                "cell_spans",
                "cell_border_fill_ids",
                "table_border_fill_id",
            },
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
                _kw("extension", block.extension),
                _kw("width", block.width, default=7200),
                _kw("height", block.height, default=7200),
                _kw("original_width", block.original_width),
                _kw("original_height", block.original_height),
                _kw("current_width", block.current_width),
                _kw("current_height", block.current_height),
                f"section_index={section_index}",
            ]
        )
        _line(lines, f"block = doc.append_picture({_py(block.name)}, asset_data, {kwargs})", indent=1)
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
        kwargs = _kwargs(
            [
                _kw("display_text", block.display_text),
                _kw("metadata_fields", block.metadata_fields, default=[]),
                f"section_index={section_index}",
            ]
        )
        _line(lines, f"block = doc.append_hyperlink({_py(block.target)}, {kwargs})", indent=1)
        _emit_source_paragraph_index(lines, block)
        return

    if isinstance(block, Bookmark):
        _line(lines, f"block = doc.append_bookmark({_py(block.name)}, section_index={section_index})", indent=1)
        _emit_source_paragraph_index(lines, block)
        return

    if isinstance(block, Field):
        kwargs = _kwargs(
            [
                f"field_type={_py(block.field_type)}",
                _kw("display_text", block.display_text),
                _kw("name", block.name),
                _kw("parameters", block.parameters, default={}),
                _kw("editable", block.editable, default=False),
                _kw("dirty", block.dirty, default=False),
                f"section_index={section_index}",
            ]
        )
        _line(lines, f"block = doc.append_field({kwargs})", indent=1)
        _emit_source_paragraph_index(lines, block)
        _assign_if_value(lines, "block", "native_field_type", block.native_field_type, default=None)
        return

    if isinstance(block, AutoNumber):
        kwargs = _kwargs(
            [
                _kw("number", block.number, default=1),
                _kw("number_type", block.number_type, default="PAGE"),
                _kw("kind", block.kind, default="newNum"),
                f"section_index={section_index}",
            ]
        )
        _line(lines, f"block = doc.append_auto_number({kwargs})", indent=1)
        _emit_source_paragraph_index(lines, block)
        return

    if isinstance(block, Note):
        kwargs = _kwargs(
            [
                _kw("kind", block.kind, default="footNote"),
                _kw("number", block.number),
                f"section_index={section_index}",
            ]
        )
        _line(lines, f"block = doc.append_note({_py(block.text)}, {kwargs})", indent=1)
        _emit_source_paragraph_index(lines, block)
        _emit_nested_blocks(lines, block)
        return

    if isinstance(block, Equation):
        kwargs = _kwargs(
            [
                _kw("width", block.width, default=4800),
                _kw("height", block.height, default=2300),
                _kw("shape_comment", block.shape_comment),
                _kw("text_color", block.text_color, default="#000000"),
                _kw("base_unit", block.base_unit, default=1100),
                _kw("font", block.font, default="HYhwpEQ"),
                f"section_index={section_index}",
            ]
        )
        _line(lines, f"block = doc.append_equation({_py(block.script)}, {kwargs})", indent=1)
        _emit_source_paragraph_index(lines, block)
        _assign_if_value(lines, "block", "layout", block.layout, default={})
        _assign_if_value(lines, "block", "out_margins", block.out_margins, default={})
        _assign_if_value(lines, "block", "rotation", block.rotation, default={})
        return

    if isinstance(block, Shape):
        kwargs = _kwargs(
            [
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
                f"section_index={section_index}",
            ]
        )
        _line(lines, f"block = doc.append_shape({kwargs})", indent=1)
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
                f"section_index={section_index}",
            ]
        )
        _line(lines, f"block = doc.append_ole({_py(block.name)}, asset_data, {kwargs})", indent=1)
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
                _kw("form_type", block.form_type, default="INPUT"),
                _kw("name", block.name),
                _kw("value", block.value),
                _kw("checked", block.checked, default=False),
                _kw("items", block.items, default=[]),
                _kw("editable", block.editable, default=True),
                _kw("locked", block.locked, default=False),
                _kw("placeholder", block.placeholder),
                f"section_index={section_index}",
            ]
        )
        _line(lines, f"block = doc.append_form({_py(block.label)}, {kwargs})", indent=1)
        _emit_source_paragraph_index(lines, block)
        return

    if isinstance(block, Memo):
        kwargs = _kwargs(
            [
                _kw("author", block.author),
                _kw("memo_id", block.memo_id),
                _kw("anchor_id", block.anchor_id),
                _kw("order", block.order),
                _kw("visible", block.visible, default=True),
                f"section_index={section_index}",
            ]
        )
        _line(lines, f"block = doc.append_memo({_py(block.text)}, {kwargs})", indent=1)
        _emit_source_paragraph_index(lines, block)
        return

    if isinstance(block, Chart):
        kwargs = _kwargs(
            [
                _kw("chart_type", block.chart_type, default="BAR"),
                _kw("categories", block.categories, default=[]),
                _kw("series", block.series, default=[]),
                _kw("data_ref", block.data_ref),
                _kw("legend_visible", block.legend_visible, default=True),
                _kw("width", block.width, default=12000),
                _kw("height", block.height, default=3200),
                _kw("shape_comment", block.shape_comment),
                f"section_index={section_index}",
            ]
        )
        _line(lines, f"block = doc.append_chart({_py(block.title)}, {kwargs})", indent=1)
        _emit_source_paragraph_index(lines, block)
        _assign_if_value(lines, "block", "layout", block.layout, default={})
        _assign_if_value(lines, "block", "out_margins", block.out_margins, default={})
        _assign_if_value(lines, "block", "rotation", block.rotation, default={})
        return

    _line(lines, f"# Skipped unsupported block: {type(block).__name__}", indent=1)


def _checked_hwpx_document(
    source: Path,
    *,
    validate_input: bool,
    strict: bool,
) -> HwpxDocument:
    hwpx_document = HwpxDocument.open(source)
    if validate_input:
        hwpx_document.validate()
    if strict:
        errors = hwpx_document.strict_lint_errors()
        if errors:
            raise HwpxValidationError(errors)
    return hwpx_document


def _generate_raw_hwpx_script(
    source: Path,
    hwpx_document: HwpxDocument,
    *,
    default_output_name: str | None,
) -> str:
    output_name = default_output_name or f"{source.stem}.recreated.hwpx"
    lines: list[str] = [
        "# Generated by jakal_hwpx.hwpx2py.",
        "# This script recreates the source HWPX from embedded package parts.",
        "from __future__ import annotations",
        "",
        "import argparse",
        "import base64",
        "from pathlib import Path",
        "",
        "from jakal_hwpx import HwpxDocument",
        "",
        f"_DEFAULT_OUTPUT = {_py(output_name)}",
        "",
        "",
        "def _decode_asset(value: str) -> bytes:",
        "    return base64.b64decode(value.encode(\"ascii\"))",
        "",
        "",
        "def build_document() -> HwpxDocument:",
        "    doc = HwpxDocument.blank()",
        "    for part_path in list(doc.list_part_paths()):",
        "        doc.remove_part(part_path)",
        "",
        "    parts: list[tuple[str, bytes]] = []",
    ]
    for part_path in hwpx_document.list_part_paths():
        part = hwpx_document.get_part(part_path)
        _emit_bytes_data(lines, part.raw_bytes, name="part_data", indent=1)
        _line(lines, f"parts.append(({_py(part_path)}, part_data))", indent=1)
        _line(lines)

    lines.extend(
        [
            "    for part_path, part_data in parts:",
            "        doc.add_part(part_path, part_data)",
            "    return doc",
            "",
            "",
            "def write_hwpx(path: str | Path = _DEFAULT_OUTPUT, *, validate: bool = True) -> Path:",
            "    return build_document().save(path, validate=validate)",
            "",
            "",
            "def main(argv: list[str] | None = None) -> int:",
            "    parser = argparse.ArgumentParser(description=\"Recreate the generated HWPX document.\")",
            "    parser.add_argument(\"output\", nargs=\"?\", default=_DEFAULT_OUTPUT, help=\"output .hwpx path\")",
            "    parser.add_argument(\"--no-validate\", action=\"store_true\", help=\"skip HWPX validation while saving\")",
            "    args = parser.parse_args(argv)",
            "    write_hwpx(args.output, validate=not args.no_validate)",
            "    return 0",
            "",
            "",
            "if __name__ == \"__main__\":",
            "    raise SystemExit(main())",
            "",
        ]
    )
    return "\n".join(lines)


def _generate_semantic_hancom_script(
    document: HancomDocument,
    *,
    source: Path,
    default_output_name: str | None,
    include_binary_assets: bool,
    generator_module: str,
    source_label: str,
    output_format: Literal["hwpx", "hwp"] = "hwpx",
) -> str:
    output_name = default_output_name or f"{source.stem}.recreated.{output_format}"
    uses_binary_assets = include_binary_assets and _document_has_binary_assets(document)
    lines: list[str] = [
        f"# Generated by {generator_module}.",
        f"# This script recreates the source {source_label} from public document APIs.",
        "from __future__ import annotations",
        "",
        "import argparse",
    ]
    if uses_binary_assets:
        lines.append("import base64")
    lines.extend(
        [
            "from pathlib import Path",
            "",
            "from jakal_hwpx import HancomDocument",
            "from jakal_hwpx.hancom_document import AutoNumber, Bookmark, Field, Hyperlink, Paragraph",
            "",
            f"_DEFAULT_OUTPUT = {_py(output_name)}",
            "",
        ]
    )
    if uses_binary_assets:
        lines.extend(
            [
                "",
                "def _decode_asset(value: str) -> bytes:",
                "    return base64.b64decode(value.encode(\"ascii\"))",
                "",
            ]
        )
    lines.extend(
        [
            "",
            "def build_document() -> HancomDocument:",
            "    doc = HancomDocument.blank()",
            "",
        ]
    )
    _emit_metadata(lines, document)
    _emit_style_definitions(lines, document)
    _emit_sections(lines, document, include_binary_assets=include_binary_assets)
    _line(lines, "return doc", indent=1)
    if output_format == "hwp":
        lines.extend(
            [
                "",
                "",
                "def write_hwp(path: str | Path = _DEFAULT_OUTPUT) -> Path:",
                "    return build_document().write_to_hwp(path)",
                "",
                "",
                "def write_hwpx(path: str | Path, *, validate: bool = True) -> Path:",
                "    return build_document().write_to_hwpx(path, validate=validate)",
                "",
                "",
                "def main(argv: list[str] | None = None) -> int:",
                "    parser = argparse.ArgumentParser(description=\"Recreate the generated HWP document.\")",
                "    parser.add_argument(\"output\", nargs=\"?\", default=_DEFAULT_OUTPUT, help=\"output .hwp path\")",
                "    args = parser.parse_args(argv)",
                "    write_hwp(args.output)",
                "    return 0",
                "",
                "",
                "if __name__ == \"__main__\":",
                "    raise SystemExit(main())",
                "",
            ]
        )
        return "\n".join(lines)

    lines.extend(
        [
            "",
            "",
            "def write_hwpx(path: str | Path = _DEFAULT_OUTPUT, *, validate: bool = True) -> Path:",
            "    return build_document().write_to_hwpx(path, validate=validate)",
            "",
            "",
            "def main(argv: list[str] | None = None) -> int:",
            "    parser = argparse.ArgumentParser(description=\"Recreate the generated HWPX document.\")",
            "    parser.add_argument(\"output\", nargs=\"?\", default=_DEFAULT_OUTPUT, help=\"output .hwpx path\")",
            "    parser.add_argument(\"--no-validate\", action=\"store_true\", help=\"skip HWPX validation while saving\")",
            "    args = parser.parse_args(argv)",
            "    write_hwpx(args.output, validate=not args.no_validate)",
            "    return 0",
            "",
            "",
            "if __name__ == \"__main__\":",
            "    raise SystemExit(main())",
            "",
        ]
    )
    return "\n".join(lines)


def _emit_authoring_helpers(lines: list[str]) -> None:
    lines.extend(
        [
            "",
            "def hdoc() -> HancomDocument:",
            "    return HancomDocument.blank()",
            "",
            "",
            "def _set_attrs(block, **values):",
            "    for name, value in values.items():",
            "        if value is not None:",
            "            setattr(block, name, value)",
            "    return block",
            "",
            "",
            "def control(doc: HancomDocument, name: str, *args, section: int = 0, **options):",
            "    kind = name.strip().lower().replace('_', '-').replace(' ', '-')",
            "    if kind in {'section', 'sec'}:",
            "        index = int(options.pop('index', section))",
            "        while len(doc.sections) <= index:",
            "            doc.append_section()",
            "        target = doc.sections[index]",
            "        for attr_name, value in options.items():",
            "            if value not in (None, {}, []):",
            "                setattr(target.settings, attr_name, value)",
            "        return target",
            "    if kind in {'header', 'head'}:",
            "        text = args[0] if args else options.pop('text', '')",
            "        return doc.append_header(text, section_index=section, **options)",
            "    if kind in {'footer', 'foot'}:",
            "        text = args[0] if args else options.pop('text', '')",
            "        return doc.append_footer(text, section_index=section, **options)",
            "    if kind in {'p', 'paragraph', 'text'}:",
            "        text = args[0] if args else options.pop('text', '')",
            "        style = options.pop('style', None)",
            "        para = options.pop('para', None)",
            "        char = options.pop('char', None)",
            "        block = doc.append_paragraph(text, section_index=section)",
            "        _set_attrs(block, style_id=style, para_pr_id=para, char_pr_id=char, **options)",
            "        return block",
            "    if kind in {'eq', 'equation', 'math'}:",
            "        script = args[0] if args else options.pop('script')",
            "        return doc.append_equation(script, section_index=section, **options)",
            "    if kind in {'table', 'tabular'}:",
            "        cells = args[0] if args else options.pop('cells')",
            "        rows = options.pop('rows', None)",
            "        cols = options.pop('cols', None)",
            "        row_count = rows if rows is not None else len(cells)",
            "        col_count = cols if cols is not None else max((len(row) for row in cells), default=0)",
            "        return doc.append_table(rows=row_count, cols=col_count, cell_texts=cells, section_index=section, **options)",
            "    if kind in {'image', 'picture', 'figure'}:",
            "        name_value = args[0] if args else options.pop('name')",
            "        data_value = args[1] if len(args) > 1 else options.pop('data')",
            "        return doc.append_picture(name_value, data_value, section_index=section, **options)",
            "    if kind == 'ole':",
            "        name_value = args[0] if args else options.pop('name')",
            "        data_value = args[1] if len(args) > 1 else options.pop('data')",
            "        return doc.append_ole(name_value, data_value, section_index=section, **options)",
            "    if kind in {'shape', 'draw'}:",
            "        return doc.append_shape(section_index=section, **options)",
            "    if kind in {'link', 'hyperlink', 'href'}:",
            "        target = args[0] if args else options.pop('target')",
            "        return doc.append_hyperlink(target, section_index=section, **options)",
            "    if kind in {'bookmark', 'label'}:",
            "        value = args[0] if args else options.pop('name')",
            "        return doc.append_bookmark(value, section_index=section)",
            "    if kind in {'field', 'ref'}:",
            "        field_type = args[0] if args else options.pop('field_type')",
            "        return doc.append_field(field_type=field_type, section_index=section, **options)",
            "    if kind in {'autonum', 'auto-number'}:",
            "        return doc.append_auto_number(section_index=section, **options)",
            "    if kind in {'note', 'footnote', 'endnote'}:",
            "        text = args[0] if args else options.pop('text', '')",
            "        if kind == 'footnote':",
            "            options.setdefault('kind', 'footNote')",
            "        elif kind == 'endnote':",
            "            options.setdefault('kind', 'endNote')",
            "        return doc.append_note(text, section_index=section, **options)",
            "    if kind == 'form':",
            "        label = args[0] if args else options.pop('label', '')",
            "        return doc.append_form(label, section_index=section, **options)",
            "    if kind == 'memo':",
            "        text = args[0] if args else options.pop('text', '')",
            "        return doc.append_memo(text, section_index=section, **options)",
            "    if kind == 'chart':",
            "        title = args[0] if args else options.pop('title', '')",
            "        return doc.append_chart(title, section_index=section, **options)",
            "    raise ValueError(f'Unsupported control: {name}')",
            "",
            "",
            "def section(doc: HancomDocument, index: int = 0, **settings):",
            "    return control(doc, 'section', index=index, **settings)",
            "",
            "",
            "def header(doc: HancomDocument, text: str, *, section: int = 0, apply_page_type: str = 'BOTH'):",
            "    return control(doc, 'header', text, section=section, apply_page_type=apply_page_type)",
            "",
            "",
            "def footer(doc: HancomDocument, text: str, *, section: int = 0, apply_page_type: str = 'BOTH'):",
            "    return control(doc, 'footer', text, section=section, apply_page_type=apply_page_type)",
            "",
            "",
            "def p(",
            "    doc: HancomDocument,",
            "    text: str,",
            "    *,",
            "    section: int = 0,",
            "    style: str | None = None,",
            "    para: str | None = None,",
            "    char: str | None = None,",
            "):",
            "    return control(doc, 'paragraph', text, section=section, style=style, para=para, char=char)",
            "",
            "",
            "def eq(doc: HancomDocument, script: str, *, section: int = 0, **options):",
            "    return control(doc, 'equation', script, section=section, **options)",
            "",
            "",
            "def table(doc: HancomDocument, cells: list[list[str]], *, section: int = 0, rows: int | None = None, cols: int | None = None, **options):",
            "    return control(doc, 'table', cells, section=section, rows=rows, cols=cols, **options)",
            "",
            "",
            "def image(doc: HancomDocument, name: str, data: bytes, *, section: int = 0, **options):",
            "    return control(doc, 'image', name, data, section=section, **options)",
            "",
            "",
            "def ole(doc: HancomDocument, name: str, data: bytes, *, section: int = 0, **options):",
            "    return control(doc, 'ole', name, data, section=section, **options)",
            "",
            "",
            "def shape(doc: HancomDocument, *, section: int = 0, **options):",
            "    return control(doc, 'shape', section=section, **options)",
            "",
            "",
            "def link(doc: HancomDocument, target: str, *, section: int = 0, **options):",
            "    return control(doc, 'link', target, section=section, **options)",
            "",
            "",
            "def bookmark(doc: HancomDocument, name: str, *, section: int = 0):",
            "    return control(doc, 'bookmark', name, section=section)",
            "",
            "",
            "def field(doc: HancomDocument, field_type: str, *, section: int = 0, **options):",
            "    return control(doc, 'field', field_type, section=section, **options)",
            "",
            "",
            "def autonum(doc: HancomDocument, *, section: int = 0, **options):",
            "    return control(doc, 'autonum', section=section, **options)",
            "",
            "",
            "def note(doc: HancomDocument, text: str, *, section: int = 0, **options):",
            "    return control(doc, 'note', text, section=section, **options)",
            "",
            "",
            "def form(doc: HancomDocument, label: str = '', *, section: int = 0, **options):",
            "    return control(doc, 'form', label, section=section, **options)",
            "",
            "",
            "def memo(doc: HancomDocument, text: str, *, section: int = 0, **options):",
            "    return control(doc, 'memo', text, section=section, **options)",
            "",
            "",
            "def chart(doc: HancomDocument, title: str = '', *, section: int = 0, **options):",
            "    return control(doc, 'chart', title, section=section, **options)",
            "",
        ]
    )


def _emit_macro_helpers(lines: list[str]) -> None:
    _emit_authoring_helpers(lines)
    lines.extend(
        [
            "",
            "def item(name: str, *args, attrs: dict[str, object] | None = None, **options):",
            "    return (name, args, options, attrs or {})",
            "",
            "",
            "def text(value: str, **options):",
            "    return item('paragraph', value, **options)",
            "",
            "",
            "def math(script: str, **options):",
            "    return item('equation', script, **options)",
            "",
            "",
            "def tabular(cells: list[list[str]], **options):",
            "    return item('table', cells, **options)",
            "",
            "",
            "def draw(**options):",
            "    return item('shape', **options)",
            "",
            "",
            "def figure(name: str, data: bytes, **options):",
            "    return item('image', name, data, **options)",
            "",
            "",
            "def object_data(name: str, data: bytes, **options):",
            "    return item('ole', name, data, **options)",
            "",
            "",
            "def href(target: str, **options):",
            "    return item('link', target, **options)",
            "",
            "",
            "def label(name: str):",
            "    return item('label', name)",
            "",
            "",
            "def control_item(name: str, *args, **options):",
            "    return item(name, *args, **options)",
            "",
            "",
            "def choices(*values: str, marker_style: str = 'number'):", 
            "    cells = [[str(index + 1), str(value)] for index, value in enumerate(values)]",
            "    return item('table', cells, attrs={'macro': {'kind': 'choices', 'marker_style': marker_style}})",
            "",
            "",
            "def problem(*items, number: int | str | None = None, title: str | None = None, label_name: str | None = None):",
            "    return ('problem', items, {'number': number, 'title': title, 'label_name': label_name}, {})",
            "",
            "",
            "def emit(doc: HancomDocument, node, *, section: int = 0):",
            "    if node is None:",
            "        return None",
            "    kind, args, options, attrs = node",
            "    if kind == 'problem':",
            "        for child in args:",
            "            emit(doc, child, section=section)",
            "        return None",
            "    call_options = dict(options)",
            "    call_options.setdefault('section', section)",
            "    if kind == 'label':",
            "        block = control(doc, 'label', *args, section=call_options.pop('section'))",
            "    else:",
            "        block = control(doc, kind, *args, **call_options)",
            "    for attr_name, value in attrs.items():",
            "        setattr(block, attr_name, value)",
            "    return block",
            "",
            "",
            "def document(doc: HancomDocument, *nodes, section: int = 0):",
            "    for node in nodes:",
            "        emit(doc, node, section=section)",
            "    return doc",
            "",
        ]
    )


def _section_setting_kwargs(settings: SectionSettings, section_index: int) -> list[str | None]:
    return [
        _kw("index", section_index, default=0),
        _kw("page_width", settings.page_width),
        _kw("page_height", settings.page_height),
        _kw("landscape", settings.landscape),
        _kw("margins", settings.margins, default={}),
        _kw("page_border_fills", settings.page_border_fills, default=[]),
        _kw("visibility", settings.visibility, default={}),
        _kw("grid", settings.grid, default={}),
        _kw("start_numbers", settings.start_numbers, default={}),
        _kw("page_numbers", settings.page_numbers, default=[]),
        _kw("page_number_positions", settings.page_number_positions, default=[]),
        _kw("footnote_pr", settings.footnote_pr, default={}),
        _kw("endnote_pr", settings.endnote_pr, default={}),
        _kw("line_number_shape", settings.line_number_shape, default={}),
        _kw("numbering_shape_id", settings.numbering_shape_id),
        _kw("memo_shape_id", settings.memo_shape_id),
    ]


def _emit_authoring_section_settings(lines: list[str], settings: SectionSettings, section_index: int) -> None:
    kwargs = _kwargs(_section_setting_kwargs(settings, section_index))
    _line(lines, f"section(doc{', ' + kwargs if kwargs else ''})", indent=1)


def _emit_authoring_header_footer(lines: list[str], block: HeaderFooter, section_index: int) -> None:
    function_name = "header" if block.kind == "header" else "footer"
    kwargs = _kwargs([_kw("section", section_index, default=0), _kw("apply_page_type", block.apply_page_type, default="BOTH")])
    suffix = f", {kwargs}" if kwargs else ""
    _line(lines, f"block = {function_name}(doc, {_py(block.text)}{suffix})", indent=1)
    _emit_source_paragraph_index(lines, block)
    _emit_nested_blocks(lines, block)


def _table_size_matches_cells(block: Table) -> bool:
    return len(block.cell_texts) == block.rows and all(len(row) == block.cols for row in block.cell_texts)


def _authoring_cell_border_fill_ids(block: Table) -> dict[tuple[int, int], int] | None:
    if not block.cell_border_fill_ids:
        return None
    if all(value == 0 for value in block.cell_border_fill_ids.values()):
        return None
    return block.cell_border_fill_ids


def _macro_attrs(block: object, defaults: dict[str, object] | None = None, *, skip: set[str] | None = None) -> dict[str, object]:
    attrs: dict[str, object] = {}
    if is_dataclass(block):
        skip_names = skip or set()
        default_values = defaults or {}
        for field_info in fields(block):
            name = field_info.name
            if name in skip_names:
                continue
            value = getattr(block, name)
            if isinstance(value, bytes):
                continue
            if value != default_values.get(name):
                attrs[name] = value
    for name in ("source_paragraph_index", "source_order", "source_text_owned_by_container"):
        if hasattr(block, name):
            value = getattr(block, name)
            if value not in (None, False):
                attrs[name] = value
    return attrs


def _macro_attrs_kw(attrs: dict[str, object]) -> str | None:
    return _kw("attrs", attrs, default={})


def _macro_expr(function_name: str, args: Sequence[str], kwargs: Sequence[str | None]) -> str:
    body = _kwargs([*args, *kwargs])
    return f"{function_name}({body})"


def _macro_expr_for_block(
    lines: list[str],
    block: object,
    section_index: int,
    *,
    include_binary_assets: bool,
    asset_counter: list[int],
) -> str | None:
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


def _emit_macro_document_call(lines: list[str], expressions: list[str], section_index: int) -> None:
    if not expressions:
        return
    _line(lines, "document(", indent=1)
    _line(lines, "doc,", indent=2)
    _line(lines, "problem(", indent=2)
    for expression in expressions:
        _line(lines, f"{expression},", indent=3)
    _line(lines, "),", indent=2)
    if section_index == 0:
        _line(lines, ")", indent=1)
    else:
        _line(lines, f"section={section_index},", indent=2)
        _line(lines, ")", indent=1)


def _emit_macro_sections(lines: list[str], document: HancomDocument, *, include_binary_assets: bool) -> None:
    asset_counter = [0]
    for section_index, section_model in enumerate(document.sections):
        function_name = f"section_{section_index + 1}"
        _line(lines, f"def {function_name}(doc: HancomDocument) -> None:")
        _emit_authoring_section_settings(lines, section_model.settings, section_index)
        for header_footer in section_model.header_footer_blocks:
            _emit_authoring_header_footer(lines, header_footer, section_index)
        pending: list[str] = []
        for block_index, block in enumerate(section_model.blocks):
            next_block = section_model.blocks[block_index + 1] if block_index + 1 < len(section_model.blocks) else None
            if (
                isinstance(block, Paragraph)
                and block.text
                and block.text == _block_display_text(next_block)
                and getattr(block, "source_paragraph_index", None)
                == getattr(next_block, "source_paragraph_index", None)
            ):
                _line(lines, f"# Paragraph text is owned by the following inline control: {_py(block.text)}", indent=1)
                continue
            expression = _macro_expr_for_block(
                lines,
                block,
                section_index,
                include_binary_assets=include_binary_assets,
                asset_counter=asset_counter,
            )
            if expression is not None:
                pending.append(expression)
                continue
            _emit_macro_document_call(lines, pending, section_index)
            pending = []
            _emit_authoring_block(lines, block, section_index, include_binary_assets=include_binary_assets)
        _emit_macro_document_call(lines, pending, section_index)
        _line(lines)


def _emit_authoring_block(lines: list[str], block: object, section_index: int, *, include_binary_assets: bool) -> None:
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


def _emit_authoring_sections(lines: list[str], document: HancomDocument, *, include_binary_assets: bool) -> None:
    for section_index, section_model in enumerate(document.sections):
        function_name = f"section_{section_index + 1}"
        _line(lines, f"def {function_name}(doc: HancomDocument) -> None:")
        _emit_authoring_section_settings(lines, section_model.settings, section_index)
        for header_footer in section_model.header_footer_blocks:
            _emit_authoring_header_footer(lines, header_footer, section_index)
        for block_index, block in enumerate(section_model.blocks):
            next_block = section_model.blocks[block_index + 1] if block_index + 1 < len(section_model.blocks) else None
            if (
                isinstance(block, Paragraph)
                and block.text
                and block.text == _block_display_text(next_block)
                and getattr(block, "source_paragraph_index", None)
                == getattr(next_block, "source_paragraph_index", None)
            ):
                _line(lines, f"# Paragraph text is owned by the following inline control: {_py(block.text)}", indent=1)
                continue
            _emit_authoring_block(lines, block, section_index, include_binary_assets=include_binary_assets)
        _line(lines)


def _generate_authoring_hancom_script(
    document: HancomDocument,
    *,
    source: Path,
    default_output_name: str | None,
    include_binary_assets: bool,
    generator_module: str,
    source_label: str,
    output_format: Literal["hwpx", "hwp"] = "hwpx",
    macro: bool = False,
) -> str:
    output_name = default_output_name or f"{source.stem}.recreated.{output_format}"
    uses_binary_assets = include_binary_assets and _document_has_binary_assets(document)
    description = "a LaTeX-like macro DSL" if macro else "a compact authoring DSL"
    lines: list[str] = [
        f"# Generated by {generator_module}.",
        f"# This script recreates the source {source_label} with {description}.",
        "from __future__ import annotations",
        "",
        "import argparse",
    ]
    if uses_binary_assets:
        lines.append("import base64")
    lines.extend(
        [
            "from pathlib import Path",
            "",
            "from jakal_hwpx import HancomDocument",
            "from jakal_hwpx.hancom_document import AutoNumber, Bookmark, Field, Hyperlink, Paragraph",
            "",
            f"_DEFAULT_OUTPUT = {_py(output_name)}",
        ]
    )
    if uses_binary_assets:
        lines.extend(
            [
                "",
                "",
                "def _decode_asset(value: str) -> bytes:",
                "    return base64.b64decode(value.encode(\"ascii\"))",
            ]
        )
    if macro:
        _emit_macro_helpers(lines)
    else:
        _emit_authoring_helpers(lines)
    lines.extend(
        [
            "",
            "def configure_document(doc: HancomDocument) -> None:",
        ]
    )
    _emit_metadata(lines, document)
    _emit_style_definitions(lines, document)
    if macro:
        _emit_macro_sections(lines, document, include_binary_assets=include_binary_assets)
    else:
        _emit_authoring_sections(lines, document, include_binary_assets=include_binary_assets)
    lines.extend(
        [
            "",
            "def build_document() -> HancomDocument:",
            "    doc = hdoc()",
            "    configure_document(doc)",
        ]
    )
    for section_index in range(len(document.sections)):
        _line(lines, f"section_{section_index + 1}(doc)", indent=1)
    _line(lines, "return doc", indent=1)
    if output_format == "hwp":
        lines.extend(
            [
                "",
                "",
                "def write_hwp(path: str | Path = _DEFAULT_OUTPUT) -> Path:",
                "    return build_document().write_to_hwp(path)",
                "",
                "",
                "def write_hwpx(path: str | Path, *, validate: bool = True) -> Path:",
                "    return build_document().write_to_hwpx(path, validate=validate)",
                "",
                "",
                "def main(argv: list[str] | None = None) -> int:",
                "    parser = argparse.ArgumentParser(description=\"Recreate the generated HWP document.\")",
                "    parser.add_argument(\"output\", nargs=\"?\", default=_DEFAULT_OUTPUT, help=\"output .hwp path\")",
                "    args = parser.parse_args(argv)",
                "    write_hwp(args.output)",
                "    return 0",
                "",
                "",
                "if __name__ == \"__main__\":",
                "    raise SystemExit(main())",
                "",
            ]
        )
        return "\n".join(lines)

    lines.extend(
        [
            "",
            "",
            "def write_hwpx(path: str | Path = _DEFAULT_OUTPUT, *, validate: bool = True) -> Path:",
            "    return build_document().write_to_hwpx(path, validate=validate)",
            "",
            "",
            "def main(argv: list[str] | None = None) -> int:",
            "    parser = argparse.ArgumentParser(description=\"Recreate the generated HWPX document.\")",
            "    parser.add_argument(\"output\", nargs=\"?\", default=_DEFAULT_OUTPUT, help=\"output .hwpx path\")",
            "    parser.add_argument(\"--no-validate\", action=\"store_true\", help=\"skip HWPX validation while saving\")",
            "    args = parser.parse_args(argv)",
            "    write_hwpx(args.output, validate=not args.no_validate)",
            "    return 0",
            "",
            "",
            "if __name__ == \"__main__\":",
            "    raise SystemExit(main())",
            "",
        ]
    )
    return "\n".join(lines)


def generate_hwpx_script(
    source_path: str | Path,
    *,
    default_output_name: str | None = None,
    validate_input: bool = True,
    strict: bool = False,
    include_binary_assets: bool = True,
    mode: Hwpx2PyMode | str = "semantic",
) -> str:
    """Generate a from-scratch Python script for a HWPX document.

    The default ``semantic`` mode recreates the document through public
    ``HancomDocument`` APIs. ``authoring`` mode uses the same model but emits
    a compact helper-based script that is easier to edit by hand. ``macro``
    mode emits a LaTeX-like item tree around the same controls. Picture and OLE
    payloads are embedded as base64 only when ``include_binary_assets`` is true.

    ``raw`` mode embeds every package part and rebuilds the HWPX package
    through ``HwpxDocument``. It preserves unsupported XML/control ordering but
    requires binary assets to be included.
    """

    selected_mode = _normalize_mode(str(mode))
    if selected_mode == "raw" and not include_binary_assets:
        raise ValueError("raw mode requires include_binary_assets=True.")

    source = Path(source_path)
    hwpx_document = _checked_hwpx_document(source, validate_input=validate_input, strict=strict)
    if selected_mode == "raw":
        return _generate_raw_hwpx_script(
            source,
            hwpx_document,
            default_output_name=default_output_name,
        )

    document = HancomDocument.from_hwpx_document(hwpx_document)
    if selected_mode in {"authoring", "macro"}:
        return _generate_authoring_hancom_script(
            document,
            source=source,
            default_output_name=default_output_name,
            include_binary_assets=include_binary_assets,
            generator_module="jakal_hwpx.hwpx2py",
            source_label="HWPX",
            macro=selected_mode == "macro",
        )
    return _generate_semantic_hancom_script(
        document,
        source=source,
        default_output_name=default_output_name,
        include_binary_assets=include_binary_assets,
        generator_module="jakal_hwpx.hwpx2py",
        source_label="HWPX",
    )


def default_script_path(source_path: str | Path) -> Path:
    source = Path(source_path)
    return source.with_name(f"{source.stem}_hwpx2py.py")


def write_hwpx_script(
    source_path: str | Path,
    script_path: str | Path | None = None,
    *,
    default_output_name: str | None = None,
    validate_input: bool = True,
    strict: bool = False,
    include_binary_assets: bool = True,
    mode: Hwpx2PyMode | str = "semantic",
) -> Path:
    target = Path(script_path) if script_path is not None else default_script_path(source_path)
    target.parent.mkdir(parents=True, exist_ok=True)
    script = generate_hwpx_script(
        source_path,
        default_output_name=default_output_name,
        validate_input=validate_input,
        strict=strict,
        include_binary_assets=include_binary_assets,
        mode=mode,
    )
    target.write_text(script, encoding="utf-8", newline="\n")
    return target


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Generate a from-scratch Python script for a HWPX file.")
    parser.add_argument("input", help="source .hwpx file")
    parser.add_argument("-o", "--output", help="generated Python script path")
    parser.add_argument(
        "--default-output",
        help="default .hwpx path used by the generated script",
    )
    parser.add_argument(
        "--no-validate-input",
        action="store_true",
        help="skip validating the source HWPX before generating the script",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="run strict_lint_errors() on the source before generating the script",
    )
    parser.add_argument(
        "--mode",
        choices=("semantic", "authoring", "dsl", "macro", "latex", "raw"),
        default="semantic",
        help="generation mode: public API reconstruction, authoring DSL, macro DSL, or embedded package parts",
    )
    parser.add_argument(
        "--skip-binary-assets",
        action="store_true",
        help="leave Picture/OLE payloads out of the generated script",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_arg_parser()
    args = parser.parse_args(argv)
    if args.mode == "raw" and args.skip_binary_assets:
        parser.error("--skip-binary-assets cannot be used with --mode raw")
    write_hwpx_script(
        args.input,
        args.output,
        default_output_name=args.default_output,
        validate_input=not args.no_validate_input,
        strict=args.strict,
        include_binary_assets=not args.skip_binary_assets,
        mode=args.mode,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
