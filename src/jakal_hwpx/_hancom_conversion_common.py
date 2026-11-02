"""Shared scalar normalization and text ownership for format conversion."""
from __future__ import annotations

import json
from typing import Any

from ._hancom_models import (
    Chart,
    Field,
    Form,
    HancomBlock,
    HeaderFooter,
    Hyperlink,
    Memo,
    Ole,
    Paragraph,
    Shape,
    Table,
)


def _blank_cell_texts(rows: int, cols: int) -> list[list[str]]:
    return [["" for _ in range(cols)] for _ in range(rows)]


def _normalize_hwp_text(value: str) -> str:
    return _xml_safe_scalar(value.replace("\r", "")).strip()


def _drop_surrogate_codepoints(value: str) -> str:
    return "".join(character for character in value if not 0xD800 <= ord(character) <= 0xDFFF)


def _xml_safe_scalar(value: object | None) -> str:
    if value is None:
        return ""
    text = str(value)
    def _is_xml_char(character: str) -> bool:
        codepoint = ord(character)
        return (
            codepoint in {0x09, 0x0A, 0x0D}
            or 0x20 <= codepoint <= 0xD7FF
            or 0xE000 <= codepoint <= 0xFFFD
            or 0x10000 <= codepoint <= 0x10FFFF
        )
    return "".join(
        character
        for character in text
        if _is_xml_char(character)
    )


def _normalize_str_map(values: dict[str, object] | None) -> dict[str, str]:
    if not values:
        return {}
    return {str(key): str(value) for key, value in values.items() if value is not None}


def _normalize_int_map(values: dict[str, Any] | None) -> dict[str, int]:
    if not values:
        return {}
    result: dict[str, int] = {}
    for key, value in values.items():
        if value is None:
            continue
        try:
            result[str(key)] = int(value)
        except (TypeError, ValueError):
            continue
    return result


def _block_owned_text(block: HancomBlock | None) -> str | None:
    if block is None:
        return None
    if isinstance(block, Hyperlink):
        return block.display_text
    if isinstance(block, Field):
        return block.display_text
    if isinstance(block, Form):
        return block.label
    if isinstance(block, Memo):
        return block.text
    if isinstance(block, Chart):
        return block.title
    if isinstance(block, Shape):
        return block.text
    if isinstance(block, Ole):
        return block.name
    return None


def _block_owned_text_for_write(block: HancomBlock | HeaderFooter | None) -> str | None:
    if block is None or getattr(block, "source_text_owned_by_container", False):
        return None
    if isinstance(block, Hyperlink):
        return block.display_text
    if isinstance(block, Field):
        return block.display_text
    return None


def _is_owned_text_shadow_paragraph(block: object, next_block: object | None) -> bool:
    if not isinstance(block, Paragraph):
        return False
    if not block.text:
        return False
    owned_text = _block_owned_text_for_write(next_block)  # type: ignore[arg-type]
    if block.text != owned_text:
        return False
    return (
        getattr(block, "source_paragraph_index", None) == getattr(next_block, "source_paragraph_index", None)
        and getattr(block, "source_order", None) == getattr(next_block, "source_order", None)
    )


def _normalize_hwp_setting_scalar(value: object | None) -> str | None:
    if value is None:
        return None
    return str(value)


def _has_word_feature_carriers(sections) -> bool:
    def scan(blocks):
        for block in blocks:
            if isinstance(block, Field) and block.field_type == "JAKAL_WORD_FEATURE":
                return True
            if scan(getattr(block, "nested_blocks", [])):
                return True
            if isinstance(block, Table) and any(scan(values) for baseline, values in getattr(block, "_docx_cell_content", {}).values()):
                return True
        return False
    return any(scan([*section.header_footer_blocks, *section.blocks]) for section in sections)


def _json_list_parameter(value: str | None) -> list[str]:
    if not value:
        return []
    try:
        decoded = json.loads(value)
    except json.JSONDecodeError:
        return []
    if not isinstance(decoded, list):
        return []
    return [str(item) for item in decoded if item is not None]


def _nested_blocks(block: object) -> list[HancomBlock]:
    from ._hancom_hwpx_write import _is_supported_nested_hwpx_block

    values = getattr(block, "nested_blocks", [])
    return [value for value in values if _is_supported_nested_hwpx_block(value)]


def _has_nested_blocks(block: object) -> bool:
    return bool(_nested_blocks(block))


def _media_type_for_extension(extension: str | None) -> str | None:
    if extension is None:
        return None
    normalized = extension.lower().lstrip(".")
    if normalized in {"jpg", "jpeg"}:
        return "image/jpeg"
    if normalized == "png":
        return "image/png"
    if normalized == "gif":
        return "image/gif"
    if normalized == "bmp":
        return "image/bmp"
    return None
