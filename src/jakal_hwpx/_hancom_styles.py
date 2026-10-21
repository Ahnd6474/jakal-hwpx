from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from ._hancom_models import (
    BulletDefinition,
    CharacterStyle,
    HancomSection,
    MemoShapeDefinition,
    NumberingDefinition,
    Paragraph,
    ParagraphStyle,
    StyleDefinition,
)
from .hwp_binary import (
    TAG_BULLET,
    TAG_CHAR_SHAPE,
    TAG_MEMO_SHAPE,
    TAG_NUMBERING,
    TAG_PARA_SHAPE,
    TAG_STYLE,
    TypedRecord,
    _build_memo_shape_native_metadata_payload,
)
from .namespaces import NS

if TYPE_CHECKING:
    from .hancom_document import HancomDocument
    from .hwp_document import HwpDocument

_HWP_STYLE_TYPE_CODES = {
    "PARA": 0,
    "CHAR": 1,
}
_HWP_STYLE_TYPE_NAMES = {value: key for key, value in _HWP_STYLE_TYPE_CODES.items()}
_HWP_DOCINFO_COUNT_INDEX_BY_TAG = {
    TAG_CHAR_SHAPE: 9,
    TAG_NUMBERING: 11,
    TAG_BULLET: 12,
    TAG_PARA_SHAPE: 13,
    TAG_STYLE: 14,
    TAG_MEMO_SHAPE: 15,
}
_DEFAULT_HWP_NUMBERING_PAYLOAD = bytes.fromhex(
    "0c00000000003200ffffffff03005e0031002e000c01000000003200ffffffff03005e0032002e000c00000000003200ffffffff"
    "03005e00330029000c01000000003200ffffffff03005e00340029000c00000000003200ffffffff040028005e00350029000c01"
    "000000003200ffffffff040028005e00360029002c00000000003200ffffffff02005e0037000000010000000100000001000000"
    "01000000010000000100000001000000"
)
_DEFAULT_HWP_BULLET_PAYLOAD = bytes.fromhex("0800000000003200ffffffffa1f0000000000000000000")
_DEFAULT_HWP_PARA_SHAPE_PAYLOAD = bytes.fromhex(
    "800100000000000000000000000000000000000000000000a000000000000000050000000000000000000000000000000000a0000000"
)
_DEFAULT_HWP_CHAR_SHAPE_PAYLOAD = bytes.fromhex(
    "0100010001000100000002000000646464646464640000000000000064646464646464000000000000004c040000020000000a0a0000000000000000ffffffffc0c0c000040000000000"
)

_HWP_PARA_ALIGN_CODES = {
    "JUSTIFY": 0,
    "LEFT": 1,
    "RIGHT": 2,
    "CENTER": 3,
    "DISTRIBUTE": 4,
    "DISTRIBUTE_SPACE": 5,
}
_HWP_PARA_ALIGN_NAMES = {value: key for key, value in _HWP_PARA_ALIGN_CODES.items()}
_HWP_NUMBERING_LEVEL_COUNT = 7
_HWP_CHAR_FONT_ORDER = ("hangul", "latin", "hanja", "japanese", "other", "symbol", "user")

def _parse_optional_int(value: str | None) -> int | None:
    if value is None:
        return None
    stripped = value.strip()
    return int(stripped) if stripped.isdigit() else None

def _coerce_bool(value) -> bool | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, int):
        return bool(value)
    text = str(value).strip().lower()
    if text in {"1", "true", "yes"}:
        return True
    if text in {"0", "false", "no"}:
        return False
    return None

def _extract_style_definition(style) -> StyleDefinition:
    return StyleDefinition(
        name=style.name or "",
        style_id=style.style_id,
        english_name=style.english_name,
        para_pr_id=style.para_pr_id,
        char_pr_id=style.char_pr_id,
        style_type=style.element.get("type", "PARA"),
        next_style_id=style.element.get("nextStyleIDRef"),
        lang_id=style.element.get("langID", "1042"),
        lock_form=style.element.get("lockForm") == "1",
    )


def _extract_paragraph_style(style) -> ParagraphStyle:
    line_spacing_types = style.element.xpath(".//hh:lineSpacing/@type", namespaces=NS)
    return ParagraphStyle(
        style_id=style.style_id,
        alignment_horizontal=style.alignment_horizontal,
        alignment_vertical=style.element.xpath("./hh:align/@vertical", namespaces=NS)[0]
        if style.element.xpath("./hh:align/@vertical", namespaces=NS)
        else None,
        line_spacing=style.line_spacing,
        line_spacing_type=line_spacing_types[0] if line_spacing_types else None,
        margins=_extract_hwpx_margin_map(style.element),
        tab_pr_id=style.element.get("tabPrIDRef"),
        snap_to_grid=style.element.get("snapToGrid"),
        condense=style.element.get("condense"),
        font_line_height=style.element.get("fontLineHeight"),
        suppress_line_numbers=style.element.get("suppressLineNumbers"),
        checked=style.element.get("checked"),
        heading=_extract_hwpx_attribute_map(
            style.element,
            "./hh:heading",
            "type",
            "idRef",
            "level",
        ),
        break_setting=_extract_hwpx_attribute_map(
            style.element,
            "./hh:breakSetting",
            "breakLatinWord",
            "breakNonLatinWord",
            "widowOrphan",
            "keepWithNext",
            "keepLines",
            "pageBreakBefore",
            "lineWrap",
        ),
        auto_spacing=_extract_hwpx_attribute_map(
            style.element,
            "./hh:autoSpacing",
            "eAsianEng",
            "eAsianNum",
        ),
    )


def _extract_character_style(style) -> CharacterStyle:
    return CharacterStyle(
        style_id=style.style_id,
        text_color=style.text_color,
        shade_color=style.element.get("shadeColor"),
        height=style.height,
        use_font_space=style.element.get("useFontSpace"),
        use_kerning=style.element.get("useKerning"),
        sym_mark=style.element.get("symMark"),
        font_refs=_extract_hwpx_attribute_map(
            style.element,
            "./hh:fontRef",
            "hangul",
            "latin",
            "hanja",
            "japanese",
            "other",
            "symbol",
            "user",
        ),
    )


def _extract_hwpx_margin_map(element: etree._Element) -> dict[str, str]:
    values: dict[str, str] = {}
    for key in ("intent", "left", "right", "prev", "next"):
        nodes = element.xpath(f"./hh:margin/hc:{key}", namespaces=NS)
        if not nodes:
            continue
        value = nodes[0].get("value")
        if value is not None:
            values[key] = value
        unit = nodes[0].get("unit")
        if unit is not None and "unit" not in values:
            values["unit"] = unit
    return values


def _extract_hwpx_attribute_map(element: etree._Element, xpath: str, *attributes: str) -> dict[str, str]:
    nodes = element.xpath(xpath, namespaces=NS)
    if not nodes:
        return {}
    return {
        attribute: value
        for attribute in attributes
        if (value := nodes[0].get(attribute)) is not None
    }


def _extract_hwp_style_definitions(docinfo_model) -> list[StyleDefinition]:
    definitions: list[StyleDefinition] = []
    for style_id, record in enumerate(docinfo_model.style_records()):
        definition = _parse_hwp_style_definition_payload(record.payload, style_id=style_id)
        if definition is not None:
            definitions.append(definition)
    return definitions


def _parse_hwp_style_definition_payload(payload: bytes, *, style_id: int | None = None) -> StyleDefinition | None:
    if len(payload) < 4:
        return None
    cursor = 0
    name_length = int.from_bytes(payload[cursor : cursor + 2], "little", signed=False)
    cursor += 2
    name_end = cursor + (name_length * 2)
    if name_end > len(payload):
        return None
    name = payload[cursor:name_end].decode("utf-16-le", errors="ignore")
    cursor = name_end
    if cursor + 2 > len(payload):
        return None
    english_name_length = int.from_bytes(payload[cursor : cursor + 2], "little", signed=False)
    cursor += 2
    english_name_end = cursor + (english_name_length * 2)
    if english_name_end > len(payload):
        return None
    english_name = payload[cursor:english_name_end].decode("utf-16-le", errors="ignore")
    cursor = english_name_end
    if cursor + 10 > len(payload):
        return None
    style_word = int.from_bytes(payload[cursor : cursor + 2], "little", signed=False)
    cursor += 2
    lang_id = int.from_bytes(payload[cursor : cursor + 2], "little", signed=False)
    cursor += 2
    next_style_id = int.from_bytes(payload[cursor : cursor + 2], "little", signed=False)
    cursor += 2
    para_pr_id = int.from_bytes(payload[cursor : cursor + 2], "little", signed=False)
    cursor += 2
    char_pr_id = int.from_bytes(payload[cursor : cursor + 2], "little", signed=False)
    style_type_code = style_word & 0x00FF
    resolved_style_id = style_word >> 8 if style_id is None else style_id
    return StyleDefinition(
        name=name or f"Style {resolved_style_id}",
        style_id=str(resolved_style_id),
        english_name=english_name or None,
        style_type=_HWP_STYLE_TYPE_NAMES.get(style_type_code, "PARA"),
        para_pr_id=str(para_pr_id),
        char_pr_id=str(char_pr_id),
        next_style_id=str(next_style_id),
        lang_id=str(lang_id),
        lock_form=False,
        native_hwp_payload=bytes(payload),
    )


def _build_hwp_placeholder_paragraph_styles(
    docinfo_model,
    sections: list[HancomSection],
    style_definitions: list[StyleDefinition],
) -> list[ParagraphStyle]:
    style_ids = {
        value
        for value in (_parse_optional_int(style.para_pr_id) for style in style_definitions)
        if value is not None
    }
    for section in sections:
        for block in section.blocks:
            if not isinstance(block, Paragraph):
                continue
            candidate = block.hwp_para_shape_id
            if candidate is None:
                candidate = _parse_optional_int(block.para_pr_id)
            if candidate is not None:
                style_ids.add(candidate)
    payloads = docinfo_model.para_shape_records()
    styles: list[ParagraphStyle] = []
    for style_id in sorted(style_ids):
        payload = bytes(payloads[style_id].payload) if 0 <= style_id < len(payloads) else None
        style = ParagraphStyle(style_id=str(style_id), native_hwp_payload=payload)
        if payload is not None:
            _apply_hwp_paragraph_style_payload(style, payload)
        styles.append(style)
    return styles


def _build_hwp_placeholder_character_styles(docinfo_model, style_definitions: list[StyleDefinition]) -> list[CharacterStyle]:
    style_ids = {
        value
        for value in (_parse_optional_int(style.char_pr_id) for style in style_definitions)
        if value is not None
    }
    payloads = docinfo_model.char_shape_records()
    styles: list[CharacterStyle] = []
    for style_id in sorted(style_ids):
        payload = bytes(payloads[style_id].payload) if 0 <= style_id < len(payloads) else None
        style = CharacterStyle(style_id=str(style_id), native_hwp_payload=payload)
        if payload is not None:
            _apply_hwp_character_style_payload(style, payload)
        styles.append(style)
    return styles


def _extract_hwp_numbering_definitions(docinfo_model) -> list[NumberingDefinition]:
    definitions: list[NumberingDefinition] = []
    for index, record in enumerate(docinfo_model.numbering_records()):
        definition = NumberingDefinition(style_id=str(index), native_hwp_payload=bytes(record.payload))
        _apply_hwp_numbering_payload(definition, record.payload)
        fields = record.fields()
        definition.para_head_flags = [dict(level.get("para_head_flags", {})) for level in fields.get("levels", [])]
        definition.unknown_short_bits = dict(fields.get("unknown_short_bits", {}))
        definitions.append(definition)
    return definitions


def _extract_hwp_bullet_definitions(docinfo_model) -> list[BulletDefinition]:
    definitions: list[BulletDefinition] = []
    for index, record in enumerate(docinfo_model.bullet_records()):
        definition = BulletDefinition(style_id=str(index), native_hwp_payload=bytes(record.payload))
        _apply_hwp_bullet_payload(definition, record.payload)
        fields = record.fields()
        definition.flags = int(fields.get("flags", 0))
        definition.flag_bits = dict(fields.get("flag_bits", {}))
        definitions.append(definition)
    return definitions


def _extract_hwp_memo_shape_definitions(docinfo_model) -> list[MemoShapeDefinition]:
    definitions: list[MemoShapeDefinition] = []
    for index, record in enumerate(docinfo_model.memo_shape_records()):
        fields = record.fields()
        definitions.append(
            MemoShapeDefinition(
                memo_shape_id=str(index),
                width=_coerce_int(fields.get("width")),
                line_width=_coerce_int(fields.get("line_width")),
                line_type=str(fields["line_type"]) if fields.get("line_type") is not None else None,
                line_color=str(fields["line_color"]) if fields.get("line_color") is not None else None,
                fill_color=str(fields["fill_color"]) if fields.get("fill_color") is not None else None,
                active_color=str(fields["active_color"]) if fields.get("active_color") is not None else None,
                memo_type=str(fields["memo_type"]) if fields.get("memo_type") is not None else None,
                native_hwp_payload=bytes(record.payload),
            )
        )
    return definitions


def _coerce_int(value: object | None, *, default: int | None = None) -> int | None:
    if value is None:
        return default
    if isinstance(value, int):
        return value
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return default


def _seed_hwp_payload(
    native_payload: bytes | None,
    fallback_payload: bytes | None,
    default_payload: bytes,
    *,
    min_len: int | None = None,
) -> bytearray:
    resolved = bytes(native_payload or fallback_payload or default_payload)
    target_length = max(len(default_payload), min_len or 0)
    if len(resolved) < target_length:
        resolved += b"\x00" * (target_length - len(resolved))
    return bytearray(resolved)


def _parse_hwp_colorref(payload: bytes) -> str:
    values = payload[:4].ljust(4, b"\x00")
    return f"#{values[2]:02X}{values[1]:02X}{values[0]:02X}"


def _build_hwp_colorref(value: object | None, *, default: str = "#000000") -> bytes:
    normalized = str(value or default).strip().lstrip("#")
    if len(normalized) != 6:
        normalized = default.lstrip("#")
    try:
        red = int(normalized[0:2], 16)
        green = int(normalized[2:4], 16)
        blue = int(normalized[4:6], 16)
    except ValueError:
        red = green = blue = 0
    return bytes((blue, green, red, 0))


def _apply_hwp_paragraph_style_payload(style: ParagraphStyle, payload: bytes) -> None:
    values = bytes(payload)
    style.native_hwp_payload = values
    if len(values) >= 4:
        attributes = int.from_bytes(values[0:4], "little", signed=False)
        style.alignment_horizontal = _HWP_PARA_ALIGN_NAMES.get(attributes & 0x07)
        style.snap_to_grid = "1" if attributes & (1 << 7) else "0"
        style.break_setting = {
            "breakLatinWord": "KEEP_WORD" if attributes & (1 << 8) else "BREAK_WORD",
            "breakNonLatinWord": "KEEP_WORD" if attributes & (1 << 8) else "BREAK_WORD",
            "widowOrphan": "1" if attributes & (1 << 9) else "0",
            "keepWithNext": "1" if attributes & (1 << 10) else "0",
            "keepLines": "1" if attributes & (1 << 11) else "0",
            "pageBreakBefore": "1" if attributes & (1 << 12) else "0",
            "lineWrap": "SQUEEZE" if attributes & (1 << 15) else "BREAK",
        }
        style.auto_spacing = {
            "eAsianEng": "1" if attributes & (1 << 23) else "0",
            "eAsianNum": "1" if attributes & (1 << 24) else "0",
        }
    if len(values) >= 24:
        margins = {
            "intent": str(int.from_bytes(values[4:8], "little", signed=True)),
            "left": str(int.from_bytes(values[8:12], "little", signed=True)),
            "right": str(int.from_bytes(values[12:16], "little", signed=True)),
            "prev": str(int.from_bytes(values[16:20], "little", signed=True)),
            "next": str(int.from_bytes(values[20:24], "little", signed=True)),
            "unit": "HWPUNIT",
        }
        style.margins = margins
    if len(values) >= 30:
        style.tab_pr_id = str(int.from_bytes(values[28:30], "little", signed=False))
    if len(values) >= 32:
        numbering_id = int.from_bytes(values[30:32], "little", signed=False)
        style.heading = {
            "type": "NUMBER" if numbering_id > 0 else "NONE",
            "idRef": str(numbering_id),
            "level": "0",
        }
    if len(values) >= 54:
        style.line_spacing = str(int.from_bytes(values[50:54], "little", signed=True))


def _build_hwp_paragraph_style_payload(style: ParagraphStyle, fallback_payload: bytes | None = None) -> bytes:
    payload = _seed_hwp_payload(style.native_hwp_payload, fallback_payload, _DEFAULT_HWP_PARA_SHAPE_PAYLOAD)
    attributes = int.from_bytes(payload[0:4], "little", signed=False)
    alignment_code = _HWP_PARA_ALIGN_CODES.get(str(style.alignment_horizontal or "").upper())
    if alignment_code is not None:
        attributes = (attributes & ~0x07) | alignment_code
    snap_to_grid = _coerce_bool(style.snap_to_grid)
    if snap_to_grid is not None:
        attributes = (attributes | (1 << 7)) if snap_to_grid else (attributes & ~(1 << 7))
    keep_word = str(style.break_setting.get("breakLatinWord", "")).upper() == "KEEP_WORD" or str(style.break_setting.get("breakNonLatinWord", "")).upper() == "KEEP_WORD"
    attributes = (attributes | (1 << 8)) if keep_word else (attributes & ~(1 << 8))
    for key, bit in {
        "widowOrphan": 9,
        "keepWithNext": 10,
        "keepLines": 11,
        "pageBreakBefore": 12,
    }.items():
        enabled = _coerce_bool(style.break_setting.get(key))
        if enabled is None:
            continue
        attributes = (attributes | (1 << bit)) if enabled else (attributes & ~(1 << bit))
    line_wrap = str(style.break_setting.get("lineWrap", "")).upper()
    if line_wrap:
        attributes = (attributes | (1 << 15)) if line_wrap == "SQUEEZE" else (attributes & ~(1 << 15))
    for key, bit in {
        "eAsianEng": 23,
        "eAsianNum": 24,
    }.items():
        enabled = _coerce_bool(style.auto_spacing.get(key))
        if enabled is None:
            continue
        attributes = (attributes | (1 << bit)) if enabled else (attributes & ~(1 << bit))
    payload[0:4] = attributes.to_bytes(4, "little", signed=False)
    for key, offset in {
        "intent": 4,
        "left": 8,
        "right": 12,
        "prev": 16,
        "next": 20,
    }.items():
        if key not in style.margins:
            continue
        value = _coerce_int(style.margins.get(key))
        if value is None:
            continue
        payload[offset : offset + 4] = int(value).to_bytes(4, "little", signed=True)
    tab_pr_id = _coerce_int(style.tab_pr_id)
    if tab_pr_id is not None:
        payload[28:30] = int(tab_pr_id).to_bytes(2, "little", signed=False)
    heading_id = _coerce_int(style.heading.get("idRef")) if style.heading else None
    if heading_id is not None:
        payload[30:32] = int(heading_id).to_bytes(2, "little", signed=False)
    line_spacing = _coerce_int(style.line_spacing)
    if line_spacing is not None:
        payload[50:54] = int(line_spacing).to_bytes(4, "little", signed=True)
    return bytes(payload)


def _apply_hwp_character_style_payload(style: CharacterStyle, payload: bytes) -> None:
    values = bytes(payload)
    style.native_hwp_payload = values
    if len(values) >= 14:
        style.font_refs = {
            key: str(int.from_bytes(values[index * 2 : (index + 1) * 2], "little", signed=False))
            for index, key in enumerate(_HWP_CHAR_FONT_ORDER)
        }
    if len(values) >= 46:
        style.height = str(int.from_bytes(values[42:46], "little", signed=False))
    if len(values) >= 50:
        attributes = int.from_bytes(values[46:50], "little", signed=False)
        style.use_font_space = "1" if attributes & 0x20000000 else "0"
        style.use_kerning = "1" if attributes & 0x40000000 else "0"
    if len(values) >= 58:
        style.text_color = _parse_hwp_colorref(values[54:58])
    if len(values) >= 64:
        style.shade_color = "none" if values[60:64] == b"\xff\xff\xff\xff" else _parse_hwp_colorref(values[60:64])


def _build_hwp_character_style_payload(style: CharacterStyle, fallback_payload: bytes | None = None) -> bytes:
    payload = _seed_hwp_payload(style.native_hwp_payload, fallback_payload, _DEFAULT_HWP_CHAR_SHAPE_PAYLOAD)
    for index, key in enumerate(_HWP_CHAR_FONT_ORDER):
        if key not in style.font_refs:
            continue
        value = _coerce_int(style.font_refs.get(key))
        if value is None:
            continue
        payload[index * 2 : (index + 1) * 2] = int(value).to_bytes(2, "little", signed=False)
    height = _coerce_int(style.height)
    if height is not None:
        payload[42:46] = int(height).to_bytes(4, "little", signed=False)
    attributes = int.from_bytes(payload[46:50], "little", signed=False)
    use_font_space = _coerce_bool(style.use_font_space)
    if use_font_space is not None:
        attributes = (attributes | 0x20000000) if use_font_space else (attributes & ~0x20000000)
    use_kerning = _coerce_bool(style.use_kerning)
    if use_kerning is not None:
        attributes = (attributes | 0x40000000) if use_kerning else (attributes & ~0x40000000)
    payload[46:50] = int(attributes).to_bytes(4, "little", signed=False)
    if style.text_color is not None:
        payload[54:58] = _build_hwp_colorref(style.text_color)
    if style.shade_color is not None:
        payload[60:64] = b"\xff\xff\xff\xff" if str(style.shade_color).lower() == "none" else _build_hwp_colorref(style.shade_color)
    return bytes(payload)


def _apply_hwp_numbering_payload(definition: NumberingDefinition, payload: bytes) -> None:
    values = bytes(payload)
    definition.native_hwp_payload = values
    definition.formats = []
    definition.para_heads = []
    definition.para_head_flags = []
    definition.width_adjusts = []
    definition.text_offsets = []
    definition.char_pr_ids = []
    definition.start_numbers = []
    definition.unknown_short = 0
    definition.unknown_short_bits = {}
    cursor = 0
    for _ in range(_HWP_NUMBERING_LEVEL_COUNT):
        if cursor + 14 > len(values):
            break
        para_head = int.from_bytes(values[cursor : cursor + 4], "little", signed=False)
        definition.para_heads.append(para_head)
        definition.para_head_flags.append({f"bit_{bit}": bool(para_head & (1 << bit)) for bit in range(32)})
        cursor += 4
        definition.width_adjusts.append(int.from_bytes(values[cursor : cursor + 2], "little", signed=True))
        cursor += 2
        definition.text_offsets.append(int.from_bytes(values[cursor : cursor + 2], "little", signed=True))
        cursor += 2
        char_pr_id = int.from_bytes(values[cursor : cursor + 4], "little", signed=True)
        definition.char_pr_ids.append(None if char_pr_id < 0 else char_pr_id)
        cursor += 4
        format_length = int.from_bytes(values[cursor : cursor + 2], "little", signed=False)
        cursor += 2
        format_byte_length = min(format_length * 2, max(len(values) - cursor, 0))
        format_bytes = values[cursor : cursor + format_byte_length]
        if len(format_bytes) % 2:
            format_bytes = format_bytes[:-1]
        definition.formats.append(format_bytes.decode("utf-16-le", errors="ignore"))
        cursor += format_byte_length
    if cursor + 2 <= len(values):
        definition.unknown_short = int.from_bytes(values[cursor : cursor + 2], "little", signed=False)
        definition.unknown_short_bits = {f"bit_{bit}": bool(definition.unknown_short & (1 << bit)) for bit in range(16)}
        cursor += 2
    remaining_levels = min(_HWP_NUMBERING_LEVEL_COUNT, (len(values) - cursor) // 4)
    for _ in range(remaining_levels):
        definition.start_numbers.append(int.from_bytes(values[cursor : cursor + 4], "little", signed=False))
        cursor += 4


def _build_hwp_numbering_payload(definition: NumberingDefinition, fallback_payload: bytes | None = None) -> bytes:
    template = NumberingDefinition()
    _apply_hwp_numbering_payload(template, bytes(definition.native_hwp_payload or fallback_payload or _DEFAULT_HWP_NUMBERING_PAYLOAD))
    payload = bytearray()
    for level in range(_HWP_NUMBERING_LEVEL_COUNT):
        flag_value = _bit_dict_to_int(definition.para_head_flags[level]) if level < len(definition.para_head_flags) else 0
        para_head = (
            definition.para_heads[level]
            if level < len(definition.para_heads)
            else (flag_value if flag_value else (template.para_heads[level] if level < len(template.para_heads) else 12))
        )
        width_adjust = (
            definition.width_adjusts[level]
            if level < len(definition.width_adjusts)
            else (template.width_adjusts[level] if level < len(template.width_adjusts) else 0)
        )
        text_offset = (
            definition.text_offsets[level]
            if level < len(definition.text_offsets)
            else (template.text_offsets[level] if level < len(template.text_offsets) else 50)
        )
        if level < len(definition.char_pr_ids):
            char_pr_id = definition.char_pr_ids[level]
        else:
            char_pr_id = template.char_pr_ids[level] if level < len(template.char_pr_ids) else None
        if level < len(definition.formats) and definition.formats[level]:
            format_text = definition.formats[level]
        else:
            format_text = template.formats[level] if level < len(template.formats) else f"^{level + 1}."
        encoded_format = str(format_text).encode("utf-16-le")
        payload.extend(int(para_head).to_bytes(4, "little", signed=False))
        payload.extend(int(width_adjust).to_bytes(2, "little", signed=True))
        payload.extend(int(text_offset).to_bytes(2, "little", signed=True))
        payload.extend(int(-1 if char_pr_id is None else char_pr_id).to_bytes(4, "little", signed=True))
        payload.extend((len(encoded_format) // 2).to_bytes(2, "little", signed=False))
        payload.extend(encoded_format)
    unknown_short = definition.unknown_short if not definition.unknown_short_bits else _bit_dict_to_int(definition.unknown_short_bits)
    payload.extend(int(unknown_short).to_bytes(2, "little", signed=False))
    for level in range(_HWP_NUMBERING_LEVEL_COUNT):
        start_number = (
            definition.start_numbers[level]
            if level < len(definition.start_numbers)
            else (template.start_numbers[level] if level < len(template.start_numbers) else 1)
        )
        payload.extend(int(start_number).to_bytes(4, "little", signed=False))
    return bytes(payload)


def _apply_hwp_bullet_payload(definition: BulletDefinition, payload: bytes) -> None:
    values = bytes(payload)
    definition.native_hwp_payload = values
    if len(values) >= 4:
        definition.flags = int.from_bytes(values[0:4], "little", signed=False)
        definition.flag_bits = {f"bit_{bit}": bool(definition.flags & (1 << bit)) for bit in range(32)}
    if len(values) >= 6:
        definition.width_adjust = int.from_bytes(values[4:6], "little", signed=True)
    if len(values) >= 8:
        definition.text_offset = int.from_bytes(values[6:8], "little", signed=True)
    if len(values) >= 12:
        char_pr_id = int.from_bytes(values[8:12], "little", signed=True)
        definition.char_pr_id = None if char_pr_id < 0 else char_pr_id
    if len(values) >= 14:
        definition.bullet_char = values[12:14].decode("utf-16-le", errors="ignore") or "\uf0a1"
    definition.unknown_tail = bytes(values[14:])


def _build_hwp_bullet_payload(definition: BulletDefinition, fallback_payload: bytes | None = None) -> bytes:
    payload = _seed_hwp_payload(definition.native_hwp_payload, fallback_payload, _DEFAULT_HWP_BULLET_PAYLOAD, min_len=14)
    template = BulletDefinition()
    _apply_hwp_bullet_payload(template, bytes(payload))
    flags = definition.flags
    if flags is None and definition.flag_bits:
        flags = _bit_dict_to_int(definition.flag_bits)
    if flags is None:
        flags = template.flags if template.flags is not None else int.from_bytes(payload[0:4], "little", signed=False)
    payload[0:4] = int(flags).to_bytes(4, "little", signed=False)
    payload[4:6] = int(definition.width_adjust).to_bytes(2, "little", signed=True)
    payload[6:8] = int(definition.text_offset).to_bytes(2, "little", signed=True)
    char_pr_id = template.char_pr_id if definition.char_pr_id is None else definition.char_pr_id
    payload[8:12] = int(-1 if char_pr_id is None else char_pr_id).to_bytes(4, "little", signed=True)
    bullet_char = definition.bullet_char or template.bullet_char or "\uf0a1"
    encoded_char = bullet_char[:1].encode("utf-16-le", errors="ignore")[:2].ljust(2, b"\x00")
    payload[12:14] = encoded_char
    if definition.unknown_tail is not None:
        payload[14:] = definition.unknown_tail
    return bytes(payload)


def _build_hwp_memo_shape_payload(definition: MemoShapeDefinition, fallback_payload: bytes | None = None) -> bytes:
    semantic_fields = {
        "width": definition.width,
        "lineWidth": definition.line_width,
        "lineType": definition.line_type,
        "lineColor": definition.line_color,
        "fillColor": definition.fill_color,
        "activeColor": definition.active_color,
        "memoType": definition.memo_type,
    }
    if definition.native_hwp_payload is not None and not any(value is not None for value in semantic_fields.values()):
        return bytes(definition.native_hwp_payload)
    return _build_memo_shape_native_metadata_payload(**semantic_fields)


def _sync_hwp_memo_shape_payloads(records, nodes: list[MemoShapeDefinition]) -> None:
    for node in nodes:
        memo_shape_id = _parse_optional_int(node.memo_shape_id)
        if memo_shape_id is None or not 0 <= memo_shape_id < len(records):
            continue
        fallback_payload = bytes(records[memo_shape_id].payload)
        records[memo_shape_id].payload = bytes(_build_hwp_memo_shape_payload(node, fallback_payload))


def _merge_hwp_style_defaults_into_paragraphs(document: HancomDocument) -> None:
    style_by_id = {
        style.style_id: style
        for style in document.style_definitions
        if style.style_id is not None
    }
    for section in document.sections:
        for block in section.blocks:
            if not isinstance(block, Paragraph) or block.style_id is None:
                continue
            style = style_by_id.get(block.style_id)
            if style is None:
                continue
            if block.para_pr_id is None and style.para_pr_id is not None:
                block.para_pr_id = style.para_pr_id
            if block.char_pr_id is None and style.char_pr_id is not None:
                block.char_pr_id = style.char_pr_id


def _sync_hwp_docinfo_styles(document: HwpDocument, hancom_document: HancomDocument) -> None:
    docinfo_model = document.docinfo_model()
    initial_style_count = len(docinfo_model.style_records())
    _assign_missing_hwp_style_ids(
        hancom_document.paragraph_styles,
        existing_ids={
            value
            for value in (
                _parse_optional_int(style.style_id)
                for style in hancom_document.paragraph_styles
            )
            if value is not None
        },
    )
    _assign_missing_hwp_style_ids(
        hancom_document.character_styles,
        existing_ids={
            value
            for value in (
                _parse_optional_int(style.style_id)
                for style in hancom_document.character_styles
            )
            if value is not None
        },
    )
    _assign_missing_hwp_style_ids(
        hancom_document.style_definitions,
        existing_ids={
            value
            for value in (
                _parse_optional_int(style.style_id)
                for style in hancom_document.style_definitions
            )
            if value is not None and 0 <= value <= 0xFF
        },
        max_id=0xFF,
    )
    _assign_missing_hwp_style_ids(
        hancom_document.numbering_definitions,
        existing_ids={
            value
            for value in (
                _parse_optional_int(definition.style_id)
                for definition in hancom_document.numbering_definitions
            )
            if value is not None
        },
    )
    _assign_missing_hwp_style_ids(
        hancom_document.bullet_definitions,
        existing_ids={
            value
            for value in (
                _parse_optional_int(definition.style_id)
                for definition in hancom_document.bullet_definitions
            )
            if value is not None
        },
    )

    required_char_shape_ids = {
        value
        for value in (
            _parse_optional_int(style.style_id)
            for style in hancom_document.character_styles
        )
        if value is not None
    }
    required_para_shape_ids = {
        value
        for value in (
            _parse_optional_int(style.style_id)
            for style in hancom_document.paragraph_styles
        )
        if value is not None
    }
    required_style_ids = {
        value
        for value in (
            _parse_optional_int(style.style_id)
            for style in hancom_document.style_definitions
        )
        if value is not None and 0 <= value <= 0xFF
    }
    required_numbering_ids = {
        value
        for value in (
            _parse_optional_int(definition.style_id)
            for definition in hancom_document.numbering_definitions
        )
        if value is not None
    }
    required_bullet_ids = {
        value
        for value in (
            _parse_optional_int(definition.style_id)
            for definition in hancom_document.bullet_definitions
        )
        if value is not None
    }
    required_memo_shape_ids = {
        value
        for value in (
            _parse_optional_int(definition.memo_shape_id)
            for definition in hancom_document.memo_shape_definitions
        )
        if value is not None
    }

    for style in hancom_document.style_definitions:
        para_pr_id = _parse_optional_int(style.para_pr_id)
        char_pr_id = _parse_optional_int(style.char_pr_id)
        if para_pr_id is not None:
            required_para_shape_ids.add(para_pr_id)
        if char_pr_id is not None:
            required_char_shape_ids.add(char_pr_id)
    for definition in hancom_document.numbering_definitions:
        for char_pr_id in definition.char_pr_ids:
            if char_pr_id is not None:
                required_char_shape_ids.add(char_pr_id)
    for definition in hancom_document.bullet_definitions:
        if definition.char_pr_id is not None:
            required_char_shape_ids.add(definition.char_pr_id)

    for section in hancom_document.sections:
        for block in section.blocks:
            if not isinstance(block, Paragraph):
                continue
            style_id = block.hwp_style_id
            if style_id is None:
                style_id = _parse_optional_int(block.style_id)
            if style_id is not None and 0 <= style_id <= 0xFF:
                required_style_ids.add(style_id)
            para_shape_id = block.hwp_para_shape_id
            if para_shape_id is None:
                para_shape_id = _parse_optional_int(block.para_pr_id)
            if para_shape_id is not None:
                required_para_shape_ids.add(para_shape_id)
            char_pr_id = _parse_optional_int(block.char_pr_id)
            if char_pr_id is not None:
                required_char_shape_ids.add(char_pr_id)
        numbering_shape_id = _parse_optional_int(section.settings.numbering_shape_id)
        if numbering_shape_id is not None:
            required_numbering_ids.add(numbering_shape_id)
    if required_char_shape_ids:
        _ensure_hwp_docinfo_record_capacity(docinfo_model, TAG_CHAR_SHAPE, max(required_char_shape_ids) + 1)
    if required_para_shape_ids:
        _ensure_hwp_docinfo_record_capacity(docinfo_model, TAG_PARA_SHAPE, max(required_para_shape_ids) + 1)
    if required_style_ids:
        _ensure_hwp_docinfo_record_capacity(docinfo_model, TAG_STYLE, max(required_style_ids) + 1)
    if required_numbering_ids:
        _ensure_hwp_docinfo_record_capacity(
            docinfo_model,
            TAG_NUMBERING,
            max(required_numbering_ids) + 1,
            default_payload=_DEFAULT_HWP_NUMBERING_PAYLOAD,
        )
    if required_bullet_ids:
        _ensure_hwp_docinfo_record_capacity(
            docinfo_model,
            TAG_BULLET,
            max(required_bullet_ids) + 1,
            default_payload=_DEFAULT_HWP_BULLET_PAYLOAD,
        )
    if required_memo_shape_ids:
        _ensure_hwp_docinfo_record_capacity(
            docinfo_model,
            TAG_MEMO_SHAPE,
            max(required_memo_shape_ids) + 1,
            default_payload=_build_memo_shape_native_metadata_payload(),
        )

    _sync_hwp_docinfo_slot_payloads(
        docinfo_model.char_shape_records(),
        hancom_document.character_styles,
        builder=_build_hwp_character_style_payload,
    )
    _sync_hwp_docinfo_slot_payloads(
        docinfo_model.para_shape_records(),
        hancom_document.paragraph_styles,
        builder=_build_hwp_paragraph_style_payload,
    )
    _sync_hwp_docinfo_slot_payloads(
        docinfo_model.records_by_tag_id(TAG_NUMBERING),
        hancom_document.numbering_definitions,
        builder=_build_hwp_numbering_payload,
    )
    _sync_hwp_docinfo_slot_payloads(
        docinfo_model.records_by_tag_id(TAG_BULLET),
        hancom_document.bullet_definitions,
        builder=_build_hwp_bullet_payload,
    )
    _sync_hwp_memo_shape_payloads(
        docinfo_model.memo_shape_records(),
        hancom_document.memo_shape_definitions,
    )

    # Explicit style definitions still need direct HWP style record authoring.
    explicit_style_ids = {
        style_id
        for style_id in (_parse_optional_int(style.style_id) for style in hancom_document.style_definitions)
        if style_id is not None and 0 <= style_id <= 0xFF
    }
    for style in hancom_document.style_definitions:
        style_id = _parse_optional_int(style.style_id)
        if style_id is None or not 0 <= style_id <= 0xFF:
            continue
        docinfo_model.style_records()[style_id].payload = _build_hwp_style_definition_payload(style, style_id=style_id)

    for style_id in range(initial_style_count, len(docinfo_model.style_records())):
        if style_id in explicit_style_ids or style_id > 0xFF:
            continue
        docinfo_model.style_records()[style_id].payload = _build_hwp_style_definition_payload(
            StyleDefinition(name=f"Style {style_id}", style_id=str(style_id)),
            style_id=style_id,
        )

    document.binary_document().replace_docinfo_model(docinfo_model)


def _assign_missing_hwp_style_ids(nodes, *, existing_ids: set[int], max_id: int | None = None) -> None:
    next_id = max(existing_ids, default=-1) + 1
    for node in nodes:
        if getattr(node, "style_id", None) is not None:
            continue
        while next_id in existing_ids:
            next_id += 1
        if max_id is not None and next_id > max_id:
            break
        node.style_id = str(next_id)
        existing_ids.add(next_id)
        next_id += 1


def _sync_hwp_docinfo_slot_payloads(records, nodes, *, builder=None) -> None:
    for node in nodes:
        style_id = _parse_optional_int(getattr(node, "style_id", None))
        if style_id is None:
            continue
        if not 0 <= style_id < len(records):
            continue
        fallback_payload = bytes(records[style_id].payload)
        if builder is None:
            payload = getattr(node, "native_hwp_payload", None)
            if payload is None:
                continue
            records[style_id].payload = bytes(payload)
            continue
        records[style_id].payload = bytes(builder(node, fallback_payload))


def _ensure_hwp_docinfo_record_capacity(
    docinfo_model,
    tag_id: int,
    count: int,
    *,
    default_payload: bytes | None = None,
) -> None:
    if count <= 0:
        return
    records = docinfo_model.records_by_tag_id(tag_id)
    parent = docinfo_model.id_mappings_record()
    if records:
        template_level = records[-1].level
        template_payload = bytes(records[-1].payload)
        insert_at = max(index for index, child in enumerate(parent.children) if child.tag_id == tag_id)
    else:
        if default_payload is None:
            return
        template_level = 1
        template_payload = bytes(default_payload)
        insert_at = max((index for index, child in enumerate(parent.children) if child.tag_id <= tag_id), default=-1)
    while len(records) < count:
        clone = TypedRecord(tag_id=tag_id, level=template_level, payload=template_payload)
        insert_at += 1
        parent.children.insert(insert_at, clone)
        records.append(clone)
    count_index = _HWP_DOCINFO_COUNT_INDEX_BY_TAG.get(tag_id)
    if count_index is not None:
        docinfo_model.id_mappings_record().set_count(count_index, len(records))


def _build_hwp_style_definition_payload(style: StyleDefinition, *, style_id: int) -> bytes:
    name = style.name or f"Style {style_id}"
    english_name = style.english_name or name
    payload = bytearray()
    payload.extend(len(name).to_bytes(2, "little", signed=False))
    payload.extend(name.encode("utf-16-le"))
    payload.extend(len(english_name).to_bytes(2, "little", signed=False))
    payload.extend(english_name.encode("utf-16-le"))
    style_type_code = _HWP_STYLE_TYPE_CODES.get(str(style.style_type or "PARA").upper(), 0)
    payload.extend((((style_id & 0xFF) << 8) | (style_type_code & 0xFF)).to_bytes(2, "little", signed=False))
    payload.extend(_coerce_hwp_docinfo_word(style.lang_id, default=1042).to_bytes(2, "little", signed=False))
    payload.extend(_coerce_hwp_docinfo_word(style.next_style_id).to_bytes(2, "little", signed=False))
    payload.extend(_coerce_hwp_docinfo_word(style.para_pr_id).to_bytes(2, "little", signed=False))
    payload.extend(_coerce_hwp_docinfo_word(style.char_pr_id).to_bytes(2, "little", signed=False))
    trailing_payload = _extract_hwp_style_definition_trailing_payload(style.native_hwp_payload)
    if trailing_payload:
        payload.extend(trailing_payload)
    return bytes(payload)


def _extract_hwp_style_definition_trailing_payload(payload: bytes | None) -> bytes:
    if not payload or len(payload) < 4:
        return b""
    cursor = 0
    name_length = int.from_bytes(payload[cursor : cursor + 2], "little", signed=False)
    cursor += 2 + name_length * 2
    if cursor + 2 > len(payload):
        return b""
    english_name_length = int.from_bytes(payload[cursor : cursor + 2], "little", signed=False)
    cursor += 2 + english_name_length * 2 + 10
    if cursor >= len(payload):
        return b""
    return bytes(payload[cursor:])


def _coerce_hwp_docinfo_word(value: str | None, *, default: int = 0) -> int:
    parsed = _parse_optional_int(value)
    resolved = default if parsed is None else parsed
    return max(0, min(resolved, 0xFFFF))


def _bit_dict_to_int(bits: dict[str, bool] | None) -> int:
    if not bits:
        return 0
    value = 0
    for key, enabled in bits.items():
        if not enabled or not key.startswith("bit_"):
            continue
        try:
            bit = int(key[4:])
        except ValueError:
            continue
        if bit < 0:
            continue
        value |= 1 << bit
    return value
