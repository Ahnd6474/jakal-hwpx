"""Shared-model paragraph, character, and named styles written to HWPX."""
from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from ._hancom_styles import _coerce_bool
from .document import HwpxDocument
from .namespaces import (
    NS,
    qname,
)

if TYPE_CHECKING:
    from .hancom_document import HancomDocument


def _apply_styles(document: HwpxDocument, hancom_document: HancomDocument) -> None:
    def discard_originals(nodes):
        for old in nodes:
            parent = old.getparent()
            parent.remove(old)
            parent.set("itemCnt", str(len(parent)))
        if nodes:
            document.header.mark_modified()

    for raw in getattr(hancom_document, "_hwpx_ref_resources", []):
        node = etree.fromstring(raw)
        ref_list = document.header.root_element.find(qname("hh", "refList"))
        existing = ref_list.find(node.tag)
        if existing is None:
            ref_list.append(node)
        else:
            ref_list.replace(existing, node)
        document.header.mark_modified()
    # Ref 0 is the default tab grid. Reserve it before importing explicit stops;
    # otherwise a header's center/right stops become the default for body lists.
    ref_list = document.header.root_element.find(qname("hh", "refList"))
    tab_properties = ref_list.find(qname("hh", "tabProperties"))
    if tab_properties is None:
        tab_properties = etree.SubElement(ref_list, qname("hh", "tabProperties"), itemCnt="1")
        etree.SubElement(tab_properties, qname("hh", "tabPr"), id="0", autoTabLeft="0", autoTabRight="0")
        document.header.mark_modified()
    for style in hancom_document.paragraph_styles:
        originals = list(document.header.root_element.xpath(".//hh:paraProperties/hh:paraPr[@id=$id]", namespaces=NS, id=style.style_id or ""))
        appended = document.append_paragraph_style(
            style_id=style.style_id,
            alignment_horizontal=style.alignment_horizontal,
            alignment_vertical=style.alignment_vertical,
            line_spacing=None,
        )
        discard_originals(originals)
        if style.native_hwpx_xml is not None:
            from copy import deepcopy
            original = etree.fromstring(style.native_hwpx_xml)
            appended.element.attrib.update(original.attrib)
            appended.element.set("id", str(style.style_id))
            for child in list(appended.element):
                appended.element.remove(child)
            for child in original:
                appended.element.append(deepcopy(child))
            from dataclasses import fields
            from ._hancom_styles import _extract_paragraph_style
            baseline = _extract_paragraph_style(appended)
            if all(getattr(style, item.name) == getattr(baseline, item.name)
                   for item in fields(style) if item.name not in {"native_hwpx_xml", "native_hwp_payload", "style_id"}):
                document.header.mark_modified()
                continue
        for key, value in {
            "tabPrIDRef": style.tab_pr_id,
            "snapToGrid": style.snap_to_grid,
            "condense": style.condense,
            "fontLineHeight": style.font_line_height,
            "suppressLineNumbers": style.suppress_line_numbers,
            "checked": style.checked,
        }.items():
            if value is not None:
                appended.element.set(key, str(value))
        if any(value is not None for value in (style.tab_pr_id, style.snap_to_grid, style.condense, style.font_line_height, style.suppress_line_numbers, style.checked)):
            appended.header_part.mark_modified()
        if style.line_spacing is not None:
            appended.set_line_spacing(style.line_spacing, spacing_type=style.line_spacing_type)
        tab_stops = getattr(style, "_docx_tab_stops", None)
        if tab_stops:
            ref_list = document.header.root_element.find(qname("hh", "refList"))
            tab_properties = ref_list.find(qname("hh", "tabProperties"))
            if tab_properties is None:
                tab_properties = etree.SubElement(ref_list, qname("hh", "tabProperties"), itemCnt="0")
            tab_id = str(max((int(node.get("id", "0")) for node in tab_properties), default=-1) + 1)
            tab = etree.SubElement(tab_properties, qname("hh", "tabPr"), id=tab_id, autoTabLeft="0", autoTabRight="0")
            for stop in tab_stops:
                etree.SubElement(tab, qname("hh", "tabItem"), pos=str(stop["pos"]), type=str(stop["type"]), leader=str(stop["leader"]), unit="HWPUNIT")
            tab_properties.set("itemCnt", str(len(tab_properties)))
            appended.element.set("tabPrIDRef", tab_id)
            document.header.mark_modified()
        if style.margins:
            margins = style.margins
            if getattr(style, "_docx_list_layout", False):
                margins = {key: str(int(value) * 2) if key != "unit" else value for key, value in margins.items()}
            appended.set_margin(
                intent=margins.get("intent"),
                left=margins.get("left"),
                right=margins.get("right"),
                prev=margins.get("prev"),
                next=margins.get("next"),
                unit=margins.get("unit"),
            )
        if style.heading:
            heading_nodes = appended.element.xpath("./hh:heading", namespaces=NS)
            heading = heading_nodes[0] if heading_nodes else etree.SubElement(appended.element, qname("hh", "heading"))
            for attribute in ("type", "idRef", "level"):
                value = style.heading.get(attribute)
                if value is not None:
                    heading.set(attribute, str(value))
            appended.header_part.mark_modified()
        if style.break_setting:
            appended.set_break_setting(
                break_latin_word=style.break_setting.get("breakLatinWord"),
                break_non_latin_word=style.break_setting.get("breakNonLatinWord"),
                widow_orphan=_coerce_bool(style.break_setting.get("widowOrphan")),
                keep_with_next=_coerce_bool(style.break_setting.get("keepWithNext")),
                keep_lines=_coerce_bool(style.break_setting.get("keepLines")),
                page_break_before=_coerce_bool(style.break_setting.get("pageBreakBefore")),
                line_wrap=style.break_setting.get("lineWrap"),
            )
        if style.auto_spacing:
            appended.set_auto_spacing(
                e_asian_eng=_coerce_bool(style.auto_spacing.get("eAsianEng")),
                e_asian_num=_coerce_bool(style.auto_spacing.get("eAsianNum")),
            )
        if style.style_id is None:
            style.style_id = appended.style_id

    for style in hancom_document.character_styles:
        originals = list(document.header.root_element.xpath(".//hh:charProperties/hh:charPr[@id=$id]", namespaces=NS, id=style.style_id or ""))
        appended = document.append_character_style(
            style_id=style.style_id,
            text_color=style.text_color,
            height=style.height,
        )
        discard_originals(originals)
        if style.native_hwpx_xml is not None:
            from copy import deepcopy

            original = etree.fromstring(style.native_hwpx_xml)
            appended.element.attrib.update(original.attrib)
            appended.element.set("id", str(style.style_id))
            for child in list(appended.element):
                appended.element.remove(child)
            for child in original:
                appended.element.append(deepcopy(child))
            if style.text_color is not None:
                appended.set_text_color(style.text_color)
            if style.height is not None:
                appended.set_height(style.height)
        for name, enabled in (("bold", style.bold), ("italic", style.italic)):
            if enabled is None:
                continue
            if style.native_hwpx_xml is not None and bool(appended.element.find(qname("hh", name)) is not None) == enabled:
                continue
            for node in appended.element.findall(qname("hh", name)):
                appended.element.remove(node)
            if enabled:
                node = etree.Element(qname("hh", name))
                underline = next((child for child in appended.element if etree.QName(child).localname in {"underline", "strikeout", "outline", "shadow"}), None)
                if underline is not None:
                    underline.addprevious(node)
                else:
                    appended.element.append(node)
        for name, attributes in (("underline", style.underline), ("strikeout", style.strikeout)):
            if attributes:
                node = appended.element.find(qname("hh", name))
                if node is None:
                    node = etree.SubElement(appended.element, qname("hh", name))
                node.attrib.update(attributes)
        if style.script_position is not None:
            for node in appended.element.xpath("./hh:supscript | ./hh:subscript", namespaces=NS):
                appended.element.remove(node)
            if style.script_position in {"SUPERSCRIPT", "SUBSCRIPT"}:
                etree.SubElement(appended.element, qname("hh", "supscript" if style.script_position == "SUPERSCRIPT" else "subscript"))
        font_refs = dict(style.font_refs)
        if style.font_names:
            fontfaces = document.header.root_element.find(".//" + qname("hh", "fontfaces"))
            if fontfaces is None:
                fontfaces = etree.Element(qname("hh", "fontfaces"), itemCnt="0")
                document.header.root_element.find(qname("hh", "refList")).insert(0, fontfaces)
            for lang, face_name in style.font_names.items():
                family = fontfaces.find(f"{qname('hh', 'fontface')}[@lang='{lang.upper()}']")
                if family is None:
                    family = etree.SubElement(fontfaces, qname("hh", "fontface"), lang=lang.upper(), fontCnt="0")
                fonts = list(family)
                font = next((font for font in fonts if font.get("id") == font_refs.get(lang) and font.get("face") == face_name), None)
                if font is None:
                    font = next((font for font in fonts if font.get("face") == face_name), None)
                if font is None:
                    font = etree.SubElement(family, qname("hh", "font"), id=str(len(fonts)), face=face_name, type="TTF", isEmbedded="0")
                font_refs[lang] = font.get("id", "0")
                family.set("fontCnt", str(len(family)))
            fontfaces.set("itemCnt", str(len(fontfaces)))
        document.header.mark_modified()
        if any(value is not None for value in (style.shade_color, style.use_font_space, style.use_kerning, style.sym_mark)):
            appended.set_effects(
                shade_color=style.shade_color,
                use_font_space=_coerce_bool(style.use_font_space),
                use_kerning=_coerce_bool(style.use_kerning),
                sym_mark=style.sym_mark,
            )
        if font_refs:
            appended.set_font_refs(
                hangul=font_refs.get("hangul"),
                latin=font_refs.get("latin"),
                hanja=font_refs.get("hanja"),
                japanese=font_refs.get("japanese"),
                other=font_refs.get("other"),
                symbol=font_refs.get("symbol"),
                user=font_refs.get("user"),
            )
        if style.font_ratios:
            ratio = appended.element.find(qname("hh", "ratio"))
            if ratio is None:
                ratio = etree.SubElement(appended.element, qname("hh", "ratio"))
            for lang, value in style.font_ratios.items():
                ratio.set(lang, str(value))
            appended.header_part.mark_modified()
        if getattr(style, "_word_extra", None):
            from .addons.docx.formatting import apply_hwpx_character

            apply_hwpx_character(appended.element, style)
            appended.header_part.mark_modified()
        if style.style_id is None:
            style.style_id = appended.style_id

    for style in hancom_document.style_definitions:
        originals = list(document.header.root_element.xpath(".//hh:styles/hh:style[@id=$id]", namespaces=NS, id=style.style_id or ""))
        document.append_style(
            style.name,
            english_name=style.english_name,
            style_id=style.style_id,
            style_type=style.style_type,
            para_pr_id=style.para_pr_id,
            char_pr_id=style.char_pr_id,
            next_style_id=style.next_style_id,
            lang_id=style.lang_id,
            lock_form=style.lock_form,
        )
        discard_originals(originals)

    for memo_shape in hancom_document.memo_shape_definitions:
        appended = document.append_memo_shape(
            memo_shape_id=memo_shape.memo_shape_id,
            width=memo_shape.width if memo_shape.width is not None else 12000,
            line_width=memo_shape.line_width if memo_shape.line_width is not None else 12,
            line_type=memo_shape.line_type or "SOLID",
            line_color=memo_shape.line_color or "#808080",
            fill_color=memo_shape.fill_color or "#FFF2CC",
            active_color=memo_shape.active_color or "#FFC000",
            memo_type=memo_shape.memo_type or "NOMAL",
        )
        if memo_shape.memo_shape_id is None:
            memo_shape.memo_shape_id = appended.memo_shape_id
