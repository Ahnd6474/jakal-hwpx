from __future__ import annotations

import zipfile
from typing import cast

from lxml import etree

from .namespaces import NS, qname
from .parts import (
    ContainerPart,
    ContentHpfPart,
    HeaderPart,
    HwpxPart,
    MimetypePart,
    SectionPart,
    SettingsPart,
    VersionPart,
)


def _full_nsmap() -> dict[str, str]:
    return dict(NS)


def _build_blank_version_part() -> VersionPart:
    root = etree.Element(qname("hv", "HCFVersion"), nsmap={"hv": NS["hv"]})
    root.set("tagetApplication", "WORDPROCESSOR")
    root.set("major", "5")
    root.set("minor", "1")
    root.set("micro", "1")
    root.set("buildNumber", "0")
    root.set("os", "1")
    root.set("xmlVersion", "1.5")
    root.set("application", "Hancom Office Hangul")
    root.set("appVersion", "12, 0, 0, 535 WIN32LEWindows_10")
    return cast(VersionPart, VersionPart.from_root("version.xml", root))

def _build_blank_container_part() -> ContainerPart:
    root = etree.Element(
        qname("ocf", "container"),
        nsmap={"ocf": NS["ocf"], "hpf": NS["hpf"]},
    )
    rootfiles = etree.SubElement(root, qname("ocf", "rootfiles"))
    rootfile = etree.SubElement(rootfiles, qname("ocf", "rootfile"))
    rootfile.set("full-path", "Contents/content.hpf")
    rootfile.set("media-type", "application/hwpml-package+xml")
    return cast(ContainerPart, ContainerPart.from_root("META-INF/container.xml", root))

def _build_blank_content_hpf_part() -> ContentHpfPart:
    root = etree.Element(qname("opf", "package"), nsmap=_full_nsmap())
    root.set("version", "")
    root.set("unique-identifier", "")
    root.set("id", "")

    metadata = etree.SubElement(root, qname("opf", "metadata"))
    etree.SubElement(metadata, qname("opf", "title"))
    language = etree.SubElement(metadata, qname("opf", "language"))
    language.text = "ko"

    manifest = etree.SubElement(root, qname("opf", "manifest"))
    for item_id, href in (
        ("header", "Contents/header.xml"),
        ("section0", "Contents/section0.xml"),
        ("settings", "settings.xml"),
    ):
        item = etree.SubElement(manifest, qname("opf", "item"))
        item.set("id", item_id)
        item.set("href", href)
        item.set("media-type", "application/xml")

    spine = etree.SubElement(root, qname("opf", "spine"))
    for item_id in ("header", "section0"):
        itemref = etree.SubElement(spine, qname("opf", "itemref"))
        itemref.set("idref", item_id)
        itemref.set("linear", "yes")
    return cast(ContentHpfPart, ContentHpfPart.from_root("Contents/content.hpf", root))

def _build_blank_header_part() -> HeaderPart:
    root = etree.Element(qname("hh", "head"), nsmap=_full_nsmap())
    root.set("version", "1.5")
    root.set("secCnt", "1")

    begin_num = etree.SubElement(root, qname("hh", "beginNum"))
    begin_num.set("page", "1")
    begin_num.set("footnote", "1")
    begin_num.set("endnote", "1")
    begin_num.set("pic", "1")
    begin_num.set("tbl", "1")
    begin_num.set("equation", "1")

    ref_list = etree.SubElement(root, qname("hh", "refList"))

    char_properties = etree.SubElement(ref_list, qname("hh", "charProperties"))
    char_properties.set("itemCnt", "1")
    char_pr = etree.SubElement(char_properties, qname("hh", "charPr"))
    char_pr.set("id", "0")
    char_pr.set("height", "1000")
    char_pr.set("textColor", "#000000")
    char_pr.set("shadeColor", "none")
    char_pr.set("useFontSpace", "0")
    char_pr.set("useKerning", "0")
    char_pr.set("symMark", "NONE")
    font_ref = etree.SubElement(char_pr, qname("hh", "fontRef"))
    for key in ("hangul", "latin", "hanja", "japanese", "other", "symbol", "user"):
        font_ref.set(key, "0")
    ratio = etree.SubElement(char_pr, qname("hh", "ratio"))
    spacing = etree.SubElement(char_pr, qname("hh", "spacing"))
    rel_sz = etree.SubElement(char_pr, qname("hh", "relSz"))
    offset = etree.SubElement(char_pr, qname("hh", "offset"))
    for node in (ratio, rel_sz):
        for key in ("hangul", "latin", "hanja", "japanese", "other", "symbol", "user"):
            node.set(key, "100")
    for node in (spacing, offset):
        for key in ("hangul", "latin", "hanja", "japanese", "other", "symbol", "user"):
            node.set(key, "0")
    underline = etree.SubElement(char_pr, qname("hh", "underline"))
    underline.set("type", "NONE")
    underline.set("shape", "SOLID")
    underline.set("color", "#000000")
    strikeout = etree.SubElement(char_pr, qname("hh", "strikeout"))
    strikeout.set("shape", "NONE")
    strikeout.set("color", "#000000")
    outline = etree.SubElement(char_pr, qname("hh", "outline"))
    outline.set("type", "NONE")
    shadow = etree.SubElement(char_pr, qname("hh", "shadow"))
    shadow.set("type", "NONE")
    shadow.set("color", "#C0C0C0")
    shadow.set("offsetX", "10")
    shadow.set("offsetY", "10")

    para_properties = etree.SubElement(ref_list, qname("hh", "paraProperties"))
    para_properties.set("itemCnt", "1")
    para_pr = etree.SubElement(para_properties, qname("hh", "paraPr"))
    para_pr.set("id", "0")
    para_pr.set("tabPrIDRef", "0")
    para_pr.set("condense", "0")
    para_pr.set("fontLineHeight", "0")
    para_pr.set("snapToGrid", "1")
    para_pr.set("suppressLineNumbers", "0")
    para_pr.set("checked", "0")
    align = etree.SubElement(para_pr, qname("hh", "align"))
    align.set("horizontal", "JUSTIFY")
    align.set("vertical", "BASELINE")
    heading = etree.SubElement(para_pr, qname("hh", "heading"))
    heading.set("type", "NONE")
    heading.set("idRef", "0")
    heading.set("level", "0")
    break_setting = etree.SubElement(para_pr, qname("hh", "breakSetting"))
    break_setting.set("breakLatinWord", "KEEP_WORD")
    break_setting.set("breakNonLatinWord", "KEEP_WORD")
    break_setting.set("widowOrphan", "0")
    break_setting.set("keepWithNext", "0")
    break_setting.set("keepLines", "0")
    break_setting.set("pageBreakBefore", "0")
    break_setting.set("lineWrap", "BREAK")
    auto_spacing = etree.SubElement(para_pr, qname("hh", "autoSpacing"))
    auto_spacing.set("eAsianEng", "0")
    auto_spacing.set("eAsianNum", "0")
    margin = etree.SubElement(para_pr, qname("hh", "margin"))
    for key in ("intent", "left", "right", "prev", "next"):
        node = etree.SubElement(margin, qname("hc", key))
        node.set("value", "0")
        node.set("unit", "HWPUNIT")
    line_spacing = etree.SubElement(para_pr, qname("hh", "lineSpacing"))
    line_spacing.set("type", "PERCENT")
    line_spacing.set("value", "160")
    line_spacing.set("unit", "HWPUNIT")

    styles = etree.SubElement(ref_list, qname("hh", "styles"))
    styles.set("itemCnt", "1")
    style = etree.SubElement(styles, qname("hh", "style"))
    style.set("id", "0")
    style.set("type", "PARA")
    style.set("name", "Normal")
    style.set("engName", "Normal")
    style.set("paraPrIDRef", "0")
    style.set("charPrIDRef", "0")
    style.set("nextStyleIDRef", "0")
    style.set("langID", "1042")
    style.set("lockForm", "0")
    return cast(HeaderPart, HeaderPart.from_root("Contents/header.xml", root))

def _build_blank_settings_part() -> SettingsPart:
    root = etree.Element(
        qname("ha", "HWPApplicationSetting"),
        nsmap={"ha": NS["ha"], "config": NS["config"]},
    )
    caret = etree.SubElement(root, qname("ha", "CaretPosition"))
    caret.set("listIDRef", "0")
    caret.set("paraIDRef", "0")
    caret.set("pos", "0")
    return cast(SettingsPart, SettingsPart.from_root("settings.xml", root))

def _build_blank_section_part() -> SectionPart:
    root = etree.Element(qname("hs", "sec"), nsmap=_full_nsmap())
    paragraph = etree.SubElement(root, qname("hp", "p"))
    paragraph.set("id", "0")
    paragraph.set("paraPrIDRef", "0")
    paragraph.set("styleIDRef", "0")
    paragraph.set("pageBreak", "0")
    paragraph.set("columnBreak", "0")
    paragraph.set("merged", "0")

    sec_run = etree.SubElement(paragraph, qname("hp", "run"))
    sec_run.set("charPrIDRef", "0")
    sec_pr = etree.SubElement(sec_run, qname("hp", "secPr"))
    sec_pr.set("id", "")
    sec_pr.set("textDirection", "HORIZONTAL")
    sec_pr.set("spaceColumns", "1134")
    sec_pr.set("tabStop", "8000")
    sec_pr.set("tabStopVal", "4000")
    sec_pr.set("tabStopUnit", "HWPUNIT")
    sec_pr.set("outlineShapeIDRef", "0")
    sec_pr.set("memoShapeIDRef", "0")
    sec_pr.set("textVerticalWidthHead", "0")
    sec_pr.set("masterPageCnt", "0")
    grid = etree.SubElement(sec_pr, qname("hp", "grid"))
    grid.set("lineGrid", "0")
    grid.set("charGrid", "0")
    grid.set("wonggojiFormat", "0")
    start_num = etree.SubElement(sec_pr, qname("hp", "startNum"))
    start_num.set("pageStartsOn", "BOTH")
    start_num.set("page", "0")
    start_num.set("pic", "0")
    start_num.set("tbl", "0")
    start_num.set("equation", "0")
    visibility = etree.SubElement(sec_pr, qname("hp", "visibility"))
    visibility.set("hideFirstHeader", "0")
    visibility.set("hideFirstFooter", "0")
    visibility.set("hideFirstMasterPage", "0")
    visibility.set("border", "SHOW_ALL")
    visibility.set("fill", "SHOW_ALL")
    visibility.set("hideFirstPageNum", "0")
    visibility.set("hideFirstEmptyLine", "0")
    visibility.set("showLineNumber", "0")
    line_number_shape = etree.SubElement(sec_pr, qname("hp", "lineNumberShape"))
    line_number_shape.set("restartType", "0")
    line_number_shape.set("countBy", "0")
    line_number_shape.set("distance", "0")
    line_number_shape.set("startNumber", "0")
    page_pr = etree.SubElement(sec_pr, qname("hp", "pagePr"))
    page_pr.set("landscape", "WIDELY")
    page_pr.set("width", "59528")
    page_pr.set("height", "84186")
    page_pr.set("gutterType", "LEFT_ONLY")
    page_margin = etree.SubElement(page_pr, qname("hp", "margin"))
    page_margin.set("header", "4252")
    page_margin.set("footer", "4252")
    page_margin.set("gutter", "0")
    page_margin.set("left", "8504")
    page_margin.set("right", "8504")
    page_margin.set("top", "5668")
    page_margin.set("bottom", "4252")

    text_run = etree.SubElement(paragraph, qname("hp", "run"))
    text_run.set("charPrIDRef", "0")
    etree.SubElement(text_run, qname("hp", "t"))
    return cast(SectionPart, SectionPart.from_root("Contents/section0.xml", root))

def _build_blank_document_state() -> tuple[dict[str, HwpxPart], list[str], dict[str, zipfile.ZipInfo], list[str]]:
    parts: dict[str, HwpxPart] = {
        "mimetype": MimetypePart("mimetype", b"application/hwp+zip"),
        "version.xml": _build_blank_version_part(),
        "Contents/content.hpf": _build_blank_content_hpf_part(),
        "Contents/header.xml": _build_blank_header_part(),
        "Contents/section0.xml": _build_blank_section_part(),
        "settings.xml": _build_blank_settings_part(),
        "META-INF/container.xml": _build_blank_container_part(),
    }
    part_order = [
        "mimetype",
        "version.xml",
        "Contents/content.hpf",
        "Contents/header.xml",
        "Contents/section0.xml",
        "settings.xml",
        "META-INF/container.xml",
    ]
    return parts, part_order, {}, []

