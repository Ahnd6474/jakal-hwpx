"""Native Hancom numbering and automatic fields for Word-origin documents."""
from __future__ import annotations

from copy import deepcopy
import json
import re
from typing import Any

from lxml import etree

from ..._hancom_models import Field, Table
from ...namespaces import NS, qname

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NUMBER_FORMATS = {"decimal": "DIGIT", "upperRoman": "ROMAN_CAPITAL", "lowerRoman": "ROMAN_SMALL",
                  "upperLetter": "LATIN_CAPITAL", "lowerLetter": "LATIN_SMALL", "decimalEnclosedCircle": "CIRCLED_DIGIT",
                  "koreanCounting": "HANGUL_SYLLABLE", "ganada": "HANGUL_SYLLABLE", "chosung": "HANGUL_JAMO",
                  "chineseCounting": "IDEOGRAPH", "ideographDigital": "IDEOGRAPH"}


def native_field(instruction: str, display: str = "") -> dict[str, Any] | None:
    """Return verified HWPX field types/parameters, keeping Word syntax in the name."""
    parts = re.findall(r'"[^"]*"|\S+', instruction)
    if not parts:
        return None
    kind = parts[0].upper()
    if kind == "MERGEFIELD" and len(parts) > 1 and not any(flag in parts for flag in (r"\b", r"\f")):
        name = parts[1].strip('"')
        return {"field_type": "MAILMERGE", "editable": False, "dirty": False,
                "parameters": {"Fiexde": 1, "Prop": 8, "Command": name, "FieldType": "USER_DEFINE", "FieldValue": name}}
    if kind == "=":
        expression = re.split(r"\s+\\(?:#|\*)", instruction, maxsplit=1)[0].strip()
        expression = re.sub(r'"[^"]*"|\s+', lambda match: match[0] if match[0].startswith('"') else "", expression)
        picture = re.search(r'\\#\s+(?:"([^"]*)"|(\S+))', instruction)
        result_format = "%g"
        if picture:
            value = next(item for item in picture.groups() if item is not None).split(";")[0]
            match = re.fullmatch(r"([^0#]*)([0#,]+)(?:\.([0#]+))?([^0#]*)", value)
            if match is None:
                return None
            prefix, integer, fraction, suffix = match.groups()
            result_format = prefix.replace("%", "%%") + (f"%.{len(fraction)}f" if fraction else "%.0f") + ("," if "," in integer else "") + suffix.replace("%", "%%")
        last = display if re.fullmatch(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?", display) else "0"
        return {"field_type": "FORMULA", "editable": False, "dirty": False,
                "parameters": {"Prop": 8, "Command": expression + "??" + result_format + ";;" + last,
                               "Formula": expression, "ResultFormat": result_format, "LastResult": last}}
    summary = {"TITLE": "$title", "SUBJECT": "$subject", "AUTHOR": "$author", "KEYWORDS": "$keywords",
               "COMMENTS": "$comments", "LASTSAVEDBY": "$lastsaveby"}
    if kind == "DOCPROPERTY" and len(parts) > 1:
        token = summary.get(parts[1].strip('"').upper())
    else:
        token = summary.get(kind)
    if token:
        return {"field_type": "SUMMERY", "editable": False, "dirty": False,
                "parameters": {"Prop": 8, "Command": token, "Property": token}}
    if kind in {"REF", "PAGEREF"} and len(parts) > 1:
        target = parts[1].strip('"')
        if not target or any(char in target for char in ";?\r\n") or any(switch in parts for switch in (r"\n", r"\r", r"\w")):
            return None
        content = 0 if kind == "PAGEREF" else 3 if r"\p" in parts else 2
        link = r"\h" in parts
        command = f"?{target};6;{content};{int(link)};0;"
        return {"field_type": "CROSSREF", "editable": True, "dirty": False,
                "parameters": {"Fiexde": 1, "Prop": 0, "Command": command, "RefPath": f"?{target};",
                    "RefType": "TARGET_BOOKMARK", "RefContentType": {0: "OBJECT_TYPE_PAGE", 2: "OBJECT_TYPE_CONTENTS", 3: "OBJECT_TYPE_UPDOWNPOS"}[content],
                    "RefHyperLink": "true" if link else "false", "RefOpenType": "HWPHYPERLINK_JUMP_CURRENTTAB"}}
    if kind in {"DATE", "TIME"}:
        format_match = re.search(r'\\@\s+(?:"([^"]*)"|(\S+))', instruction)
        # Word's implicit format depends on its locale/settings. In particular,
        # an unformatted TIME must never turn into an ISO calendar date.
        pattern = next((value for value in format_match.groups() if value is not None), "") if format_match else ""
        known = {"yyyy-MM-dd": (r"\:1-02-03;0;", "YYYY-MM-DD"),
                 "yyyy/MM/dd": (r"\:1/02/03;0;", "YYYY/MM/DD")}
        if pattern in known:
            command, date_format = known[pattern]
            return {"field_type": "DATE", "editable": False, "dirty": False,
                    "parameters": {"Prop": 8, "Command": command, "DateNation": "KOR", "DateFormat": date_format}}
    return None


def _definition(parameters):
    root = etree.fromstring(str(parameters["NumberingXML"]).encode("utf-8"))
    abstract, number = root.find(f"{{{W}}}abstractNum"), root.find(f"{{{W}}}num")
    levels = {int(node.get(f"{{{W}}}ilvl", "0")): deepcopy(node) for node in abstract.findall(f"{{{W}}}lvl")}
    for override in number.findall(f"{{{W}}}lvlOverride"):
        level = int(override.get(f"{{{W}}}ilvl", "0"))
        replacement = override.find(f"{{{W}}}lvl")
        if replacement is not None:
            levels[level] = deepcopy(replacement)
        start = override.find(f"{{{W}}}startOverride")
        if start is not None and level in levels:
            previous = levels[level].find(f"{{{W}}}start")
            if previous is not None:
                levels[level].remove(previous)
            levels[level].insert(0, deepcopy(start))
            levels[level][0].tag = f"{{{W}}}start"
    return levels


def activate_lists(document: Any) -> None:
    from ...hancom_document import _hwpx_field_kwargs, _word_effect_from_field_name

    ref_list = document.header.root_element.find(qname("hh", "refList"))
    groups = {}
    styles = {}
    named_styles = {}
    lists = {}
    changed = False
    for field in list(document.fields()):
        restored = _word_effect_from_field_name(field.name)
        if restored is None or restored["kind"] != "JAKAL_LIST_LABEL":
            continue
        parameters = restored["parameters"]
        if parameters.get("NativeList") == "1":
            continue
        raw = str(parameters.get("NumberingXML", ""))
        if not raw:
            continue
        level = int(parameters.get("Level", "0"))
        definitions = _definition(parameters)
        current = definitions.get(level)
        if current is None:
            continue
        format_node = current.find(f"{{{W}}}numFmt")
        kind = format_node.get(f"{{{W}}}val") if format_node is not None else "decimal"
        # Word's zero-padded/locale-specific counters need a separate evaluator;
        # never silently replace their number format with ordinary decimals.
        if kind != "bullet" and kind not in NUMBER_FORMATS:
            continue
        paragraph = field.element.xpath("ancestor::hp:p[1]", namespaces=NS)[0]
        # A Word multilevel definition can mix counters and bullets, and each
        # bullet level can use a different character. HWPX stores those in
        # separate resource tables, so they must not share one resource ID.
        group = (raw, "bullet", level) if kind == "bullet" else (raw, "number")
        if group not in groups:
            tag = "bullets" if kind == "bullet" else "numberings"
            container = ref_list.find(qname("hh", tag))
            if container is None:
                container = etree.SubElement(ref_list, qname("hh", tag), itemCnt="0")
            identity = str(max((int(node.get("id", "0")) for node in container), default=0) + 1)
            if kind == "bullet":
                text = current.find(f"{{{W}}}lvlText")
                char = text.get(f"{{{W}}}val", "•").replace("\uf0b7", "•").replace("\uf0a7", "▪") if text is not None else "•"
                node = etree.SubElement(container, qname("hh", "bullet"), id=identity, char=char[:1], useImage="0")
                etree.SubElement(node, qname("hh", "paraHead"), align="RIGHT", useInstWidth="1", autoIndent="0",
                                 widthAdjust="225", textOffsetType="HWPUNIT", textOffset="900", numFormat="DIGIT", charPrIDRef="4294967295", checkable="0")
            else:
                node = etree.SubElement(container, qname("hh", "numbering"), id=identity, start="1")
                for index in range(10):
                    source = definitions.get(index)
                    fmt = source.find(f"{{{W}}}numFmt") if source is not None else None
                    text = source.find(f"{{{W}}}lvlText") if source is not None else None
                    start = source.find(f"{{{W}}}start") if source is not None else None
                    head = etree.SubElement(node, qname("hh", "paraHead"), start=start.get(f"{{{W}}}val", "1") if start is not None else "1",
                        level=str(index + 1), align="RIGHT", useInstWidth="1", autoIndent="0", widthAdjust="225",
                        textOffsetType="HWPUNIT", textOffset="900", numFormat=NUMBER_FORMATS.get(fmt.get(f"{{{W}}}val") if fmt is not None else "decimal", "DIGIT"),
                        charPrIDRef="4294967295", checkable="0")
                    head.text = re.sub(r"%(\d+)", r"^\1", text.get(f"{{{W}}}val", "^N") if text is not None else "^N")
            container.set("itemCnt", str(len(container)))
            groups[group] = (identity, "BULLET" if kind == "bullet" else "NUMBER")
        identity, heading_kind = groups[group]
        original_style = paragraph.get("paraPrIDRef", "0")
        key = (original_style, identity, heading_kind, level)
        if key not in styles:
            properties = ref_list.find(qname("hh", "paraProperties"))
            original = next(node for node in properties if node.get("id") == original_style)
            clone = deepcopy(original)
            style_id = str(max(int(node.get("id", "0")) for node in properties) + 1)
            clone.set("id", style_id)
            heading = clone.find(qname("hh", "heading"))
            if heading is None:
                heading = etree.SubElement(clone, qname("hh", "heading"))
            heading.attrib.update({"type": heading_kind, "idRef": identity, "level": str(level)})
            properties.append(clone)
            properties.set("itemCnt", str(len(properties)))
            styles[key] = style_id
        paragraph.set("paraPrIDRef", styles[key])
        original_named = paragraph.get("styleIDRef", "0")
        named_key = (original_named, styles[key])
        if named_key not in named_styles:
            container = ref_list.find(qname("hh", "styles"))
            original = next((node for node in container if node.get("id") == original_named), None)
            if original is not None:
                clone = deepcopy(original)
                named_id = str(max(int(node.get("id", "0")) for node in container) + 1)
                clone.set("id", named_id)
                clone.set("paraPrIDRef", styles[key])
                if clone.get("nextStyleIDRef") == original_named:
                    clone.set("nextStyleIDRef", named_id)
                container.append(clone)
                container.set("itemCnt", str(len(container)))
                named_styles[named_key] = named_id
        if named_key in named_styles:
            paragraph.set("styleIDRef", named_styles[named_key])
        # The native paragraph heading draws the label. The field is an empty
        # identity record, so no duplicate static number is left in the text.
        field.set_display_text("")
        parameters.update(NativeList="1", NativeNumberingID=identity)
        record = Field("JAKAL_LIST_LABEL", "", name=restored["name"], parameters=parameters)
        field.element.set("name", str(_hwpx_field_kwargs(record)["name"]))
        lists[heading_kind + ":" + identity] = dict(parameters)
        field.section.mark_modified()
        changed = True
    if changed:
        # Persist definitions even if every original list item is deleted later.
        for field in document.fields():
            record = _word_effect_from_field_name(field.name)
            if record is None or record["kind"] != "JAKAL_WORD_FEATURE":
                continue
            data = json.loads(str(record["parameters"].get("Record", "{}")))
            if data.get("kind") == "resources":
                data["properties"]["native_lists"] = lists
                record["parameters"]["Record"] = json.dumps(data)
                field.element.set("name", str(_hwpx_field_kwargs(Field(record["kind"], name=record["name"], parameters=record["parameters"]))["name"]))
                field.section.mark_modified()
        document.header.mark_modified()


def restore_list_styles(document: Any) -> None:
    definitions = dict(getattr(document, "_word_resources", {}).get("native_lists", {}))
    styles = {style.style_id: style for style in document.paragraph_styles}

    def collect(blocks):
        for block in blocks:
            if isinstance(block, Field) and block.parameters.get("NativeList") == "1":
                style = styles.get(getattr(block, "_word_native_para_pr_id", None))
                if style is not None:
                    definitions[style.heading.get("type", "NUMBER") + ":" + style.heading.get("idRef", "")] = dict(block.parameters)
            for nested in getattr(block, "nested_blocks", []):
                collect([nested])
            if isinstance(block, Table):
                for baseline, values in getattr(block, "_docx_cell_content", {}).values():
                    collect(values)
    for section in document.sections:
        collect([*section.blocks, *section.header_footer_blocks])
    for style in document.paragraph_styles:
        definition = definitions.get(style.heading.get("type", "") + ":" + style.heading.get("idRef", ""))
        if definition:
            value = {**definition, "Level": style.heading.get("level", definition.get("Level", "0"))}
            setattr(style, "_word_native_list", value)
