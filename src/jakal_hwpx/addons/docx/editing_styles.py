"""Style catalogue, inheritance and theme edits without rebuilding a document."""
from __future__ import annotations

from .editing import _font, _paragraph_format, _tag, _color
from .editing_navigation import containers


def create_style(editor, name, *, kind="paragraph", base=None, next_style=None, font=None, paragraph=None):
    from docx.enum.style import WD_STYLE_TYPE
    editor._check()
    kinds = {"paragraph": WD_STYLE_TYPE.PARAGRAPH, "character": WD_STYLE_TYPE.CHARACTER, "table": WD_STYLE_TYPE.TABLE}
    styles = editor._backend.styles
    if not isinstance(name, str) or not name.strip() or name in styles or kind not in kinds:
        raise ValueError("Use a unique, non-empty style name and paragraph/character/table kind.")
    _font(None, font or {})
    _paragraph_format(None, paragraph or {})
    if paragraph and kind == "character":
        raise ValueError("Character styles cannot contain paragraph formatting.")
    for value in (base, next_style):
        if value is not None and (value not in styles or styles[value].type != kinds[kind]):
            raise ValueError("Related styles must exist and have the same type.")
    if next_style is not None and kind != "paragraph":
        raise ValueError("Only paragraph styles have a next style.")
    target = styles.add_style(name, kinds[kind])
    if base is not None:
        target.base_style = styles[base]
    if next_style is not None:
        target.next_paragraph_style = styles[next_style]
    _font(target, font or {})
    if paragraph:
        _paragraph_format(target, paragraph)
    return name


def configure_style(editor, name, **values):
    editor._check()
    if values.keys() - {"base", "next_style"}:
        raise ValueError("Style relations are base and next_style.")
    styles = editor._backend.styles
    target = styles[name]
    for key, value in values.items():
        if value is not None and (value not in styles or styles[value].type != target.type):
            raise ValueError("Related styles must exist and have the same type.")
        if key == "next_style" and target.type != 1:
            raise ValueError("Only paragraph styles have a next style.")
        if key == "base" and value is not None:
            cursor, seen = styles[value], set()
            while cursor is not None:
                if cursor.style_id == target.style_id or cursor.style_id in seen:
                    raise ValueError("Style inheritance must not contain a cycle.")
                seen.add(cursor.style_id)
                cursor = cursor.base_style
    for key, value in values.items():
        if key == "base":
            target.base_style = None if value is None else styles[value]
        else:
            target.next_paragraph_style = None if value is None else styles[value]


def delete_style(editor, name, *, replacement):
    editor._check()
    styles = editor._backend.styles
    old, new = styles[name], styles[replacement]
    if old.style_id == new.style_id or old.type != new.type:
        raise ValueError("Replacement must be a different style of the same type.")
    if old.element.get(_tag("default")) in {"1", "true"}:
        raise ValueError("A document's default style cannot be deleted.")
    # Reparent dependent styles before rewriting their XML references.
    cursor, seen = new, set()
    while cursor is not None:
        if cursor.style_id == old.style_id or cursor.style_id in seen:
            raise ValueError("Replacement must not inherit the deleted style.")
        seen.add(cursor.style_id)
        cursor = cursor.base_style
    owners = list(containers(editor))
    roots = [part.element for part in editor._backend.part.package.parts if getattr(part, "element", None) is not None] + [owner._root for owner in owners]
    for root in roots:
        for node in root.iter():
            if node.tag in {_tag(tag) for tag in ("pStyle", "rStyle", "tblStyle", "basedOn", "next", "link")} and node.get(_tag("val")) == old.style_id:
                node.set(_tag("val"), new.style_id)
    old.delete()


def format_theme(editor, *, colors=None, major_fonts=None, minor_fonts=None):
    from lxml import etree
    editor._check()
    part = next((rel.target_part for rel in editor._backend.part.rels.values() if not rel.is_external and rel.reltype.endswith("/theme")), None)
    if part is None:
        raise ValueError("Document has no theme part.")
    a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    names = {"dark1": "dk1", "light1": "lt1", "dark2": "dk2", "light2": "lt2", "hyperlink": "hlink", "followed_hyperlink": "folHlink", **{f"accent{i}": f"accent{i}" for i in range(1, 7)}}
    colors = colors or {}
    if colors.keys() - names.keys():
        raise ValueError("Unknown theme color slot.")
    converted = {names[key]: _color(value) for key, value in colors.items()}
    for fonts in (major_fonts, minor_fonts):
        if fonts is not None and (fonts.keys() - {"latin", "eastAsia", "complexScript"} or any(not isinstance(value, str) or not value for value in fonts.values())):
            raise ValueError("Theme fonts use latin, eastAsia and complexScript.")
    root = etree.fromstring(part.blob)
    for key, value in converted.items():
        node = root.find(".//" + a + "clrScheme/" + a + key)
        if node is None:
            raise ValueError("Theme is missing a required color slot.")
        for child in list(node):
            node.remove(child)
        etree.SubElement(node, a + "srgbClr", val=value)
    for name, fonts in (("majorFont", major_fonts), ("minorFont", minor_fonts)):
        for script, value in (fonts or {}).items():
            tag = {"latin": "latin", "eastAsia": "ea", "complexScript": "cs"}[script]
            node = root.find(".//" + a + name + "/" + a + tag)
            if node is None:
                raise ValueError("Theme is missing a required font slot.")
            node.set("typeface", value)
    part._blob = etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
