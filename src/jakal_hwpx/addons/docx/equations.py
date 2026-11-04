"""Bidirectional native equation mapping for the supported math grammar.

Unsupported structures fail explicitly instead of concatenating math leaves,
which changes fractions, radicals and scripts into a different expression.
"""
from __future__ import annotations

import re
from lxml import etree

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NARY = {"sum": "∑", "prod": "∏", "coprod": "∐", "int": "∫", "oint": "∮", "iint": "∬", "iiint": "∭", "bigcup": "⋃", "bigcap": "⋂"}
SEMANTIC_PROPERTIES = {"naryPr": {"chr", "subHide", "supHide"}, "dPr": {"begChr", "endChr", "sepChr"},
                       "fPr": {"type"}, "mPr": {"mcs"}, "rPr": {"aln"}}


def _tag(name: str) -> str:
    return f"{{{M}}}{name}"


def _children(node):
    return [child for child in node if not etree.QName(child).localname.endswith("Pr")]


def _shape(node):
    return [node.tag, *[tag for child in _children(node) for tag in _shape(child)]]


def _projection(node):
    """Describe reversible display substitutions without retaining operands."""
    from copy import deepcopy
    projected = deepcopy(node)
    mapped = []
    for item in reversed(list(projected.iter())):
        name = etree.QName(item).localname
        if name in {"limLow", "limUpp"}:
            base, limit = item.find(_tag("e")), item.find(_tag("lim"))
            if base is None or limit is None or name == "limLow" and "".join(base.itertext(_tag("t"))).strip() == "lim":
                continue
            for child in list(item):
                item.remove(child)
            item.tag = _tag("eqArr")
            limit.tag = _tag("e")
            item.extend([base, limit] if name == "limLow" else [limit, base])
            mapped.append((item, name))
        elif name == "d" and len(item.findall(_tag("e"))) > 1:
            bodies = item.findall(_tag("e"))
            separator = item.find("./" + _tag("dPr") + "/" + _tag("sepChr"))
            symbol = separator.get(_tag("val"), "|") if separator is not None else "|"
            body = etree.Element(_tag("e"))
            matrix = etree.SubElement(body, _tag("m"))
            row = etree.SubElement(matrix, _tag("mr"))
            for index, entry in enumerate(bodies):
                if index:
                    cell = etree.SubElement(row, _tag("e"))
                    etree.SubElement(etree.SubElement(cell, _tag("r")), _tag("t")).text = symbol
                row.append(entry)
            item.append(body)
            mapped.append((item, "delimiter_arguments"))
    records = []
    for item, kind in mapped:
        path, child = [], item
        while child is not projected:
            parent = child.getparent()
            path.insert(0, _children(parent).index(child))
            child = parent
        records.append({"path": path, "kind": kind})
    return projected, records


def _restore_projection(node, record):
    from copy import deepcopy
    if not record.get("mappings") or _shape(node) != record.get("projection_shape"):
        return
    work = deepcopy(node)
    for mapping in sorted(record["mappings"], key=lambda value: len(value["path"]), reverse=True):
        item = work
        for index in mapping["path"]:
            item = _children(item)[index]
        kind = mapping["kind"]
        if kind in {"limLow", "limUpp"}:
            entries = item.findall(_tag("e"))
            if item.tag != _tag("eqArr") or len(entries) != 2:
                return
            base, limit = entries if kind == "limLow" else entries[::-1]
            item.tag = _tag(kind)
            for child in list(item):
                item.remove(child)
            limit.tag = _tag("lim")
            item.extend([base, limit])
        else:
            bodies = item.findall(_tag("e"))
            if item.tag != _tag("d") or len(bodies) != 1:
                return
            matrices = bodies[0].findall(_tag("m"))
            if len(matrices) != 1 or len(matrices[0].findall(_tag("mr"))) != 1:
                return
            cells = matrices[0].find(_tag("mr")).findall(_tag("e"))
            separators = ["".join(cell.itertext(_tag("t"))) for cell in cells[1::2]]
            if len(cells) < 3 or len(cells) % 2 != 1 or len(set(separators)) != 1 or len(separators[0]) != 1:
                return
            item.remove(bodies[0])
            item.extend(cells[::2])
            properties = item.find(_tag("dPr"))
            if properties is None:
                properties = etree.Element(_tag("dPr"))
                item.insert(0, properties)
            separator = properties.find(_tag("sepChr"))
            if separator is None:
                separator = etree.SubElement(properties, _tag("sepChr"))
            separator.set(_tag("val"), separators[0])
    if _shape(work) == record.get("shape"):
        node[:] = list(work)


def capture_properties(node):
    """Capture only appearance properties, never formula text or operands."""
    from copy import deepcopy
    shape, properties = [], []
    def visit(item, path):
        shape.append(item.tag)
        logical = 0
        for position, child in enumerate(item):
            if etree.QName(child).localname.endswith("Pr"):
                copy = deepcopy(child)
                for semantic in list(copy):
                    if etree.QName(semantic).localname in SEMANTIC_PROPERTIES.get(etree.QName(copy).localname, set()):
                        copy.remove(semantic)
                properties.append({"path": path, "position": position, "xml": etree.tostring(copy, encoding="unicode")})
            else:
                visit(child, path + [logical])
                logical += 1
    visit(node, [])
    projected, mappings = _projection(node)
    return {"shape": shape, "properties": properties, "mappings": mappings,
            "projection_shape": _shape(projected) if mappings else []}


def restore_properties(node, record):
    if record:
        _restore_projection(node, record)
    if not record or capture_properties(node)["shape"] != record.get("shape"):
        return node
    for value in record.get("properties", []):
        parent = node
        for index in value["path"]:
            parent = [child for child in parent if not etree.QName(child).localname.endswith("Pr")][index]
        property_node = etree.fromstring(value["xml"].encode("utf-8"))
        semantics = SEMANTIC_PROPERTIES.get(etree.QName(property_node).localname, set())
        for child in list(property_node):
            if etree.QName(child).localname in semantics:
                property_node.remove(child)
        for previous in list(parent.findall(property_node.tag)):
            from copy import deepcopy
            for child in reversed(list(previous)):
                if etree.QName(child).localname in semantics:
                    property_node.insert(0, deepcopy(child))
            parent.remove(previous)
        order = {"naryPr": ["chr", "limLoc", "grow", "subHide", "supHide", "ctrlPr"],
                 "dPr": ["begChr", "sepChr", "endChr", "grow", "shp", "ctrlPr"],
                 "mPr": ["baseJc", "plcHide", "rSpRule", "cGpRule", "rSp", "cSp", "cGp", "mcs", "ctrlPr"]}.get(etree.QName(property_node).localname)
        if order:
            children = list(property_node)
            for child in children:
                property_node.remove(child)
            property_node.extend(sorted(children, key=lambda child: order.index(etree.QName(child).localname) if etree.QName(child).localname in order else len(order)))
        parent.insert(min(value["position"], len(parent)), property_node)
    return node


def omml_to_hwp(node: etree._Element) -> str:
    name = etree.QName(node).localname

    def argument(tag: str) -> str:
        child = node.find(_tag(tag))
        return "" if child is None else omml_to_hwp(child)

    if name in {"oMath", "oMathPara", "num", "den", "e", "sup", "sub", "deg", "lim"}:
        return "".join(omml_to_hwp(child) for child in node if not etree.QName(child).localname.endswith("Pr"))
    if name == "r":
        text = "".join(child.text or "" for child in node if child.tag == _tag("t"))
        if re.search(r"[{}^_\\]", text):
            raise ValueError("Unsupported literal math operator in Word equation text; use native math structures or an external converter.")
        return text
    if name == "f":
        kind = node.find("./" + _tag("fPr") + "/" + _tag("type"))
        style = kind.get(_tag("val"), "bar") if kind is not None else "bar"
        if style not in {"bar", "noBar"}:
            raise ValueError("Unsupported Word fraction layout: " + style)
        return "{" + argument("num") + ("} atop {" if style == "noBar" else "} over {") + argument("den") + "}"
    if name in {"sSup", "sSub", "sSubSup"}:
        value = "{" + argument("e") + "}"
        if name != "sSup":
            value += " _ {" + argument("sub") + "}"
        if name != "sSub":
            value += " ^ {" + argument("sup") + "}"
        return value
    if name == "rad":
        degree = argument("deg")
        return ("root {" + degree + "} of " if degree else "sqrt ") + "{" + argument("e") + "}"
    if name == "m":
        rows = node.findall(_tag("mr"))
        counts = [len(row.findall(_tag("e"))) for row in rows]
        if not counts or not counts[0] or len(set(counts)) != 1:
            raise ValueError("Unsupported Word equation structure m:m with a ragged/empty matrix.")
        alignments = []
        for column in node.findall("./" + _tag("mPr") + "/" + _tag("mcs") + "/" + _tag("mc")):
            count = column.find("./" + _tag("mcPr") + "/" + _tag("count"))
            alignment = column.find("./" + _tag("mcPr") + "/" + _tag("mcJc"))
            alignments.extend([alignment.get(_tag("val"), "center") if alignment is not None else "center"] * (int(count.get(_tag("val"), "1")) if count is not None else 1))
        alignments = (alignments + ["center"] * counts[0])[:counts[0]]
        if any(value not in {"center", "left", "right"} for value in alignments):
            raise ValueError("Unsupported matrix column alignment.")
        if counts[0] == 1 and alignments[0] != "center":
            return ("lpile" if alignments[0] == "left" else "rpile") + " {" + " # ".join(omml_to_hwp(row.findall(_tag("e"))[0]) for row in rows) + "}"
        if any(value != "center" for value in alignments):
            return "matrix {" + " ".join({"left": "lcol", "right": "rcol", "center": "col"}[align] + " {" + " # ".join(omml_to_hwp(row.findall(_tag("e"))[index]) for row in rows) + "}" for index, align in enumerate(alignments)) + "}"
        return "matrix {" + " # ".join(" & ".join(omml_to_hwp(cell) for cell in row.findall(_tag("e"))) for row in rows) + "}"
    if name == "eqArr":
        rows = node.findall(_tag("e"))
        if not rows:
            raise ValueError("An equation array needs at least one row.")
        aligned = any(item.find("./" + _tag("rPr") + "/" + _tag("aln")) is not None for row in rows for item in row)
        if aligned:
            lines = []
            for row in rows:
                text = ""
                for item in _children(row):
                    alignment = item.find("./" + _tag("rPr") + "/" + _tag("aln"))
                    if alignment is not None and alignment.get(_tag("val"), "1") not in {"0", "false", "off"}:
                        text += " & "
                    text += omml_to_hwp(item)
                lines.append(text)
            return "eqalign {" + " # ".join(lines) + "}"
        return "pile {" + " # ".join(omml_to_hwp(row) for row in rows) + "}"
    if name == "nary":
        char = node.find("./" + _tag("naryPr") + "/" + _tag("chr"))
        symbol = char.get(_tag("val")) if char is not None else "∫"
        command = next((key for key, value in NARY.items() if value == symbol), None)
        if command is None or node.find(_tag("e")) is None:
            raise ValueError("Unsupported Word equation structure m:nary without a supported operator/body.")
        value = command
        for name, op in (("sub", "_"), ("sup", "^")):
            content = argument(name)
            hidden = node.find("./" + _tag("naryPr") + "/" + _tag(name + "Hide"))
            if content and hidden is not None and hidden.get(_tag("val"), "1") in {"1", "true", "on"}:
                raise ValueError("Unsupported hidden non-empty Word operator limit.")
            if content:
                value += " " + op + " {" + content + "}"
        return value + " {" + argument("e") + "}"
    if name == "d":
        bodies = node.findall(_tag("e"))
        if not bodies:
            raise ValueError("A Word delimiter needs at least one argument.")
        delimiters = []
        for tag, default in (("begChr", "("), ("endChr", ")")):
            prop = node.find("./" + _tag("dPr") + "/" + _tag(tag))
            delimiters.append(prop.get(_tag("val"), default) if prop is not None else default)
        if any(value not in {"", "(", ")", "[", "]", "{", "}", "|", "‖", "⌈", "⌉", "⌊", "⌋"} for value in delimiters):
            raise ValueError("Unsupported Word math delimiter.")
        body = argument("e")
        if len(bodies) > 1:
            prop = node.find("./" + _tag("dPr") + "/" + _tag("sepChr"))
            separator = prop.get(_tag("val"), "|") if prop is not None else "|"
            if len(separator) != 1 or separator in "{}^_&#\\":
                raise ValueError("Unsupported multi-argument delimiter separator.")
            body = "matrix {" + (" & " + separator + " & ").join("{" + omml_to_hwp(value) + "}" for value in bodies) + "}"
        return "left " + (delimiters[0] or ".") + " {" + body + "} right " + (delimiters[1] or ".")
    if name in {"limLow", "limUpp"}:
        if name == "limLow" and argument("e").strip() == "lim":
            return "lim _ {" + argument("lim") + "}"
        entries = [argument("e"), argument("lim")]
        if name == "limUpp":
            entries.reverse()
        return "pile {" + " # ".join("{" + value + "}" for value in entries) + "}"
    raise ValueError(f"Unsupported Word equation structure m:{name}; use an external converter.")


def hwp_to_omml(script: str) -> etree._Element:
    if re.search(r"\\|\b(color|bold|italic|lbound|upperlim)\b", script, re.IGNORECASE):
        raise ValueError("HWP equations require an external DOCX converter for this unsupported math grammar.")
    keywords = "over|atop|choose|binom|sqrt|root|of|matrix|pmatrix|bmatrix|dmatrix|cases|pile|lpile|rpile|eqalign|col|lcol|rcol|" + "|".join(NARY) + "|lim|left|right|from|to"
    pattern = r"\b(?:" + keywords + r")\b"
    tokens = re.findall(pattern + r"|[{}^_&#()\[\]|‖⌈⌉⌊⌋]|[^{}^_&#()\[\]|‖⌈⌉⌊⌋]+?(?=" + pattern + r"|[{}^_&#()\[\]|‖⌈⌉⌊⌋]|$)", script, re.IGNORECASE)
    cursor = 0

    def run(value: str) -> etree._Element:
        node = etree.Element(_tag("r"))
        etree.SubElement(node, _tag("t")).text = value
        return node

    def container(name: str, children: list[etree._Element]) -> etree._Element:
        node = etree.Element(_tag(name))
        node.extend(children)
        return node

    def whitespace() -> None:
        nonlocal cursor
        while cursor < len(tokens) and not tokens[cursor].strip():
            cursor += 1

    def atom() -> list[etree._Element]:
        nonlocal cursor
        whitespace()
        if cursor >= len(tokens):
            raise ValueError("Incomplete HWP equation; use an external converter.")
        token = tokens[cursor].strip()
        if re.fullmatch(pattern, token, re.IGNORECASE):
            token = token.lower()
        cursor += 1
        if token == "{":
            value = expression()
            whitespace()
            if cursor >= len(tokens) or tokens[cursor].strip() != "}":
                raise ValueError("Unbalanced HWP equation braces.")
            cursor += 1
            return value
        if token in {"sqrt", "root"}:
            degree = []
            if token == "root":
                degree = atom()
                whitespace()
                if cursor >= len(tokens) or tokens[cursor].strip() != "of":
                    raise ValueError("HWP root requires an 'of' argument.")
                cursor += 1
            value = atom()
            radical = etree.Element(_tag("rad"))
            if not degree:
                properties = etree.SubElement(radical, _tag("radPr"))
                etree.SubElement(properties, _tag("degHide")).set(_tag("val"), "1")
            radical.append(container("deg", degree))
            radical.append(container("e", value))
            return [radical]
        if token == "binom":
            numerator, denominator = atom(), atom()
            return [delimiter("(", ")", [fraction(numerator, denominator, "noBar")])]
        if token in {"matrix", "pmatrix", "bmatrix", "dmatrix", "cases", "pile", "lpile", "rpile", "eqalign"}:
            whitespace()
            if cursor >= len(tokens) or tokens[cursor].strip() != "{":
                raise ValueError("A matrix requires a braced row/cell body.")
            cursor += 1
            rows, cells = [], []
            whitespace()
            alignments = []
            if cursor < len(tokens) and tokens[cursor].strip().lower() in {"col", "lcol", "rcol"}:
                columns = []
                while cursor < len(tokens) and tokens[cursor].strip() != "}":
                    column = tokens[cursor].strip().lower()
                    if column not in {"col", "lcol", "rcol"}:
                        raise ValueError("A column matrix requires col/lcol/rcol entries.")
                    alignments.append({"col": "center", "lcol": "left", "rcol": "right"}[column])
                    cursor += 1
                    whitespace()
                    if cursor >= len(tokens) or tokens[cursor].strip() != "{":
                        raise ValueError("Matrix columns require braced rows.")
                    cursor += 1
                    values = []
                    while True:
                        values.append(expression({"#", "}"}))
                        whitespace()
                        if cursor >= len(tokens):
                            raise ValueError("Unbalanced column matrix braces.")
                        separator = tokens[cursor].strip()
                        cursor += 1
                        if separator == "}":
                            break
                    columns.append(values)
                    whitespace()
                if cursor >= len(tokens) or len({len(column) for column in columns}) != 1:
                    raise ValueError("Matrix columns must have equal row counts.")
                cursor += 1
                rows = [list(row) for row in zip(*columns)]
            column_mode = bool(rows)
            while not column_mode:
                cells.append(expression({"&", "#", "}"}))
                whitespace()
                if cursor >= len(tokens):
                    raise ValueError("Unbalanced matrix braces.")
                separator = tokens[cursor].strip()
                cursor += 1
                if separator in {"#", "}"}:
                    rows.append(cells)
                    cells = []
                if separator == "}":
                    break
            if not rows or not rows[0] or len({len(row) for row in rows}) != 1:
                raise ValueError("Matrix rows must have equal column counts.")
            if token == "eqalign":
                array = etree.Element(_tag("eqArr"))
                for row in rows:
                    entry = etree.SubElement(array, _tag("e"))
                    for index, cell in enumerate(row):
                        if index:
                            marker = run("")
                            properties = etree.Element(_tag("rPr"))
                            etree.SubElement(properties, _tag("aln")).set(_tag("val"), "1")
                            marker.insert(0, properties)
                            entry.append(marker)
                        entry.extend(cell)
                return [array]
            if token in {"pile", "lpile", "rpile"}:
                if any(len(row) != 1 for row in rows):
                    raise ValueError("Pile rows take one expression; use matrix for multiple columns.")
                if token == "pile":
                    return [container("eqArr", [container("e", row[0]) for row in rows])]
                alignments = ["left" if token == "lpile" else "right"]
            matrix = container("m", [container("mr", [container("e", cell) for cell in row]) for row in rows])
            if alignments or token == "cases":
                properties = etree.Element(_tag("mPr"))
                matrix.insert(0, properties)
                definitions = etree.SubElement(properties, _tag("mcs"))
                for align in alignments or ["left"] * len(rows[0]):
                    column = etree.SubElement(etree.SubElement(definitions, _tag("mc")), _tag("mcPr"))
                    etree.SubElement(column, _tag("count")).set(_tag("val"), "1")
                    etree.SubElement(column, _tag("mcJc")).set(_tag("val"), align)
            pair = {"pmatrix": ("(", ")"), "bmatrix": ("[", "]"), "dmatrix": ("|", "|"), "cases": ("{", "")}.get(token)
            return [delimiter(*pair, [matrix])] if pair is not None else [matrix]
        if token in NARY or token == "lim":
            limits: dict[str, list[etree._Element]] = {"sub": [], "sup": []}
            whitespace()
            while cursor < len(tokens) and tokens[cursor].strip().lower() in {"_", "^", "from", "to"}:
                op = tokens[cursor].strip().lower()
                cursor += 1
                key = "sub" if op in {"_", "from"} else "sup"
                if limits[key]:
                    raise ValueError("Duplicate operator limit.")
                limits[key] = atom()
                whitespace()
            body = [run("lim")] if token == "lim" else atom()
            if token == "lim":
                key = "sub"
                if limits["sup" if key == "sub" else "sub"]:
                    raise ValueError("A limit function takes only its declared limit side.")
                return [container("limLow" if token == "lim" else "limUpp", [container("e", body), container("lim", limits[key])])]
            node = etree.Element(_tag("nary"))
            pr = etree.SubElement(node, _tag("naryPr"))
            etree.SubElement(pr, _tag("chr")).set(_tag("val"), NARY[token])
            etree.SubElement(pr, _tag("limLoc")).set(_tag("val"), "subSup" if "int" in token else "undOvr")
            for key in ("sub", "sup"):
                etree.SubElement(pr, _tag(key + "Hide")).set(_tag("val"), "0" if limits[key] else "1")
                node.append(container(key, limits[key]))
            node.append(container("e", body))
            return [node]
        if token == "left":
            whitespace()
            if cursor >= len(tokens):
                raise ValueError("Missing opening delimiter.")
            opening = tokens[cursor].strip()
            cursor += 1
            body = expression({"right"})
            whitespace()
            if cursor >= len(tokens) or tokens[cursor].strip().lower() != "right":
                raise ValueError("A left delimiter requires a matching right delimiter.")
            cursor += 1
            whitespace()
            if cursor >= len(tokens):
                raise ValueError("Missing closing delimiter.")
            closing = tokens[cursor].strip()
            cursor += 1
            if opening not in {"(", "[", "{", "|", "‖", "⌈", "⌊", "."} or closing not in {")", "]", "}", "|", "‖", "⌉", "⌋", "."}:
                raise ValueError("Unsupported delimiter pair.")
            node = etree.Element(_tag("d"))
            pr = etree.SubElement(node, _tag("dPr"))
            etree.SubElement(pr, _tag("begChr")).set(_tag("val"), "" if opening == "." else opening)
            etree.SubElement(pr, _tag("endChr")).set(_tag("val"), "" if closing == "." else closing)
            node.append(container("e", body))
            return [node]
        if token in {"}", "over", "atop", "choose", "^", "_", "of", "&", "#", "right", "from", "to", "col", "lcol", "rcol"}:
            raise ValueError("Unexpected HWP equation operator: " + token)
        return [run(tokens[cursor - 1].strip())]

    def expression(stops=None) -> list[etree._Element]:
        nonlocal cursor
        result: list[etree._Element] = []
        while cursor < len(tokens):
            whitespace()
            if cursor >= len(tokens) or tokens[cursor].strip().lower() in (stops or {"}"}):
                break
            value = atom()
            whitespace()
            while cursor < len(tokens) and tokens[cursor].strip() in {"^", "_"}:
                token = tokens[cursor].strip()
                cursor += 1
                argument = atom()
                previous = value[0] if len(value) == 1 else None
                opposite = "sSub" if token == "^" else "sSup"
                if previous is not None and previous.tag == _tag(opposite):
                    node = etree.Element(_tag("sSubSup"))
                    base = previous.find(_tag("e"))
                    assert base is not None
                    node.append(base)
                    sub = previous.find(_tag("sub")) if token == "^" else container("sub", argument)
                    sup = container("sup", argument) if token == "^" else previous.find(_tag("sup"))
                    assert sub is not None and sup is not None
                    node.extend([sub, sup])
                else:
                    node = etree.Element(_tag("sSup" if token == "^" else "sSub"))
                    node.extend([container("e", value), container("sup" if token == "^" else "sub", argument)])
                value = [node]
                whitespace()
            if cursor < len(tokens) and tokens[cursor].strip().lower() in {"over", "atop", "choose"}:
                operator = tokens[cursor].strip().lower()
                cursor += 1
                denominator = atom()
                node = fraction(value, denominator, "bar" if operator == "over" else "noBar")
                value = [delimiter("(", ")", [node]) if operator == "choose" else node]
            result.extend(value)
        return result

    def fraction(numerator, denominator, style):
        node = etree.Element(_tag("f"))
        properties = etree.SubElement(node, _tag("fPr"))
        etree.SubElement(properties, _tag("type")).set(_tag("val"), style)
        node.extend([container("num", numerator), container("den", denominator)])
        return node

    def delimiter(opening, closing, children):
        node = etree.Element(_tag("d"))
        properties = etree.SubElement(node, _tag("dPr"))
        etree.SubElement(properties, _tag("begChr")).set(_tag("val"), opening)
        etree.SubElement(properties, _tag("endChr")).set(_tag("val"), closing)
        node.append(container("e", children))
        return node

    math = etree.Element(_tag("oMath"), nsmap={"m": M})
    math.extend(expression())
    if cursor != len(tokens):
        raise ValueError("Unexpected trailing HWP equation tokens.")
    return math
