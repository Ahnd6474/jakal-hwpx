"""Bidirectional native equation mapping for the supported math grammar.

Unsupported structures fail explicitly instead of concatenating math leaves,
which changes fractions, radicals and scripts into a different expression.
"""
from __future__ import annotations

import re
from lxml import etree

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _tag(name: str) -> str:
    return f"{{{M}}}{name}"


def omml_to_hwp(node: etree._Element) -> str:
    name = etree.QName(node).localname

    def argument(tag: str) -> str:
        child = node.find(_tag(tag))
        return "" if child is None else omml_to_hwp(child)

    if name in {"oMath", "oMathPara", "num", "den", "e", "sup", "sub", "deg"}:
        return "".join(omml_to_hwp(child) for child in node if not etree.QName(child).localname.endswith("Pr"))
    if name == "r":
        text = "".join(child.text or "" for child in node if child.tag == _tag("t"))
        if re.search(r"[{}^_\\]", text):
            raise ValueError("Unsupported literal math operator in Word equation text; use native math structures or an external converter.")
        return text
    if name == "f":
        return "{" + argument("num") + "} over {" + argument("den") + "}"
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
    raise ValueError(f"Unsupported Word equation structure m:{name}; use an external converter.")


def hwp_to_omml(script: str) -> etree._Element:
    if re.search(r"\\|\b(matrix|cases|pile|sum|prod|int|lim|left|right|atop|choose)\b", script, re.IGNORECASE):
        raise ValueError("HWP equations require an external DOCX converter for this unsupported math grammar.")
    tokens = re.findall(r"\b(?:over|sqrt|root|of)\b|[{}^_]|[^{}^_]+?(?=\b(?:over|sqrt|root|of)\b|[{}^_]|$)", script)
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
        if token in {"}", "over", "^", "_", "of"}:
            raise ValueError("Unexpected HWP equation operator: " + token)
        return [run(tokens[cursor - 1].strip())]

    def expression() -> list[etree._Element]:
        nonlocal cursor
        result: list[etree._Element] = []
        while cursor < len(tokens):
            whitespace()
            if cursor >= len(tokens) or tokens[cursor].strip() == "}":
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
            if cursor < len(tokens) and tokens[cursor].strip() == "over":
                cursor += 1
                denominator = atom()
                fraction = etree.Element(_tag("f"))
                fraction.extend([container("num", value), container("den", denominator)])
                value = [fraction]
            result.extend(value)
        return result

    math = etree.Element(_tag("oMath"), nsmap={"m": M})
    math.extend(expression())
    if cursor != len(tokens):
        raise ValueError("Unexpected trailing HWP equation tokens.")
    return math
