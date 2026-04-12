from __future__ import annotations

import zipfile
from pathlib import Path

from lxml import etree


NS = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
}


def _sample_matches(path: Path, expression: str) -> bool:
    with zipfile.ZipFile(path) as zf:
        for name in [n for n in zf.namelist() if n.startswith("Contents/section") and n.endswith(".xml")]:
            data = zf.read(name)
            if not data.lstrip().startswith(b"<"):
                continue
            root = etree.fromstring(data)
            if root.xpath(expression, namespaces=NS):
                return True
    return False


def test_sample_corpus_covers_critical_feature_gaps(valid_hwpx_files: list[Path]) -> None:
    coverage = {
        "footer": ".//hp:footer",
        "footnote": ".//hp:footNote",
        "endnote": ".//hp:endNote",
        "bookmark": ".//hp:bookmark",
        "hyperlink": ".//hp:fieldBegin[@type='HYPERLINK']",
    }

    missing = [
        label
        for label, expression in coverage.items()
        if not any(_sample_matches(path, expression) for path in valid_hwpx_files)
    ]

    assert missing == []


def test_sample_corpus_includes_hard_example_like_fixture(valid_hwpx_files: list[Path]) -> None:
    target = next((path for path in valid_hwpx_files if path.name == "generated_hard_example_like.hwpx"), None)
    assert target is not None

    expectations = {
        "picture": ".//hp:pic",
        "equation": ".//hp:equation",
        "footnote": ".//hp:footNote",
        "field": ".//hp:fieldBegin",
        "auto_number": ".//hp:autoNum",
        "rect": ".//hp:rect",
        "textart": ".//hp:textart",
    }

    missing = [label for label, expression in expectations.items() if not _sample_matches(target, expression)]
    assert missing == []


def test_sample_corpus_includes_native_control_bridge_fixture(valid_hwpx_files: list[Path]) -> None:
    target = next((path for path in valid_hwpx_files if path.name == "generated_native_control_authoring_bridge.hwpx"), None)
    assert target is not None

    expectations = {
        "equation": ".//hp:equation",
        "rect": ".//hp:rect",
        "ellipse": ".//hp:ellipse",
        "arc": ".//hp:arc",
        "polygon": ".//hp:polygon",
        "textart": ".//hp:textart",
        "line": ".//hp:line",
        "ole": ".//hp:ole",
    }

    missing = [label for label, expression in expectations.items() if not _sample_matches(target, expression)]
    assert missing == []


def test_native_control_bridge_fixture_captures_full_parity_metadata(valid_hwpx_files: list[Path]) -> None:
    target = next((path for path in valid_hwpx_files if path.name == "generated_native_control_authoring_bridge.hwpx"), None)
    assert target is not None

    expectations = {
        "equation_comment": ".//hp:equation[hp:shapeComment='Generated native equation']",
        "shape_comment": ".//hp:rect[hp:shapeComment='Generated native shape']",
        "shape_fill": ".//hp:rect/hc:fillBrush/hc:winBrush[@faceColor='#ABCDEF']",
        "shape_line": ".//hp:rect/hp:lineShape[@color='#123456']",
        "ellipse_fill": ".//hp:ellipse/hc:fillBrush/hc:winBrush[@faceColor='#A1B2C3']",
        "ellipse_line": ".//hp:ellipse/hp:lineShape[@color='#102030']",
        "arc_fill": ".//hp:arc/hc:fillBrush/hc:winBrush[@faceColor='#C3B2A1']",
        "polygon_fill": ".//hp:polygon/hc:fillBrush/hc:winBrush[@faceColor='#DDEEFF']",
        "textart_fill": ".//hp:textart/hc:fillBrush/hc:winBrush[@faceColor='#FFEEDD']",
        "textart_comment": ".//hp:textart[hp:shapeComment='Generated native textart']",
        "ole_comment": ".//hp:ole[hp:shapeComment='Generated native OLE']",
        "ole_object_type": ".//hp:ole[@objectType='LINK' and @drawAspect='ICON' and @hasMoniker='1' and @eqBaseLine='12']",
        "ole_line": ".//hp:ole/hp:lineShape[@color='#445566' and @width='77']",
    }

    missing = [label for label, expression in expectations.items() if not _sample_matches(target, expression)]
    assert missing == []
