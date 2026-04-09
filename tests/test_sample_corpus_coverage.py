from __future__ import annotations

import zipfile
from pathlib import Path

from lxml import etree


NS = {"hp": "http://www.hancom.co.kr/hwpml/2011/paragraph"}


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
