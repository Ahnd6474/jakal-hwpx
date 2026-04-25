from __future__ import annotations

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "c14": "http://schemas.microsoft.com/office/drawing/2007/8/2/chart",
    "ha": "http://www.hancom.co.kr/hwpml/2011/app",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
    "hh": "http://www.hancom.co.kr/hwpml/2011/head",
    "hhs": "http://www.hancom.co.kr/hwpml/2011/history",
    "hm": "http://www.hancom.co.kr/hwpml/2011/master-page",
    "ho": "http://schemas.haansoft.com/office/8.0",
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hp10": "http://www.hancom.co.kr/hwpml/2016/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hv": "http://www.hancom.co.kr/hwpml/2011/version",
    "hpf": "http://www.hancom.co.kr/schema/2011/hpf",
    "config": "urn:oasis:names:tc:opendocument:xmlns:config:1.0",
    "dc": "http://purl.org/dc/elements/1.1/",
    "epub": "http://www.idpf.org/2007/ops",
    "hwpunitchar": "http://www.hancom.co.kr/hwpml/2016/HwpUnitChar",
    "jakalchart": "urn:jakal-hwpx:chart-metadata",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "ocf": "urn:oasis:names:tc:opendocument:xmlns:container",
    "odf": "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
    "ooxmlchart": "http://www.hancom.co.kr/hwpml/2016/ooxmlchart",
    "opf": "http://www.idpf.org/2007/opf/",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

SECTION_PATTERN = r"^Contents/section(\d+)\.xml$"


def qname(prefix: str, local_name: str) -> str:
    return f"{{{NS[prefix]}}}{local_name}"


def resolve_tag(tag: str) -> str:
    if tag.startswith("{"):
        return tag
    if ":" in tag:
        prefix, local_name = tag.split(":", 1)
        return qname(prefix, local_name)
    return tag


def local_name(tag: str) -> str:
    if tag.startswith("{"):
        return tag.split("}", 1)[1]
    return tag
