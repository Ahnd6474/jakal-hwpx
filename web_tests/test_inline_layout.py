"""Anonymous regression cases for inline breaks in converted HWP paragraphs."""
from lxml import etree
import pytest

from jakal_hwpx import HwpxDocument
from jakal_hwpx.namespaces import NS
from server.conversion import inspect_converted
from server.engine import paragraph_text
from web_tests.test_api import client, upload, wait


def document_with_inline_text(content, positions):
    document = HwpxDocument.blank()
    node = document.append_paragraph("").element
    run = node.find("hp:run", NS)
    for child in list(run):
        run.remove(child)
    run.append(etree.fromstring(f'<hp:t xmlns:hp="{NS["hp"]}">{content}</hp:t>'))
    for cache in node.findall("hp:linesegarray", NS):
        node.remove(cache)
    cache = etree.SubElement(node, f'{{{NS["hp"]}}}linesegarray')
    for position in positions:
        etree.SubElement(cache, f'{{{NS["hp"]}}}lineseg', textpos=str(position))
    return document, node


@pytest.mark.parametrize("content,positions,expected", [
    ("첫 줄<hp:lineBreak/>둘째 줄<hp:lineBreak/>", [0, 4, 9], "첫 줄\n둘째 줄\n"),
    ("<hp:lineBreak/>", [0, 1], "\n"),
    ("   ", [0, 2], "   "),
    ("가😀<hp:lineBreak/>나", [0, 4], "가😀\n나"),
])
def test_valid_inline_positions_are_not_flagged_or_repaired(tmp_path, content, positions, expected):
    doc, node = document_with_inline_text(content, positions)
    original = etree.tostring(node)
    assert paragraph_text(node) == expected
    assert not doc.stale_paragraph_layout_validation_errors()
    assert not doc.repair_stale_paragraph_layout()
    assert etree.tostring(node) == original
    raw = doc.compile(validate=False, auto_repair=False)
    path = tmp_path / "converted.hwpx"
    path.write_bytes(raw)
    inspect_converted(path)
    assert path.read_bytes() == raw


def test_real_overflow_after_inline_break_still_fails_conversion(tmp_path):
    doc, _ = document_with_inline_text("가<hp:lineBreak/>나", [0, 4])
    path = tmp_path / "broken.hwpx"
    path.write_bytes(doc.compile(validate=False, auto_repair=False))
    with pytest.raises(ValueError, match="변환 결과"):
        inspect_converted(path)
    assert doc.repair_stale_paragraph_layout()
    assert not doc.stale_paragraph_layout_validation_errors()


def test_web_snapshot_keeps_inline_text_tails_and_noop_download_bytes(client):
    document, node = document_with_inline_text("첫 줄<hp:lineBreak/>둘째 줄<hp:lineBreak/>마지막 줄", [0, 4, 9])
    expected = paragraph_text(node)
    raw = document.compile(validate=False, auto_repair=False)
    identifier = upload(client, raw)
    base = f"/api/documents/{identifier}"
    p = client.get(base).json()["sections"][0]["paragraphs"][1]
    assert p["text"] == expected == "첫 줄\n둘째 줄\n마지막 줄"
    assert not p["editable"]  # Inline controls retain their original layout.
    job = client.post(base + "/validations", json={"revision": 0}).json()
    assert wait(client, job["job_id"])["status"] == "completed"
    assert client.get(base + "/download?revision=0").content == raw
