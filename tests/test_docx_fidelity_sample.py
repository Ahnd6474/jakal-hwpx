"""A real DOCX fixture exercising the pure Python conversion paths."""
from pathlib import Path

import pytest

from jakal_hwpx import Document, HwpDocument, HwpxDocument
from jakal_hwpx.hwp_binary import TAG_BORDER_FILL


SOURCE = Path(__file__).resolve().parents[1] / "examples" / "samples" / "docx" / "representative.docx"


def test_representative_docx_retains_table_colors_borders_and_simple_equation(tmp_path: Path) -> None:
    source = Document.open(SOURCE)
    assert any("automatic numbering" in warning for warning in source.import_warnings)
    assert any(block.apply_page_type == "FIRST" for block in source.sections[0].header_footer_blocks)
    assert not any("cell shading" in warning for warning in source.import_warnings)
    assert not any("Word equations" in warning for warning in source.import_warnings)
    assert source.save(tmp_path / "copy.docx").read_bytes() == SOURCE.read_bytes()

    with pytest.warns(UserWarning):
        hwpx_path = source.save(tmp_path / "representative.hwpx")
    with pytest.warns(UserWarning):
        hwp_path = source.save(tmp_path / "representative.hwp")

    hwpx = HwpxDocument.open(hwpx_path)
    hwpx.strict_validate()
    assert [equation.script for equation in hwpx.equations()] == ["x+y=3"]
    assert "•\tBullet one" in hwpx.get_document_text()
    assert "2.\tNumber two" in hwpx.get_document_text()
    fills = {
        int(node.get("id")): node
        for node in hwpx.header.root_element.xpath('//*[local-name()="borderFill"]')
    }
    first = fills[hwpx.tables()[0].cell(0, 0).border_fill_id]
    brush = first.xpath('.//*[local-name()="winBrush"]')[0]
    assert brush.get("faceColor") == "#D9EAF7"
    assert first.xpath('./*[local-name()="leftBorder"]')[0].get("color") == "#4472C4"

    hwp = HwpDocument.open(hwp_path)
    hwp.strict_validate()
    assert [equation.script for equation in hwp.equations()] == ["x+y=3"]
    assert "•\tBullet one" in hwp.get_document_text()
    first_cell = hwp.tables()[0].cells()[0]
    record = hwp.docinfo_model().records_by_tag_id(TAG_BORDER_FILL)[first_cell.border_fill_id]
    assert record.payload[36:40] == bytes.fromhex("f7ead900")
    assert record.payload[4:8] == bytes.fromhex("c4724400")
