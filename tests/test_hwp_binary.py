from __future__ import annotations

from pathlib import Path

import pytest

from jakal_hwpx import HwpBinaryDocument, HwpBinaryFileHeader, HwpParagraph, HwpRecord


REPO_ROOT = Path(__file__).resolve().parents[1]
HWP_SAMPLE_DIR = REPO_ROOT / "examples" / "samples" / "hwp"


@pytest.fixture(scope="session")
def sample_hwp_path() -> Path:
    paths = sorted(HWP_SAMPLE_DIR.glob("*.hwp"))
    assert paths, f"No .hwp samples were found under {HWP_SAMPLE_DIR}"
    return paths[0]


def test_open_hwp_binary_document_exposes_file_header_and_streams(sample_hwp_path: Path) -> None:
    document = HwpBinaryDocument.open(sample_hwp_path)

    file_header = document.file_header()
    assert isinstance(file_header, HwpBinaryFileHeader)
    assert file_header.signature == "HWP Document File"
    assert file_header.version == "5.1.1.0"
    assert file_header.compressed is True

    stream_paths = document.list_stream_paths()
    assert "FileHeader" in stream_paths
    assert "DocInfo" in stream_paths
    assert "BodyText/Section0" in stream_paths
    assert "PrvText" in stream_paths

    preview = document.preview_text()
    assert "2027학년도" in preview
    assert "수학 영역" in preview


def test_docinfo_and_section_records_parse_from_sample_hwp(sample_hwp_path: Path) -> None:
    document = HwpBinaryDocument.open(sample_hwp_path)

    docinfo_records = document.docinfo_records()
    assert docinfo_records
    assert isinstance(docinfo_records[0], HwpRecord)
    assert docinfo_records[0].tag_id == 0x010

    paragraphs = document.paragraphs(0)
    assert paragraphs
    assert isinstance(paragraphs[0], HwpParagraph)
    assert any("2027학년도" in paragraph.text for paragraph in paragraphs)
    assert "2027학년도" in document.get_document_text()


def test_hwp_binary_document_can_update_preview_and_same_length_body_text(
    sample_hwp_path: Path,
    tmp_path: Path,
) -> None:
    document = HwpBinaryDocument.open(sample_hwp_path)
    document.set_preview_text("JAKAL HWP PREVIEW")

    replacements = document.replace_text_same_length("2027", "2028", count=1)
    assert replacements == 1

    output_path = tmp_path / "edited_sample.hwp"
    document.save(output_path)

    reopened = HwpBinaryDocument.open(output_path)
    assert reopened.preview_text().startswith("JAKAL HWP PREVIEW")
    assert "2028학년도" in reopened.get_document_text()
