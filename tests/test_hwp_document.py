from __future__ import annotations

from pathlib import Path

import pytest

from jakal_hwpx import HwpDocument, HwpDocumentProperties, HwpParagraphObject, HwpSection


REPO_ROOT = Path(__file__).resolve().parents[1]
HWP_SAMPLE_DIR = REPO_ROOT / "examples" / "samples" / "hwp"


@pytest.fixture(scope="session")
def sample_hwp_path() -> Path:
    paths = sorted(HWP_SAMPLE_DIR.glob("*.hwp"))
    assert paths, f"No .hwp samples were found under {HWP_SAMPLE_DIR}"
    return paths[0]


def test_hwp_document_exposes_object_model_for_sections_and_paragraphs(sample_hwp_path: Path) -> None:
    document = HwpDocument.open(sample_hwp_path)

    properties = document.document_properties()
    assert isinstance(properties, HwpDocumentProperties)
    assert properties.section_count >= 1

    sections = document.sections()
    assert sections
    assert isinstance(sections[0], HwpSection)

    paragraphs = sections[0].paragraphs()
    assert paragraphs
    assert isinstance(paragraphs[0], HwpParagraphObject)
    assert any("2027" in paragraph.text for paragraph in paragraphs)


def test_hwp_paragraph_object_can_replace_same_length_text_and_roundtrip(
    sample_hwp_path: Path,
    tmp_path: Path,
) -> None:
    document = HwpDocument.open(sample_hwp_path)
    target = next(paragraph for paragraph in document.section(0).paragraphs() if "2027" in paragraph.text)

    replacements = target.replace_text_same_length("2027", "2028", count=1)
    assert replacements == 1

    output_path = tmp_path / "object_model_roundtrip.hwp"
    document.save(output_path)

    reopened = HwpDocument.open(output_path)
    assert "2028학년도" in reopened.get_document_text()


def test_hwp_section_object_can_scope_same_length_replacements(sample_hwp_path: Path, tmp_path: Path) -> None:
    document = HwpDocument.open(sample_hwp_path)
    section = document.section(0)

    replacements = section.replace_text_same_length("2027", "2029", count=1)
    assert replacements == 1

    output_path = tmp_path / "section_scope_roundtrip.hwp"
    document.save(output_path)

    reopened = HwpDocument.open(output_path)
    assert "2029학년도" in reopened.get_document_text()
