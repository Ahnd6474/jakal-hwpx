from __future__ import annotations

from pathlib import Path

from jakal_hwpx import BridgeTextBlock, PdfDocument, hwpx_to_pdf, pdf_to_hwpx

from conftest import find_sample_with_section_xpath


REPO_ROOT = Path(__file__).resolve().parents[1]


def _sample_pdf_path() -> Path:
    pdfs = sorted((REPO_ROOT / "examples" / "samples" / "pdf").glob("*.pdf"))
    if not pdfs:
        pdfs = sorted(REPO_ROOT.glob("*.pdf"))
    if not pdfs:
        raise AssertionError("No sample PDF file was found in the repository root.")
    return pdfs[0]


def _first_text_fragment(value: str) -> str:
    for line in value.splitlines():
        stripped = line.strip()
        if len(stripped) >= 4:
            return stripped[:30]
    return value.strip()[:30]


def test_blank_pdf_document_can_be_written_and_reopened(tmp_path: Path) -> None:
    document = PdfDocument.blank()
    document.set_metadata(title="Bridge PDF", author="jakal-hwpx", subject="pdf smoke")
    page = document.add_page(width=300, height=200)
    page.add_text("Hello PDF\nSecond line", x=24, y=160, font_size=12)

    output_path = tmp_path / "generated.pdf"
    document.save(output_path)

    reopened = PdfDocument.open(output_path)
    assert len(reopened.pages) == 1
    assert reopened.metadata().title == "Bridge PDF"
    assert reopened.pages[0].extract_text() == "Hello PDF\nSecond line"


def test_sample_pdf_exposes_non_text_analysis_and_image_placements() -> None:
    document = PdfDocument.open(_sample_pdf_path())
    first_page = document.pages[0]
    analysis = first_page.analyze_non_text()
    placements = first_page.image_placements()

    assert len(document.pages) >= 1
    assert first_page.width > 0
    assert first_page.height > 0
    assert analysis.image_count >= 1
    assert analysis.vector_summary.path_operator_count > 0
    assert analysis.vector_summary.paint_operator_count > 0
    assert placements
    assert placements[0].width > 0
    assert placements[0].height > 0


def test_pdf_to_hwpx_bridge_embeds_images_and_supports_ocr_blocks() -> None:
    pdf_document = PdfDocument.open(_sample_pdf_path())
    first_page_text = [
        BridgeTextBlock(text="OCR BLOCK TOP", left=36, top=48, width=120, height=12),
        BridgeTextBlock(text="OCR BLOCK LOWER", left=40, top=92, width=140, height=12),
    ]

    hwpx_document = pdf_to_hwpx(pdf_document, ocr_blocks_by_page={0: first_page_text})

    assert len(hwpx_document.sections) == len(pdf_document.pages)
    section_text = hwpx_document.sections[0].extract_text(paragraph_separator="\n")
    assert "OCR BLOCK TOP" in section_text
    assert "OCR BLOCK LOWER" in section_text
    assert section_text.index("OCR BLOCK TOP") < section_text.index("OCR BLOCK LOWER")
    assert "[NON_TEXT]" in hwpx_document.get_document_text()
    assert len(hwpx_document.pictures()) >= len(pdf_document.pages[0].image_placements())
    settings = hwpx_document.section_settings(0)
    assert settings.page_width == int(round(pdf_document.pages[0].width * 100))
    assert settings.page_height == int(round(pdf_document.pages[0].height * 100))


def test_hwpx_to_pdf_bridge_writes_readable_pdf(sample_hwpx_path: Path, tmp_path: Path) -> None:
    from jakal_hwpx import HwpxDocument

    hwpx_document = HwpxDocument.open(sample_hwpx_path)
    expected_fragment = _first_text_fragment(hwpx_document.sections[0].extract_text(paragraph_separator="\n"))

    pdf_document = hwpx_to_pdf(hwpx_document)
    output_path = tmp_path / "from-hwpx.pdf"
    pdf_document.save(output_path)

    reopened = PdfDocument.open(output_path)
    assert len(reopened.pages) == len(hwpx_document.sections)
    assert reopened.metadata().title == hwpx_document.metadata().title
    assert expected_fragment in reopened.pages[0].extract_text()


def test_hwpx_to_pdf_renders_tables_pictures_and_shapes(valid_hwpx_files: list[Path], tmp_path: Path) -> None:
    from jakal_hwpx import HwpxDocument

    picture_table_source = find_sample_with_section_xpath(valid_hwpx_files, ".//hp:pic and .//hp:tbl")
    shape_source = find_sample_with_section_xpath(valid_hwpx_files, ".//hp:rect or .//hp:line or .//hp:textart")

    picture_table_pdf = hwpx_to_pdf(HwpxDocument.open(picture_table_source))
    picture_table_path = tmp_path / "picture-table.pdf"
    picture_table_pdf.save(picture_table_path)
    reopened_picture_table = PdfDocument.open(picture_table_path)
    picture_table_analysis = reopened_picture_table.pages[0].analyze_non_text()
    assert picture_table_analysis.image_count >= 1
    assert picture_table_analysis.vector_summary.path_operator_count > 0

    shape_pdf = hwpx_to_pdf(HwpxDocument.open(shape_source))
    shape_path = tmp_path / "shape.pdf"
    shape_pdf.save(shape_path)
    reopened_shape = PdfDocument.open(shape_path)
    shape_analysis = reopened_shape.pages[0].analyze_non_text()
    assert shape_analysis.vector_summary.path_operator_count > 0
