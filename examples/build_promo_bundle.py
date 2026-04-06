from __future__ import annotations

import argparse
import json
import sys
from dataclasses import asdict, dataclass
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = REPO_ROOT / "src"

if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from jakal_hwpx import BridgeTextBlock, HwpxDocument, PdfDocument, hwpx_to_pdf, pdf_to_hwpx


@dataclass
class PromoArtifact:
    source_name: str
    output_name: str
    output_path: str
    summary: dict[str, object]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build promo-ready conversion outputs for hwpx2pdf and pdf2hwpx.")
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=REPO_ROOT / "examples" / "promo" / "output",
        help="Directory where promo outputs will be written.",
    )
    return parser.parse_args()


def hwpx_candidates() -> list[Path]:
    return sorted((REPO_ROOT / "examples" / "samples" / "hwpx").glob("*.hwpx"))


def pdf_candidates() -> list[Path]:
    return sorted((REPO_ROOT / "examples" / "samples" / "pdf").glob("*.pdf"))


def score_hwpx_candidate(path: Path) -> tuple[int, int, int, int]:
    document = HwpxDocument.open(path)
    return (
        len(document.tables()) * 5 + len(document.pictures()) * 2 + len(document.shapes()) * 2,
        len(document.tables()),
        len(document.pictures()),
        len(document.shapes()),
    )


def select_hwpx_source() -> Path:
    candidates = hwpx_candidates()
    if not candidates:
        raise FileNotFoundError("No HWPX sample files were found under examples/samples/hwpx.")
    return max(candidates, key=score_hwpx_candidate)


def select_pdf_source() -> Path:
    candidates = pdf_candidates()
    if not candidates:
        raise FileNotFoundError("No PDF sample files were found under examples/samples/pdf.")
    return candidates[0]


def build_ocr_blocks(pdf_document: PdfDocument) -> tuple[list[str], dict[int, list[BridgeTextBlock]]]:
    page_text: list[str] = []
    block_map: dict[int, list[BridgeTextBlock]] = {}

    for index, page in enumerate(pdf_document.pages):
        text = page.extract_text().strip()
        page_text.append(text)
        lines = [line.strip() for line in text.splitlines() if line.strip()][:8]
        if not lines:
            continue
        block_map[index] = [
            BridgeTextBlock(
                text=line,
                left=36.0,
                top=48.0 + offset * 18.0,
                width=220.0,
                height=12.0,
            )
            for offset, line in enumerate(lines)
        ]
    return page_text, block_map


def build_hwpx_to_pdf_artifact(output_dir: Path) -> PromoArtifact:
    source = select_hwpx_source()
    source_doc = HwpxDocument.open(source)
    converted = hwpx_to_pdf(source_doc)
    output_path = output_dir / "promo_hwpx2pdf_complex_layout.pdf"
    converted.save(output_path)

    reopened = PdfDocument.open(output_path)
    first_page_analysis = asdict(reopened.pages[0].analyze_non_text()) if reopened.pages else {}
    summary = {
        "source_tables": len(source_doc.tables()),
        "source_pictures": len(source_doc.pictures()),
        "source_shapes": len(source_doc.shapes()),
        "source_sections": len(source_doc.sections),
        "output_pages": len(reopened.pages),
        "page1_non_text": first_page_analysis,
    }
    return PromoArtifact(
        source_name=source.name,
        output_name=output_path.name,
        output_path=str(output_path.relative_to(REPO_ROOT)),
        summary=summary,
    )


def build_pdf_to_hwpx_artifact(output_dir: Path) -> PromoArtifact:
    source = select_pdf_source()
    source_pdf = PdfDocument.open(source)
    ocr_text, ocr_blocks = build_ocr_blocks(source_pdf)
    converted = pdf_to_hwpx(source_pdf, ocr_text_by_page=ocr_text, ocr_blocks_by_page=ocr_blocks)
    output_path = output_dir / "promo_pdf2hwpx_weird_template.hwpx"
    converted.save(output_path)

    reopened = HwpxDocument.open(output_path)
    summary = {
        "source_pages": len(source_pdf.pages),
        "source_page1_image_placements": len(source_pdf.pages[0].image_placements()) if source_pdf.pages else 0,
        "output_sections": len(reopened.sections),
        "output_pictures": len(reopened.pictures()),
        "output_validation_errors": reopened.validation_errors(),
        "output_preview_text": bool(reopened.preview_text and reopened.preview_text.text.strip()),
    }
    return PromoArtifact(
        source_name=source.name,
        output_name=output_path.name,
        output_path=str(output_path.relative_to(REPO_ROOT)),
        summary=summary,
    )


def write_report(output_dir: Path, hwpx_artifact: PromoArtifact, pdf_artifact: PromoArtifact) -> None:
    report = "\n".join(
        [
            "# Promo Conversion Bundle",
            "",
            "## hwpx2pdf",
            f"- Source: `{hwpx_artifact.source_name}`",
            "- Why this sample: table-heavy and visually mixed HWPX with pictures and shapes",
            f"- Output: `{hwpx_artifact.output_path}`",
            f"- Source tables: {hwpx_artifact.summary['source_tables']}",
            f"- Source pictures: {hwpx_artifact.summary['source_pictures']}",
            f"- Source shapes: {hwpx_artifact.summary['source_shapes']}",
            f"- Output pages: {hwpx_artifact.summary['output_pages']}",
            "",
            "## pdf2hwpx",
            f"- Source: `{pdf_artifact.source_name}`",
            "- Why this sample: irregular PDF layout with many images and dense non-text structure",
            f"- Output: `{pdf_artifact.output_path}`",
            f"- Source pages: {pdf_artifact.summary['source_pages']}",
            f"- Source page 1 image placements: {pdf_artifact.summary['source_page1_image_placements']}",
            f"- Output sections: {pdf_artifact.summary['output_sections']}",
            f"- Output pictures: {pdf_artifact.summary['output_pictures']}",
            "",
            "## Notes",
            "- `pdf2hwpx` uses supplied OCR-like text blocks for the text layer and imports PDF images into HWPX picture objects.",
            "- `hwpx2pdf` focuses on readable text plus simple rendering for tables, pictures, and shapes.",
        ]
    )
    (output_dir / "promo_report.md").write_text(report, encoding="utf-8")


def write_manifest(output_dir: Path, hwpx_artifact: PromoArtifact, pdf_artifact: PromoArtifact) -> None:
    payload = {
        "hwpx2pdf": asdict(hwpx_artifact),
        "pdf2hwpx": asdict(pdf_artifact),
    }
    (output_dir / "promo_manifest.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def main() -> None:
    args = parse_args()
    output_dir = args.output_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    hwpx_artifact = build_hwpx_to_pdf_artifact(output_dir)
    pdf_artifact = build_pdf_to_hwpx_artifact(output_dir)
    write_report(output_dir, hwpx_artifact, pdf_artifact)
    write_manifest(output_dir, hwpx_artifact, pdf_artifact)

    print(f"Promo bundle written to {output_dir}")


if __name__ == "__main__":
    main()
