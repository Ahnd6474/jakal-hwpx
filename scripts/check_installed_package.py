from __future__ import annotations

import argparse
from pathlib import Path

from jakal_hwpx import HwpxDocument
from jakal_hwpx import __all__ as public_api


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run a minimal smoke test against an installed jakal-hwpx distribution."
    )
    parser.add_argument(
        "--sample-path",
        type=Path,
        required=True,
        help="Path to a repository fixture HWPX file used for open/save validation.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        required=True,
        help="Directory where smoke-test outputs will be written.",
    )
    return parser.parse_args()


def require_public_exports() -> None:
    required_exports = {
        "__version__",
        "BinaryDataPart",
        "Bookmark",
        "DocumentMetadata",
        "Equation",
        "Field",
        "HeaderFooterBlock",
        "HwpxDocument",
        "HwpxError",
        "HwpxPart",
        "HwpxValidationError",
        "InvalidHwpxFileError",
        "Note",
        "Picture",
        "SectionSettings",
        "ShapeObject",
        "Table",
        "TableCell",
    }
    missing = sorted(required_exports.difference(public_api))
    if missing:
        raise AssertionError(f"Missing expected public exports: {missing}")


def smoke_open_save(sample_path: Path, output_dir: Path) -> None:
    document = HwpxDocument.open(sample_path)
    assert document.validation_errors() == []

    roundtrip_path = output_dir / "smoke_roundtrip.hwpx"
    document.save(roundtrip_path)

    reopened = HwpxDocument.open(roundtrip_path)
    assert reopened.validation_errors() == []


def smoke_blank_document(output_dir: Path) -> None:
    document = HwpxDocument.blank()
    document.set_metadata(title="Smoke Title", creator="ci-smoke")
    document.append_paragraph("installed package smoke")

    blank_path = output_dir / "smoke_blank.hwpx"
    document.save(blank_path)

    reopened = HwpxDocument.open(blank_path)
    assert reopened.metadata().title == "Smoke Title"
    assert "installed package smoke" in reopened.get_document_text()
    assert reopened.validation_errors() == []


def main() -> None:
    args = parse_args()
    sample_path = args.sample_path.resolve()
    output_dir = args.output_dir.resolve()

    if not sample_path.exists():
        raise FileNotFoundError(sample_path)

    output_dir.mkdir(parents=True, exist_ok=True)
    require_public_exports()
    smoke_open_save(sample_path, output_dir)
    smoke_blank_document(output_dir)
    print(f"[ok] installed package smoke passed: {sample_path}")


if __name__ == "__main__":
    main()
