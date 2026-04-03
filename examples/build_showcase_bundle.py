from __future__ import annotations

import argparse
import json
import sys
import zipfile
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path

from lxml import etree

REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = REPO_ROOT / "src"

if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from jakal_hwpx import HwpxDocument  # noqa: E402


NS = {"hp": "http://www.hancom.co.kr/hwpml/2011/paragraph"}


@dataclass
class ShowcaseResult:
    name: str
    description: str
    source: str
    output: str
    features: list[str]
    validations: dict[str, list[str] | None]
    notes: list[str] = field(default_factory=list)


def valid_hwpx_files(corpus_dir: Path) -> list[Path]:
    return [path for path in sorted(corpus_dir.glob("*.hwpx")) if zipfile.is_zipfile(path)]


def find_sample_with_section_xpath(files: list[Path], expression: str) -> Path:
    fallback: Path | None = None
    for path in files:
        with zipfile.ZipFile(path) as archive:
            matched = False
            for name in [item for item in archive.namelist() if item.startswith("Contents/section") and item.endswith(".xml")]:
                data = archive.read(name)
                if not data.lstrip().startswith(b"<"):
                    continue
                root = etree.fromstring(data)
                if root.xpath(expression, namespaces=NS):
                    matched = True
                    break
            if not matched:
                continue
        if fallback is None:
            fallback = path
        try:
            document = HwpxDocument.open(path)
        except Exception:
            continue
        if document.reference_validation_errors() == []:
            return path
    if fallback is not None:
        return fallback
    raise LookupError(f"No sample matched xpath: {expression}")


def preferred_style_id(items: list[object]) -> str | None:
    for item in items:
        style_id = getattr(item, "style_id", None)
        if style_id not in {None, "0"}:
            return style_id
    if not items:
        return None
    return getattr(items[0], "style_id", None)


def validate_document(document: HwpxDocument, *, verify_hancom: bool) -> dict[str, list[str] | None]:
    hancom_open = document.hancom_open_validation_errors() if verify_hancom else None
    return {
        "xml": document.xml_validation_errors(),
        "reference": document.reference_validation_errors(),
        "save_reopen": document.save_reopen_validation_errors(),
        "hancom_open": hancom_open,
    }


def finalize_document(
    document: HwpxDocument,
    *,
    name: str,
    description: str,
    source: Path,
    output_path: Path,
    features: list[str],
    verify_hancom: bool,
    notes: list[str] | None = None,
) -> ShowcaseResult:
    document.set_metadata(title=name, creator="jakal_hwpx showcase")
    document.set_preview_text(description)
    document.save(output_path)
    return ShowcaseResult(
        name=name,
        description=description,
        source=str(source),
        output=str(output_path),
        features=features,
        validations=validate_document(document, verify_hancom=verify_hancom),
        notes=notes or [],
    )


def build_layout_showcase(files: list[Path], output_dir: Path, *, verify_hancom: bool) -> ShowcaseResult:
    source = find_sample_with_section_xpath(files, ".//hp:header and .//hp:footer")
    document = HwpxDocument.open(source)

    document.headers()[0].set_text("JAKAL HWPX SHOWCASE HEADER")
    document.footers()[0].set_text("Generated from Python")
    settings = document.section_settings(0)
    if settings.page_width:
        settings.set_page_size(width=settings.page_width + 200)
    settings.set_margins(left=9000, right=9000, top=7000, bottom=5000)

    style_id = preferred_style_id(document.styles())
    para_pr_id = preferred_style_id(document.paragraph_styles())
    char_pr_id = preferred_style_id(document.character_styles())

    document.append_paragraph("JAKAL HWPX SHOWCASE")
    document.append_paragraph("STYLE_TARGET_LAYOUT")
    document.apply_style_batch(
        section_index=0,
        text_contains="STYLE_TARGET_LAYOUT",
        style_id=style_id,
        para_pr_id=para_pr_id,
        char_pr_id=char_pr_id,
    )

    return finalize_document(
        document,
        name="showcase_layout_headers",
        description="Header, footer, page setup, paragraph insertion, and style batch editing.",
        source=source,
        output_path=output_dir / "showcase_layout_headers.hwpx",
        features=["header", "footer", "section_settings", "paragraph", "style_batch"],
        verify_hancom=verify_hancom,
    )


def build_table_picture_showcase(files: list[Path], output_dir: Path, *, verify_hancom: bool) -> ShowcaseResult:
    source = find_sample_with_section_xpath(files, ".//hp:pic and .//hp:tbl[hp:tr/hp:tc[2]]")
    document = HwpxDocument.open(source)

    table = next(table for table in document.tables() if table.column_count >= 2)
    table.set_cell_text(0, 0, "Python updated cell")
    table.set_cell_text(0, 1, "Round-trip safe")
    new_row = table.append_row()
    new_row[0].set_text("Merged")
    new_row[1].set_text("Row")
    merged = table.merge_cells(new_row[0].row, new_row[0].column, new_row[0].row, new_row[1].column)
    merged.set_text("Python merged row")

    picture = document.pictures()[0]
    picture.shape_comment = "Edited by jakal_hwpx showcase"

    return finalize_document(
        document,
        name="showcase_table_picture",
        description="Table cell updates, merged row editing, and picture metadata editing.",
        source=source,
        output_path=output_dir / "showcase_table_picture.hwpx",
        features=["table_editing", "table_merge", "picture_comment"],
        verify_hancom=verify_hancom,
    )


def build_field_showcase(files: list[Path], output_dir: Path, *, verify_hancom: bool) -> ShowcaseResult:
    source = files[0]
    document = HwpxDocument.open(source)

    document.append_paragraph("Dynamic fields generated from Python")
    document.append_bookmark("showcase_anchor")
    document.append_hyperlink("https://github.com/hancom-io/hwpx-owpml-model", display_text="Open schema reference")
    document.append_mail_merge_field("customer_name", display_text="CUSTOMER_NAME")
    document.append_calculation_field("40+2", display_text="42")
    document.append_cross_reference("showcase_anchor", display_text="Jump target")

    return finalize_document(
        document,
        name="showcase_fields_references",
        description="Bookmark, hyperlink, mail merge, formula, and cross-reference examples.",
        source=source,
        output_path=output_dir / "showcase_fields_references.hwpx",
        features=["bookmark", "hyperlink", "mailmerge", "formula", "cross_reference"],
        verify_hancom=verify_hancom,
    )


def build_notes_showcase(files: list[Path], output_dir: Path, *, verify_hancom: bool) -> ShowcaseResult:
    source = find_sample_with_section_xpath(files, ".//hp:footNote and .//hp:endNote")
    document = HwpxDocument.open(source)

    notes = document.notes()
    notes[0].set_text("Footnote updated by Python.")
    notes[1].set_text("Endnote updated by Python.")

    return finalize_document(
        document,
        name="showcase_notes",
        description="Footnote and endnote content editing.",
        source=source,
        output_path=output_dir / "showcase_notes.hwpx",
        features=["footnote", "endnote"],
        verify_hancom=verify_hancom,
    )


def build_numbering_showcase(files: list[Path], output_dir: Path, *, verify_hancom: bool) -> ShowcaseResult:
    source = find_sample_with_section_xpath(files, ".//hp:newNum")
    document = HwpxDocument.open(source)

    target = next(item for item in document.auto_numbers() if item.kind == "newNum")
    target.set_number(9)

    return finalize_document(
        document,
        name="showcase_numbering",
        description="Body-level automatic numbering update using hp:newNum.",
        source=source,
        output_path=output_dir / "showcase_numbering.hwpx",
        features=["newnum", "automatic_numbering"],
        verify_hancom=verify_hancom,
        notes=["Footnote/endnote marker numbers are recalculated by Hancom and are not forced by this demo."],
    )


def build_equation_showcase(files: list[Path], output_dir: Path, *, verify_hancom: bool) -> ShowcaseResult:
    source = find_sample_with_section_xpath(files, ".//hp:equation")
    document = HwpxDocument.open(source)

    document.equations()[0].script = "x=1+2"

    return finalize_document(
        document,
        name="showcase_equation",
        description="Equation script editing.",
        source=source,
        output_path=output_dir / "showcase_equation.hwpx",
        features=["equation"],
        verify_hancom=verify_hancom,
    )


def build_shape_showcase(files: list[Path], output_dir: Path, *, verify_hancom: bool) -> ShowcaseResult:
    source = find_sample_with_section_xpath(files, ".//hp:textart or .//hp:rect or .//hp:line")
    document = HwpxDocument.open(source)

    shape = document.shapes()[0]
    shape.shape_comment = "Edited by Python showcase"
    if shape.kind == "textart":
        shape.set_text("TEXTART EDITED")

    return finalize_document(
        document,
        name="showcase_shapes",
        description="Shape comment editing and textart text replacement when available.",
        source=source,
        output_path=output_dir / "showcase_shapes.hwpx",
        features=["shape", "textart"],
        verify_hancom=verify_hancom,
    )


def write_markdown_report(results: list[ShowcaseResult], report_path: Path) -> None:
    lines = [
        "# Showcase Bundle",
        "",
        "Generated demonstration documents for jakal_hwpx.",
        "",
    ]
    for result in results:
        lines.extend(
            [
                f"## {result.name}",
                "",
                result.description,
                "",
                f"- Source: `{result.source}`",
                f"- Output: `{result.output}`",
                f"- Features: {', '.join(result.features)}",
                f"- XML validation errors: {len(result.validations['xml'] or [])}",
                f"- Reference validation errors: {len(result.validations['reference'] or [])}",
                f"- Save/reopen validation errors: {len(result.validations['save_reopen'] or [])}",
                (
                    f"- Hancom open validation errors: {len(result.validations['hancom_open'] or [])}"
                    if result.validations["hancom_open"] is not None
                    else "- Hancom open validation: skipped"
                ),
            ]
        )
        if result.notes:
            for note in result.notes:
                lines.append(f"- Note: {note}")
        lines.append("")
    report_path.write_text("\n".join(lines), encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build promotion-ready HWPX showcase examples.")
    parser.add_argument(
        "--corpus-dir",
        default=str(REPO_ROOT / "all_hwpx_flat"),
        help="Directory containing valid HWPX samples.",
    )
    parser.add_argument(
        "--output-dir",
        default=str(REPO_ROOT / "examples" / "output"),
        help="Directory where showcase documents will be written.",
    )
    parser.add_argument(
        "--skip-hancom",
        action="store_true",
        help="Skip Hancom open validation even if Hwp is installed.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    corpus_dir = Path(args.corpus_dir)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    files = valid_hwpx_files(corpus_dir)
    if not files:
        raise SystemExit(f"No valid HWPX files were found in {corpus_dir}.")

    verify_hancom = not args.skip_hancom and HwpxDocument.discover_hancom_executable() is not None
    results = [
        build_layout_showcase(files, output_dir, verify_hancom=verify_hancom),
        build_table_picture_showcase(files, output_dir, verify_hancom=verify_hancom),
        build_field_showcase(files, output_dir, verify_hancom=verify_hancom),
        build_notes_showcase(files, output_dir, verify_hancom=verify_hancom),
        build_numbering_showcase(files, output_dir, verify_hancom=verify_hancom),
        build_equation_showcase(files, output_dir, verify_hancom=verify_hancom),
        build_shape_showcase(files, output_dir, verify_hancom=verify_hancom),
    ]

    manifest_path = output_dir / "showcase_manifest.json"
    report_path = output_dir / "showcase_report.md"
    manifest = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "corpus_dir": str(corpus_dir),
        "output_dir": str(output_dir),
        "hancom_validation_ran": verify_hancom,
        "documents": [asdict(item) for item in results],
    }
    manifest_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")
    write_markdown_report(results, report_path)

    for result in results:
        print(f"[ok] {result.name}: {result.output}")
    print(f"[ok] manifest: {manifest_path}")
    print(f"[ok] report:   {report_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
