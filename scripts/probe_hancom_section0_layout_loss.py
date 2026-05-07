from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = REPO_ROOT / "src"

if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from jakal_hwpx import HwpxDocument  # noqa: E402
from jakal_hwpx._hancom import convert_document  # noqa: E402


PARA_NS = {"hp": "http://www.hancom.co.kr/hwpml/2011/paragraph"}


def _paragraphs(document: HwpxDocument):
    return document.sections[0].root_element.xpath("./hp:p", namespaces=PARA_NS)


def _colpr_positions(document: HwpxDocument) -> list[int]:
    positions: list[int] = []
    for index, paragraph in enumerate(_paragraphs(document)):
        if paragraph.xpath("./hp:run/hp:ctrl/hp:colPr", namespaces=PARA_NS):
            positions.append(index)
    return positions


def _trim_section0_prefix(document: HwpxDocument, keep_count: int) -> None:
    section = document.sections[0]
    paragraphs = _paragraphs(document)
    for paragraph in paragraphs[keep_count:]:
        section.root_element.remove(paragraph)
    section.mark_modified()


def _summary(document: HwpxDocument) -> dict[str, object]:
    return {
        "text_len": len(document.get_document_text()),
        "table_count": len(document.tables()),
        "picture_count": len(document.pictures()),
        "shape_count": len(document.shapes()),
        "paragraph_count": len(_paragraphs(document)),
        "colpr_positions": _colpr_positions(document),
    }


def _probe_cutoff(source: Path, output_root: Path, cutoff: int, *, timeout_seconds: int) -> dict[str, object]:
    document = HwpxDocument.open(source)
    _trim_section0_prefix(document, cutoff)

    sliced_path = output_root / f"prefix_{cutoff}.hwpx"
    hwp_path = output_root / f"prefix_{cutoff}.hwp"
    back_path = output_root / f"prefix_{cutoff}_back.hwpx"

    document.save(sliced_path, validate=True)
    convert_document(sliced_path, hwp_path, "HWP", timeout_seconds=timeout_seconds)
    convert_document(hwp_path, back_path, "HWPX", timeout_seconds=timeout_seconds)

    source_doc = HwpxDocument.open(sliced_path)
    roundtrip_doc = HwpxDocument.open(back_path)
    roundtrip_doc.validate()

    source_summary = _summary(source_doc)
    roundtrip_summary = _summary(roundtrip_doc)

    changed = {
        key: (source_summary[key], roundtrip_summary[key])
        for key in source_summary
        if source_summary[key] != roundtrip_summary[key]
    }

    return {
        "cutoff": cutoff,
        "source_path": str(sliced_path),
        "roundtrip_path": str(back_path),
        "source": source_summary,
        "roundtrip": roundtrip_summary,
        "changed": changed,
        "stable": not changed,
    }


def run_probe(source: Path, output_root: Path, cutoffs: list[int], *, timeout_seconds: int) -> list[dict[str, object]]:
    output_root.mkdir(parents=True, exist_ok=True)
    results: list[dict[str, object]] = []
    for cutoff in cutoffs:
        results.append(_probe_cutoff(source, output_root, cutoff, timeout_seconds=timeout_seconds))
    return results


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Slice section0 paragraph prefixes and compare Hancom roundtrip stability to locate layout-loss boundaries."
    )
    parser.add_argument("source", type=Path, help="Source HWPX file to probe.")
    parser.add_argument("--output-root", type=Path, default=REPO_ROOT / ".codex-temp" / "section0_layout_loss_probe")
    parser.add_argument("--cutoffs", nargs="+", type=int, required=True, help="Prefix paragraph counts to keep.")
    parser.add_argument("--timeout-seconds", type=int, default=60)
    return parser.parse_args()


def main() -> int:
    args = _parse_args()
    results = run_probe(args.source, args.output_root, args.cutoffs, timeout_seconds=args.timeout_seconds)
    print(json.dumps(results, ensure_ascii=False, indent=2))
    unstable = [result for result in results if not result["stable"]]
    return 1 if unstable else 0


if __name__ == "__main__":
    raise SystemExit(main())
