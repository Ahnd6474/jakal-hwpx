from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = REPO_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from jakal_hwpx import run_template_lab


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate minimal legacy HWP control templates from a donor collection.")
    parser.add_argument(
        "--collection-root",
        default=str(REPO_ROOT / "hwp_collection"),
        help="Directory containing legacy .hwp donor documents.",
    )
    parser.add_argument(
        "--output-dir",
        default=str(REPO_ROOT / "build" / "hwp_template_lab"),
        help="Directory where candidate .hwp templates should be written.",
    )
    parser.add_argument(
        "--feature",
        action="append",
        dest="features",
        choices=["table", "picture", "hyperlink"],
        help="Feature to target. Can be passed multiple times. Defaults to all supported features.",
    )
    parser.add_argument(
        "--paragraph-window",
        type=int,
        default=0,
        help="How many neighboring paragraphs to keep around the target paragraph.",
    )
    parser.add_argument(
        "--validate-with-hancom",
        action="store_true",
        help="If set, validate each reduced candidate by asking Hancom to open and save it as HWPX.",
    )
    args = parser.parse_args()

    candidates = run_template_lab(
        args.collection_root,
        args.output_dir,
        features=args.features,
        paragraph_window=args.paragraph_window,
        validate_with_hancom=args.validate_with_hancom,
    )

    payload = [
        {
            "feature": item.feature,
            "source_path": str(item.source_path),
            "output_path": str(item.output_path),
            "control_id": item.control_id,
            "section_index": item.section_index,
            "occurrence_index": item.occurrence_index,
            "paragraph_window": item.paragraph_window,
            "kept_paragraph_ranges": list(item.kept_paragraph_ranges),
            "valid": item.valid,
            "validation_detail": item.validation_detail,
        }
        for item in candidates
    ]
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
