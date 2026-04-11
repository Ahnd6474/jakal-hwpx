from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = REPO_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from jakal_hwpx import build_hwp_pure_profile


def main() -> int:
    parser = argparse.ArgumentParser(description="Build a pure-python HWP template profile from hwp_collection.")
    parser.add_argument(
        "--collection-root",
        default=str(REPO_ROOT / "hwp_collection"),
        help="Directory containing donor .hwp files.",
    )
    parser.add_argument(
        "--output-dir",
        default=str(REPO_ROOT / "build" / "hwp_pure_profile"),
        help="Directory where base.hwp, feature templates, and profile.json will be written.",
    )
    args = parser.parse_args()

    profile = build_hwp_pure_profile(args.collection_root, args.output_dir)
    payload = {
        "root": str(profile.root),
        "base_path": str(profile.base_path),
        "donor_path": str(profile.donor_path),
        "template_paths": {key: str(value) for key, value in profile.template_paths.items()},
    }
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
