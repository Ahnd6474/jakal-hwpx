from __future__ import annotations

import argparse
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = REPO_ROOT / "src"

if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from jakal_hwpx._hwp_stability_lab import run_hwp_stability_matrix, write_hwp_stability_report  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description="Run the pure-python HWP stability matrix.")
    parser.add_argument(
        "--with-hancom",
        action="store_true",
        help="Also validate each generated HWP by asking Hancom to open/save it through the COM smoke path.",
    )
    args = parser.parse_args()

    output_dir = REPO_ROOT / "examples" / "output_hwp_stability_lab"
    results = run_hwp_stability_matrix(output_dir, validate_with_hancom=args.with_hancom)
    report_path = write_hwp_stability_report(results, output_dir / "stability_report.json")

    failures = [result for result in results if not result.ok]
    binary_failures = [result for result in results if not result.binary_ok]
    hancom_failures = [result for result in results if result.hancom_status == "failed"]
    hancom_skips = [result for result in results if result.hancom_status == "skipped"]
    print(
        f"cases={len(results)} binary_failures={len(binary_failures)} "
        f"hancom_failures={len(hancom_failures)} hancom_skipped={len(hancom_skips)} report={report_path}"
    )
    for result in failures:
        print(f"[FAIL] {result.name}: binary={result.binary_errors} hancom={result.hancom_errors}")
    for result in hancom_skips:
        if result.hancom_errors:
            print(f"[HANCOM-SKIP] {result.name}: {result.hancom_errors[0]}")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
