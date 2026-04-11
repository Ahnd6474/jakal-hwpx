from __future__ import annotations

import argparse
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = REPO_ROOT / "src"

if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from jakal_hwpx._bridge_stability_lab import run_bridge_stability_matrix, write_bridge_stability_report  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description="Run the HWP/HWPX bridge stability matrix.")
    parser.add_argument(
        "--with-hancom",
        action="store_true",
        help="Use the real Hancom converter instead of the sample-backed fake converter.",
    )
    args = parser.parse_args()

    mode = "hancom" if args.with_hancom else "sample"
    output_dir = REPO_ROOT / "examples" / "output_bridge_stability_lab" / mode
    results = run_bridge_stability_matrix(output_dir, validate_with_hancom=args.with_hancom)
    report_path = write_bridge_stability_report(results, output_dir / "stability_report.json")

    failures = [result for result in results if not result.ok]
    bridge_failures = [result for result in results if not result.bridge_ok]
    hancom_failures = [result for result in results if result.hancom_status == "failed"]
    hancom_skips = [result for result in results if result.hancom_status == "skipped"]
    print(
        f"mode={mode} cases={len(results)} bridge_failures={len(bridge_failures)} "
        f"hancom_failures={len(hancom_failures)} hancom_skipped={len(hancom_skips)} report={report_path}"
    )
    for result in failures:
        print(f"[FAIL] {result.name}: bridge={result.bridge_errors} hancom={result.hancom_errors}")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
