from __future__ import annotations

import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = REPO_ROOT / "src"

if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from jakal_hwpx._stability_lab import run_stability_matrix, write_stability_report  # noqa: E402


def main() -> int:
    output_dir = REPO_ROOT / "examples" / "output_stability_lab"
    results = run_stability_matrix(output_dir)
    report_path = write_stability_report(results, output_dir / "stability_report.json")

    failures = [result for result in results if not result.ok]
    print(f"cases={len(results)} failures={len(failures)} report={report_path}")
    for result in failures:
        print(f"[FAIL] {result.name}: {result.control_errors or result.edited_errors or result.reopened_errors or result.signature_mismatch}")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
