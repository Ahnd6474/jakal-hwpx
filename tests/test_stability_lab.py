from __future__ import annotations

from pathlib import Path

from jakal_hwpx._stability_lab import run_stability_matrix


def test_stability_lab_matrix(tmp_path: Path) -> None:
    results = run_stability_matrix(tmp_path / "stability_lab")
    failures = [result for result in results if not result.ok]
    assert not failures, "\n".join(
        f"{result.name}: control={result.control_errors} edited={result.edited_errors} "
        f"reopened={result.reopened_errors} mismatch={result.signature_mismatch}"
        for result in failures
    )
