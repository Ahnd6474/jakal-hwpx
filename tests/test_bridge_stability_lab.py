from __future__ import annotations

from pathlib import Path

import pytest

import jakal_hwpx._bridge_stability_lab as bridge_stability_lab

from jakal_hwpx._bridge_stability_lab import run_bridge_stability_matrix


REPO_ROOT = Path(__file__).resolve().parents[1]
HWP_SAMPLE_DIR = REPO_ROOT / "examples" / "samples" / "hwp"


@pytest.fixture(scope="session")
def sample_hwp_path() -> Path:
    paths = sorted(HWP_SAMPLE_DIR.glob("*.hwp"))
    assert paths, f"No .hwp samples were found under {HWP_SAMPLE_DIR}"
    return next((path for path in paths if not path.name.startswith("generated_")), paths[0])


def test_bridge_stability_lab_matrix(tmp_path: Path) -> None:
    results = run_bridge_stability_matrix(tmp_path / "bridge_stability_lab")
    failures = [result for result in results if not result.ok]
    assert not failures, "\n".join(f"{result.name}: bridge={result.bridge_errors} hancom={result.hancom_errors}" for result in failures)
    assert len(results) >= 20
    assert all(result.bridge_ok for result in results)
    assert all(result.hancom_status == "skipped" for result in results)


def test_bridge_stability_lab_can_record_hancom_results(
    tmp_path: Path,
    sample_hwp_path: Path,
    sample_hwpx_path: Path,
    monkeypatch,
) -> None:
    conversions: list[str] = []
    fake_convert = bridge_stability_lab._sample_backed_converter(conversions)
    monkeypatch.setattr(bridge_stability_lab, "convert_document", fake_convert)

    results = run_bridge_stability_matrix(tmp_path / "bridge_stability_lab_hancom", validate_with_hancom=True)
    assert results
    assert len(results) >= 20
    assert all(result.bridge_ok for result in results)
    assert all(result.hancom_status == "passed" for result in results)
    assert all(result.hancom_ok is True for result in results)
    assert conversions
