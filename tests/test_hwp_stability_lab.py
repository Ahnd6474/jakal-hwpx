from __future__ import annotations

from pathlib import Path

import jakal_hwpx._hwp_stability_lab as hwp_stability_lab

from jakal_hwpx._hwp_stability_lab import run_hwp_stability_matrix


def test_hwp_stability_lab_matrix(tmp_path: Path) -> None:
    results = run_hwp_stability_matrix(tmp_path / "hwp_stability_lab")
    failures = [result for result in results if not result.ok]
    assert not failures, "\n".join(f"{result.name}: binary={result.binary_errors} hancom={result.hancom_errors}" for result in failures)
    assert len(results) >= 70
    assert all(result.binary_ok for result in results)
    assert all(result.hancom_status == "skipped" for result in results)


def test_hwp_stability_lab_can_record_hancom_smoke_results(tmp_path: Path, monkeypatch) -> None:
    conversions: list[tuple[Path, Path, str]] = []

    def fake_convert(input_path: str | Path, output_path: str | Path, output_format: str, **_: object) -> Path:
        source = Path(input_path)
        target = Path(output_path)
        target.parent.mkdir(parents=True, exist_ok=True)
        if output_format.upper() == "HWPX":
            target.write_text("<hwpx />", encoding="utf-8")
        else:
            target.write_bytes(source.read_bytes() if source.exists() else b"fake-hwp")
        conversions.append((source, target, output_format))
        return target

    monkeypatch.setattr(hwp_stability_lab, "convert_document", fake_convert)

    results = run_hwp_stability_matrix(tmp_path / "hwp_stability_lab_hancom", validate_with_hancom=True)
    assert results
    assert len(results) >= 70
    assert all(result.binary_ok for result in results)
    assert all(result.hancom_status == "passed" for result in results)
    assert all(result.hancom_ok is True for result in results)
    assert all(len(result.hancom_artifacts) == 2 for result in results)
    assert conversions
