from __future__ import annotations

from pathlib import Path

from jakal_hwpx import HwpxDocument


def test_roundtrip_all_valid_hwpx_samples(valid_hwpx_files: list[Path], tmp_path: Path) -> None:
    failures: list[str] = []

    for index, source_path in enumerate(valid_hwpx_files):
        try:
            document = HwpxDocument.open(source_path)
            output_path = tmp_path / f"roundtrip_{index}.hwpx"
            document.save(output_path)
            reopened = HwpxDocument.open(output_path)
            if reopened.validation_errors():
                failures.append(f"{source_path.name}: {reopened.validation_errors()}")
        except Exception as exc:  # noqa: BLE001
            failures.append(f"{source_path.name}: {exc}")

    assert not failures, "\n".join(failures)
