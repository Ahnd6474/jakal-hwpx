from __future__ import annotations

from pathlib import Path

import pytest

from jakal_hwpx import HwpDocument, HwpHwpxBridge, HwpxDocument


REPO_ROOT = Path(__file__).resolve().parents[1]
HWP_SAMPLE_DIR = REPO_ROOT / "examples" / "samples" / "hwp"


@pytest.fixture(scope="session")
def sample_hwp_path() -> Path:
    paths = sorted(HWP_SAMPLE_DIR.glob("*.hwp"))
    assert paths, f"No .hwp samples were found under {HWP_SAMPLE_DIR}"
    return paths[0]


def test_explicit_bridge_can_open_hwp_and_save_both_formats(
    sample_hwp_path: Path,
    sample_hwpx_path: Path,
    tmp_path: Path,
) -> None:
    conversions: list[str] = []

    def fake_converter(input_path: str | Path, output_path: str | Path, output_format: str) -> Path:
        target = Path(output_path)
        conversions.append(output_format)
        if output_format == "HWPX":
            target.write_bytes(sample_hwpx_path.read_bytes())
        elif output_format == "HWP":
            target.write_bytes(sample_hwp_path.read_bytes())
        else:
            raise AssertionError(output_format)
        return target

    bridge = HwpHwpxBridge.open(sample_hwp_path, converter=fake_converter)
    assert isinstance(bridge.hwp_document(), HwpDocument)
    assert isinstance(bridge.hwpx_document(), HwpxDocument)

    output_hwpx = tmp_path / "bridge_from_hwp.hwpx"
    output_hwp = tmp_path / "bridge_from_hwp.hwp"
    bridge.save(output_hwpx)
    bridge.save(output_hwp)

    assert output_hwpx.exists()
    assert output_hwp.exists()
    assert conversions == []


def test_explicit_bridge_can_open_hwpx_and_convert_back_to_hwp(
    sample_hwp_path: Path,
    sample_hwpx_path: Path,
    tmp_path: Path,
) -> None:
    conversions: list[str] = []

    def fake_converter(input_path: str | Path, output_path: str | Path, output_format: str) -> Path:
        target = Path(output_path)
        conversions.append(output_format)
        if output_format == "HWP":
            target.write_bytes(sample_hwp_path.read_bytes())
        elif output_format == "HWPX":
            target.write_bytes(sample_hwpx_path.read_bytes())
        else:
            raise AssertionError(output_format)
        return target

    bridge = HwpHwpxBridge.open(sample_hwpx_path, converter=fake_converter)
    assert isinstance(bridge.hwpx_document(), HwpxDocument)
    assert isinstance(bridge.hwp_document(), HwpDocument)

    output_hwp = tmp_path / "bridge_from_hwpx.hwp"
    bridge.save_hwp(output_hwp)

    assert output_hwp.exists()
    assert conversions == []


def test_hwpx_document_exposes_reverse_bridge_helpers(
    sample_hwp_path: Path,
    sample_hwpx_path: Path,
    tmp_path: Path,
) -> None:
    conversions: list[str] = []

    def fake_converter(input_path: str | Path, output_path: str | Path, output_format: str) -> Path:
        target = Path(output_path)
        conversions.append(output_format)
        if output_format == "HWP":
            target.write_bytes(sample_hwp_path.read_bytes())
        elif output_format == "HWPX":
            target.write_bytes(sample_hwpx_path.read_bytes())
        else:
            raise AssertionError(output_format)
        return target

    document = HwpxDocument.open(sample_hwpx_path)
    bridge = document.bridge(converter=fake_converter)
    assert isinstance(bridge, HwpHwpxBridge)

    hwp_document = document.to_hwp_document(converter=fake_converter)
    assert isinstance(hwp_document, HwpDocument)

    output_hwp = tmp_path / "reverse_helper.hwp"
    document.save_as_hwp(output_hwp, converter=fake_converter)
    assert output_hwp.exists()
    assert conversions == []
