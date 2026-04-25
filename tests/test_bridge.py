from __future__ import annotations

from pathlib import Path

import pytest

from jakal_hwpx import HancomDocument, HwpDocument, HwpHwpxBridge, HwpxDocument


REPO_ROOT = Path(__file__).resolve().parents[1]
HWP_SAMPLE_DIR = REPO_ROOT / "examples" / "samples" / "hwp"


@pytest.fixture(scope="session")
def sample_hwp_path() -> Path:
    paths = sorted(HWP_SAMPLE_DIR.glob("*.hwp"))
    assert paths, f"No .hwp samples were found under {HWP_SAMPLE_DIR}"
    return next((path for path in paths if not path.name.startswith("generated_")), paths[0])


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


def test_bridge_open_detects_hwp_content_even_when_extension_is_hwpx(sample_hwp_path: Path, tmp_path: Path) -> None:
    mislabeled_path = tmp_path / "mislabeled_binary.hwpx"
    mislabeled_path.write_bytes(sample_hwp_path.read_bytes())

    bridge = HwpHwpxBridge.open(mislabeled_path)

    assert isinstance(bridge.hwp_document(), HwpDocument)


def test_bridge_open_detects_hwpx_content_even_when_extension_is_hwp(sample_hwpx_path: Path, tmp_path: Path) -> None:
    mislabeled_path = tmp_path / "mislabeled_zip.hwp"
    mislabeled_path.write_bytes(sample_hwpx_path.read_bytes())

    bridge = HwpHwpxBridge.open(mislabeled_path)

    assert isinstance(bridge.hwpx_document(), HwpxDocument)


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


def test_bridge_can_use_hancom_ir_as_authoritative_edit_surface(tmp_path: Path) -> None:
    source = HwpxDocument.blank()
    source.append_paragraph("SOURCE-LINE")

    bridge = HwpHwpxBridge.from_hwpx(source)
    ir = bridge.hancom_document()

    assert isinstance(ir, HancomDocument)
    ir.metadata.title = "IR-BRIDGE-TITLE"
    ir.append_paragraph("IR-BRIDGE-LINE")
    ir.append_field(field_type="DOCPROPERTY", display_text="IR-BRIDGE-FIELD")

    output_hwpx = tmp_path / "bridge_ir.hwpx"
    output_hwp = tmp_path / "bridge_ir.hwp"
    bridge.save_hwpx(output_hwpx)
    bridge.save_hwp(output_hwp)

    reopened_hwpx = HwpxDocument.open(output_hwpx)
    reopened_hwp = HwpDocument.open(output_hwp)

    assert reopened_hwpx.metadata().title == "IR-BRIDGE-TITLE"
    assert "IR-BRIDGE-LINE" in reopened_hwpx.get_document_text()
    assert "IR-BRIDGE-FIELD" in reopened_hwpx.get_document_text()
    assert reopened_hwp.preview_text() == "IR-BRIDGE-TITLE"
    assert "IR-BRIDGE-LINE" in reopened_hwp.get_document_text()
    assert "IR-BRIDGE-FIELD" in reopened_hwp.get_document_text()
