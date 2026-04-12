from __future__ import annotations

import sys
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = REPO_ROOT / "src"
HWP_SAMPLE_DIR = REPO_ROOT / "examples" / "samples" / "hwp"
HWPX_SAMPLE_DIR = REPO_ROOT / "examples" / "samples" / "hwpx"

if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from jakal_hwpx import HwpDocument  # noqa: E402


def build_native_hwp_sample(output_path: Path) -> Path:
    document = HwpDocument.blank()
    document.set_preview_text("Generated native control authoring sample")
    document.append_paragraph("Generated native control authoring sample")
    document.append_equation("x+y=z", width=3456, height=2100, font="Batang", shape_comment="Generated native equation")
    document.append_shape(
        kind="rect",
        text="Native rect",
        width=3600,
        height=1800,
        fill_color="#ABCDEF",
        line_color="#123456",
        shape_comment="Generated native shape",
    )
    document.append_shape(
        kind="ellipse",
        text="Native ellipse",
        width=3800,
        height=2000,
        fill_color="#A1B2C3",
        line_color="#102030",
        shape_comment="Generated native ellipse",
    )
    document.append_shape(
        kind="arc",
        text="Native arc",
        width=4000,
        height=2200,
        fill_color="#C3B2A1",
        line_color="#203040",
        shape_comment="Generated native arc",
    )
    document.append_shape(
        kind="polygon",
        text="Native polygon",
        width=4200,
        height=2400,
        fill_color="#DDEEFF",
        line_color="#304050",
        shape_comment="Generated native polygon",
    )
    document.append_shape(
        kind="textart",
        text="Native textart",
        width=4400,
        height=2600,
        fill_color="#FFEEDD",
        line_color="#405060",
        shape_comment="Generated native textart",
    )
    document.append_shape(kind="line", width=4200, height=900, line_color="#345678")
    document.append_ole(
        "native-control.ole",
        b"NATIVE_CONTROL_OLE",
        width=3900,
        height=2200,
        shape_comment="Generated native OLE",
        object_type="LINK",
        draw_aspect="ICON",
        has_moniker=True,
        eq_baseline=12,
        line_color="#445566",
        line_width=77,
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    return output_path


def build_native_bridge_hwpx_sample(source_path: Path, output_path: Path) -> Path:
    document = HwpDocument.open(source_path)
    bridged = document.to_hwpx_document()
    bridged.set_metadata(
        title="Generated Native Control Authoring Bridge",
        creator="jakal_hwpx",
        description="Synthetic bridge fixture generated from native HWP equation, shape, and OLE authoring.",
        keyword="generated,corpus,native,bridge",
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    bridged.save(output_path)
    return output_path


def main() -> int:
    native_hwp = build_native_hwp_sample(HWP_SAMPLE_DIR / "generated_native_control_authoring.hwp")
    native_hwpx = build_native_bridge_hwpx_sample(
        native_hwp,
        HWPX_SAMPLE_DIR / "generated_native_control_authoring_bridge.hwpx",
    )
    print(f"[ok] {native_hwp}")
    print(f"[ok] {native_hwpx}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
