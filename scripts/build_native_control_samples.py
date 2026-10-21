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
from jakal_hwpx._ole_payloads import default_ole_cfb_bytes  # noqa: E402


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
        default_ole_cfb_bytes(),
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
    shape_styles = {
        "rect": ("#ABCDEF", "#123456"),
        "ellipse": ("#A1B2C3", "#102030"),
        "arc": ("#C3B2A1", "#203040"),
        "polygon": ("#DDEEFF", "#304050"),
        "textart": ("#FFEEDD", "#405060"),
    }
    for shape in bridged.shapes():
        if shape.kind in shape_styles:
            fill_color, line_color = shape_styles[shape.kind]
            shape.set_fill_style(face_color=fill_color, hatch_color=line_color, alpha=0)
            shape.set_line_style(color=line_color, width=33, style="SOLID")
        elif shape.kind == "line":
            shape.set_line_style(color="#345678", width=33, style="SOLID")
    ole = bridged.oles()[0]
    ole.set_line_style(
        color="#445566",
        width=77,
        style="SOLID",
        end_cap="ROUND",
        head_fill=False,
        tail_fill=False,
        head_size="SMALL_SMALL",
        tail_size="SMALL_SMALL",
        outline_style="OUTER",
    )
    ole.set_object_metadata(object_type="LINK", draw_aspect="ICON", has_moniker=True, eq_baseline=12)
    ole.shape_comment = "Generated native OLE"
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
