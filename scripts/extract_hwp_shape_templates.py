from __future__ import annotations

import argparse
import json
from pathlib import Path

from jakal_hwpx import HwpDocument


def _shape_payloads(path: Path) -> dict[str, object]:
    document = HwpDocument.open(path)
    shapes = document.shapes()
    if not shapes:
        raise RuntimeError(f"{path} does not contain any shape controls.")
    if len(shapes) != 1:
        raise RuntimeError(f"{path} contains {len(shapes)} shapes; expected exactly one.")

    shape = shapes[0]
    component = shape._shape_component_record()
    specific = shape._shape_specific_record()
    if component is None or specific is None:
        raise RuntimeError(f"{path} does not expose both shape component and specific payload records.")

    return {
        "path": str(path),
        "kind": shape.kind,
        "size": shape.size(),
        "shape_comment": shape.shape_comment,
        "text": shape.text,
        "ctrl_payload_hex": shape.control_node.payload.hex(),
        "component_payload_hex": component.payload.hex(),
        "specific_payload_hex": specific.payload.hex(),
        "ctrl_payload_len": len(shape.control_node.payload),
        "component_payload_len": len(component.payload),
        "specific_payload_len": len(specific.payload),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Extract native HWP shape template payloads from donor HWP files.")
    parser.add_argument("paths", nargs="+", help="One or more donor HWP files containing exactly one shape.")
    parser.add_argument("--output", help="Optional JSON output path.")
    args = parser.parse_args()

    payloads = [_shape_payloads(Path(path).expanduser().resolve()) for path in args.paths]
    text = json.dumps(payloads, ensure_ascii=False, indent=2)
    if args.output:
        output_path = Path(args.output).expanduser().resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(text, encoding="utf-8")
    else:
        print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
