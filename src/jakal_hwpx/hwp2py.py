from __future__ import annotations

import argparse
from pathlib import Path
from typing import Sequence

from .hancom_document import HancomConverter, HancomDocument
from .hwpx2py import (
    HancomScriptMode,
    _generate_authoring_hancom_script,
    _generate_semantic_hancom_script,
    _normalize_hancom_mode,
)


def generate_hwp_script(
    source_path: str | Path,
    *,
    default_output_name: str | None = None,
    include_binary_assets: bool = True,
    converter: HancomConverter | None = None,
    mode: HancomScriptMode | str = "semantic",
) -> str:
    """Generate a Python script that recreates a HWP file from scratch.

    HWP input is decoded into the public ``HancomDocument`` model first. The
    generated script then rebuilds that model from scratch and writes HWP.
    """

    source = Path(source_path)
    document = HancomDocument.read_hwp(source, converter=converter)
    selected_mode = _normalize_hancom_mode(str(mode))
    if selected_mode in {"authoring", "macro"}:
        return _generate_authoring_hancom_script(
            document,
            source=source,
            default_output_name=default_output_name,
            include_binary_assets=include_binary_assets,
            generator_module="jakal_hwpx.hwp2py",
            source_label="HWP",
            output_format="hwp",
            macro=selected_mode == "macro",
        )
    return _generate_semantic_hancom_script(
        document,
        source=source,
        default_output_name=default_output_name,
        include_binary_assets=include_binary_assets,
        generator_module="jakal_hwpx.hwp2py",
        source_label="HWP",
        output_format="hwp",
    )


def default_hwp_script_path(source_path: str | Path) -> Path:
    source = Path(source_path)
    return source.with_name(f"{source.stem}_hwp2py.py")


def default_script_path(source_path: str | Path) -> Path:
    return default_hwp_script_path(source_path)


def write_hwp_script(
    source_path: str | Path,
    script_path: str | Path | None = None,
    *,
    default_output_name: str | None = None,
    include_binary_assets: bool = True,
    converter: HancomConverter | None = None,
    mode: HancomScriptMode | str = "semantic",
) -> Path:
    target = Path(script_path) if script_path is not None else default_hwp_script_path(source_path)
    target.parent.mkdir(parents=True, exist_ok=True)
    script = generate_hwp_script(
        source_path,
        default_output_name=default_output_name,
        include_binary_assets=include_binary_assets,
        converter=converter,
        mode=mode,
    )
    target.write_text(script, encoding="utf-8", newline="\n")
    return target


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Generate a from-scratch Python script for a HWP file.")
    parser.add_argument("input", help="source .hwp file")
    parser.add_argument("-o", "--output", help="generated Python script path")
    parser.add_argument(
        "--default-output",
        help="default .hwp path used by the generated script",
    )
    parser.add_argument(
        "--skip-binary-assets",
        action="store_true",
        help="leave Picture/OLE payloads out of the generated script",
    )
    parser.add_argument(
        "--mode",
        choices=("semantic", "authoring", "dsl", "macro", "latex"),
        default="semantic",
        help="generation mode: public API reconstruction, compact authoring DSL, or macro DSL",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_arg_parser()
    args = parser.parse_args(argv)
    write_hwp_script(
        args.input,
        args.output,
        default_output_name=args.default_output,
        include_binary_assets=not args.skip_binary_assets,
        mode=args.mode,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
