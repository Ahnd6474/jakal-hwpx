"""Compare equal-size document page PNGs without rewarding white background.

CLI image loading requires Pillow. The metric itself uses only the standard
library so it can be regression-tested without an imaging dependency.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path


def compare_rgb_pixels(reference: bytes, candidate: bytes, width: int, height: int,
                       *, ink_threshold: int = 245, difference_threshold: int = 24) -> dict:
    if width < 1 or height < 1 or len(reference) != width * height * 3 or len(candidate) != len(reference):
        raise ValueError("Both RGB buffers must have the same positive page dimensions.")
    if not 0 <= difference_threshold < 256 or not 0 < ink_threshold <= 256:
        raise ValueError("Thresholds must be valid RGB channel values.")
    changed = changed_ink = source_ink = target_ink = union = intersection = maximum = 0
    for offset in range(0, len(reference), 3):
        a, b = reference[offset:offset + 3], candidate[offset:offset + 3]
        delta = max(abs(a[index] - b[index]) for index in range(3))
        maximum = max(maximum, delta)
        differs = delta > difference_threshold
        painted_a, painted_b = min(a) < ink_threshold, min(b) < ink_threshold
        painted = painted_a or painted_b
        changed += differs
        changed_ink += differs and painted
        source_ink += painted_a
        target_ink += painted_b
        union += painted
        intersection += painted_a and painted_b
    return {
        "width": width, "height": height, "pixel_exact": reference == candidate,
        "ink_threshold": ink_threshold, "difference_threshold": difference_threshold,
        "maximum_channel_difference": maximum, "changed_pixels": changed,
        "reference_ink_pixels": source_ink, "candidate_ink_pixels": target_ink,
        "ink_union_pixels": union, "whole_page_changed_fraction": changed / (width * height),
        "foreground_changed_fraction": changed_ink / union if union else None,
        "ink_iou": intersection / union if union else None,
        "foreground_comparison_available": bool(union),
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("reference", type=Path)
    parser.add_argument("candidate", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--ink-threshold", type=int, default=245)
    parser.add_argument("--difference-threshold", type=int, default=24)
    args = parser.parse_args()
    from PIL import Image

    def load(path):
        with Image.open(path) as image:
            image = image.convert("RGBA")
            background = Image.new("RGBA", image.size, "white")
            return Image.alpha_composite(background, image).convert("RGB")

    source, target = load(args.reference), load(args.candidate)
    if source.size != target.size:
        parser.error("Page dimensions differ; images must not be resized or aligned to hide a layout change.")
    result = compare_rgb_pixels(source.tobytes(), target.tobytes(), *source.size,
                                ink_threshold=args.ink_threshold, difference_threshold=args.difference_threshold)
    result.update(reference=str(args.reference), candidate=str(args.candidate))
    text = json.dumps(result, indent=2)
    if args.output:
        args.output.write_text(text + "\n", encoding="utf-8")
    print(text)


if __name__ == "__main__":
    main()
