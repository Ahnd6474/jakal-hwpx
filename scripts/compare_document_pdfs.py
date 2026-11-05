"""Compare every PDF page without resizing, alignment or white-area credit.

Requires Pillow, numpy, pdfplumber, and Poppler's pdftoppm. Export both PDFs in
the same application first. Missing pages or different page sizes invalidate
the aggregate score; they must not become a partial-page fidelity percentage.
"""

from pathlib import Path
import argparse
import json
import subprocess
import shutil
import numpy as np
from PIL import Image, ImageDraw
import pdfplumber

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("reference", type=Path)
parser.add_argument("candidate", type=Path)
parser.add_argument("--output-dir", type=Path, required=True)
parser.add_argument("--pdftoppm", default=shutil.which("pdftoppm"))
parser.add_argument("--dpi", type=int, default=120)
args = parser.parse_args()
if not args.pdftoppm or args.dpi <= 0:
    parser.error("Provide --pdftoppm or put Poppler on PATH, and use a positive DPI.")
root = args.output_dir
root.mkdir(parents=True, exist_ok=True)
poppler = args.pdftoppm
cases = {"comparison": args.reference}
reports = []
for name, original in cases.items():
    folder = root
    restored = args.candidate
    if not original.exists() or not restored.exists():
        parser.error("Both PDF input files must exist.")
    for label, path in [("source", original), ("returned", restored)]:
        subprocess.run(
            [str(poppler), "-r", str(args.dpi), "-png", str(path), str(folder / label)],
            check=True,
            stdout=subprocess.DEVNULL,
        )
    with pdfplumber.open(original) as a, pdfplumber.open(restored) as b:
        item = {
            "reference": str(original),
            "candidate": str(restored),
            "pages": [len(a.pages), len(b.pages)],
            "dpi": args.dpi,
            "ink_threshold": 245,
            "difference_threshold": 24,
            "page_results": [],
        }
        union_total = changed_total = intersection_total = 0
        thumbs = []
        for i in range(max(len(a.pages), len(b.pages))):
            if i >= min(len(a.pages), len(b.pages)):
                item["page_results"].append({"page": i + 1, "missing_page": True})
                for label, count in (("source", len(a.pages)), ("returned", len(b.pages))):
                    tile = Image.new("RGB", (260, 375), "#dddddd")
                    if i < count:
                        path = folder / f"{label}-{i + 1:0{len(str(count))}d}.png"
                        with Image.open(path) as page_image:
                            preview = page_image.convert("RGB")
                        preview.thumbnail((248, 350))
                        tile.paste(preview, (6, 20))
                    else:
                        ImageDraw.Draw(tile).text((30, 160), "MISSING PAGE", fill="red")
                    ImageDraw.Draw(tile).text((6, 4), f"{label} / page {i + 1}", fill="black")
                    thumbs.append(tile)
                continue
            # Poppler zero-pads to the number of digits in page count.
            pa = folder / f"source-{i + 1:0{len(str(len(a.pages)))}d}.png"
            pb = folder / f"returned-{i + 1:0{len(str(len(b.pages)))}d}.png"
            ia, ib = Image.open(pa).convert("RGB"), Image.open(pb).convert("RGB")
            row = {
                "page": i + 1,
                "dimensions_equal": ia.size == ib.size
                and (a.pages[i].width, a.pages[i].height) == (b.pages[i].width, b.pages[i].height),
            }
            row.update(
                reference_size_pt=[a.pages[i].width, a.pages[i].height],
                candidate_size_pt=[b.pages[i].width, b.pages[i].height],
                reference_size_px=list(ia.size),
                candidate_size_px=list(ib.size),
            )
            words_a, words_b = a.pages[i].extract_words(), b.pages[i].extract_words()
            row.update(
                word_counts=[len(words_a), len(words_b)],
                word_text_equal=[w["text"] for w in words_a] == [w["text"] for w in words_b],
            )
            if row["word_text_equal"]:
                row["maximum_word_position_delta_pt"] = {
                    key: max([abs(x[key] - y[key]) for x, y in zip(words_a, words_b)] or [0]) for key in ["x0", "top"]
                }
            if row["dimensions_equal"]:
                aa, bb = np.asarray(ia).astype(np.int16), np.asarray(ib).astype(np.int16)
                ink_a, ink_b = np.min(aa, axis=2) < 245, np.min(bb, axis=2) < 245
                union = ink_a | ink_b
                change = (np.max(np.abs(aa - bb), axis=2) > 24) & union
                u, c, t = int(union.sum()), int(change.sum()), int((ink_a & ink_b).sum())
                union_total += u
                changed_total += c
                intersection_total += t
                row.update(
                    pixel_exact=bool(np.array_equal(aa, bb)),
                    foreground_changed_fraction=c / u if u else None,
                    ink_iou=t / u if u else None,
                    ink_union_pixels=u,
                    changed_ink_pixels=c,
                )
                diff = np.full(aa.shape, 255, dtype=np.uint8)
                diff[union] = np.asarray(ia)[union] // 3 + 160
                diff[change] = [220, 0, 60]
                Image.fromarray(diff).save(folder / f"difference-{i + 1}.png")
            for label, im in [("source", ia), ("returned", ib)]:
                im.thumbnail((248, 350))
                tile = Image.new("RGB", (260, 375), "#dddddd")
                tile.paste(im, (6, 20))
                ImageDraw.Draw(tile).text((6, 4), f"{label} / page {i + 1}", fill="black")
                thumbs.append(tile)
            item["page_results"].append(row)
        valid = len(a.pages) == len(b.pages) and all(p.get("dimensions_equal") for p in item["page_results"])
        item.update(
            structure_equal=valid,
            foreground_changed_fraction=changed_total / union_total if union_total and valid else None,
            ink_iou=intersection_total / union_total if union_total and valid else None,
            pixel_exact=bool(valid and all(p.get("pixel_exact") for p in item["page_results"])),
        )
        sheet = Image.new("RGB", (520, 375 * ((len(thumbs) + 1) // 2)), "white")
        for j, im in enumerate(thumbs):
            sheet.paste(im, ((j % 2) * 260, (j // 2) * 375))
        sheet.save(folder / "contact.png")
    (folder / "comparison.json").write_text(json.dumps(item, indent=2), encoding="utf-8")
    reports.append(item)
    print(json.dumps({k: v for k, v in item.items() if k != "page_results"}), flush=True)
