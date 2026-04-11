from __future__ import annotations

import base64
import sys
import zipfile
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = REPO_ROOT / "src"
SAMPLE_DIR = REPO_ROOT / "examples" / "samples" / "hwpx"

if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from jakal_hwpx import HwpxDocument  # noqa: E402


PNG_1X1 = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO8B3ioAAAAASUVORK5CYII="
)
OLE_SAMPLE_PATH = REPO_ROOT / "ole_test.hwpx"
UNICODE_STRESS_TEXT = "안성민은 자칼이다つㄹשׁשׁㄹㅇㄹㅇ 😱🎅🏦🌩⏩💵🔥"


def _extract_first_ole_payload(sample_path: Path) -> bytes | None:
    if not sample_path.exists() or not zipfile.is_zipfile(sample_path):
        return None
    document = HwpxDocument.open(sample_path)
    oles = document.oles()
    if not oles:
        return None
    return oles[0].binary_data()


def build_feature_matrix_sample(output_path: Path) -> Path:
    document = HwpxDocument.blank()
    document.set_metadata(
        title="Generated Feature Matrix",
        creator="jakal_hwpx",
        description="Synthetic corpus sample covering header/footer, notes, bookmark, hyperlink, table, picture, shape, numbering, and equation.",
        keyword="generated,corpus,feature-matrix",
    )
    document.set_preview_text("Generated sample corpus fixture for missing feature coverage.")

    para_style = document.append_paragraph_style(alignment_horizontal="CENTER", line_spacing=180)
    para_style.set_margin(left=120, right=240, prev=360, next=480, intent=600)
    para_style.set_break_setting(keep_with_next=True, keep_lines=True, page_break_before=True, line_wrap="BREAK")
    char_style = document.append_character_style(text_color="#1F3A5F", height=1200)
    char_style.set_font_refs(hangul="1", latin="1")
    char_style.set_relative_shape("spacing", hangul=5, latin=6)
    char_style.set_relative_shape("ratio", hangul=95, latin=96)
    char_style.set_relative_shape("offset", hangul=2, latin=3)
    style = document.append_style(
        "Generated Matrix Style",
        english_name="Generated Matrix Style",
        para_pr_id=para_style.style_id,
        char_pr_id=char_style.style_id,
    )
    style.configure(next_style_id=style.style_id, lang_id="1033", lock_form=True)

    document.set_paragraph_text(0, 0, "Generated feature matrix sample")
    document.append_paragraph(
        "Synthetic corpus fixture",
        style_id=style.style_id,
        para_pr_id=para_style.style_id,
        char_pr_id=char_style.style_id,
    )
    document.append_header("Generated header block", apply_page_type="EVEN", hide_first=True)
    document.append_footer("Generated footer block", apply_page_type="ODD", hide_first=False)
    document.section_settings(0).set_margins(header=321, footer=654)
    document.section_settings(0).set_visibility(hide_first_header=False, hide_first_footer=True)
    document.append_footnote("Generated footnote", number=1)
    document.append_endnote("Generated endnote", number=1)
    document.append_bookmark("generated_anchor")
    document.append_hyperlink("https://example.com/generated", display_text="Generated Link")
    document.append_mail_merge_field("generated_name", display_text="GENERATED_NAME")
    document.append_calculation_field("40+2", display_text="42")
    document.append_cross_reference("generated_anchor", display_text="Generated Anchor Ref")
    document.append_auto_number(number=5, number_type="PAGE", kind="newNum")
    document.append_auto_number(number=2, number_type="FIGURE", kind="autoNum")
    document.append_equation(
        "x=1+2",
        shape_comment="Generated equation",
        treat_as_char=False,
        allow_overlap=True,
        vert_align="CENTER",
        horz_align="RIGHT",
        vert_offset=77,
        horz_offset=88,
        out_margin_left=13,
        out_margin_right=26,
    )
    document.append_table(
        2,
        2,
        cell_texts=[["R1C1", "R1C2"], ["R2C1", "R2C2"]],
        section_index=0,
        treat_as_char=False,
        vert_align="CENTER",
        horz_align="RIGHT",
        vert_offset=120,
        horz_offset=240,
        out_margin_left=10,
        out_margin_right=20,
    )
    document.append_picture(
        "generated-pixel.png",
        PNG_1X1,
        width=2400,
        height=2400,
        section_index=0,
        shape_comment="Generated picture",
        treat_as_char=False,
        allow_overlap=True,
        vert_align="CENTER",
        horz_align="RIGHT",
        vert_offset=333,
        horz_offset=444,
        out_margin_left=11,
        out_margin_right=22,
        out_margin_top=33,
        out_margin_bottom=44,
    )
    document.append_shape(
        kind="rect",
        text="Generated rect",
        width=9000,
        height=2400,
        shape_comment="Generated rect",
        treat_as_char=False,
        allow_overlap=True,
        vert_align="BOTTOM",
        horz_align="CENTER",
        vert_offset=555,
        horz_offset=666,
        out_margin_left=12,
        out_margin_right=24,
    )
    document.append_shape(kind="line", width=6000, height=1200)
    document.append_shape(kind="textart", text="Generated textart", width=9000, height=2400)

    ole_payload = _extract_first_ole_payload(OLE_SAMPLE_PATH)
    if ole_payload is not None:
        document.append_ole(
            "generated-sample.ole",
            ole_payload,
            width=42001,
            height=13501,
            shape_comment="Generated OLE",
            allow_overlap=True,
            vert_align="CENTER",
            horz_align="RIGHT",
            vert_offset=111,
            horz_offset=222,
            out_margin_left=10,
            out_margin_right=20,
            out_margin_top=30,
            out_margin_bottom=40,
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    return output_path


def build_hard_example_like_sample(output_path: Path) -> Path:
    document = HwpxDocument.blank()
    document.set_metadata(
        title="Generated Hard Example Like",
        creator="jakal_hwpx",
        description="Synthetic hard-example-like sample covering mixed controls, unicode-rich text, and embedded objects.",
        keyword="generated,hard-example-like,unicode,ole,picture,equation",
    )
    document.set_preview_text("Generated hard-example-like fixture with unicode-rich body text and mixed controls.")

    document.set_paragraph_text(0, 0, f" {UNICODE_STRESS_TEXT}")
    document.append_paragraph("너 안성민")
    document.append_paragraph("")
    document.append_paragraph("너 안성민은 안성민이 아니다")

    document.append_footnote(f"각주: {UNICODE_STRESS_TEXT}", number=1)
    document.append_hyperlink("https://example.com/hard", display_text="hard-link")
    document.append_auto_number(number=1, number_type="FIGURE", kind="autoNum")
    document.append_equation("x=1+2", shape_comment="Hard-like equation", treat_as_char=False, allow_overlap=True)
    document.append_picture(
        "hard-like.bmp",
        PNG_1X1,
        width=4200,
        height=4200,
        shape_comment="Hard-like picture",
        treat_as_char=False,
        allow_overlap=True,
    )
    document.append_shape(
        kind="rect",
        text=f"rect: {UNICODE_STRESS_TEXT}",
        width=12000,
        height=2600,
        shape_comment="Hard-like rect",
        treat_as_char=False,
        allow_overlap=True,
    )
    document.append_shape(
        kind="textart",
        text="너 안성민",
        width=9000,
        height=2400,
        shape_comment="Hard-like textart",
    )

    ole_payload = _extract_first_ole_payload(OLE_SAMPLE_PATH)
    if ole_payload is not None:
        document.append_ole(
            "hard-like.ole",
            ole_payload,
            width=42001,
            height=13501,
            shape_comment="Hard-like OLE",
            allow_overlap=False,
            treat_as_char=False,
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    return output_path


def main() -> int:
    outputs = [
        build_feature_matrix_sample(SAMPLE_DIR / "generated_feature_matrix.hwpx"),
        build_hard_example_like_sample(SAMPLE_DIR / "generated_hard_example_like.hwpx"),
    ]
    for output in outputs:
        print(f"[ok] {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
