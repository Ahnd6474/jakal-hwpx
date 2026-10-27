import shutil
from pathlib import Path

import pytest

from jakal_hwpx import HwpxDocument
from server.engine import process

ROOT = Path(__file__).resolve().parents[1]


@pytest.mark.parametrize("name", ["generated_feature_matrix.hwpx", "math_template_2_formula_fixed.hwpx", "social_security_application_example.hwpx"])
def test_real_corpus_analysis_and_noop_validation(tmp_path, name):
    source = ROOT / "examples" / "samples" / "hwpx" / name
    destination = tmp_path / "revision-0.hwpx"
    shutil.copyfile(source, destination)
    result = process(tmp_path, "analyze", {}, name)
    assert result["report"]["new_issues"] == []
    assert process(tmp_path, "validate", {"revision": 0}, name)["revision"] == 0
    assert destination.read_bytes() == source.read_bytes()


def test_merged_table_edit_preserves_span_and_other_cells(tmp_path):
    document = HwpxDocument.blank()
    table = document.append_table(rows=3, columns=3, cell_texts=[["병합 제목", "", ""], ["A", "B", "C"], ["D", "E", "F"]])
    table.merge_cells(0, 0, 0, 2)
    document.save(tmp_path / "revision-0.hwpx")
    process(tmp_path, "analyze", {}, "merged.hwpx")
    result = process(tmp_path, "edit", {"revision": 0, "operations": [{"type": "cell", "section": 0, "table": 0, "row": 0, "column": 0, "text": "수정한 병합 제목"}]}, "merged.hwpx")
    assert result["revision"] == 1
    reopened = HwpxDocument.open(tmp_path / "revision-1.hwpx")
    assert reopened.tables()[0].cell(0, 0).col_span == 3
    assert reopened.tables()[0].cell(0, 0).text == "수정한 병합 제목"
    assert reopened.tables()[0].cell(1, 1).text == "B"


def test_multiline_cell_can_be_replaced_with_single_line_and_cleared(tmp_path):
    from server.engine import cell_text
    document = HwpxDocument.blank()
    table = document.append_table(rows=1, columns=1, cell_texts=[["첫 줄"]])
    table.cell(0, 0).set_text("첫 줄\n둘째 줄\n셋째 줄")
    document.save(tmp_path / "revision-0.hwpx")
    process(tmp_path, "analyze", {}, "multiline.hwpx")
    for revision, text in enumerate(["한 줄로 바꾸기", "", "다시\n여러 줄\n"]):
        process(tmp_path, "edit", {"revision": revision, "operations": [{"type": "cell", "section": 0, "table": 0, "row": 0, "column": 0, "text": text}]}, "multiline.hwpx")
        reopened = HwpxDocument.open(tmp_path / f"revision-{revision + 1}.hwpx")
        assert cell_text(reopened.tables()[0].cell(0, 0)) == text
