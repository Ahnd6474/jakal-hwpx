from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

import pytest
from lxml import etree

from jakal_hwpx import HancomDocument, HwpxDocument, generate_hwpx_script, write_hwpx_script
from jakal_hwpx.hwpx2py import main as hwpx2py_main
from jakal_hwpx.namespaces import NS, qname


REPO_ROOT = Path(__file__).resolve().parents[1]


def _subprocess_env() -> dict[str, str]:
    env = os.environ.copy()
    src_root = str(REPO_ROOT / "src")
    env["PYTHONPATH"] = src_root if not env.get("PYTHONPATH") else f"{src_root}{os.pathsep}{env['PYTHONPATH']}"
    return env


def _append_text_run(paragraph: etree._Element, text: str) -> None:
    run = etree.SubElement(paragraph, qname("hp", "run"))
    run.set("charPrIDRef", "0")
    etree.SubElement(run, qname("hp", "t")).text = text


def _append_bookmark_run(paragraph: etree._Element, name: str) -> None:
    run = etree.SubElement(paragraph, qname("hp", "run"))
    run.set("charPrIDRef", "0")
    ctrl = etree.SubElement(run, qname("hp", "ctrl"))
    etree.SubElement(ctrl, qname("hp", "bookmark")).set("name", name)


def _append_hyperlink_field_runs(paragraph: etree._Element, display_text: str) -> None:
    begin_run = etree.SubElement(paragraph, qname("hp", "run"))
    begin_run.set("charPrIDRef", "0")
    begin_ctrl = etree.SubElement(begin_run, qname("hp", "ctrl"))
    field_begin = etree.SubElement(begin_ctrl, qname("hp", "fieldBegin"))
    field_begin.set("id", "9001")
    field_begin.set("type", "HYPERLINK")
    field_begin.set("name", "")
    field_begin.set("editable", "0")
    field_begin.set("dirty", "0")
    field_begin.set("zorder", "0")
    field_begin.set("fieldid", "19001")

    _append_text_run(paragraph, display_text)

    end_run = etree.SubElement(paragraph, qname("hp", "run"))
    end_run.set("charPrIDRef", "0")
    end_ctrl = etree.SubElement(end_run, qname("hp", "ctrl"))
    field_end = etree.SubElement(end_ctrl, qname("hp", "fieldEnd"))
    field_end.set("beginIDRef", "9001")
    field_end.set("fieldid", "19001")


def _replace_paragraph_with_bookmark_field_text(paragraph: etree._Element, prefix: str) -> None:
    for child in list(paragraph):
        paragraph.remove(child)
    _append_bookmark_run(paragraph, f"{prefix}-bookmark")
    _append_hyperlink_field_runs(paragraph, f"{prefix}-FIELD")
    _append_text_run(paragraph, f"{prefix}-TEXT")


def test_hwpx2py_writes_from_scratch_script_and_recreates_supported_blocks(tmp_path: Path) -> None:
    source_doc = HancomDocument.blank()
    source_doc.metadata.title = "hwpx2py sample"
    source_doc.append_header("Header")
    source_doc.append_paragraph("Body paragraph")
    source_doc.append_hyperlink("https://example.com", display_text="Example")
    source_doc.append_bookmark("body-anchor")
    source_doc.append_cross_reference("body-anchor", display_text="See body")
    source_doc.append_auto_number(number=7)
    source_doc.append_note("Footnote")
    source_doc.append_equation("x+y")
    source_doc.append_shape(
        kind="ellipse",
        text="Shape",
        fill_color="#EEEEEE",
        line_width=88,
        original_width=13000,
        original_height=3600,
        current_width=11000,
        current_height=3000,
    )
    source_doc.append_form("Approval", name="approval", value="pending")
    source_doc.append_memo("Memo")
    source_doc.append_chart("Revenue", categories=["Q1"], series=[{"name": "Sales", "values": [10]}])
    source_doc.append_table(rows=2, cols=2, cell_texts=[["A", "B"], ["1", "2"]])
    source_doc.append_picture(
        "sample.png",
        b"not-a-real-png",
        extension="png",
        width=2400,
        height=1600,
        original_width=3200,
        original_height=2100,
        current_width=1200,
        current_height=800,
    )
    source_doc.append_ole(
        "sample.ole",
        b"OLE-DATA",
        width=42001,
        height=13501,
        original_width=44000,
        original_height=15000,
        current_width=321,
        current_height=654,
    )

    source_path = tmp_path / "source.hwpx"
    source_doc.write_to_hwpx(source_path)

    script_path = write_hwpx_script(
        source_path,
        tmp_path / "recreate.py",
        default_output_name="recreated.hwpx",
        strict=True,
    )
    script_text = script_path.read_text(encoding="utf-8")

    assert "base64" in script_text
    assert "from_bytes" not in script_text
    assert "doc.append_picture(" in script_text
    assert "doc.append_ole(" in script_text
    assert "original_width=3200" in script_text
    assert "line_width=88" in script_text
    assert "current_height=654" in script_text
    assert "Skipped Picture asset" not in script_text
    assert "Skipped OLE asset" not in script_text
    assert "doc.append_paragraph('Body paragraph', section_index=0)" in script_text

    output_path = tmp_path / "recreated.hwpx"
    subprocess.run(
        [sys.executable, str(script_path), str(output_path)],
        check=True,
        cwd=tmp_path,
        env=_subprocess_env(),
    )

    recreated = HwpxDocument.open(output_path)
    recreated.validate()

    assert recreated.metadata().title == "hwpx2py sample"
    recreated_text = recreated.get_document_text()
    assert "Body paragraph" in recreated_text
    assert recreated_text.count("Example") == 1
    assert recreated_text.count("See body") == 1
    assert len(recreated.headers()) == 1
    assert len(recreated.tables()) == 1
    assert len(recreated.bookmarks()) == 1
    assert len(recreated.hyperlinks()) == 1
    assert len(recreated.cross_references()) == 1
    assert len(recreated.auto_numbers()) == 1
    assert len(recreated.notes()) == 1
    assert len(recreated.equations()) == 1
    assert len(recreated.forms()) == 1
    assert len(recreated.memos()) == 1
    assert len(recreated.charts()) == 1
    assert recreated.pictures()[0].binary_data() == b"not-a-real-png"
    assert recreated.oles()[0].binary_data() == b"OLE-DATA"
    assert recreated.shapes()[0].line_style()["width"] == "88"
    assert recreated.shapes()[0].element.xpath("./hp:orgSz/@width", namespaces=NS) == ["13000"]
    assert recreated.shapes()[0].element.xpath("./hp:curSz/@height", namespaces=NS) == ["3000"]
    assert recreated.pictures()[0].element.xpath("./hp:orgSz/@width", namespaces=NS) == ["3200"]
    assert recreated.pictures()[0].element.xpath("./hp:curSz/@height", namespaces=NS) == ["800"]
    assert recreated.oles()[0].element.xpath("./hp:orgSz/@width", namespaces=NS) == ["44000"]
    assert recreated.oles()[0].element.xpath("./hp:curSz/@height", namespaces=NS) == ["654"]


def test_hwpx2py_authoring_mode_emits_editable_dsl_script(tmp_path: Path) -> None:
    source_doc = HancomDocument.blank()
    source_doc.metadata.title = "authoring sample"
    source_doc.append_paragraph("Body paragraph")
    source_doc.append_equation(r"\frac{1}{2}+x")
    source_doc.append_table(rows=1, cols=2, cell_texts=[["A", "B"]])
    source_doc.append_shape(kind="rect", text="Shape", line_width=88)

    source_path = tmp_path / "authoring_source.hwpx"
    source_doc.write_to_hwpx(source_path)

    script_path = write_hwpx_script(
        source_path,
        tmp_path / "recreate_authoring.py",
        default_output_name="authoring_recreated.hwpx",
        mode="authoring",
    )
    script_text = script_path.read_text(encoding="utf-8")

    assert "compact authoring DSL" in script_text
    assert "def control(doc: HancomDocument, name: str" in script_text
    assert "return control(doc, 'equation', script" in script_text
    assert "def p(" in script_text
    assert "def eq(" in script_text
    assert "def section_1(doc: HancomDocument) -> None:" in script_text
    assert "p(doc, 'Body paragraph')" in script_text
    assert "eq(doc, r'\\frac{1}{2}+x')" in script_text
    assert "table(doc, [['A', 'B']])" in script_text
    assert "shape(doc, text='Shape', line_width=88)" in script_text

    output_path = tmp_path / "authoring_recreated.hwpx"
    subprocess.run(
        [sys.executable, str(script_path), str(output_path)],
        check=True,
        cwd=tmp_path,
        env=_subprocess_env(),
    )

    recreated = HwpxDocument.open(output_path)
    recreated.validate()
    assert recreated.get_document_text().count("Body paragraph") == 1
    assert len(recreated.equations()) == 1
    assert len(recreated.tables()) == 1
    assert recreated.shapes()[0].line_style()["width"] == "88"


def test_hwpx2py_macro_mode_emits_latex_like_problem_tree(tmp_path: Path) -> None:
    source_doc = HancomDocument.blank()
    source_doc.metadata.title = "macro sample"
    source_doc.append_paragraph("Body paragraph")
    source_doc.append_equation(r"\frac{1}{2}+x")
    source_doc.append_table(rows=1, cols=2, cell_texts=[["A", "B"]])
    source_doc.append_shape(kind="rect", text="", line_width=88)

    source_path = tmp_path / "macro_source.hwpx"
    source_doc.write_to_hwpx(source_path)

    script_path = write_hwpx_script(
        source_path,
        tmp_path / "recreate_macro.py",
        default_output_name="macro_recreated.hwpx",
        mode="macro",
    )
    script_text = script_path.read_text(encoding="utf-8")

    assert "LaTeX-like macro DSL" in script_text
    assert "def document(doc: HancomDocument, *nodes, section: int = 0):" in script_text
    assert "def problem(*items" in script_text
    assert "text('Body paragraph'" in script_text
    assert "math(r'\\frac{1}{2}+x'" in script_text
    assert "tabular([['A', 'B']]" in script_text
    assert "draw(line_width=88" in script_text

    output_path = tmp_path / "macro_recreated.hwpx"
    subprocess.run(
        [sys.executable, str(script_path), str(output_path)],
        check=True,
        cwd=tmp_path,
        env=_subprocess_env(),
    )

    recreated = HwpxDocument.open(output_path)
    recreated.validate()
    assert recreated.get_document_text().count("Body paragraph") == 1
    assert len(recreated.equations()) == 1
    assert len(recreated.tables()) == 1
    assert recreated.shapes()[0].line_style()["width"] == "88"


def test_hwpx2py_can_still_skip_binary_assets(tmp_path: Path) -> None:
    source_doc = HancomDocument.blank()
    source_doc.append_picture("sample.png", b"PNG", extension="png")
    source_doc.append_ole("sample.ole", b"OLE")
    source_path = tmp_path / "source.hwpx"
    source_doc.write_to_hwpx(source_path)

    script_text = generate_hwpx_script(source_path, include_binary_assets=False)

    assert "base64" not in script_text
    assert "doc.append_picture(" not in script_text
    assert "doc.append_ole(" not in script_text
    assert "Skipped Picture asset" in script_text
    assert "Skipped OLE asset" in script_text


def test_hwpx2py_preserves_absent_picture_size_and_line_nodes(tmp_path: Path) -> None:
    source_doc = HwpxDocument.blank()
    picture = source_doc.append_picture("no-size.png", b"PNG", media_type="image/png", width=2400, height=1600)
    for node in picture.element.xpath("./hp:sz | ./hp:pos | ./hp:outMargin", namespaces=NS):
        picture.element.remove(node)
    picture.section.mark_modified()

    source_path = tmp_path / "source.hwpx"
    source_doc.save(source_path)
    script_path = write_hwpx_script(source_path, tmp_path / "recreate.py")
    script_text = script_path.read_text(encoding="utf-8")

    assert "block.has_size_node = False" in script_text
    assert "block.has_position_node = False" in script_text
    assert "block.has_out_margin_node = False" in script_text
    assert "block.has_line_node" not in script_text

    output_path = tmp_path / "recreated.hwpx"
    subprocess.run(
        [sys.executable, str(script_path), str(output_path)],
        check=True,
        cwd=tmp_path,
        env=_subprocess_env(),
    )

    recreated = HwpxDocument.open(output_path)
    recreated.validate()
    recreated_picture = recreated.pictures()[0]
    assert recreated_picture.element.xpath("./hp:sz", namespaces=NS) == []
    assert recreated_picture.element.xpath("./hp:pos", namespaces=NS) == []
    assert recreated_picture.element.xpath("./hp:outMargin", namespaces=NS) == []
    assert recreated_picture.element.xpath("./hp:lineShape", namespaces=NS) == []


def test_hwpx2py_raw_mode_recreates_embedded_package_parts(tmp_path: Path) -> None:
    source_doc = HwpxDocument.blank()
    source_doc.set_paragraph_text(0, 0, "앞 문장")
    source_doc.append_hyperlink("https://example.com", display_text="링크", paragraph_index=0)
    source_doc.append_table(1, 2, cell_texts=[["A", "B"]], paragraph_index=0)
    source_doc.append_picture(
        "inline.png",
        b"PNG-DATA",
        media_type="image/png",
        width=2400,
        height=1600,
        paragraph_index=0,
    )
    source_doc.append_paragraph("다음 문단")

    source_path = tmp_path / "source.hwpx"
    source_doc.save(source_path)

    script_path = write_hwpx_script(source_path, tmp_path / "recreate_raw.py", mode="raw")
    script_text = script_path.read_text(encoding="utf-8")

    assert "from jakal_hwpx import HwpxDocument" in script_text
    assert "def build_document() -> HwpxDocument:" in script_text
    assert "HancomDocument" not in script_text
    assert "parts.append(('Contents/section0.xml', part_data))" in script_text
    assert "doc.add_part(part_path, part_data)" in script_text

    output_path = tmp_path / "recreated.hwpx"
    subprocess.run(
        [sys.executable, str(script_path), str(output_path)],
        check=True,
        cwd=tmp_path,
        env=_subprocess_env(),
    )

    source = HwpxDocument.open(source_path)
    recreated = HwpxDocument.open(output_path)
    recreated.validate()

    assert recreated.list_part_paths() == source.list_part_paths()
    for part_path in source.list_part_paths():
        assert recreated.get_part(part_path).raw_bytes == source.get_part(part_path).raw_bytes
    assert recreated.paragraph_xml(0, 0) == source.paragraph_xml(0, 0)


def test_hwpx2py_raw_mode_requires_embedded_binary_assets(tmp_path: Path) -> None:
    source_doc = HwpxDocument.blank()
    source_path = tmp_path / "source.hwpx"
    source_doc.save(source_path)

    with pytest.raises(ValueError, match="raw mode requires"):
        generate_hwpx_script(source_path, mode="raw", include_binary_assets=False)


def test_hwpx2py_semantic_mode_preserves_source_paragraph_grouping(tmp_path: Path) -> None:
    source_doc = HwpxDocument.blank()
    source_doc.set_paragraph_text(0, 0, "앞")
    source_doc.append_hyperlink("https://example.com", display_text="링크", paragraph_index=0)
    source_doc.append_table(1, 1, cell_texts=[["셀"]], paragraph_index=0)
    source_doc.append_picture("inline.png", b"PNG", media_type="image/png", paragraph_index=0)
    source_doc.append_paragraph("뒤")
    source_path = tmp_path / "source.hwpx"
    source_doc.save(source_path)

    script_path = write_hwpx_script(source_path, tmp_path / "recreate.py")
    script_text = script_path.read_text(encoding="utf-8")

    assert "block.source_paragraph_index = 0" in script_text
    assert "block.source_paragraph_index = 1" in script_text

    output_path = tmp_path / "recreated.hwpx"
    subprocess.run(
        [sys.executable, str(script_path), str(output_path)],
        check=True,
        cwd=tmp_path,
        env=_subprocess_env(),
    )

    source = HwpxDocument.open(source_path)
    recreated = HwpxDocument.open(output_path)
    recreated.validate()

    assert recreated.paragraph_count(0) == source.paragraph_count(0)
    assert recreated.sections[0].text_fragments() == source.sections[0].text_fragments()
    assert len(recreated.tables()) == 1
    assert len(recreated.pictures()) == 1
    assert len(recreated.hyperlinks()) == 1


def test_hwpx2py_semantic_mode_preserves_table_caption_text(tmp_path: Path) -> None:
    source_doc = HwpxDocument.blank()
    source_doc.set_paragraph_text(0, 0, "")
    table = source_doc.append_table(1, 1, cell_texts=[["CELL"]], paragraph_index=0)

    caption = etree.Element(f"{{{NS['hp']}}}caption", side="BOTTOM", fullSz="0", width="1000", gap="0", lastWidth="1000")
    sub_list = etree.SubElement(caption, f"{{{NS['hp']}}}subList")
    paragraph = etree.SubElement(sub_list, f"{{{NS['hp']}}}p", id="1", paraPrIDRef="0", styleIDRef="0")
    run = etree.SubElement(paragraph, f"{{{NS['hp']}}}run", charPrIDRef="0")
    text = etree.SubElement(run, f"{{{NS['hp']}}}t")
    text.text = "Table . caption"
    in_margin = table.element.xpath("./hp:inMargin[1]", namespaces=NS)[0]
    in_margin.addprevious(caption)
    table.section.mark_modified()

    source_path = tmp_path / "source.hwpx"
    source_doc.save(source_path)
    script_path = write_hwpx_script(source_path, tmp_path / "recreate.py")

    output_path = tmp_path / "recreated.hwpx"
    subprocess.run(
        [sys.executable, str(script_path), str(output_path)],
        check=True,
        cwd=tmp_path,
        env=_subprocess_env(),
    )

    source = HwpxDocument.open(source_path)
    recreated = HwpxDocument.open(output_path)

    assert recreated.sections[0].text_fragments() == source.sections[0].text_fragments()


def test_hwpx2py_semantic_mode_keeps_textart_text_out_of_body_text(tmp_path: Path) -> None:
    source_doc = HwpxDocument.blank()
    source_doc.set_paragraph_text(0, 0, "")
    source_doc.append_shape(kind="textart", text="TEXTART", paragraph_index=0)

    source_path = tmp_path / "source.hwpx"
    source_doc.save(source_path)
    script_path = write_hwpx_script(source_path, tmp_path / "recreate.py")

    output_path = tmp_path / "recreated.hwpx"
    subprocess.run(
        [sys.executable, str(script_path), str(output_path)],
        check=True,
        cwd=tmp_path,
        env=_subprocess_env(),
    )

    source = HwpxDocument.open(source_path)
    recreated = HwpxDocument.open(output_path)

    assert source.shapes()[0].text == "TEXTART"
    assert recreated.shapes()[0].text == "TEXTART"
    assert source.sections[0].text_fragments() == [""]
    assert recreated.sections[0].text_fragments() == source.sections[0].text_fragments()


def test_hwpx2py_semantic_mode_preserves_legacy_textart_draw_text(tmp_path: Path) -> None:
    source_doc = HwpxDocument.blank()
    source_doc.set_paragraph_text(0, 0, "")
    textart = source_doc.append_shape(kind="textart", text="", paragraph_index=0)
    textart.element.attrib.pop("text", None)

    draw_text = etree.SubElement(textart.element, f"{{{NS['hp']}}}drawText", lastWidth="12000", name="", editable="0")
    sub_list = etree.SubElement(draw_text, f"{{{NS['hp']}}}subList")
    paragraph = etree.SubElement(sub_list, f"{{{NS['hp']}}}p", id="1", paraPrIDRef="0", styleIDRef="0")
    run = etree.SubElement(paragraph, f"{{{NS['hp']}}}run", charPrIDRef="0")
    text = etree.SubElement(run, f"{{{NS['hp']}}}t")
    text.text = "LEGACY"
    etree.SubElement(draw_text, f"{{{NS['hp']}}}textMargin", left="283", right="283", top="283", bottom="283")
    textart.section.mark_modified()

    source_path = tmp_path / "source.hwpx"
    source_doc.save(source_path)
    script_path = write_hwpx_script(source_path, tmp_path / "recreate.py")

    output_path = tmp_path / "recreated.hwpx"
    subprocess.run(
        [sys.executable, str(script_path), str(output_path)],
        check=True,
        cwd=tmp_path,
        env=_subprocess_env(),
    )

    source = HwpxDocument.open(source_path)
    recreated = HwpxDocument.open(output_path)

    assert source.sections[0].text_fragments() == ["LEGACY"]
    assert recreated.sections[0].text_fragments() == source.sections[0].text_fragments()


def test_hwpx2py_semantic_mode_does_not_duplicate_nested_shape_text(tmp_path: Path) -> None:
    source_doc = HwpxDocument.blank()
    source_doc.set_paragraph_text(0, 0, "")
    container = source_doc.append_shape(kind="container", text="", paragraph_index=0)
    child = source_doc.append_shape(kind="rect", text="INNER", paragraph_index=0)

    child_run = child.element.getparent()
    target_runs = container.element.xpath("./hp:drawText/hp:subList/hp:p/hp:run", namespaces=NS)
    assert target_runs
    target_runs[0].append(child.element)
    if child_run is not None and len(child_run) == 0:
        parent = child_run.getparent()
        if parent is not None:
            parent.remove(child_run)
    container.section.mark_modified()

    source_path = tmp_path / "source.hwpx"
    source_doc.save(source_path)
    script_path = write_hwpx_script(source_path, tmp_path / "recreate.py")

    output_path = tmp_path / "recreated.hwpx"
    subprocess.run(
        [sys.executable, str(script_path), str(output_path)],
        check=True,
        cwd=tmp_path,
        env=_subprocess_env(),
    )

    source = HwpxDocument.open(source_path)
    recreated = HwpxDocument.open(output_path)

    assert len(source.shapes()) == 2
    assert len(recreated.shapes()) == 2
    assert source.sections[0].text_fragments() == ["INNER"]
    assert recreated.sections[0].text_fragments() == source.sections[0].text_fragments()


@pytest.mark.parametrize("container", ["header", "endnote", "shape"])
def test_hwpx2py_semantic_mode_preserves_nested_field_text_once(container: str, tmp_path: Path) -> None:
    source_doc = HwpxDocument.blank()
    source_doc.set_paragraph_text(0, 0, "")
    prefix = f"{container}-nested"

    if container == "header":
        owner = source_doc.append_header("", paragraph_index=0)
        paragraph = owner.element.xpath("./hp:subList/hp:p[1]", namespaces=NS)[0]
    elif container == "endnote":
        owner = source_doc.append_note("", kind="endNote", paragraph_index=0)
        paragraph = owner.element.xpath("./hp:subList/hp:p[1]", namespaces=NS)[0]
    else:
        owner = source_doc.append_shape(kind="rect", text="", paragraph_index=0)
        paragraph = owner.element.xpath("./hp:drawText/hp:subList/hp:p[1]", namespaces=NS)[0]

    _replace_paragraph_with_bookmark_field_text(paragraph, prefix)
    owner.section.mark_modified()

    source_path = tmp_path / f"{container}.hwpx"
    source_doc.save(source_path)
    script_path = write_hwpx_script(source_path, tmp_path / "recreate.py")
    script_text = script_path.read_text(encoding="utf-8")

    assert "block.nested_blocks = []" in script_text

    output_path = tmp_path / "recreated.hwpx"
    subprocess.run(
        [sys.executable, str(script_path), str(output_path)],
        check=True,
        cwd=tmp_path,
        env=_subprocess_env(),
    )

    source = HwpxDocument.open(source_path)
    recreated = HwpxDocument.open(output_path)

    assert recreated.sections[0].text_fragments() == source.sections[0].text_fragments()
    assert recreated.get_document_text() == source.get_document_text()
    assert len(recreated.bookmarks()) == len(source.bookmarks()) == 1
    assert len(recreated.hyperlinks()) == len(source.hyperlinks()) == 1
    assert len(recreated.headers()) == len(source.headers())
    assert len(recreated.notes()) == len(source.notes())
    assert len(recreated.shapes()) == len(source.shapes())
    if container == "header":
        assert recreated.headers()[0].element.xpath(".//hp:fieldBegin", namespaces=NS)
    elif container == "endnote":
        assert recreated.notes()[0].element.xpath(".//hp:fieldBegin", namespaces=NS)
    else:
        assert recreated.shapes()[0].element.xpath(".//hp:fieldBegin", namespaces=NS)


def test_hancom_document_write_hwpx_does_not_duplicate_source_field_text(tmp_path: Path) -> None:
    source_doc = HwpxDocument.blank()
    source_doc.set_paragraph_text(0, 0, "")
    source_doc.append_paragraph("", section_index=0)
    _replace_paragraph_with_bookmark_field_text(source_doc.sections[0].root_element.xpath("./hp:p[2]", namespaces=NS)[0], "section")
    source_doc.sections[0].mark_modified()

    source_path = tmp_path / "source.hwpx"
    output_path = tmp_path / "recreated.hwpx"
    source_doc.save(source_path)

    HancomDocument.read_hwpx(source_path).write_to_hwpx(output_path)

    source = HwpxDocument.open(source_path)
    recreated = HwpxDocument.open(output_path)

    assert recreated.sections[0].text_fragments() == source.sections[0].text_fragments()
    assert recreated.get_document_text() == source.get_document_text()
    assert len(recreated.hyperlinks()) == len(source.hyperlinks()) == 1


def test_hwpx2py_cli_writes_script(tmp_path: Path) -> None:
    source_doc = HancomDocument.blank()
    source_doc.append_paragraph("CLI paragraph")
    source_path = tmp_path / "source.hwpx"
    script_path = tmp_path / "source_to_py.py"
    source_doc.write_to_hwpx(source_path)

    result = hwpx2py_main(
        [
            str(source_path),
            "-o",
            str(script_path),
            "--default-output",
            "cli-output.hwpx",
            "--strict",
            "--skip-binary-assets",
        ]
    )

    assert result == 0
    script_text = script_path.read_text(encoding="utf-8")
    assert "_DEFAULT_OUTPUT = 'cli-output.hwpx'" in script_text
    assert "def build_document() -> HancomDocument:" in script_text
    assert "doc.append_paragraph('CLI paragraph', section_index=0)" in script_text


def test_hwpx2py_cli_writes_raw_mode_script(tmp_path: Path) -> None:
    source_doc = HancomDocument.blank()
    source_doc.append_paragraph("CLI raw paragraph")
    source_path = tmp_path / "source.hwpx"
    script_path = tmp_path / "source_raw_to_py.py"
    source_doc.write_to_hwpx(source_path)

    result = hwpx2py_main(
        [
            str(source_path),
            "-o",
            str(script_path),
            "--mode",
            "raw",
        ]
    )

    assert result == 0
    script_text = script_path.read_text(encoding="utf-8")
    assert "def build_document() -> HwpxDocument:" in script_text
    assert "parts.append(('Contents/section0.xml', part_data))" in script_text
