"""Real document corpus: HTTP upload, worker edit, validate, download and reopen."""
from collections import Counter
import hashlib
import io
from pathlib import Path
from zipfile import ZipFile

from lxml import etree
import pytest

from jakal_hwpx import HwpxDocument
from server.engine import cell_text, paragraph_text
from server.policies import targets
from web_tests.test_api import client, wait

ROOT = Path(__file__).resolve().parents[1]
CORPUS = [
    "examples/samples/hwp/기출킬러양식.hwp",
    "examples/samples/hwpx/math_template_2_formula_fixed.hwpx",
    "examples/samples/hwpx/social_security_application_example.hwpx",
]


def preserved_content(raw):
    """Include master pages, merged-cell geometry and original media bytes."""
    objects, cells, binaries = Counter(), [], {}
    with ZipFile(io.BytesIO(raw)) as archive:
        for name in sorted(archive.namelist()):
            data = archive.read(name)
            if name.startswith("BinData/"):
                binaries[name] = hashlib.sha256(data).hexdigest()
            if not name.endswith(".xml"):
                continue
            root = etree.fromstring(data, parser=etree.XMLParser(resolve_entities=False, no_network=True))
            for node in root.iter():
                if not isinstance(node.tag, str):
                    continue
                tag = etree.QName(node).localname
                if tag in {"tbl", "equation", "pic", "ole", "rect", "ellipse", "fieldBegin", "fieldEnd"}:
                    objects[tag] += 1
                if tag in {"cellAddr", "cellSpan"}:
                    cells.append((name, tag, tuple(sorted(node.attrib.items()))))
    return objects, cells, binaries


def completed(client, response):
    assert response.status_code == 202, response.text
    job = wait(client, response.json()["job_id"], timeout=135)
    assert job["status"] == "completed", job


@pytest.mark.parametrize("relative", CORPUS, ids=lambda value: Path(value).name)
def test_real_document_full_roundtrip(client, tmp_path, relative):
    source = ROOT / relative
    original = source.read_bytes()
    uploaded = client.post("/api/documents", files={"file": (source.name, original)})
    completed(client, uploaded)
    base = f"/api/documents/{uploaded.json()['document_id']}"
    doc = client.get(base).json()
    assert doc["revision"] == 0 and doc["format"] == "hwpx"
    assert doc["report"]["issues"] == []
    if source.suffix == ".hwp":
        assert doc["conversion"]["source_format"] == "hwp"
        assert doc["conversion"]["count"] == 0
    assert client.get(base + "/original").content == original
    converted = client.get(base + "/source?revision=0").content
    if source.suffix == ".hwpx":
        assert converted == original
    signature = preserved_content(converted)
    assert sum(signature[0].values()) > 0

    # Pick real editable content, falling back to an existing empty paragraph
    # in math documents whose nonempty paragraphs contain protected controls.
    before = targets(doc)
    chosen = []
    for section in doc["sections"]:
        for kind in ("paragraph", "cell"):
            candidates = [t for t in before.values() if t["editable"] and t["operation"]["type"] == kind and t["operation"]["section"] == section["index"]]
            if candidates:
                chosen.append(next((t for t in candidates if t["text"].strip()), candidates[0]))
    assert chosen, f"No editable content in {source.name}"
    changes = {t["id"]: f"{t['text']}\n{t['operation']['type']} 수정 확인 & <안내> 😀\n둘째 줄" for t in chosen}
    operations = [{**t["operation"], "text": changes[t["id"]]} for t in chosen]
    assert client.post(base + "/exports", json={"revision": 0}).status_code == 409
    completed(client, client.post(base + "/operations", json={"revision": 0, "operations": operations}))
    assert client.get(base + "/download?revision=1").status_code == 409
    assert client.post(base + "/validations", json={"revision": 0}).status_code == 409
    completed(client, client.post(base + "/validations", json={"revision": 1}))
    exported = client.post(base + "/exports", json={"revision": 1})
    assert exported.status_code == 200, exported.text
    downloaded = client.get(exported.json()["url"])
    assert downloaded.status_code == 200
    assert downloaded.headers["content-type"] == "application/hwp+zip"
    assert ".hwpx" in downloaded.headers["content-disposition"]

    destination = tmp_path / f"{source.stem}_수정.hwpx"
    destination.write_bytes(downloaded.content)
    reopened = HwpxDocument.open(destination)
    assert reopened.validation_errors() == []
    assert reopened.xml_validation_errors() == []
    assert reopened.reference_validation_errors() == []
    assert preserved_content(downloaded.content) == signature
    for op in operations:
        if op["type"] == "paragraph":
            actual = paragraph_text(reopened.sections[op["section"]].paragraphs()[op["index"]]._element)
        else:
            actual = cell_text(reopened.tables(op["section"])[op["table"]].cell(op["row"], op["column"]))
        assert actual == op["text"]
    fresh = client.get(base).json()
    assert fresh["validated"] == fresh["revision"] == 1
    assert fresh["report"]["new_issues"] == []
    after = targets(fresh)
    assert before.keys() == after.keys()
    for key, target in before.items():
        assert after[key]["text"] == changes.get(key, target["text"])
    assert client.get(base + "/original").content == original
    assert client.get(base + "/source?revision=0").content == converted

    reuploaded = client.post("/api/documents", files={"file": (destination.name, downloaded.content)})
    completed(client, reuploaded)
    imported = client.get(f"/api/documents/{reuploaded.json()['document_id']}").json()
    assert imported["sections"] == fresh["sections"]
    assert imported["report"]["issues"] == []

    # A later save must revoke the URL for the previously validated version.
    again = {**operations[0], "text": operations[0]["text"] + "\n추가 수정"}
    completed(client, client.post(base + "/operations", json={"revision": 1, "operations": [again]}))
    assert client.get(exported.json()["url"]).status_code == 409
    assert client.get(base).json()["validated"] is None


def test_invalid_converted_hwp_is_rejected_without_publishing_a_broken_base(client):
    # This real HWP has placeholder OLE/shape controls. The converter reports
    # no loss but emits empty binaryItemIDRef values with no manifest targets.
    source = ROOT / "examples/samples/hwp/generated_native_control_authoring.hwp"
    original = source.read_bytes()
    response = client.post("/api/documents", files={"file": (source.name, original)})
    assert response.status_code == 202
    result = response.json()
    job = wait(client, result["job_id"], timeout=135)
    assert job["status"] == "failed", job
    assert "변환 결과" in job["error"] and "참조 오류" in job["error"]
    base = f"/api/documents/{result['document_id']}"
    assert client.get(base).json()["status"] == "failed"
    assert client.get(base + "/original").content == original
    assert client.get(base + "/source?revision=0").status_code == 409
    assert client.post(base + "/exports", json={"revision": 0}).status_code == 409
    directory = client.app.state.store.directory(result["document_id"])
    assert not (directory / "revision-0.hwpx").exists()
    assert not (directory / "conversion.json").exists()
    assert not (directory / "conversion-pending.hwpx").exists()
    assert not (directory / "conversion-pending.json").exists()
