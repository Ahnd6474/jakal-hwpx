from __future__ import annotations

import itertools
import os
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

import pytest
from lxml import etree


REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = REPO_ROOT / "src"
SAMPLE_CORPUS_DIR = REPO_ROOT / "examples" / "samples" / "hwpx"

if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

_TMP_PATH_COUNTER = itertools.count()
collect_ignore_glob = ["pytest-cache-files-*", "_pytest_cache*", "_pytest_basetemp*", "_tmp_pytest*", "_release_tmp*"]


def _valid_hwpx_files(root: Path) -> list[Path]:
    if not root.exists():
        return []
    return [path for path in sorted(root.glob("*.hwpx")) if zipfile.is_zipfile(path)]


def find_sample_with_section_xpath(files: list[Path], expression: str) -> Path:
    namespaces = {"hp": "http://www.hancom.co.kr/hwpml/2011/paragraph"}
    for path in files:
        with zipfile.ZipFile(path) as zf:
            for name in [n for n in zf.namelist() if n.startswith("Contents/section") and n.endswith(".xml")]:
                data = zf.read(name)
                if not data.lstrip().startswith(b"<"):
                    continue
                root = etree.fromstring(data)
                if root.xpath(expression, namespaces=namespaces):
                    return path
    raise LookupError(f"No sample matched xpath: {expression}")


@pytest.fixture(scope="session")
def sample_corpus_dir() -> Path:
    assert SAMPLE_CORPUS_DIR.exists(), f"Sample corpus directory is missing: {SAMPLE_CORPUS_DIR}"
    return SAMPLE_CORPUS_DIR


@pytest.fixture(scope="session")
def valid_hwpx_files(sample_corpus_dir: Path) -> list[Path]:
    files = _valid_hwpx_files(sample_corpus_dir)
    assert files, f"No valid HWPX samples were found under the fixed test corpus: {sample_corpus_dir}."
    return files


@pytest.fixture(scope="session")
def sample_hwpx_path(valid_hwpx_files: list[Path]) -> Path:
    return valid_hwpx_files[0]


def _tmp_dir_name(nodeid: str) -> str:
    sanitized = re.sub(r"[^A-Za-z0-9_.-]+", "-", nodeid).strip("-")
    return sanitized[:80] or "tmp"


@pytest.fixture(scope="session")
def _workspace_tmp_root() -> Path:
    root = REPO_ROOT / "tests" / "_tmp_pytest"
    shutil.rmtree(root, ignore_errors=True)
    root.mkdir(parents=True, exist_ok=True)
    return root


@pytest.fixture(scope="session", autouse=True)
def _workspace_temp_env(_workspace_tmp_root: Path) -> None:
    temp_root = _workspace_tmp_root / "session-temp"
    temp_root.mkdir(parents=True, exist_ok=True)
    original_env = {key: os.environ.get(key) for key in ("TMP", "TEMP", "TMPDIR", "JAKAL_HWPX_TEMP_ROOT")}
    original_tempdir = tempfile.tempdir
    os.environ["TMP"] = str(temp_root)
    os.environ["TEMP"] = str(temp_root)
    os.environ["TMPDIR"] = str(temp_root)
    os.environ["JAKAL_HWPX_TEMP_ROOT"] = str(temp_root)
    tempfile.tempdir = str(temp_root)
    try:
        yield
    finally:
        for key, value in original_env.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value
        tempfile.tempdir = original_tempdir


@pytest.fixture
def tmp_path(_workspace_tmp_root: Path, request: pytest.FixtureRequest) -> Path:
    path = _workspace_tmp_root / f"{next(_TMP_PATH_COUNTER):04d}-{_tmp_dir_name(request.node.nodeid)}"
    path.mkdir(parents=True, exist_ok=False)
    return path
