from __future__ import annotations

import sys
import zipfile
from pathlib import Path

import pytest
from lxml import etree


REPO_ROOT = Path(__file__).resolve().parents[1]
SAMPLE_ROOT = REPO_ROOT / "all_hwpx_flat"
SRC_ROOT = REPO_ROOT / "src"

if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))


def _valid_hwpx_files() -> list[Path]:
    if not SAMPLE_ROOT.exists():
        return []
    return [path for path in sorted(SAMPLE_ROOT.glob("*.hwpx")) if zipfile.is_zipfile(path)]


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
def valid_hwpx_files() -> list[Path]:
    files = _valid_hwpx_files()
    assert files, "No valid HWPX samples were found under all_hwpx_flat."
    return files


@pytest.fixture(scope="session")
def sample_hwpx_path(valid_hwpx_files: list[Path]) -> Path:
    return valid_hwpx_files[0]
