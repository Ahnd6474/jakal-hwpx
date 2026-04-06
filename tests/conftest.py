from __future__ import annotations

import os
import sys
import zipfile
from pathlib import Path

import pytest
from lxml import etree


REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = REPO_ROOT / "src"

if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))


def _candidate_sample_roots() -> list[Path]:
    roots: list[Path] = []
    env_root = os.environ.get("JAKAL_HWPX_SAMPLE_DIR")
    if env_root:
        roots.append(Path(env_root).expanduser())
    roots.extend(
        [
            REPO_ROOT / "all_hwpx_flat",
            REPO_ROOT / "examples" / "output_smoke",
            REPO_ROOT / "examples" / "output",
            REPO_ROOT / "examples" / "samples" / "hwpx",
            REPO_ROOT,
        ]
    )
    return roots


def _valid_hwpx_files(root: Path) -> list[Path]:
    if not root.exists():
        return []
    return [path for path in sorted(root.glob("*.hwpx")) if zipfile.is_zipfile(path)]


def _discover_sample_corpus() -> tuple[Path, list[Path]]:
    checked_roots: list[str] = []
    seen: set[Path] = set()

    for root in _candidate_sample_roots():
        resolved = root.resolve()
        if resolved in seen:
            continue
        seen.add(resolved)
        checked_roots.append(str(root))
        files = _valid_hwpx_files(root)
        if files:
            return root, files

    searched = ", ".join(checked_roots)
    raise AssertionError(f"No valid HWPX samples were found. Searched: {searched}")


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
    root, _ = _discover_sample_corpus()
    return root


@pytest.fixture(scope="session")
def valid_hwpx_files(sample_corpus_dir: Path) -> list[Path]:
    files = _valid_hwpx_files(sample_corpus_dir)
    assert files, f"No valid HWPX samples were found under {sample_corpus_dir}."
    return files


@pytest.fixture(scope="session")
def sample_hwpx_path(valid_hwpx_files: list[Path]) -> Path:
    return valid_hwpx_files[0]
