from __future__ import annotations

import re
import shutil
import subprocess
import sys
import tarfile
import zipfile
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
DIST_DIR = REPO_ROOT / "dist"
PYPROJECT = REPO_ROOT / "pyproject.toml"
INIT_FILE = REPO_ROOT / "src" / "jakal_hwpx" / "__init__.py"
REQUIRED_WHEEL_PATTERNS = (
    "jakal_hwpx/__init__.py",
    "jakal_hwpx/document.py",
    "jakal_hwpx/elements.py",
    "jakal_hwpx/parts.py",
    "jakal_hwpx/py.typed",
)
REQUIRED_SDIST_PATTERNS = tuple(f"src/{path}" for path in REQUIRED_WHEEL_PATTERNS)
OPTIONAL_TOP_LEVEL_FILES = ("LICENSE", "THIRD_PARTY_NOTICES.md")


def extract_version(pattern: str, text: str, source: Path) -> str:
    match = re.search(pattern, text, re.MULTILINE)
    if match is None:
        raise RuntimeError(f"Could not find version in {source}")
    return match.group(1)


def read_versions() -> str:
    pyproject_version = extract_version(r'^version = "([^"]+)"$', PYPROJECT.read_text(encoding="utf-8"), PYPROJECT)
    init_version = extract_version(r'^__version__ = "([^"]+)"$', INIT_FILE.read_text(encoding="utf-8"), INIT_FILE)
    if pyproject_version != init_version:
        raise RuntimeError(
            f"Version mismatch: pyproject.toml has {pyproject_version}, src/jakal_hwpx/__init__.py has {init_version}"
        )
    return pyproject_version


def run(*args: str) -> None:
    print("+", " ".join(args))
    subprocess.run(args, cwd=REPO_ROOT, check=True)


def rebuild_dist() -> list[Path]:
    shutil.rmtree(DIST_DIR, ignore_errors=True)
    DIST_DIR.mkdir(parents=True, exist_ok=True)
    run(sys.executable, "-m", "build", "--sdist", "--wheel", "--outdir", str(DIST_DIR), "--no-isolation")
    artifacts = sorted(DIST_DIR.iterdir())
    if not artifacts:
        raise RuntimeError("No artifacts were produced in dist/")
    run(sys.executable, "-m", "twine", "check", *(str(path) for path in artifacts))
    return artifacts


def ensure_required_entries(artifacts: list[Path], version: str) -> None:
    wheel_name = f"jakal_hwpx-{version}-py3-none-any.whl"
    sdist_name = f"jakal_hwpx-{version}.tar.gz"
    wheel = next((path for path in artifacts if path.name == wheel_name), None)
    sdist = next((path for path in artifacts if path.name == sdist_name), None)
    if wheel is None or sdist is None:
        raise RuntimeError(f"Expected {wheel_name} and {sdist_name} in dist/, found {[path.name for path in artifacts]}")

    with zipfile.ZipFile(wheel) as archive:
        names = set(archive.namelist())
        for required in REQUIRED_WHEEL_PATTERNS:
            if required not in names:
                raise RuntimeError(f"Wheel is missing required file: {required}")

    prefix = f"jakal_hwpx-{version}/"
    with tarfile.open(sdist, "r:gz") as archive:
        names = set(archive.getnames())
        for required in REQUIRED_SDIST_PATTERNS:
            candidate = prefix + required
            if candidate not in names:
                raise RuntimeError(f"sdist is missing required file: {candidate}")

    missing_optional = [name for name in OPTIONAL_TOP_LEVEL_FILES if not (REPO_ROOT / name).exists()]
    if missing_optional:
        print(f"[warn] Missing top-level release files: {', '.join(missing_optional)}")


def main() -> None:
    version = read_versions()
    print(f"[ok] version consistency: {version}")
    artifacts = rebuild_dist()
    ensure_required_entries(artifacts, version)
    print(f"[ok] release artifacts validated: {[path.name for path in artifacts]}")


if __name__ == "__main__":
    main()
