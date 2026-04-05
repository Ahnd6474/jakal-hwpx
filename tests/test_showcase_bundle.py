from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest


REPO_ROOT = Path(__file__).resolve().parents[1]


def test_showcase_bundle_script(tmp_path: Path) -> None:
    corpus_dir = REPO_ROOT / "all_hwpx_flat"
    if not corpus_dir.exists():
        pytest.skip("showcase corpus is not available in this checkout")

    output_dir = tmp_path / "showcase_output"
    command = [
        sys.executable,
        str(REPO_ROOT / "examples" / "build_showcase_bundle.py"),
        "--corpus-dir",
        str(corpus_dir),
        "--output-dir",
        str(output_dir),
    ]
    completed = subprocess.run(command, cwd=REPO_ROOT, capture_output=True, text=True, check=True)

    manifest_path = output_dir / "showcase_manifest.json"
    report_path = output_dir / "showcase_report.md"

    assert "[ok] manifest:" in completed.stdout
    assert manifest_path.exists()
    assert report_path.exists()

    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert len(manifest["documents"]) == 7

    expected_names = {
        "showcase_layout_headers",
        "showcase_table_picture",
        "showcase_fields_references",
        "showcase_notes",
        "showcase_numbering",
        "showcase_equation",
        "showcase_shapes",
    }
    assert {item["name"] for item in manifest["documents"]} == expected_names

    for item in manifest["documents"]:
        assert Path(item["output"]).exists()
        assert item["validations"]["xml"] == []
        assert item["validations"]["reference"] == []
        assert item["validations"]["save_reopen"] == []
