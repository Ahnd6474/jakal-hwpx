from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]


def test_showcase_bundle_script(sample_corpus_dir: Path, tmp_path: Path) -> None:
    corpus_dir = sample_corpus_dir
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
