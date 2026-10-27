from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
from collections import Counter
from pathlib import Path
from zipfile import ZipFile

from lxml import etree


def inspect_converted(path: Path) -> None:
    from jakal_hwpx import HwpxDocument
    from .engine import inspect_archive, issues

    inspect_archive(path)
    # A conversion creates a new document. XML and reference errors cannot be
    # treated as pre-existing HWPX issues by the editor's baseline comparison.
    if issues(HwpxDocument.open(path, validate=False)):
        raise ValueError("HWP → HWPX 변환 결과에 문서 구조·첨부 참조 오류가 있어 중단했습니다. 한글에서 HWPX로 저장한 파일을 올려 주세요. 업로드한 HWP 원본은 그대로 보관됩니다.")


def convert_hwp(directory: Path) -> dict:
    """Keep the upload immutable; publish a converted base only after preflight."""
    from .hwp_engine import inspect_hwp

    source = directory / "revision-0.hwp"
    inspect_hwp(source)
    metadata = directory / "conversion.json"
    if metadata.exists() and (directory / "revision-0.hwpx").exists():
        inspect_converted(directory / "revision-0.hwpx")
        return json.loads(metadata.read_text(encoding="utf-8"))
    node = os.environ.get("JAKAL_NODE") or shutil.which("node")
    if not node:
        raise ValueError("HWP → HWPX 변환에는 Node.js 22 이상과 web 폴더의 npm install이 필요합니다.")
    pending = directory / "conversion-pending.hwpx"
    report = directory / "conversion-pending.json"
    pending.unlink(missing_ok=True)
    report.unlink(missing_ok=True)
    try:
        subprocess.run(
            [node, "--max-old-space-size=512", str(Path(__file__).with_name("convert_hwp.mjs")), str(source), str(pending), str(report)],
            stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            check=True, timeout=60,
            creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
        )
        inspect_converted(pending)
        # Count across all XML parts, including master pages and nested objects.
        from jakal_hwpx import HwpDocument
        native = HwpDocument.open(source)
        counts = Counter()
        with ZipFile(pending) as archive:
            for name in archive.namelist():
                if name.endswith(".xml"):
                    root = etree.fromstring(archive.read(name), parser=etree.XMLParser(resolve_entities=False, no_network=True))
                    counts.update(etree.QName(n).localname for n in root.iter() if isinstance(n.tag, str))
        expected = {"tbl": len(native.tables()), "equation": len(native.equations()), "pic": len(native.pictures())}
        if any(counts[tag] != count for tag, count in expected.items()):
            raise ValueError("HWP → HWPX 변환에서 표·수식·그림 수가 일치하지 않아 중단했습니다. 한글에서 HWPX로 저장한 파일을 올려 주세요. HWP 원본은 그대로 보관됩니다.")
        details = json.loads(report.read_text(encoding="utf-8"))
        details.update(source_format="hwp", target_format="hwpx", preserved_objects=expected)
        report.write_text(json.dumps(details, ensure_ascii=False), encoding="utf-8")
        pending.replace(directory / "revision-0.hwpx")
        report.replace(metadata)
        return details
    except (subprocess.SubprocessError, OSError) as exc:
        raise ValueError("HWP를 HWPX로 안전하게 변환하지 못했습니다. 지원하지 않는 개체가 있거나 변환 한도를 초과했습니다. 한글에서 HWPX로 저장한 파일을 올려 주세요. 업로드한 HWP 원본은 그대로 보관됩니다.") from exc
    finally:
        pending.unlink(missing_ok=True)
        report.unlink(missing_ok=True)
