"""Publish complete files by replacement, preserving the previous file on error."""
from __future__ import annotations

import json
import os
from pathlib import Path


def atomic_write(path: Path, data: bytes) -> None:
    # Jobs are serialized per document; the fixed suffix also allows recovery
    # to remove staging files after the subprocess is forcibly terminated.
    temporary = path.with_name(path.name + ".tmp")
    try:
        with temporary.open("wb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)


def atomic_json(path: Path, value: object) -> None:
    atomic_write(path, json.dumps(value, ensure_ascii=False).encode("utf-8"))
