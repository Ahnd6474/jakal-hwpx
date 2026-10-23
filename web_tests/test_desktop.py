from pathlib import Path
import importlib.util

import pytest

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location('desktop_bootstrap', ROOT / 'desktop/bootstrap.py')
bootstrap = importlib.util.module_from_spec(spec)
spec.loader.exec_module(bootstrap)


def test_local_listener_uses_an_exclusive_loopback_port_and_falls_back_if_busy():
    first = bootstrap.bind_local(0)
    try:
        second = bootstrap.bind_local(first.getsockname()[1])
        try:
            assert second.getsockname()[0] == '127.0.0.1'
            assert first.getsockname()[1] != second.getsockname()[1]
        finally:
            second.close()
    finally:
        first.close()


def test_only_one_process_can_use_the_same_document_store(tmp_path):
    path = tmp_path / 'server.lock'
    first = bootstrap.acquire_lock(path)
    try:
        with pytest.raises(RuntimeError, match='already'):
            bootstrap.acquire_lock(path)
    finally:
        first.close()
    bootstrap.acquire_lock(path).close()
