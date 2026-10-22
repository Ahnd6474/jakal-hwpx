from __future__ import annotations

import os
from pathlib import Path
import shutil
import socket
import subprocess
import sys
import time

import pytest

from server.runtime import check_runtime, server_options

REPO = Path(__file__).resolve().parents[1]


def test_platform_port_and_explicit_override(monkeypatch):
    monkeypatch.delenv('JAKAL_PORT', raising=False)
    monkeypatch.setenv('PORT', '12345')
    assert server_options()['port'] == 12345
    monkeypatch.setenv('JAKAL_PORT', '8001')
    assert server_options()['port'] == 8001


@pytest.mark.parametrize('port', ['0', '-1', '65536', 'oops', ''])
def test_invalid_platform_port_fails_before_serving(monkeypatch, port):
    monkeypatch.delenv('JAKAL_PORT', raising=False)
    monkeypatch.setenv('PORT', port)
    with pytest.raises(ValueError, match='between'):
        server_options()


def test_incomplete_production_image_cannot_serve_a_broken_editor(tmp_path, monkeypatch):
    monkeypatch.setenv('JAKAL_ENV', 'production')
    with pytest.raises(RuntimeError, match='incomplete'):
        check_runtime(tmp_path)
    for name in ['web/dist/index.html', 'web/dist/studio/index.html',
            'web/node_modules/@rhwp/core/rhwp.js', 'web/node_modules/@rhwp/core/rhwp_bg.wasm']:
        file = tmp_path / name
        file.parent.mkdir(parents=True, exist_ok=True)
        file.touch()
    monkeypatch.setattr(shutil, 'which', lambda _: None)
    with pytest.raises(RuntimeError, match='Node.js'):
        check_runtime(tmp_path)


def test_environment_typo_cannot_disable_production_guards(monkeypatch):
    monkeypatch.setenv('JAKAL_ENV', 'prodution')
    with pytest.raises(ValueError, match='JAKAL_ENV'):
        check_runtime()


def test_real_production_process_and_loopback_healthcheck(tmp_path, monkeypatch):
    # An actual uvicorn process exercises proxy trust, Secure cookies, worker
    # subprocesses and Node/WASM conversion together, outside TestClient.
    with socket.socket() as reserved:
        reserved.bind(('127.0.0.1', 0))
        port = str(reserved.getsockname()[1])
    env = os.environ | {
        'JAKAL_ENV': 'production', 'JAKAL_HOST': '127.0.0.1', 'JAKAL_PORT': port,
        'JAKAL_PUBLIC_ORIGIN': 'https://deployment.invalid',
        'JAKAL_TRUSTED_PROXIES': '127.0.0.1', 'JAKAL_DATA_DIR': str(tmp_path / 'private'),
        'PYTHONPATH': str(REPO / 'src') + os.pathsep + str(REPO),
    }
    for key, value in env.items():
        if key.startswith('JAKAL_'):
            monkeypatch.setenv(key, value)
    from server.healthcheck import probe
    from scripts.check_deployment import check_deployment
    with (tmp_path / 'server.log').open('w', encoding='utf-8') as log:
        process = subprocess.Popen([sys.executable, '-m', 'server'], cwd=REPO, env=env,
            stdin=subprocess.DEVNULL, stdout=log, stderr=log)
        try:
            deadline = time.monotonic() + 30
            while time.monotonic() < deadline:
                assert process.poll() is None, (tmp_path / 'server.log').read_text(encoding='utf-8')
                if probe():
                    break
                time.sleep(.15)
            else:
                pytest.fail('Production process did not become healthy')
            check_deployment()
            # Probe neither follows redirects nor bypasses public host checks.
            monkeypatch.setenv('JAKAL_PUBLIC_ORIGIN', 'https://wrong.invalid')
            assert not probe()
        finally:
            process.terminate()
            try:
                process.wait(timeout=15)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait(timeout=5)
