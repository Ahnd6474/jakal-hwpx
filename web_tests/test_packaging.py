import hashlib
import io
import os
from pathlib import Path
import shutil
import subprocess
import sys
import zipfile

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
import build_desktop as builder


def test_server_archive_keeps_shell_control_executable(tmp_path):
    root = tmp_path / 'Jakal-Server'
    root.mkdir()
    (root / 'manage.sh').write_bytes(b'#!/bin/sh\nexit 0\n')
    archive = builder.archive_release(root, tmp_path, 'server.zip', {'version': 'test'})
    with zipfile.ZipFile(archive) as contents:
        script = contents.getinfo('Jakal-Server/manage.sh')
        assert script.create_system == 3
        assert script.external_attr >> 16 == 0o100755
        assert b'\r' not in contents.read(script)


@pytest.mark.parametrize('name', ['.env', 'app/secret.sqlite3', 'runtime/private.pem', 'app/web/dist/upload.hwpx'])
def test_release_rejects_private_runtime_data_and_uploaded_documents(tmp_path, name):
    file = tmp_path / name
    file.parent.mkdir(parents=True, exist_ok=True)
    file.write_bytes(b'private')
    with pytest.raises(RuntimeError):
        builder.validate_bundle(tmp_path)


def test_runtime_checksum_is_verified_before_use(tmp_path, monkeypatch):
    monkeypatch.setattr(builder.urllib.request, 'urlopen', lambda *a, **kw: io.BytesIO(b'wrong binary'))
    with pytest.raises(RuntimeError, match='checksum'):
        builder.download_verified({'url': 'https://example.invalid/runtime.zip', 'sha256': hashlib.sha256(b'expected').hexdigest()}, tmp_path)
    assert not list(tmp_path.glob('*.zip'))


@pytest.mark.parametrize('version,port,valid', [('28.0.1', '18988', True), ('27.5.0', '18988', False), ('28.0.1', '0', False), ('28.0.1', '18988;echo unsafe', False)])
def test_local_controls_validate_engine_and_port_and_preserve_data_on_stop(tmp_path, version, port, valid):
    controls = tmp_path / 'controls'
    controls.mkdir()
    (controls / 'compose.local.yaml').write_text('services: {}')
    executables = tmp_path / 'tools'
    executables.mkdir()
    calls = tmp_path / 'calls.txt'
    if sys.platform == 'win32':
        shutil.copy2(ROOT / 'deploy/local.ps1', controls / 'local.ps1')
        (executables / 'docker.cmd').write_text('@echo off\necho %*>>"%TASK_DOCKER_CALLS%"\nif "%~1"=="version" echo %TASK_DOCKER_VERSION%\nexit /b 0\n', encoding='ascii')
        base = ['powershell.exe', '-NoProfile', '-NonInteractive', '-File', str(controls / 'local.ps1'), '-Action']
    else:
        shutil.copy2(ROOT / 'deploy/local.sh', controls / 'local.sh')
        stub = executables / 'docker'
        stub.write_text('#!/bin/sh\nprintf "%s\\n" "$*" >> "$TASK_DOCKER_CALLS"\nif [ "$1" = version ]; then printf "%s\\n" "$TASK_DOCKER_VERSION"; fi\n', encoding='ascii')
        stub.chmod(0o755)
        base = ['sh', str(controls / 'local.sh')]
    env = os.environ | {'PATH': str(executables) + os.pathsep + os.environ['PATH'], 'TASK_DOCKER_CALLS': str(calls),
        'TASK_DOCKER_VERSION': version, 'JAKAL_LOCAL_PORT': port, 'JAKAL_LOCAL_PROJECT': 'jakal-test'}
    result = subprocess.run([*base, 'start'], env=env, capture_output=True, text=True)
    assert (result.returncode == 0) is valid, result.stderr
    if valid:
        assert subprocess.run([*base, 'stop'], env=env, capture_output=True).returncode == 0
        recorded = calls.read_text()
        assert '--project-name jakal-test' in recorded and 'stop' in recorded and '--wait' in recorded
        assert ' down' not in recorded and '--volumes' not in recorded
    else:
        assert ' up ' not in calls.read_text()


@pytest.mark.parametrize('domain,valid', [('editor.example.com', True), ('https://editor.example.com/a', False), ('x.example.com;echo unsafe', False)])
def test_server_controls_validate_domain_and_stop_keeps_document_volume(tmp_path, domain, valid):
    controls = tmp_path / 'controls'
    controls.mkdir()
    tools = tmp_path / 'tools'
    tools.mkdir()
    calls = tmp_path / 'docker-calls.txt'
    if sys.platform == 'win32':
        shutil.copy2(ROOT / 'deploy/manage.ps1', controls / 'manage.ps1')
        (tools / 'docker.cmd').write_text('@echo off\necho %*>>"%TASK_DOCKER_CALLS%"\nexit /b 0\n', encoding='ascii')
        base = ['powershell.exe', '-NoProfile', '-NonInteractive', '-File', str(controls / 'manage.ps1')]
        start = [*base, '-Action', 'start', '-Domain', domain]
        stop = [*base, '-Action', 'stop']
    else:
        shutil.copy2(ROOT / 'deploy/manage.sh', controls / 'manage.sh')
        stub = tools / 'docker'
        stub.write_text('#!/bin/sh\nprintf "%s\\n" "$*" >> "$TASK_DOCKER_CALLS"\n', encoding='ascii')
        stub.chmod(0o755)
        base = ['sh', str(controls / 'manage.sh')]
        start, stop = [*base, 'start', domain], [*base, 'stop']
    environment = os.environ | {'PATH': str(tools) + os.pathsep + os.environ['PATH'], 'TASK_DOCKER_CALLS': str(calls)}
    result = subprocess.run(start, env=environment, capture_output=True, text=True)
    assert (result.returncode == 0) is valid, result.stderr
    config = controls / '.env.production'
    if not valid:
        assert not config.exists()
        return
    assert 'JAKAL_DOMAIN=editor.example.com' in config.read_text(encoding='utf-8-sig')
    before = config.read_bytes()
    assert subprocess.run(stop, env=environment, capture_output=True).returncode == 0
    assert config.read_bytes() == before
    recorded = calls.read_text()
    assert 'stop' in recorded and '--wait' in recorded
    assert ' down' not in recorded and 'volume rm' not in recorded
