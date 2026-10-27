"""Check the extracted Windows release without relying on installed runtimes."""
from __future__ import annotations

import argparse
import asyncio
import hashlib
import http.client
import io
import json
import os
from pathlib import Path
import secrets
import subprocess
import sys
import time
import zipfile

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO))
from check_deployment import DeploymentClient, check_deployment
from jakal_hwpx import HwpxDocument
from check_mcp import check_site


def start(package, data, environment):
    process = subprocess.Popen([str(package / 'Jakal.exe'), '--headless', '--data-root', str(data)],
        stdin=subprocess.PIPE, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=environment)
    deadline = time.monotonic() + 90
    state = data / 'state/ready.json'
    while time.monotonic() < deadline:
        if process.poll() is not None:
            log = data / 'logs/application.log'
            failure = data / 'logs/launcher-error.log'
            if failure.exists():
                log = failure
            raise RuntimeError(f'Packaged launcher exited ({process.returncode}). ' + (log.read_text(encoding='utf-8-sig')[-3000:] if log.exists() else 'No startup log.'))
        try:
            ready = json.loads(state.read_text(encoding='utf-8'))
            response = http.client.HTTPConnection('127.0.0.1', ready['port'], timeout=2)
            response.request('GET', '/api/health')
            healthy = response.getresponse()
            valid = healthy.status == 200 and json.loads(healthy.read())['status'] == 'ok'
            response.close()
            if valid:
                return process, ready
        except (OSError, ValueError, KeyError, http.client.HTTPException):
            pass
        time.sleep(.1)
    process.kill()
    process.wait(timeout=10)
    raise RuntimeError('Packaged launcher did not become ready.')


def stop(process, data):
    process.stdin.close()
    assert process.wait(timeout=30) == 0
    assert not (data / 'state/ready.json').exists()


def client_for(ready, saved=None):
    port = ready['port']
    os.environ['JAKAL_PUBLIC_ORIGIN'] = f'http://127.0.0.1:{port}'
    DeploymentClient.secure_cookie = False
    DeploymentClient.connection_factory = staticmethod(lambda: http.client.HTTPConnection('127.0.0.1', port, timeout=10))
    client = DeploymentClient()
    if saved:
        client.headers.update({key: value for key, value in saved.items() if key.lower() != 'host'})
    return client


def check_imports(client):
    from PIL import Image
    image = io.BytesIO()
    Image.new('RGB', (32, 48), 'white').save(image, 'JPEG')
    pdf = next((REPO / 'examples/samples/pdf').glob('*.pdf')).read_bytes()
    for name, raw, source in [('image.jpg', image.getvalue(), 'images'), ('sample.pdf', pdf, 'pdf')]:
        boundary = secrets.token_hex(12)
        content = (f'--{boundary}\r\nContent-Disposition: form-data; name="files"; filename="{name}"\r\n'
            'Content-Type: application/octet-stream\r\n\r\n').encode() + raw + f'\r\n--{boundary}--\r\n'.encode()
        status, _, body = client.request('POST', '/api/page-templates/extract', content,
            {'Content-Type': f'multipart/form-data; boundary={boundary}'})
        assert status == 200, ('template import', source, status, body[:300])
        assert json.loads(body)['source'] == source


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('archive', type=Path)
    args = parser.parse_args()
    work = REPO / '.artifacts/desktop-checks' / ('한글 공백 경로 ' + secrets.token_hex(5))
    work.mkdir(parents=True)
    with zipfile.ZipFile(args.archive) as archive:
        archive.extractall(work)
    package = work / 'Jakal-Workspace'
    assert (package / 'guide.html').is_file(), 'Offline startup guide is missing.'
    assert (package / '시작.cmd').is_file(), 'Double-click startup helper is missing.'
    data = work / '사용자 자료'
    environment = os.environ | {'PATH': str(work / 'no-system-runtimes'), 'NODE_OPTIONS': '--require nonexistent-file',
        'PYTHONPATH': str(work / 'wrong-python-path'), 'JAKAL_ENV': 'production',
        'JAKAL_PUBLIC_ORIGIN': 'https://wrong.invalid'}
    ui = subprocess.run([str(package / 'Jakal.exe'), '--self-test-ui', '--data-root', str(work / 'UI 검사')],
        stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=environment, timeout=90)
    assert ui.returncode == 0, 'Native launcher window did not start and stop correctly.'
    process, ready = start(package, data, environment)
    try:
        client_for(ready)
        check_deployment(desktop=True)
        asyncio.run(check_site(f"http://127.0.0.1:{ready['port']}"))
        client = client_for(ready)
        client.begin()
        check_imports(client)
        created = HwpxDocument.blank()
        created.set_paragraph_text(0, 0, '재실행 후에도 유지할 문서')
        original = created.compile()
        base = client.upload('persistence.hwpx', original)
        saved = dict(client.headers)
    finally:
        stop(process, data)
    process, ready = start(package, data, environment)
    try:
        client = client_for(ready, saved)
        assert client.json('GET', '/api/session')['workspace'] == saved['X-Jakal-Workspace']
        assert client.request('GET', base + '/original')[2] == original
        client.json('DELETE', base)
        client.json('DELETE', '/api/session')
    finally:
        stop(process, data)
    manifest = json.loads((package / 'package-manifest.json').read_text(encoding='utf-8'))
    actual = {p.relative_to(package).as_posix() for p in package.rglob('*') if p.is_file()}
    assert actual == {*manifest['files'], 'package-manifest.json'}, actual - set(manifest['files'])
    for name, digest in manifest['files'].items():
        assert hashlib.sha256((package / name).read_bytes()).hexdigest() == digest, name
    assert not any(package.rglob('*.sqlite3'))
    print('Windows package passed: no system Python/Node, Korean/spaced path, HWP/HWPX editing, PDF/JPG import, graceful exit and persistent sessions/documents.')


if __name__ == '__main__':
    main()
