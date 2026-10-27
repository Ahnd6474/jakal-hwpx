"""Exercise a running loopback deployment with generated/public documents only."""
from __future__ import annotations

import http.client
from http.cookies import SimpleCookie
import json
import os
import secrets
import time
from urllib.parse import urlsplit

from jakal_hwpx import HwpxDocument
from server.runtime import REPO, server_options


class DeploymentClient:
    secure_cookie = True
    connection_factory = staticmethod(lambda: http.client.HTTPConnection(
        '127.0.0.1', server_options()['port'], timeout=10))

    def __init__(self):
        origin = urlsplit(os.environ['JAKAL_PUBLIC_ORIGIN'])
        self.headers = {'Host': origin.netloc, 'X-Forwarded-Proto': 'https'}

    def request(self, method, path, body=None, headers=None):
        connection = self.connection_factory()
        try:
            connection.request(method, path, body=body, headers=self.headers | (headers or {}))
            response = connection.getresponse()
            return response.status, {key.lower(): value for key, value in response.getheaders()}, response.read()
        finally:
            connection.close()

    def json(self, method, path, payload=None, *, status=200):
        raw = json.dumps(payload, ensure_ascii=False).encode() if payload is not None else None
        code, _, body = self.request(method, path, raw, {'Content-Type': 'application/json'})
        assert code == status, (method, path, code)
        return json.loads(body)

    def upload(self, name, raw):
        boundary = secrets.token_hex(16)
        body = (f'--{boundary}\r\nContent-Disposition: form-data; name="file"; filename="{name}"\r\n'
            'Content-Type: application/octet-stream\r\n\r\n').encode() + raw + f'\r\n--{boundary}--\r\n'.encode()
        status, _, data = self.request('POST', '/api/documents', body,
            {'Content-Type': f'multipart/form-data; boundary={boundary}'})
        assert status == 202, ('upload', name, status)
        result = json.loads(data)
        self.wait(result['job_id'])
        return '/api/documents/' + result['document_id']

    def wait(self, identifier):
        deadline = time.monotonic() + 120
        while time.monotonic() < deadline:
            result = self.json('GET', '/api/jobs/' + identifier)
            if result['status'] == 'completed':
                return result
            assert result['status'] != 'failed', ('job failed', result.get('error'))
            time.sleep(.1)
        raise AssertionError('Deployment job timed out')

    def begin(self):
        status, headers, raw = self.request('GET', '/api/session')
        assert status == 200
        cookie = SimpleCookie()
        cookie.load(headers['set-cookie'])
        value = cookie['__Host-jakal_session' if self.secure_cookie else 'jakal_session']
        assert bool(value['secure']) == self.secure_cookie and value['httponly'] and value['samesite'] == 'strict'
        self.headers['Cookie'] = f'{value.key}={value.value}'
        self.headers['X-Jakal-Workspace'] = json.loads(raw)['workspace']


def check_deployment(*, via_gateway=False, desktop=False):
    # The loopback transport models the proxy's HTTP hop; public HTTPS is
    # terminated by Caddy. No request targets an external host.
    client = DeploymentClient()
    status, headers, _ = client.request('GET', '/api/health')
    assert status == 200 and 'set-cookie' not in headers
    for path in ['/', '/studio/index.html']:
        assert client.request('GET', path)[0] == 200, path
    rejected, blocked_headers, blocked_body = client.request('GET', '/api/health', headers={'Host': 'other.invalid'})
    if via_gateway:
        assert 'set-cookie' not in blocked_headers and b'"status":"ok"' not in blocked_body
    else:
        assert rejected == 400
    forwarded_status = client.request('GET', '/api/health', headers={'X-Forwarded-Proto': 'http'})[0]
    assert forwarded_status == (200 if via_gateway or desktop else 400)
    client.begin()
    assert client.request('GET', '/api/documents', headers={'Origin': 'https://other.invalid'})[0] == 403
    assert client.request('GET', '/api/documents', headers={'X-Jakal-Workspace': '0' * 32})[0] == 401
    created = HwpxDocument.blank()
    created.set_paragraph_text(0, 0, 'Deployment input')
    public_hwp = REPO / 'web/dist/sample.hwp'
    for name, raw in [('deployment.hwpx', created.compile()), ('deployment.hwp', public_hwp.read_bytes())]:
        base = client.upload(name, raw)
        stranger = DeploymentClient()
        stranger.begin()
        assert stranger.request('GET', base + '/original')[0] == 404
        document = client.json('GET', base)
        paragraph = next(p for section in document['sections'] for p in section['paragraphs'] if p['editable'])
        edited = 'Deployment edited ' + name
        job = client.json('POST', base + '/operations', {'revision': 0, 'operations': [{
            'type': 'paragraph', 'section': paragraph['section'], 'index': paragraph['index'], 'text': edited,
        }]}, status=202)
        client.wait(job['job_id'])
        validation = client.json('POST', base + '/validations', {'revision': 1}, status=202)
        client.wait(validation['job_id'])
        exported = client.json('POST', base + '/exports', {'revision': 1})
        status, headers, content = client.request('GET', exported['url'])
        assert status == 200 and headers['cache-control'] == 'no-store'
        reopened = HwpxDocument.from_bytes(content)
        assert edited in reopened.sections[paragraph['section']].text_fragments()
        assert client.request('DELETE', base)[0] == 200
        assert client.request('GET', base + '/original')[0] == 404
        stranger.json('DELETE', '/api/session')
    client.json('DELETE', '/api/session')
    assert client.request('GET', '/api/documents')[0] == 401
    print(('Desktop' if desktop else 'Production') + ' HWP/HWPX upload, edit, validate, download and session isolation passed.')


if __name__ == '__main__':
    check_deployment()
