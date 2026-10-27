import hashlib
import time
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from server.app import create_app
from server.privacy import COOKIE, SECURE_COOKIE
from server.schemas import PageLayout
from server.storage import Store
from web_tests.test_api import client, raw_document, upload, wait


def test_client_chosen_session_is_replaced_and_workspace_is_not_a_credential(tmp_path):
    app = create_app(tmp_path / 'data', run_worker=False)
    with TestClient(app) as browser:
        chosen = '1' * 64
        browser.cookies.set(COOKIE, chosen, domain='testserver.local', path='/')
        response = browser.get('/api/session')
        issued = response.cookies.get(COOKIE)
        assert issued != chosen and len(issued) == 64
        assert len(response.json()['workspace']) == 32
        rows = []
        with app.state.store.connect() as db:
            rows = db.execute('SELECT * FROM sessions').fetchall()
        assert len(rows) == 1 and rows[0]['owner'] == hashlib.sha256(issued.encode()).hexdigest()
        assert issued not in str([dict(row) for row in rows])
        with TestClient(app) as stranger:
            denied = stranger.get('/api/documents', headers={'X-Jakal-Workspace': response.json()['workspace']})
            assert denied.status_code == 401 and 'set-cookie' not in denied.headers


def test_all_private_read_write_and_file_routes_hide_documents_from_other_sessions(client, raw_document):
    identifier = upload(client, raw_document)
    base = f'/api/documents/{identifier}'
    job = client.post(base + '/validations', json={'revision': 0}).json()['job_id']
    assert wait(client, job)['status'] == 'completed'
    other = TestClient(client.app)
    try:
        for path in ['', '/elements', '/state', '/page-layout?revision=0', '/revisions', '/revisions/0',
                '/source?revision=0', '/original', '/download?revision=0', '/report', '/text']:
            for method in ['GET', 'HEAD']:
                response = other.request(method, base + path, headers={'Range': 'bytes=0-100'})
                assert response.status_code in {404, 405}, (method, path, response.text)
                assert response.headers['Cache-Control'] == 'no-store'
                assert 'content-disposition' not in response.headers
                assert raw_document[:100] not in response.content
        assert other.get('/api/documents').json() == []
        assert other.get(f'/api/jobs/{job}?wait_ms=1500').status_code == 404
        commands = [
            ('POST', '/operations', {'revision': 0, 'operations': [{'type': 'paragraph', 'section': 0, 'index': 0, 'text': 'stolen'}]}),
            ('POST', '/validations', {'revision': 0}), ('POST', '/restorations', {'revision': 0, 'source_revision': 0}),
            ('POST', '/exports', {'revision': 0}), ('POST', '/copies', {'revision': 0, 'name': 'copy'}),
            ('PATCH', '/name', {'revision': 0, 'expected_name': '시험.hwpx', 'name': 'changed'}),
            ('PUT', '/settings', {'version': 0, 'settings': {}}),
            ('POST', '/page-layout', {'revision': 0, 'layout': PageLayout().model_dump()}),
        ]
        for method, path, body in commands:
            assert other.request(method, base + path, json=body).status_code == 404, path
        assert other.post(base + '/studio?revision=0', files={'file': ('x.hwpx', raw_document)}).status_code == 404
        assert other.delete(base).status_code == 404
        for path in [f'/.web-data/{identifier}/revision-0.hwpx', f'/{identifier}/revision-0.hwpx', '/app.sqlite3']:
            assert other.get(path).status_code == 404
        assert client.get(base + '/download?revision=0').content == raw_document
    finally:
        other.close()


def test_end_session_revokes_copied_cookie_and_fences_stale_tabs_before_mutations(client, raw_document):
    identifier = upload(client, raw_document)
    info = client.get('/api/session').json()
    stolen = client.cookies.get(COOKIE)
    ended = client.delete('/api/session', headers={'X-Jakal-Workspace': info['workspace']})
    assert ended.status_code == 200 and 'Max-Age=0' in ended.headers['set-cookie']
    client.cookies.clear(); client.cookies.set(COOKIE, stolen, domain='testserver.local', path='/')
    response = client.post('/api/documents', headers={'X-Jakal-Workspace': info['workspace']}, files={'file': ('old.hwpx', raw_document)})
    assert response.status_code == 401 and response.json()['code'] == 'session_changed'
    assert client.get(f'/api/documents/{identifier}/original').status_code == 404
    assert client.get('/api/documents').json() == []
    store = client.app.state.store
    store.cleanup()
    assert not store.directory(identifier).exists()


def test_session_survives_restart_but_expired_and_revoked_credentials_do_not(tmp_path):
    root = tmp_path / 'data'
    first = create_app(root, run_worker=False)
    with TestClient(first) as browser:
        info = browser.get('/api/session').json(); cookie = browser.cookies.get(COOKIE)
    second = create_app(root, run_worker=False)
    with TestClient(second) as browser:
        browser.cookies.set(COOKIE, cookie, domain='testserver.local', path='/')
        assert browser.get('/api/session').json() == info
        with second.state.store.connect() as db:
            db.execute('UPDATE sessions SET expires=?', (time.time() - 1,))
        response = browser.get('/api/session', headers={'X-Jakal-Workspace': info['workspace']})
        assert response.status_code == 401
        assert browser.get('/api/session').json()['workspace'] != info['workspace']


def test_public_deployment_requires_https_exact_host_and_secure_host_cookie(tmp_path, monkeypatch):
    monkeypatch.setenv('JAKAL_ENV', 'production')
    monkeypatch.delenv('JAKAL_PUBLIC_ORIGIN', raising=False)
    with pytest.raises(ValueError, match='requires'):
        create_app(tmp_path / 'data', run_worker=False)
    monkeypatch.setenv('JAKAL_PUBLIC_ORIGIN', 'https://editor.example')
    app = create_app(tmp_path / 'data', run_worker=False)
    with TestClient(app, base_url='https://editor.example') as browser:
        response = browser.get('/api/session')
        cookie = response.headers['set-cookie']
        assert cookie.startswith(SECURE_COOKIE + '=') and 'Secure' in cookie and 'Domain=' not in cookie
        assert response.headers['Strict-Transport-Security'] == 'max-age=31536000'
        assert browser.get('/api/openapi.json').status_code == 404
        assert browser.get('/api/session', headers={'Host': 'evil.example'}).status_code == 400
        for headers in [{'Origin': 'https://sibling.example'}, {'Sec-Fetch-Site': 'same-site'}, {'Sec-Fetch-Site': 'cross-site'}]:
            assert browser.get('/api/documents', headers=headers).status_code == 403
    with TestClient(app, base_url='http://editor.example') as browser:
        assert browser.get('/api/session').status_code == 400


def test_data_root_cannot_overlap_public_directories_and_parsers_do_not_inherit_hosting_secrets(tmp_path, monkeypatch):
    for public in ['web/public/private', 'web/dist/private', 'website/private', '.']:
        with pytest.raises(ValueError, match='storage'):
            create_app(Path(public), run_worker=False)
    from server.process_environment import parser_environment
    monkeypatch.setenv('OPENAI_API_KEY', 'private-secret')
    monkeypatch.setenv('DATABASE_URL', 'private-database')
    monkeypatch.setenv('NODE_OPTIONS', '--require attacker.js')
    env = parser_environment(Path.cwd())
    assert not {'OPENAI_API_KEY', 'DATABASE_URL', 'NODE_OPTIONS'} & env.keys()
    assert 'PYTHONPATH' in env and 'PATH' in {key.upper() for key in env}


def test_local_upgrade_only_adopts_existing_document_owner(tmp_path):
    root = tmp_path / 'data'; store = Store(root)
    token = '2' * 64; owner = hashlib.sha256(token.encode()).hexdigest()
    store.create_document(owner, 'previous.hwpx', 'a' * 32)
    app = create_app(root, run_worker=False)
    with TestClient(app) as browser:
        browser.cookies.set(COOKIE, token, domain='testserver.local', path='/')
        assert len(browser.get('/api/documents').json()) == 1
        assert browser.cookies.get(COOKIE) == token
        browser.delete('/api/session')
    # A restart must not resurrect the legacy credential after use was ended.
    newer = create_app(root, run_worker=False)
    with TestClient(newer) as browser:
        browser.cookies.set(COOKIE, token, domain='testserver.local', path='/')
        assert browser.get('/api/documents').json() == []
        assert browser.cookies.get(COOKIE) != token


def test_static_server_refuses_accidentally_copied_private_documents_and_secrets(tmp_path):
    from fastapi import FastAPI
    from server.static import WebStaticFiles, PUBLIC_DOCUMENTS
    from scripts.check_private_artifacts import PUBLIC_DEMOS, private_artifact
    for name in ['uploads/copied.hwpx', '.web-data/original.hwp', '.env', 'app.sqlite3',
            'app.sqlite3-wal', 'server.key', 'private.pdf', 'private.hwpx', 'sample.hwpx']:
        file = tmp_path / name; file.parent.mkdir(parents=True, exist_ok=True); file.write_bytes(b'private document bytes')
    (tmp_path / 'legitimate.js').write_text('public script')
    app = FastAPI(); app.mount('/', WebStaticFiles(directory=tmp_path))
    with TestClient(app) as browser:
        for name in ['uploads/copied.hwpx', '.web-data/original.hwp', '.env', 'app.sqlite3',
                'app.sqlite3-wal', 'server.key', 'private.pdf', 'private.hwpx', 'sample.hwpx']:
            assert browser.get('/' + name).status_code == 404
        assert browser.get('/legitimate.js').status_code == 200
    assert {f'web/public/{name}': digest for name, digest in PUBLIC_DOCUMENTS.items()} == PUBLIC_DEMOS
    assert private_artifact('private-uploads/x.hwpx')
    assert private_artifact('.env.production')
    assert not private_artifact('.env.production.example')
