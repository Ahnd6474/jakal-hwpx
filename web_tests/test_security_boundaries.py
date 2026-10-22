import asyncio

import pytest
from fastapi import FastAPI
from fastapi.testclient import TestClient

from server.app import create_app
from server.privacy import COOKIE, SessionIssuance, SessionRateLimited
from server.static import WebStaticFiles


@pytest.fixture
def site(tmp_path):
    with TestClient(create_app(tmp_path / 'data', run_worker=False)) as browser:
        yield browser


@pytest.mark.parametrize('name,first,second', [
    ('Host', 'testserver', 'other.example'),
    ('Origin', 'http://testserver', 'https://other.example'),
    ('Sec-Fetch-Site', 'same-origin', 'cross-site'),
    ('X-Jakal-Workspace', 'a' * 32, 'b' * 32),
    ('Content-Length', '0', '1'),
    ('Transfer-Encoding', 'chunked', 'identity'),
])
def test_ambiguous_headers_are_rejected_before_session_or_body_processing(site, name, first, second):
    response = site.get('/api/session', headers=[(name, first), (name, second)])
    assert response.status_code == 400
    assert 'set-cookie' not in response.headers and response.headers['cache-control'] == 'no-store'
    with site.app.state.store.connect() as db:
        assert db.execute('SELECT count(*) FROM sessions').fetchone()[0] == 0


def test_duplicate_cookies_and_transfer_length_conflicts_fail_closed(site):
    site.get('/api/session')
    token = site.cookies.get(COOKIE)
    for headers in [
            {'Cookie': f'{COOKIE}={token}; {COOKIE}=' + '1' * 64},
            [('Cookie', f'{COOKIE}={token}'), ('Cookie', f'{COOKIE}={token}')],
            {'Content-Length': '0', 'Transfer-Encoding': 'chunked'}]:
        assert site.get('/api/documents', headers=headers).status_code == 400
    assert site.get('/api/documents').status_code == 200
    response = site.get('/api/documents', headers=[('Cookie', 'other=value'), ('Cookie', f'{COOKIE}={token}')])
    assert response.status_code == 200 and response.cookies.get(COOKIE) == token


@pytest.mark.parametrize('value', ['localhost:wrong', 'localhost:70000', 'localhost ', 'localhost\t', 'localhost,other.example'])
def test_invalid_host_authorities_are_not_normalized_into_trusted_hosts(site, value):
    assert site.get('/api/session', headers={'Host': value}).status_code == 400


def test_public_resources_do_not_create_or_refresh_private_sessions(site):
    for path in ['/api/health', '/api/templates', '/', '/studio/', '/does-not-exist']:
        response = site.get(path)
        assert 'set-cookie' not in response.headers
        assert 'x-jakal-workspace' not in response.headers
    with site.app.state.store.connect() as db:
        assert db.execute('SELECT count(*) FROM sessions').fetchone()[0] == 0


def test_cookie_churn_is_limited_without_blocking_an_existing_workspace_or_health(site):
    clock = [0.0]
    site.app.state.sessions.issuance = SessionIssuance(per_peer=2, total=3, clock=lambda: clock[0])
    first = site.get('/api/session'); token = first.cookies.get(COOKIE)
    second = TestClient(site.app)
    try:
        assert second.get('/api/session').status_code == 200
        for value in ['', '1' * 64, '2' * 64]:
            response = second.get('/api/session', headers={'Cookie': f'{COOKIE}={value}'})
            assert response.status_code == 429 and response.headers['Retry-After'] == '60'
            assert 'set-cookie' not in response.headers
        assert site.get('/api/documents').status_code == 200
        assert site.cookies.get(COOKIE) == token
        assert second.get('/api/health', headers={'Cookie': ''}).status_code == 200
        with site.app.state.store.connect() as db:
            assert db.execute('SELECT count(*) FROM sessions').fetchone()[0] == 2
        clock[0] = 60.1
        assert second.get('/api/session', headers={'Cookie': ''}).status_code == 200
    finally:
        second.close()


def test_issuance_limit_stays_bounded_and_global_limit_survives_peer_eviction():
    clock = [0.0]
    limits = SessionIssuance(per_peer=2, total=4, max_peers=2, clock=lambda: clock[0])
    for n in range(4):
        limits.consume(f'peer-{n}')
    assert len(limits.peers) == 2 and len(limits.recent) == 4
    with pytest.raises(SessionRateLimited):
        limits.consume('evicted-peer')
    clock[0] = 61
    limits.consume('new-peer')
    assert len(limits.recent) == 1 and len(limits.peers) <= 2


@pytest.mark.parametrize('path', [
    'private.hwpx::$DATA', 'private.pdf ', 'private.pdf.', 'app.sqlite3:stream', 'app.sqlite3.',
    'app.db', 'app.db-wal', 'private.hwpx.gz.br', 'snapshot-1.json.gz', 'result-' + 'a' * 32 + '.json',
    'folder\\private.txt', 'NUL.txt', 'COM1.txt', 'COM¹.txt', 'CONOUT$', 'x\x00.txt', 'x\x1f.txt',
    'APP~1.SQL', 'SNAPSH~1.JSO', ' private.pdf',
])
def test_static_aliases_never_reach_filesystem_lookup(tmp_path, monkeypatch, path):
    static = WebStaticFiles(directory=tmp_path)
    def forbidden(*args):
        pytest.fail('Ambiguous or private paths must be rejected before filesystem access')
    monkeypatch.setattr(static, 'lookup_path', forbidden)
    from starlette.exceptions import HTTPException
    with pytest.raises(HTTPException) as caught:
        asyncio.run(static.get_response(path, {'headers': []}))
    assert caught.value.status_code == 404


def test_private_static_records_and_repeated_compression_suffixes_remain_unreadable(tmp_path):
    for name in ['private.hwpx.gz.br', 'snapshot-1.json', 'app.db', 'ordinary.css']:
        (tmp_path / name).write_bytes(b'confidential bytes')
    app = FastAPI(); app.mount('/', WebStaticFiles(directory=tmp_path))
    with TestClient(app) as browser:
        for name in ['private.hwpx.gz.br', 'snapshot-1.json', 'app.db', 'ordinary.css:stream']:
            assert browser.get('/' + name).status_code == 404
        assert browser.get('/ordinary.css').status_code == 200
