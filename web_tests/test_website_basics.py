import pytest
from fastapi.testclient import TestClient

from server.app import create_app


@pytest.fixture
def site(tmp_path):
    with TestClient(create_app(tmp_path / 'data', run_worker=False), raise_server_exceptions=False) as client:
        yield client


@pytest.mark.parametrize('method,path,headers,status', [
    ('POST', '/api/documents', {'Origin': 'https://untrusted.example'}, 403),
    ('DELETE', '/api/documents/missing', {'Sec-Fetch-Site': 'cross-site'}, 403),
    ('GET', '/api/documents/missing', {}, 404),
    ('POST', '/api/documents', {}, 422),
    ('POST', '/api/documents', {'Content-Length': str(22 * 1024 * 1024)}, 413),
    ('POST', '/api/documents', {'Content-Length': '-1'}, 400),
    ('POST', '/api/documents', {'Content-Length': 'invalid'}, 400),
])
def test_error_responses_keep_private_cache_and_security_policy(site, method, path, headers, status):
    response = site.request(method, path, headers=headers)
    assert response.status_code == status
    assert response.headers['Cache-Control'] == 'no-store'
    assert response.headers['X-Content-Type-Options'] == 'nosniff'
    assert "frame-ancestors 'none'" in response.headers['Content-Security-Policy']
    assert 'camera=()' in response.headers['Permissions-Policy']
    if status == 403:
        assert 'Set-Cookie' not in response.headers


def test_unexpected_error_hides_internal_details_and_keeps_security_headers(site, monkeypatch):
    def fail(owner):
        raise RuntimeError('private document content C:/private/file.hwpx')
    monkeypatch.setattr(site.app.state.store, 'list_documents', fail)
    response = site.get('/api/documents')
    assert response.status_code == 500
    assert 'private' not in response.text
    assert response.headers['Cache-Control'] == 'no-store'
    assert response.headers['X-Content-Type-Options'] == 'nosniff'
    assert response.json()['detail']


def test_secure_cookie_follows_https_and_http_remains_usable(tmp_path, monkeypatch):
    monkeypatch.delenv('JAKAL_COOKIE_SECURE', raising=False)
    app = create_app(tmp_path / 'data', run_worker=False)
    with TestClient(app, base_url='https://site.test') as client:
        cookie = client.get('/api/health').headers['Set-Cookie']
        assert 'Secure' in cookie and 'HttpOnly' in cookie and 'SameSite=strict' in cookie
    with TestClient(app, base_url='http://127.0.0.1') as client:
        assert 'Secure' not in client.get('/api/health').headers['Set-Cookie']
        assert client.get('/api/documents').status_code == 200


def test_page_policy_and_studio_embedding_remain_distinct(site):
    page = site.get('/')
    assert 'camera=()' in page.headers['Permissions-Policy']
    assert "frame-ancestors 'none'" in page.headers['Content-Security-Policy']
    studio = site.get('/studio/index.html')
    assert "frame-ancestors 'self'" in studio.headers['Content-Security-Policy']
    assert 'no-store' not in studio.headers.get('Cache-Control', '')


def test_missing_page_returns_a_readable_404_without_immutable_cache(site):
    response = site.get('/page-that-does-not-exist', headers={'Accept-Encoding': 'gzip'})
    assert response.status_code == 404
    assert '내 문서로 돌아가기' in response.text
    assert response.headers['Content-Type'].startswith('text/html')
    assert 'immutable' not in response.headers['Cache-Control']
