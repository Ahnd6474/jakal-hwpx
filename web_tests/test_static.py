import gzip

import pytest
from fastapi import FastAPI
from fastapi.testclient import TestClient

from server.static import WebStaticFiles, accepts_gzip
from web_tests.test_api import client, raw_document, upload, wait


@pytest.mark.parametrize('header,expected', [('gzip, deflate, br', True), ('gzip;q=0, *;q=1', False), ('*;q=.5', True), ('br', False), ('gzip;q=garbage', False), ('gzip;q=2', False), ('GZIP;q=0.5', True)])
def test_encoding_negotiation(header, expected):
    assert accepts_gzip(header) == expected


def test_precompressed_asset_cache_variants_and_ranges(tmp_path):
    assets = tmp_path / 'assets'
    assets.mkdir()
    data = b'const wasmIsLocal = true;\n' * 1000
    (assets / 'test-abcdefgh.js').write_bytes(data)
    (assets / 'test-abcdefgh.js.gz').write_bytes(gzip.compress(data))
    (assets / 'uncompressed-abcdefgh.js').write_bytes(data)
    (tmp_path / 'index.html').write_text('<h1>Editor</h1>')
    (tmp_path / '404.html').write_text('<h1>Page not found</h1>')
    app = FastAPI()
    app.mount('/', WebStaticFiles(directory=tmp_path, html=True))
    with TestClient(app) as client:
        compressed = client.get('/assets/test-abcdefgh.js', headers={'Accept-Encoding': 'gzip'})
        assert compressed.content == data
        assert compressed.headers['content-encoding'] == 'gzip'
        assert 'javascript' in compressed.headers['content-type']
        assert int(compressed.headers['content-length']) < len(data) / 10
        assert 'immutable' in compressed.headers['cache-control']
        plain = client.get('/assets/test-abcdefgh.js', headers={'Accept-Encoding': 'gzip;q=0'})
        assert plain.content == data and 'content-encoding' not in plain.headers
        assert compressed.headers['vary'] == plain.headers['vary'] == 'Accept-Encoding'
        assert client.get('/assets/test-abcdefgh.js', headers={'If-None-Match': compressed.headers['etag'], 'Accept-Encoding': 'gzip'}).status_code == 304
        partial = client.get('/assets/test-abcdefgh.js', headers={'Range': 'bytes=0-4', 'Accept-Encoding': 'gzip'})
        assert partial.status_code == 206 and partial.content == data[:5] and 'content-encoding' not in partial.headers
        assert client.get('/assets/uncompressed-abcdefgh.js', headers={'Accept-Encoding': 'gzip'}).content == data
        assert client.get('/').headers['cache-control'] == 'no-cache'
        missing = client.get('/assets/missing-abcdefgh.js', headers={'Accept-Encoding': 'gzip'})
        assert missing.status_code == 404
        assert 'Page not found' in missing.text
        assert missing.headers['content-type'].startswith('text/html')
        assert 'content-encoding' not in missing.headers
        assert 'immutable' not in missing.headers['cache-control']
        assert client.get('/not-a-page').status_code == 404


def test_public_assets_do_not_issue_session_cookies_and_private_api_is_not_cached(client, raw_document):
    # The public cache must never distribute a session cookie to another visitor.
    response = client.get('/studio/assets/missing-abcdefgh.js')
    assert 'set-cookie' not in response.headers
    assert client.get('/api/health').headers['cache-control'] == 'no-store'
    identifier = upload(client, raw_document)
    response = client.get(f'/api/documents/{identifier}')
    assert response.headers['cache-control'] == 'no-store'
    assert response.json()['status'] == 'ready'


def test_bounded_job_wait_returns_completion_and_checks_ownership(client, raw_document):
    response = client.post('/api/documents', files={'file': ('wait.hwpx', raw_document, 'application/hwp+zip')})
    job = response.json()['job_id']
    result = client.get(f'/api/jobs/{job}?wait_ms=1500')
    assert result.status_code == 200
    # wait_ms is a maximum polling interval, not a parser completion deadline.
    # A cold Windows worker may legitimately still be running after 1.5 seconds.
    assert result.json()['status'] in {'queued', 'running', 'completed'}
    assert wait(client, job)['status'] == 'completed'
    stranger = TestClient(client.app)
    assert stranger.get(f'/api/jobs/{job}?wait_ms=1500').status_code == 404
    stranger.close()
    assert client.get(f'/api/jobs/{job}?wait_ms=1501').status_code == 422


def test_editor_toolbar_icons_and_page_icons_use_available_local_assets(client):
    html = client.get('/studio/').text
    assert 'href="/studio/favicon.ico"' in html
    assert 'href="/studio/icons/icon-256.png"' in html
    css = client.get('/studio/jakal.css').text
    assert "--ui-icon-sprite-url: url('/studio/images/icon_small_ko.svg')" in css
    for path in ['/studio/favicon.ico', '/studio/icons/icon-256.png', '/studio/images/icon_small_ko.svg', '/studio/print.html']:
        assert client.get(path).status_code == 200
