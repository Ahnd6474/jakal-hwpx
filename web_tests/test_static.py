import gzip
import re
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import unquote, urljoin, urlsplit

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


def test_built_frontend_asset_references_and_binary_files_are_valid():
    """Check emitted paths, including CSS fallbacks and lazy JS/worker chunks."""
    root = Path(__file__).resolve().parents[1] / 'web' / 'dist'
    assert (root / 'index.html').is_file(), 'Run npm run build in web first'
    references = []

    class AssetLinks(HTMLParser):
        def handle_starttag(self, tag, attrs):
            attrs = dict(attrs)
            if tag in {'script', 'img', 'iframe', 'source'} and attrs.get('src'):
                references.append((base, attrs['src']))
            if tag == 'link' and attrs.get('href'):
                references.append((base, attrs['href']))

    for path in root.rglob('*'):
        if not path.is_file():
            continue
        base = '/' + path.relative_to(root).as_posix()
        if path.suffix == '.html':
            AssetLinks().feed(path.read_text(encoding='utf-8'))
        elif path.suffix == '.css':
            references.extend((base, url) for url in re.findall(r'''url\(\s*["']?([^\s"')]+)''', path.read_text(encoding='utf-8')))
        elif path.suffix == '.js':
            # Vite emits imports, new URL worker/WASM paths and preload lists
            # as literal relative paths. Ignore dynamic strings and external URLs.
            source = path.read_text(encoding='utf-8')
            references.extend((base, url) for url in re.findall(
                r'''["']((?:\./|\.\./|/(?:assets|studio)/)[^\s"'`{}]+\.(?:js|css|wasm|woff2?|ttf|otf))["'](?!\s*:)''', source))
            # Font catalogs resolve against the editor document, not its JS
            # chunk. WASM import-map property keys above are not network URLs.
            document_base = '/studio/' if base.startswith('/studio/') else '/'
            references.extend((document_base, url) for url in re.findall(
                r'''["'](fonts/[^\s"']+\.(?:woff2?|ttf|otf))["']''', source))
        elif path.suffix == '.gz':
            original = path.with_suffix('')
            assert original.is_file(), f'Orphan gzip asset: {path}'
            assert gzip.decompress(path.read_bytes()) == original.read_bytes(), f'Stale gzip asset: {path}'
        elif path.suffix == '.wasm':
            assert path.read_bytes()[:8] == b'\0asm\x01\0\0\0', f'Invalid WASM asset: {path}'
        elif path.suffix in {'.woff', '.woff2', '.ttf', '.otf'}:
            assert path.read_bytes()[:4] in {b'wOFF', b'wOF2', b'OTTO', b'\0\x01\0\0', b'ttcf'}, f'Invalid font asset: {path}'

    missing = []
    for base, reference in references:
        if reference.startswith(('data:', 'http:', 'https:', '//', '#')):
            continue
        resolved = unquote(urlsplit(urljoin(base, reference)).path)
        target = root / resolved.lstrip('/')
        if not target.is_file() and not (target / 'index.html').is_file():
            missing.append((base, reference, resolved))
    assert not missing, f'Broken frontend asset references: {missing}'


def test_standalone_editor_sprites_resolve_for_both_themes(client):
    root = Path(__file__).resolve().parents[1] / 'web' / 'dist' / 'studio' / 'assets'
    css = '\n'.join(path.read_text(encoding='utf-8') for path in root.glob('*.css'))
    sprites = re.findall(r'url\((/[^)]+icon_small_ko(?:_dark)?\.svg)\)', css)
    assert set(sprites) == {'/studio/images/icon_small_ko.svg', '/studio/images/icon_small_ko_dark.svg'}
    for url in sprites:
        response = client.get(url)
        assert response.status_code == 200
        assert response.headers['content-type'].startswith('image/svg+xml')
        assert b'<svg' in response.content
