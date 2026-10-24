from __future__ import annotations

import hashlib
import json
import secrets
import time

import pytest
from fastapi.testclient import TestClient

from jakal_hwpx import HwpxDocument
from server.app import create_app
from server.errors import WorkspaceError


def wait(client, job):
    until = time.monotonic() + 30
    while time.monotonic() < until:
        result = client.get('/api/jobs/' + job).json()
        if result['status'] in {'completed', 'failed'}:
            assert result['status'] == 'completed', result
            return result
        time.sleep(.05)
    pytest.fail('Document worker timed out')


@pytest.fixture
def workspace(tmp_path):
    with TestClient(create_app(tmp_path / 'private')) as client:
        source = HwpxDocument.blank()
        source.set_paragraph_text(0, 0, '원본 문단')
        source.append_paragraph('둘째 문단')
        source.append_table(rows=1, columns=1, cell_texts=[['표 원본']])
        result = client.post('/api/documents', files={'file': ('문서.hwpx', source.compile())}).json()
        wait(client, result['job_id'])
        yield client, result['document_id']


def connect(client, document, permission='propose'):
    response = client.post('/api/connections', json={'document_id': document, 'label': '테스트 연결', 'permission': permission})
    assert response.status_code == 201, response.text
    return response.json()


def call(client, grant, name, body=None):
    return client.post('/api/connected/tools/' + name, json=body or {}, headers={'Authorization': 'Bearer ' + grant['token']})


def changes(**extra):
    return {'revision': 0, 'settings_version': 0, 'request_key': secrets.token_hex(16),
        'changes': [{'id': 's0-p0', 'text': '외부에서 수정한 문단'}], **extra}


def test_connection_is_document_scoped_hashed_and_never_a_browser_cookie(workspace):
    client, document = workspace
    grant = connect(client, document)
    with client.app.state.store.connect() as db:
        row = dict(db.execute('SELECT * FROM connections WHERE id=?', (grant['id'],)).fetchone())
        assert row['token_hash'] == hashlib.sha256(grant['token'].encode()).hexdigest()
        assert grant['token'] not in str(row)
    listing = client.get(f'/api/documents/{document}/connections')
    assert grant['token'] not in listing.text and 'token_hash' not in listing.text
    result = call(client, grant, 'read_document_range')
    assert result.status_code == 200 and result.json()['items'][0]['text'] == '원본 문단'
    assert 'set-cookie' not in result.headers
    assert result.headers['Cache-Control'] == 'no-store'
    assert client.get('/api/connected/info').status_code == 401  # Cookie cannot authorize the integration.
    outsider = TestClient(client.app)
    try:
        assert outsider.get('/api/documents/' + document, headers={'Authorization': 'Bearer ' + grant['token']}).status_code == 404
        assert outsider.delete('/api/connections/' + grant['id']).status_code == 404
        assert outsider.post('/api/connections', json={'document_id': document, 'label': '침입'}).status_code == 404
    finally:
        outsider.close()


@pytest.mark.parametrize('permission', ['read', 'propose'])
def test_lower_permissions_cannot_edit_directly(workspace, permission):
    client, document = workspace
    grant = connect(client, document, permission)
    response = call(client, grant, 'apply_changes', changes())
    assert response.status_code == 403
    if permission == 'read':
        assert call(client, grant, 'propose_changes', changes()).status_code == 403
        assert call(client, grant, 'validate_document', {'revision': 0, 'settings_version': 0, 'request_key': secrets.token_hex(16)}).status_code == 403
    assert client.get('/api/documents/' + document).json()['revision'] == 0


def test_proposal_selected_approval_export_and_retry_are_safe(workspace):
    client, document = workspace
    grant = connect(client, document)
    body = changes(changes=[{'id': 's0-p0', 'text': '선택한 수정'}, {'id': 's0-t0-0-0', 'text': '선택하지 않은 표 수정'}])
    proposed = call(client, grant, 'propose_changes', body).json()
    assert call(client, grant, 'propose_changes', body).json() == proposed
    changed_body = body | {'title': '다른 요청'}
    assert call(client, grant, 'propose_changes', changed_body).status_code == 409
    proposal = proposed['proposal_id']
    assert client.get('/api/documents/' + document).json()['revision'] == 0
    review = client.get('/api/proposals/' + proposal).json()
    assert review['changes'][0]['before'] == '원본 문단'
    assert client.get(f'/api/documents/{document}/state').json()['pending_proposals'] == 1
    approval = {'revision': 0, 'settings_version': 0, 'targets': ['s0-p0']}
    accepted = client.post(f'/api/proposals/{proposal}/approve', json=approval)
    assert accepted.status_code == 202, accepted.text
    wait(client, accepted.json()['job_id'])
    assert client.post(f'/api/proposals/{proposal}/approve', json=approval).json()['job_id'] == accepted.json()['job_id']
    fresh = client.get('/api/documents/' + document).json()
    assert fresh['revision'] == 1
    assert fresh['sections'][0]['paragraphs'][0]['text'] == '선택한 수정'
    assert fresh['sections'][0]['tables'][0]['cells'][0]['text'] == '표 원본'
    assert client.post(f'/api/proposals/{proposal}/approve', json=approval | {'targets': ['s0-t0-0-0']}).status_code == 409
    assert call(client, grant, 'export_document', {'revision': 1}).status_code == 409
    validate = call(client, grant, 'validate_document', {'revision': 1, 'settings_version': 0, 'request_key': secrets.token_hex(16)})
    assert validate.status_code == 200, validate.text
    wait(client, validate.json()['job_id'])
    exported = call(client, grant, 'export_document', {'revision': 1})
    assert exported.status_code == 200
    import base64
    raw = base64.b64decode(exported.json()['resource']['blob'])
    assert hashlib.sha256(raw).hexdigest() == exported.json()['sha256']
    assert HwpxDocument.from_bytes(raw).sections[0].text_fragments()[0] == '선택한 수정'


def test_direct_edit_version_fence_browser_unsaved_and_duplicate_replay(workspace):
    client, document = workspace
    tab = secrets.token_hex(16)
    client.post(f'/api/documents/{document}/presence', json={'tab': tab, 'dirty': True, 'revision': 0})
    assert client.post('/api/connections', json={'document_id': document, 'label': '직접 수정', 'permission': 'write'}).status_code == 409
    client.post(f'/api/documents/{document}/presence', json={'tab': tab, 'dirty': False, 'revision': 0})
    grant = connect(client, document, 'write')
    body = changes()
    client.post(f'/api/documents/{document}/presence', json={'tab': tab, 'dirty': True, 'revision': 0})
    assert call(client, grant, 'apply_changes', body).json()['code'] == 'browser_unsaved'
    client.post(f'/api/documents/{document}/presence', json={'tab': tab, 'dirty': False, 'revision': 0})
    response = call(client, grant, 'apply_changes', body)
    assert response.status_code == 200, response.text
    job = response.json()['job_id']
    wait(client, job)
    assert call(client, grant, 'apply_changes', body).json()['job_id'] == job
    assert call(client, grant, 'apply_changes', body | {'request_key': secrets.token_hex(16)}).status_code == 409
    assert call(client, grant, 'apply_changes', body | {'changes': [{'id': 's0-p0', 'text': '잘못된 재전송'}]}).status_code == 409
    other = connect(client, document, 'read')
    assert call(client, other, 'get_job', {'job_id': job}).status_code == 404
    assert call(client, grant, 'get_job', {'job_id': job}).json()['status'] == 'completed'


def test_dirty_browser_lock_survives_background_timer_throttling_and_needs_explicit_release(workspace):
    client, document = workspace
    tab = secrets.token_hex(16)
    client.post(f'/api/documents/{document}/presence', json={'tab': tab, 'dirty': True, 'revision': 0})
    with client.app.state.store.connect() as db:
        expires = db.execute('SELECT expires FROM editor_presence WHERE tab=?', (tab,)).fetchone()[0]
    assert expires > time.time() + 3600  # Background/sleep must not silently unlock unsaved work.
    assert client.delete(f'/api/documents/{document}/presence/{tab}').status_code == 409
    with client.app.state.store.connect() as db:
        db.execute('UPDATE editor_presence SET expires=expires-120 WHERE tab=?', (tab,))
    editing = client.get(f'/api/documents/{document}/connections').json()['editing']
    assert editing[0]['tab'] == tab
    assert client.delete(f'/api/documents/{document}/presence/{tab}').status_code == 200
    assert connect(client, document, 'write')['permission'] == 'write'


@pytest.mark.parametrize('invalidate', ['revoke', 'session', 'expiry', 'delete'])
def test_access_stops_when_grant_document_or_session_ends(workspace, invalidate):
    client, document = workspace
    grant = connect(client, document)
    proposal = call(client, grant, 'propose_changes', changes()).json()['proposal_id']
    if invalidate == 'revoke':
        assert client.delete('/api/connections/' + grant['id']).status_code == 200
        assert client.get('/api/proposals/' + proposal).json()['status'] == 'rejected'
    elif invalidate == 'session':
        assert client.delete('/api/session').status_code == 200
    elif invalidate == 'expiry':
        with client.app.state.store.connect() as db:
            db.execute('UPDATE connections SET expires=0 WHERE id=?', (grant['id'],))
    else:
        assert client.delete('/api/documents/' + document).status_code == 200
        with client.app.state.store.connect() as db:
            assert not db.execute('SELECT 1 FROM connection_proposals WHERE id=?', (proposal,)).fetchone()
    assert call(client, grant, 'get_document_outline').status_code == 401


def test_approval_is_rechecked_in_same_transaction_as_job_acceptance(workspace, monkeypatch):
    client, document = workspace
    grant = connect(client, document)
    proposal = call(client, grant, 'propose_changes', changes()).json()['proposal_id']
    original = client.app.state.documents.edit
    def revoke_before_accept(*args, **kwargs):
        client.app.state.connections.revoke(client.app.state.connections.authenticate(grant['token'])['owner'], grant['id'])
        return original(*args, **kwargs)
    monkeypatch.setattr(client.app.state.documents, 'edit', revoke_before_accept)
    response = client.post(f'/api/proposals/{proposal}/approve', json={'revision': 0, 'settings_version': 0, 'targets': ['s0-p0']})
    assert response.status_code == 401
    assert client.get('/api/documents/' + document).json()['revision'] == 0


def test_rejected_and_stale_proposals_cannot_apply(workspace):
    client, document = workspace
    grant = connect(client, document)
    first = call(client, grant, 'propose_changes', changes()).json()['proposal_id']
    assert client.post(f'/api/proposals/{first}/reject', json={}).status_code == 200
    approval = {'revision': 0, 'settings_version': 0, 'targets': ['s0-p0']}
    assert client.post(f'/api/proposals/{first}/approve', json=approval).status_code == 409
    second = call(client, grant, 'propose_changes', changes()).json()['proposal_id']
    job = client.post(f'/api/documents/{document}/operations', json={'revision': 0, 'operations': [{'type': 'paragraph', 'section': 0, 'index': 0, 'text': '웹에서 수정'}]}).json()['job_id']
    wait(client, job)
    assert client.post(f'/api/proposals/{second}/approve', json=approval).status_code == 409


def test_input_scoping_protected_terms_and_duplicate_headers(workspace):
    client, document = workspace
    grant = connect(client, document)
    assert call(client, grant, 'read_document_range', {'document_id': secrets.token_hex(16)}).status_code == 422
    assert call(client, grant, 'read_document_range', {'limit': 100000}).status_code == 422
    assert call(client, grant, 'propose_changes', changes(changes=[{'id': '../secret', 'text': 'no'}])).status_code == 422
    client.put(f'/api/documents/{document}/settings', json={'version': 0, 'settings': {'protected_terms': ['원본']}})
    assert call(client, grant, 'propose_changes', changes(settings_version=1)).json()['code'] == 'edit_not_allowed'
    assert client.post('/api/connected/tools/get_document_outline', json={}, headers=[('Authorization', 'Bearer ' + grant['token']), ('Authorization', 'Bearer bad')]).status_code == 400
    assert client.post('/api/connected/tools/get_document_outline', json={}, headers={'Authorization': 'Bearer ' + grant['token'], 'Origin': 'https://evil.example'}).status_code == 403


def test_real_mcp_http_initialization_tools_and_error_results(workspace):
    client, document = workspace
    grant = connect(client, document)
    headers = {'Authorization': 'Bearer ' + grant['token'], 'Accept': 'application/json, text/event-stream', 'MCP-Protocol-Version': '2025-11-25'}
    def rpc(method, params):
        response = client.post('/mcp', headers=headers, json={'jsonrpc': '2.0', 'id': 1, 'method': method, 'params': params})
        assert response.status_code == 200, response.text
        assert response.headers['Cache-Control'] == 'no-store'
        assert 'set-cookie' not in response.headers
        return response.json()
    initialized = rpc('initialize', {'protocolVersion': '2025-11-25', 'capabilities': {}, 'clientInfo': {'name': 'test-client', 'version': '1'}})
    assert initialized['result']['serverInfo']['name'] == 'Jakal Document Workspace'
    tools = rpc('tools/list', {})['result']['tools']
    assert 'propose_changes' in {tool['name'] for tool in tools}
    assert 'apply_changes' not in {tool['name'] for tool in tools}
    result = rpc('tools/call', {'name': 'get_document_outline', 'arguments': {}})['result']
    assert result['structuredContent']['id'] == document
    denied = rpc('tools/call', {'name': 'apply_changes', 'arguments': changes()})['result']
    assert denied['isError'] is True
    client.delete('/api/connections/' + grant['id'])
    assert client.post('/mcp', json={}, headers=headers).status_code == 401


def test_browser_connection_fails_closed_without_bearer(workspace):
    client, _ = workspace
    assert client.post('/mcp', json={'method': 'tools/list'}).status_code == 401
    assert client.get('/mcp').status_code == 401


def test_container_local_connection_only_offers_host_usable_http_settings(workspace, monkeypatch):
    client, document = workspace
    monkeypatch.setenv('JAKAL_STDIO_CONFIG', '0')
    response = client.post('/api/connections', json={'document_id': document, 'label': 'Docker'})
    assert response.status_code == 201
    assert response.json()['stdio_config'] is None
    assert response.json()['http_config']['mcpServers']['jakal']['url'].endswith('/mcp')


def test_long_paragraphs_can_be_read_in_chunks_and_search_returns_matching_context(workspace, monkeypatch):
    client, document = workspace
    grant = connect(client, document, 'read')
    original = client.app.state.documents.view
    def view(row):
        snapshot = original(row)
        snapshot['sections'][0]['paragraphs'][0]['text'] = '가' * 12_000 + '찾을 문구' + '나' * 12_000
        return snapshot
    monkeypatch.setattr(client.app.state.documents, 'view', view)
    first = call(client, grant, 'read_document_range', {'target_id': 's0-p0'}).json()
    assert len(first['items'][0]['text']) == 10_000
    assert first['items'][0]['next_text_offset'] == 10_000
    next_part = call(client, grant, 'read_document_range', {'target_id': 's0-p0', 'text_offset': 10_000}).json()
    assert '찾을 문구' in next_part['items'][0]['text']
    found = call(client, grant, 'find_in_document', {'query': '찾을 문구'}).json()
    assert '찾을 문구' in found['items'][0]['text']
    assert found['items'][0]['text_offset'] > 0
