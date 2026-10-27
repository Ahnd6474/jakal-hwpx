import hashlib
from fastapi.testclient import TestClient
from web_tests.test_api import client, raw_document, upload, wait


def test_rename_keeps_bytes_history_settings_and_validation(client, raw_document):
    identifier = upload(client, raw_document)
    base = f'/api/documents/{identifier}'
    doc = client.get(base).json()
    source = client.get(base + '/source?revision=0').content
    job = client.post(base + '/validations', json={'revision': 0}).json()['job_id']
    assert wait(client, job)['status'] == 'completed'
    history = client.get(base + '/revisions').json()
    response = client.patch(base + '/name', json={'revision': 0, 'expected_name': doc['name'], 'name': '회의록 최종.hwpx'})
    assert response.status_code == 200, response.text
    renamed = response.json()
    assert renamed['name'] == '회의록 최종.hwpx'
    assert renamed['revision'] == 0 and renamed['validated'] == 0
    assert renamed['sections'] == doc['sections'] and renamed['settings'] == doc['settings']
    assert client.get('/api/documents').json()[0]['name'] == renamed['name']
    assert client.get(base + '/source?revision=0').content == source
    assert client.get(base + '/revisions').json() == history
    assert client.get(base + '/original').content == raw_document
    assert 'filename*=' in client.get(base + '/download?revision=0').headers['content-disposition']
    assert client.patch(base + '/name', json={'revision': 0, 'expected_name': doc['name'], 'name': '다른 탭'}).status_code == 409


def test_copy_latest_revision_has_independent_history_and_rules(client, raw_document):
    identifier = upload(client, raw_document)
    base = f'/api/documents/{identifier}'
    rule = {'instructions': '검토용', 'protected_terms': ['원본'], 'strict_validation': True}
    assert client.put(base + '/settings', json={'version': 0, 'settings': rule}).status_code == 200
    edit = client.post(base + '/operations', json={'revision': 0, 'operations': [{'type': 'paragraph', 'section': 0, 'index': 0, 'text': '원본 문단의 현재 내용'}]}).json()
    assert wait(client, edit['job_id'])['status'] == 'completed'
    current = client.get(base + '/source?revision=1').content
    response = client.post(base + '/copies', json={'revision': 1, 'name': '새 작업.hwpx'})
    assert response.status_code == 202, response.text
    copy = response.json()
    assert wait(client, copy['job_id'])['status'] == 'completed'
    copied_base = f"/api/documents/{copy['document_id']}"
    copied = client.get(copied_base).json()
    assert copied['revision'] == 0 and copied['name'] == '새 작업.hwpx'
    assert copied['settings'] == rule
    assert hashlib.sha256(client.get(copied_base + '/original').content).digest() == hashlib.sha256(current).digest()
    assert len(client.get(copied_base + '/revisions').json()['items']) == 1
    assert client.get(base).json()['revision'] == 1
    assert len(client.get(base + '/revisions').json()['items']) == 2
    assert client.get(base + '/original').content == raw_document
    edit_copy = client.post(copied_base + '/operations', json={'revision': 0, 'operations': [{'type': 'paragraph', 'section': 0, 'index': 0, 'text': '원본 문단에서 분리한 복사본의 편집'}]}).json()
    assert wait(client, edit_copy['job_id'])['status'] == 'completed'
    check_copy = client.post(copied_base + '/validations', json={'revision': 1}).json()
    assert wait(client, check_copy['job_id'])['status'] == 'completed'
    exported = client.post(copied_base + '/exports', json={'revision': 1})
    assert exported.status_code == 200
    assert client.get(exported.json()['url']).status_code == 200
    assert client.get(base + '/source?revision=1').content == current
    assert client.post(base + '/copies', json={'revision': 0, 'name': '오래된 버전'}).status_code == 409
    assert client.delete(copied_base).status_code == 200
    assert client.get(base).status_code == 200


def test_actions_reject_unsafe_names_other_sessions_and_active_jobs(client, raw_document):
    identifier = upload(client, raw_document)
    base = f'/api/documents/{identifier}'
    name = client.get(base).json()['name']
    stranger = TestClient(client.app)
    assert stranger.patch(base + '/name', json={'revision': 0, 'expected_name': name, 'name': 'rename'}).status_code == 404
    assert stranger.post(base + '/copies', json={'revision': 0, 'name': 'copy'}).status_code == 404
    stranger.close()
    for unsafe in ['../escape', 'a\\b', 'bad:name', '...', '\x00name', ' ']:
        assert client.post(base + '/copies', json={'revision': 0, 'name': unsafe}).status_code == 422
        assert client.patch(base + '/name', json={'revision': 0, 'expected_name': name, 'name': unsafe}).status_code == 422
    store = client.app.state.store
    client.app.state.worker.close()
    with store.connect() as db:
        row = dict(db.execute('SELECT * FROM documents WHERE id=?', (identifier,)).fetchone())
    store.enqueue(identifier, row['owner'], 'validate', {'revision': 0}, 0)
    assert client.post(base + '/copies', json={'revision': 0, 'name': 'busy'}).status_code == 409
    assert client.patch(base + '/name', json={'revision': 0, 'expected_name': name, 'name': 'busy'}).status_code == 409

def test_copy_io_failure_and_quota_leave_no_partial_document(client, raw_document, monkeypatch):
    identifier = upload(client, raw_document)
    store = client.app.state.store
    original_ids = {path.name for path in store.root.iterdir() if path.is_dir()}
    def fail(*args):
        raise OSError('disk full')
    with monkeypatch.context() as patch:
        patch.setattr('server.storage.shutil.copyfile', fail)
        assert client.post(f'/api/documents/{identifier}/copies', json={'revision': 0, 'name': 'failed'}).status_code == 503
    assert {path.name for path in store.root.iterdir() if path.is_dir()} == original_ids
    assert len(client.get('/api/documents').json()) == 1
    with store.connect() as db:
        owner = db.execute('SELECT owner FROM documents WHERE id=?', (identifier,)).fetchone()[0]
    for index in range(19):
        store.create_document(owner, 'placeholder.hwpx', f'{index:032x}')
    assert client.post(f'/api/documents/{identifier}/copies', json={'revision': 0, 'name': 'over quota'}).status_code == 409
    assert {path.name for path in store.root.iterdir() if path.is_dir()} == original_ids
    assert len(client.get('/api/documents').json()) == 20


def test_converted_hwp_can_be_renamed_and_copied_without_losing_original(client):
    from pathlib import Path
    raw = Path('web/public/sample.hwp').read_bytes()
    result = client.post('/api/documents', files={'file': ('input.hwp', raw)}).json()
    assert wait(client, result['job_id'])['status'] == 'completed'
    base = f"/api/documents/{result['document_id']}"
    doc = client.get(base).json()
    assert doc['format'] == 'hwpx'
    renamed = client.patch(base + '/name', json={'revision': 0, 'expected_name': doc['name'], 'name': '회의자료'})
    assert renamed.json()['name'] == '회의자료.hwp'
    assert client.get(base + '/original').content == raw
    copy = client.post(base + '/copies', json={'revision': 0, 'name': '회의자료 사본'}).json()
    assert wait(client, copy['job_id'])['status'] == 'completed'
    copied = client.get(f"/api/documents/{copy['document_id']}").json()
    assert copied['format'] == 'hwpx' and copied['name'] == '회의자료 사본.hwpx'
    assert client.get(f"/api/documents/{copy['document_id']}/original").content == client.get(base + '/source?revision=0').content
