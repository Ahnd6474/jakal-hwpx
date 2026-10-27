import pytest
from jakal_hwpx import HwpxDocument
from server.templates import make_template, TEMPLATES
from server.engine import issues
from web_tests.test_api import client, wait


@pytest.mark.parametrize('kind', [item['id'] for item in TEMPLATES])
def test_templates_roundtrip_are_valid(kind):
    raw = make_template(kind, '탐구 <제목> 😀', 20)
    document = HwpxDocument.from_bytes(raw)
    assert issues(document) == []
    assert document.section_settings().margins()['left'] == 5669


def test_created_template_uses_normal_edit_validation_and_download_flow(client):
    catalog = client.get('/api/templates').json()
    assert len(catalog) == 4
    assert {item['id'] for item in catalog} == {'blank', 'report', 'meeting', 'table'}
    result = client.post('/api/documents/from-template', json={'template': 'table', 'title': '../나의 문서'})
    assert result.status_code == 202, result.text
    data = result.json()
    assert wait(client, data['job_id'])['status'] == 'completed'
    base = '/api/documents/' + data['document_id']
    document = client.get(base).json()
    assert document['revision'] == 0 and document['format'] == 'hwpx'
    assert document['name'] == '나의 문서.hwpx'
    assert document['report']['issues'] == []
    assert len(document['sections'][0]['tables']) == 1
    target = next(p for p in document['sections'][0]['paragraphs'] if p['editable'] and '메모' in p['text'])
    edited = client.post(base + '/operations', json={'revision': 0, 'operations': [{'type': 'paragraph', 'section': 0, 'index': target['index'], 'text': '변경한 메모'}]})
    assert edited.status_code == 202
    assert wait(client, edited.json()['job_id'])['status'] == 'completed'
    checked = client.post(base + '/validations', json={'revision': 1})
    assert wait(client, checked.json()['job_id'])['status'] == 'completed'
    exported = client.post(base + '/exports', json={'revision': 1})
    assert exported.status_code == 200
    raw = client.get(exported.json()['url']).content
    assert issues(HwpxDocument.from_bytes(raw)) == []


@pytest.mark.parametrize('payload', [
    {'template': '../file', 'title': 'test'},
    {'template': 'exam', 'title': 'test', 'question_count': 21},
    {'template': 'blank', 'title': ' /:*? '},
    {'template': 'blank', 'title': 'x', 'script': 'no'},
])
def test_invalid_template_requests_do_not_create_documents(client, payload):
    assert client.post('/api/documents/from-template', json=payload).status_code == 422
    assert client.get('/api/documents').json() == []
