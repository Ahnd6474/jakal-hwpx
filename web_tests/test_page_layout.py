import json

import pytest
from fastapi.testclient import TestClient
from jakal_hwpx import HwpxDocument
from jakal_hwpx.namespaces import NS
from server.app import create_app
from server.engine import issues, paragraph_text
from server.page_layout import apply_layout, body_signature, read_layouts, units
from server.schemas import PageLayout
from web_tests.test_api import client, raw_document, upload, wait


def test_page_template_full_flow_preserves_body_and_records_restore(client, raw_document):
    identifier = upload(client, raw_document)
    base = f'/api/documents/{identifier}'
    source = HwpxDocument.from_bytes(raw_document)
    before = body_signature(source)
    initial = client.get(base + '/page-layout?revision=0').json()
    assert initial['revision'] == 0 and initial['sections'][0]['layout']['columns'] == 1
    layout = PageLayout(columns=2, separator=True, gap_mm=10, header_text='기초 유기화학 · 김영주', footer_text='자료 확인', page_numbers='bottom-right').model_dump()
    headers = {'Idempotency-Key': 'page-layout-request-01'}
    result = client.post(base + '/page-layout', json={'revision': 0, 'settings_version': 0, 'layout': layout}, headers=headers)
    assert result.status_code == 202, result.text
    assert result.json()['kind'] == 'layout'
    assert wait(client, result.json()['job_id'])['status'] == 'completed'
    assert client.post(base + '/page-layout', json={'revision': 0, 'settings_version': 0, 'layout': layout}, headers=headers).json() == result.json()
    fresh = client.get(base).json()
    assert fresh['revision'] == 1 and fresh['validated'] is None
    assert client.get(base + '/page-layout?revision=0').status_code == 409
    state = client.get(base + '/page-layout?revision=1').json()['sections'][0]
    assert state['layout']['columns'] == 2 and state['layout']['separator'] is True
    assert state['header_text'] == layout['header_text'] and state['footer_text'] == layout['footer_text']
    assert state['page_number_position'] == 'BOTTOM_RIGHT'
    checked = client.post(base + '/validations', json={'revision': 1}).json()
    assert wait(client, checked['job_id'])['status'] == 'completed'
    exported = client.get(base + '/download?revision=1').content
    reopened = HwpxDocument.from_bytes(exported)
    assert body_signature(reopened) == before
    assert issues(reopened) == []
    assert client.get(base + '/original').content == raw_document
    assert client.get(base + '/revisions').json()['items'][0]['kind'] == 'layout'
    restored = client.post(base + '/restorations', json={'revision': 1, 'source_revision': 0}).json()
    assert wait(client, restored['job_id'])['status'] == 'completed'
    assert client.get(base + '/page-layout?revision=2').json()['sections'][0]['layout']['columns'] == 1


def test_reapplying_layout_does_not_accumulate_headers_or_column_controls(client, raw_document):
    identifier = upload(client, raw_document)
    base = f'/api/documents/{identifier}'
    for revision, columns in enumerate([2, 3, 1]):
        layout = PageLayout(columns=columns, separator=True, header_text='반복 머리말', footer_text='반복 꼬리말', page_numbers='bottom-center')
        job = client.post(base + '/page-layout', json={'revision': revision, 'layout': layout.model_dump()}).json()
        assert wait(client, job['job_id'])['status'] == 'completed'
        document = HwpxDocument.open(client.app.state.store.directory(identifier) / f'revision-{revision + 1}.hwpx')
        assert len(document.headers()) == len(document.footers()) == 1
        root = document.sections[0].root_element
        assert len(root.xpath('.//hp:colPr', namespaces=NS)) == 1
        assert len(root.xpath('.//hp:pageNum', namespaces=NS)) == 1
        assert bool(root.xpath('.//hp:colLine', namespaces=NS)) == (columns > 1)


def test_keep_headers_is_default_and_removing_them_is_explicit():
    document = HwpxDocument.blank()
    document.set_paragraph_text(0, 0, '본문을 유지합니다.')
    document.append_header('기존 머리말', paragraph_index=0)
    document.append_footer('기존 꼬리말', paragraph_index=0)
    apply_layout(document, PageLayout(columns=2))
    assert document.headers()[0].text == '기존 머리말' and document.footers()[0].text == '기존 꼬리말'
    apply_layout(document, PageLayout(header_text='', footer_text='', page_numbers='none'))
    assert document.headers() == [] and document.footers() == []
    assert paragraph_text(document.sections[0].paragraphs()[0]._element) == '본문을 유지합니다.'


def test_layout_discards_old_main_text_line_positions_only():
    from lxml import etree
    document = HwpxDocument.blank()
    document.append_paragraph('둘째 본문')
    root = document.sections[0].root_element
    for paragraph in root.xpath('./hp:p', namespaces=NS):
        etree.SubElement(paragraph, '{' + NS['hp'] + '}linesegarray')
    apply_layout(document, PageLayout(columns=2))
    assert root.xpath('./hp:p/hp:linesegarray', namespaces=NS) == []


def test_section_scope_leaves_other_page_layouts_untouched():
    document = HwpxDocument.blank()
    document.add_section(text='다음 구역')
    original = read_layouts(document)[0]
    apply_layout(document, PageLayout(columns=2, separator=True), 1)
    result = read_layouts(HwpxDocument.from_bytes(document.compile(validate=False, auto_repair=False), validate=False))
    assert result[0] == original and result[1]['layout']['columns'] == 2
    with pytest.raises(ValueError, match='구역'):
        apply_layout(document, PageLayout(), 99)


def test_new_page_template_has_no_starter_body_and_uses_validation_gate(client):
    result = client.post('/api/documents/from-page-layout', json={'title': '내 양식', 'layout': PageLayout(columns=2, separator=True, header_text='머리말').model_dump()})
    assert result.status_code == 202, result.text
    assert wait(client, result.json()['job_id'])['status'] == 'completed'
    base = '/api/documents/' + result.json()['document_id']
    doc = client.get(base).json()
    assert doc['revision'] == 0 and doc['name'] == '내 양식.hwpx'
    assert all(not paragraph['text'] for section in doc['sections'] for paragraph in section['paragraphs'])
    assert client.get(base + '/download?revision=0').status_code == 409
    job = client.post(base + '/validations', json={'revision': 0}).json()
    assert wait(client, job['job_id'])['status'] == 'completed'
    assert client.get(base + '/download?revision=0').status_code == 200


@pytest.mark.parametrize('override', [
    {'width_mm': 0}, {'width_mm': float('inf')}, {'left_mm': 99, 'right_mm': 99, 'columns': 2},
    {'height_mm': 80, 'top_mm': 40, 'bottom_mm': 40}, {'header_mm': 30, 'top_mm': 10},
    {'columns': 4}, {'header_text': '잘못된\n머리말'}, {'script': 'no'},
])
def test_invalid_page_layout_does_not_enqueue(client, raw_document, override):
    identifier = upload(client, raw_document)
    layout = {**PageLayout().model_dump(), **override}
    if override.get('width_mm') == float('inf'):
        response = client.post(f'/api/documents/{identifier}/page-layout', content=json.dumps({'revision': 0, 'layout': layout}), headers={'Content-Type': 'application/json'})
    else:
        response = client.post(f'/api/documents/{identifier}/page-layout', json={'revision': 0, 'layout': layout})
    assert response.status_code == 422, response.text
    assert client.get(f'/api/documents/{identifier}').json()['revision'] == 0


def test_header_replacement_respects_protected_terms(client, raw_document):
    document = HwpxDocument.from_bytes(raw_document)
    document.append_header('보호할 머리말', paragraph_index=0)
    identifier = upload(client, document.compile())
    base = f'/api/documents/{identifier}'
    assert client.put(base + '/settings', json={'version': 0, 'settings': {'protected_terms': ['보호할 머리말']}}).status_code == 200
    job = client.post(base + '/page-layout', json={'revision': 0, 'layout': PageLayout(header_text='').model_dump()}).json()
    failed = wait(client, job['job_id'])
    assert failed['status'] == 'failed' and '보호 문구' in failed['error']
    assert client.get(base).json()['revision'] == 0
    assert list(client.app.state.store.directory(identifier).glob('layout-*.hwpx*')) == []
    stranger = TestClient(client.app)
    try:
        assert stranger.get(base + '/page-layout?revision=0').status_code == 404
        assert stranger.post(base + '/page-layout', json={'revision': 0, 'layout': PageLayout().model_dump()}).status_code == 404
    finally:
        stranger.close()
