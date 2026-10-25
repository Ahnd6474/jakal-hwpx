import base64
import io
from pathlib import Path

import pytest
from PIL import Image
from jakal_hwpx import HwpxDocument
from server.engine import issues
from server.page_background import apply_backgrounds, background_pictures, decode_image
from server.page_layout import body_signature
from server.schemas import BackgroundPage, PageLayout
from server.template_extract import ExtractionError, run_extraction
from web_tests.test_api import client, raw_document, upload, wait


def jpeg(color='red', size=(80, 120)):
    image = Image.new('RGB', size, color)
    output = io.BytesIO(); image.save(output, format='JPEG'); image.close()
    return output.getvalue()


def extract(client, files):
    response = client.post('/api/page-templates/extract', files=[('files', (name, raw)) for name, raw in files])
    assert response.status_code == 200, response.text
    return response.json()


def test_hwpx_import_reads_sections_running_text_and_geometry_without_document_mutation(client, raw_document):
    document = HwpxDocument.from_bytes(raw_document)
    from server.page_layout import apply_layout
    apply_layout(document, PageLayout(columns=2, separator=True, header_text='머리말 <원본>', footer_text='꼬리말', page_numbers='bottom-right'))
    value = extract(client, [('쪽 양식.hwpx', document.compile(validate=False, auto_repair=False))])
    assert value['source'] == 'hwpx' and value['background_pages'] == []
    layout = value['choices'][0]['layout']
    assert layout['columns'] == 2 and layout['separator'] is True
    assert layout['header_text'] == '머리말 <원본>' and layout['footer_text'] == '꼬리말'
    assert layout['page_numbers'] == 'bottom-right'
    assert '본문 내용' in value['notes'][0]
    assert client.get('/api/documents').json() == []


def test_jpg_collection_preserves_order_and_orientation_and_creates_editable_document(client):
    value = extract(client, [('2.jpg', jpeg('blue', (120, 80))), ('1.jpg', jpeg('red'))])
    assert value['source'] == 'images' and len(value['background_pages']) == 2
    assert value['background_pages'][0]['width_mm'] > value['background_pages'][0]['height_mm']
    for page, expected in zip(value['background_pages'], [2, 0]):
        image = decode_image(BackgroundPage.model_validate(page))
        try:
            color = image.getpixel((image.width // 2, image.height // 2))
            assert color[expected] > 200
        finally:
            image.close()
    created = client.post('/api/documents/from-page-layout', json={'title': '사진 모음 양식', 'layout': value['choices'][0]['layout'], 'background_pages': value['background_pages']})
    assert created.status_code == 202, created.text
    assert wait(client, created.json()['job_id'])['status'] == 'completed'
    base = '/api/documents/' + created.json()['document_id']
    doc = client.get(base).json()
    assert len(doc['sections']) == 2 and doc['revision'] == 0
    assert all(not paragraph['text'] for section in doc['sections'] for paragraph in section['paragraphs'])
    path = client.app.state.store.directory(created.json()['document_id']) / 'revision-0.hwpx'
    reopened = HwpxDocument.open(path, validate=False)
    assert len(reopened.pictures()) == 2 and issues(reopened) == []
    assert all(section['has_background'] for section in client.get(base + '/page-layout?revision=0').json()['sections'])
    # The visual editor's complete-HWPX save path can write above the image.
    reopened.set_paragraph_text(0, 0, '배경 위에 작성')
    edited = client.post(base + '/studio?revision=0', files={'file': ('edited.hwpx', reopened.compile(validate=False, auto_repair=False))})
    assert edited.status_code == 202, edited.text
    assert wait(client, edited.json()['job_id'])['status'] == 'completed'
    checked = client.post(base + '/validations', json={'revision': 1}).json()
    assert wait(client, checked['job_id'])['status'] == 'completed'
    downloaded = client.get(base + '/download?revision=1')
    assert downloaded.status_code == 200
    saved = HwpxDocument.from_bytes(downloaded.content)
    assert len(saved.pictures()) == 2 and issues(saved) == []


def test_pdf_import_keeps_page_sizes_and_returns_preview_without_changing_source(client):
    # Use the five-page landscape fixture regardless of filesystem iteration order.
    source = Path(__file__).resolve().parents[1] / 'examples/samples/pdf/2026학년도 정시 주요사항 안내.pdf'
    original = source.read_bytes()
    value = extract(client, [(source.name, original)])
    assert value['source'] == 'pdf' and len(value['background_pages']) == 5
    assert len(value['thumbnails']) == 5
    assert value['choices'][0]['layout']['width_mm'] > value['choices'][0]['layout']['height_mm']
    assert source.read_bytes() == original
    assert client.get('/api/documents').json() == []


def test_imported_background_applies_to_current_body_preserves_text_and_can_be_removed(client, raw_document):
    identifier = upload(client, raw_document)
    base = f'/api/documents/{identifier}'
    value = extract(client, [('원본.jpg', jpeg())])
    body = {'revision': 0, 'layout': value['choices'][0]['layout'], 'background_pages': value['background_pages']}
    headers = {'Idempotency-Key': 'import-background-0001'}
    response = client.post(base + '/page-layout', json=body, headers=headers)
    assert response.status_code == 202, response.text
    assert wait(client, response.json()['job_id'])['status'] == 'completed'
    assert client.post(base + '/page-layout', json=body, headers=headers).json() == response.json()
    path = client.app.state.store.directory(identifier)
    document = HwpxDocument.open(path / 'revision-1.hwpx', validate=False)
    assert body_signature(document) == body_signature(HwpxDocument.from_bytes(raw_document))
    assert len(document.pictures()) == 1
    removed = client.post(base + '/page-layout', json={'revision': 1, 'layout': PageLayout().model_dump(), 'background_pages': []}).json()
    assert wait(client, removed['job_id'])['status'] == 'completed'
    assert client.get(base + '/page-layout?revision=2').json()['sections'][0]['has_background'] is False
    checked = client.post(base + '/validations', json={'revision': 2}).json()
    assert wait(client, checked['job_id'])['status'] == 'completed'
    assert client.get(base + '/download?revision=2').status_code == 200
    assert client.get(base + '/original').content == raw_document


@pytest.mark.parametrize('files,status', [
    ([('wrong.jpg', b'not an image')], 422), ([('wrong.pdf', b'%PDF-nothing')], 422),
    ([('x.txt', b'text')], 415), ([('empty.pdf', b'')], 422),
    ([('x.pdf', b'fake'), ('y.jpg', b'fake')], 422),
    ([(f'{i}.jpg', b'fake') for i in range(11)], 422),
])
def test_invalid_imports_do_not_create_documents_or_leave_uploads(client, files, status):
    response = client.post('/api/page-templates/extract', files=[('files', (name, raw)) for name, raw in files])
    assert response.status_code == status, response.text
    assert client.get('/api/documents').json() == []
    assert list(client.app.state.store.root.glob('layout-import-*')) == []


def test_background_json_roundtrip_is_revalidated_and_malformed_image_is_rejected(client):
    import json
    value = extract(client, [('test.jpg', jpeg())])
    payload = {'version': 2, 'name': '재사용 양식', 'layout': value['choices'][0]['layout'], 'background_pages': value['background_pages']}
    restored = extract(client, [('양식.json', json.dumps(payload).encode())])
    assert restored['name'] == '재사용 양식' and len(restored['background_pages']) == 1
    payload['background_pages'][0]['image'] = 'data:image/jpeg;base64,' + base64.b64encode(b'not jpeg data').decode()
    response = client.post('/api/page-templates/extract', files={'files': ('양식.json', json.dumps(payload).encode())})
    assert response.status_code == 422


def test_import_timeout_kills_parser_and_leaves_no_result(tmp_path):
    (tmp_path / 'input-00.jpg').write_bytes(jpeg())
    with pytest.raises(ExtractionError, match='제한 시간'):
        run_extraction(tmp_path, [{'file': 'input-00.jpg', 'name': 'test.jpg'}], timeout=0)


def test_multiple_backgrounds_cannot_overwrite_current_document(client, raw_document):
    identifier = upload(client, raw_document)
    page = {'width_mm': 210, 'height_mm': 297, 'image': 'data:image/jpeg;base64,' + base64.b64encode(jpeg()).decode()}
    response = client.post(f'/api/documents/{identifier}/page-layout', json={'revision': 0, 'layout': PageLayout().model_dump(), 'background_pages': [page, page]})
    assert response.status_code == 422
    assert client.get(f'/api/documents/{identifier}').json()['revision'] == 0


def test_background_covers_missing_even_header_and_unhides_first_page_without_removing_running_text():
    document = HwpxDocument.blank()
    document.append_header('홀수 머리말', apply_page_type='ODD', hide_first=True)
    page = BackgroundPage(width_mm=210, height_mm=297, image='data:image/jpeg;base64,' + base64.b64encode(jpeg()).decode())
    apply_backgrounds(document, [page])
    assert {header.apply_page_type for header in document.headers(0)} == {'ODD', 'EVEN'}
    assert any(header.text == '홀수 머리말' for header in document.headers(0))
    assert len(background_pictures(document.sections[0])) == 2
    from jakal_hwpx.namespaces import NS
    assert document.sections[0].root_element.xpath('.//hp:visibility', namespaces=NS)[0].get('hideFirstHeader') == '0'
    assert issues(document) == []
