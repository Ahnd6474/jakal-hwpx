"""Page geometry and recurring page controls, independent of document content."""
from __future__ import annotations

import secrets
from pathlib import Path

from lxml import etree
from jakal_hwpx import HwpxDocument
from jakal_hwpx.namespaces import NS
from jakal_hwpx.elements import _invalidate_paragraph_layout
from .schemas import PageLayout, PageLayoutRequest
from .files import atomic_write

HP = '{' + NS['hp'] + '}'


def units(mm: float) -> int:
    return round(mm * 7200 / 25.4)


def millimeters(value: int) -> float:
    return round(value * 25.4 / 7200, 2)


def read_layouts(document: HwpxDocument) -> list[dict]:
    result = []
    for index, section in enumerate(document.sections):
        settings = document.section_settings(index)
        margins = settings.margins()
        columns = section.root_element.xpath('./hp:p/hp:run/hp:ctrl/hp:colPr', namespaces=NS)
        col = columns[0] if columns else None
        headers, footers = document.headers(index), document.footers(index)
        page_nums = section.root_element.xpath('./hp:p/hp:run/hp:ctrl/hp:pageNum', namespaces=NS)
        result.append({'section': index, 'layout': {
            'width_mm': millimeters(settings.page_width or 59528), 'height_mm': millimeters(settings.page_height or 84189),
            **{f'{key}_mm': millimeters(margins.get(key, 0)) for key in ('left', 'right', 'top', 'bottom', 'header', 'footer')},
            'columns': int(col.get('colCount', '1')) if col is not None else 1,
            'gap_mm': millimeters(int(col.get('sameGap', '0'))) if col is not None else 8,
            'separator': col is not None and any(child.tag == HP + 'colLine' and child.get('type') != 'NONE' for child in col),
            'header_text': None, 'footer_text': None, 'page_numbers': 'keep',
        }, 'header_text': ' / '.join(block.text for block in headers), 'footer_text': ' / '.join(block.text for block in footers),
            'page_number_position': page_nums[0].get('pos') if page_nums else 'NONE',
            'has_background': bool(section.root_element.xpath('.//hp:header//hp:pic[hp:shapeComment="Jakal page template background"]', namespaces=NS)),
            'column_regions': len(columns)})
    return result


def remove_controls(section, tag: str):
    # Touch only main-body recurring page controls, never controls nested in
    # tables, notes or other objects. Other children of a ctrl are preserved.
    for node in section.root_element.xpath(f'./hp:p/hp:run/hp:ctrl/hp:{tag}', namespaces=NS):
        parent = node.getparent()
        parent.remove(node)
        if not len(parent):
            parent.getparent().remove(parent)
    section.mark_modified()


def apply_layout(document: HwpxDocument, layout: PageLayout, section_index: int | None = None):
    indices = range(len(document.sections)) if section_index is None else [section_index]
    if section_index is not None and section_index >= len(document.sections):
        raise ValueError('적용할 구역을 찾을 수 없습니다. 문서를 다시 열어 주세요.')
    for index in indices:
        section = document.sections[index]
        from .page_background import attach_background, background_pictures, fit_backgrounds
        background = background_pictures(section)
        settings = document.section_settings(index)
        settings.set_page_size(width=units(layout.width_mm), height=units(layout.height_mm),
            landscape='LANDSCAPE' if layout.width_mm > layout.height_mm else 'WIDELY')
        settings.set_margins(**{key: units(getattr(layout, key + '_mm')) for key in ('left', 'right', 'top', 'bottom', 'header', 'footer')})
        # Cached body line positions describe the old width/page geometry.
        # Leave table/cell geometry intact while letting main text reflow.
        for paragraph in section.root_element.xpath('./hp:p', namespaces=NS):
            _invalidate_paragraph_layout(paragraph)
        remove_controls(section, 'colPr')
        col = etree.Element(HP + 'colPr', id='', type='NEWSPAPER', layout='LEFT', colCount=str(layout.columns), sameSz='1', sameGap=str(units(layout.gap_mm)))
        if layout.separator and layout.columns > 1:
            etree.SubElement(col, HP + 'colLine', type='SOLID', width='0.12 mm', color='#000000')
        document.append_control_xml(etree.tostring(col, encoding='unicode'), section_index=index, paragraph_index=0)
        for tag, text in (('header', layout.header_text), ('footer', layout.footer_text)):
            if text is None:
                continue
            remove_controls(section, tag)
            if text:
                wrapper = getattr(document, 'append_' + tag)(text, section_index=index, paragraph_index=0, hide_first=False)
                if tag == 'footer':
                    style = document.append_paragraph_style(alignment_horizontal='CENTER').element.get('id')
                    for p in wrapper.element.xpath('./hp:subList/hp:p', namespaces=NS):
                        p.set('paraPrIDRef', style)
        if layout.page_numbers != 'keep':
            remove_controls(section, 'pageNum')
            if layout.page_numbers != 'none':
                num = etree.Element(HP + 'pageNum', pos='BOTTOM_CENTER' if layout.page_numbers == 'bottom-center' else 'BOTTOM_RIGHT', formatType='DIGIT', sideChar='')
                document.append_control_xml(etree.tostring(num, encoding='unicode'), section_index=index, paragraph_index=0)
                settings.set_visibility(hide_first_page_num=False)
        # Backgrounds are independent of editable running text. Replacing or
        # clearing a header must not silently discard an imported page form.
        if background and not background_pictures(section):
            headers = document.headers(index)
            if not headers:
                headers = [document.append_header('', section_index=index, paragraph_index=0, hide_first=False)]
            attach_background(headers[0], background[0])
        fit_backgrounds(document, index)


def body_signature(document: HwpxDocument):
    # Changing page controls must not remove a table, picture, equation or body
    # text, even when replacing an existing header/footer is explicitly allowed.
    text = []
    objects = []
    for section in document.sections:
        for node in section.root_element.xpath('.//hp:t[not(ancestor::hp:header or ancestor::hp:footer)]', namespaces=NS):
            text.append(etree.tostring(node, method='c14n'))
        for node in section.root_element.xpath('.//*[self::hp:tbl or self::hp:pic or self::hp:equation or self::hp:ole][not(ancestor::hp:header or ancestor::hp:footer)]', namespaces=NS):
            objects.append(etree.tostring(node, method='c14n'))
    return text, objects


def process_layout(directory: Path, payload: dict, name: str):
    from .engine import inspect_archive, structure_signature
    from .studio import process_studio
    request = PageLayoutRequest.model_validate({key: value for key, value in payload.items() if key != 'settings'})
    source = directory / f'revision-{request.revision}.hwpx'
    if not source.exists():
        raise ValueError('HWPX로 변환한 문서를 다시 열어 주세요.')
    inspect_archive(source)
    document = HwpxDocument.open(source, validate=False)
    if document.is_distribution_protected:
        raise ValueError('보호 문서는 쪽 양식을 변경할 수 없습니다.')
    before = body_signature(document)
    attachments = structure_signature(document)['binaries']
    apply_layout(document, request.layout, request.section)
    if request.background_pages is not None:
        from .page_background import apply_backgrounds
        apply_backgrounds(document, request.background_pages, request.section)
    raw = document.compile(validate=False, auto_repair=False)
    reopened = HwpxDocument.from_bytes(raw, validate=False)
    after_attachments = structure_signature(reopened)['binaries']
    preserved = all(after_attachments.get(key) == value for key, value in attachments.items())
    if before != body_signature(reopened) or not preserved or (request.background_pages is None and attachments != after_attachments):
        raise ValueError('본문·표·그림·첨부 보존 검사에 실패해 쪽 양식을 저장하지 않았습니다.')
    path = directory / f'layout-{secrets.token_hex(16)}.hwpx'
    try:
        atomic_write(path, raw)
        result = process_studio(directory, payload, name, source=path)
        return result
    finally:
        path.unlink(missing_ok=True)
