"""Raster page forms stored as repeating, behind-text header images."""
from __future__ import annotations

import base64
import binascii
import io
import secrets
import warnings

from PIL import Image, ImageOps
from lxml import etree
from jakal_hwpx.namespaces import NS
from .schemas import BackgroundPage
from .page_layout import units

MAX_PIXELS = 25_000_000
MAX_PAGES = 10
MAX_IMAGES = 12 * 1024 * 1024
PREFIX = 'data:image/jpeg;base64,'
MARKER = 'Jakal page template background'
Image.MAX_IMAGE_PIXELS = MAX_PIXELS


def decode_image(page: BackgroundPage) -> Image.Image:
    if not page.image.startswith(PREFIX):
        raise ValueError('배경 양식은 검증된 JPEG 이미지여야 합니다.')
    try:
        raw = base64.b64decode(page.image[len(PREFIX):], validate=True)
        with warnings.catch_warnings():
            warnings.simplefilter('error', Image.DecompressionBombWarning)
            with Image.open(io.BytesIO(raw)) as source:
                if source.format != 'JPEG' or source.width * source.height > MAX_PIXELS:
                    raise ValueError('배경 이미지의 형식 또는 크기가 지원 범위를 벗어났습니다.')
                source.load()
                return ImageOps.exif_transpose(source).convert('RGB')
    except (binascii.Error, OSError, Image.DecompressionBombError, Image.DecompressionBombWarning):
        raise ValueError('배경 이미지 파일을 읽지 못했습니다. 양식을 다시 가져와 주세요.') from None


def checked_pages(pages: list[BackgroundPage]):
    if len(pages) > MAX_PAGES or sum(len(page.image) * 3 // 4 for page in pages) > MAX_IMAGES:
        raise ValueError('배경 양식은 최대 10쪽, 이미지 합계 12MB 이하여야 합니다.')
    result = []
    for page in pages:
        image = decode_image(page)
        try:
            output = io.BytesIO()
            image.save(output, format='JPEG', quality=90, optimize=True)
            result.append((page, output.getvalue()))
        finally:
            image.close()
    if sum(len(data) for _, data in result) > MAX_IMAGES:
        raise ValueError('배경 양식의 이미지 합계가 12MB를 넘었습니다.')
    return result


def background_pictures(section):
    return section.root_element.xpath('.//hp:header//hp:pic[hp:shapeComment=$marker]', namespaces=NS, marker=MARKER)


def attach_background(header, picture):
    sublist = header.element.find('hp:subList', NS)
    paragraph = etree.SubElement(sublist, '{' + NS['hp'] + '}p', id=str(secrets.randbits(31)),
        paraPrIDRef='0', styleIDRef='0', pageBreak='0', columnBreak='0', merged='0')
    etree.SubElement(paragraph, '{' + NS['hp'] + '}run', charPrIDRef='0').append(picture)
    # The renderer handles picture paragraphs separately from running text.
    # A final, page-relative paragraph preserves text and avoids shifting it.


def fit_backgrounds(document, section_index: int):
    """Keep physical-page coverage when paper size or header margins change."""
    settings = document.section_settings(section_index)
    margins = settings.margins()
    for pic in background_pictures(document.sections[section_index]):
        original = pic.find('hp:orgSz', NS)
        width, height = int(original.get('width')), int(original.get('height'))
        ratio = min(settings.page_width / width, settings.page_height / height)
        pw, ph = round(width * ratio), round(height * ratio)
        for tag in ('curSz', 'sz'):
            node = pic.find('hp:' + tag, NS)
            node.set('width', str(pw)); node.set('height', str(ph))
        scale = pic.find('hp:renderingInfo/hc:scaMatrix', NS)
        scale.set('e1', f'{pw / width:.6f}'); scale.set('e5', f'{ph / height:.6f}')
        position = pic.find('hp:pos', NS)
        position.set('horzRelTo', 'COLUMN'); position.set('vertRelTo', 'PAGE')
        position.set('horzOffset', str((settings.page_width - pw) // 2 - margins.get('left', 0)))
        position.set('vertOffset', str((settings.page_height - ph) // 2 - margins.get('top', 0)))
    document.sections[section_index].mark_modified()


def replace_background(document, section_index: int, page: BackgroundPage | None, raw: bytes | None = None):
    section = document.sections[section_index]
    for pic in background_pictures(section):
        run = pic.getparent()
        run.remove(pic)
        paragraph = run.getparent()
        if len(run) == 0 and len(paragraph) == 1:
            paragraph.getparent().remove(paragraph)
    if page is None:
        section.mark_modified()
        return
    headers = document.headers(section_index)
    if not headers:
        headers = [document.append_header('', section_index=section_index, paragraph_index=0, hide_first=False)]
    types = {header.apply_page_type for header in headers}
    if 'BOTH' not in types:
        for missing in {'ODD', 'EVEN'} - types:
            headers.append(document.append_header('', section_index=section_index, paragraph_index=0, apply_page_type=missing))
    settings = document.section_settings(section_index)
    settings.set_visibility(hide_first_header=False)
    margins = settings.margins()
    width, height = settings.page_width, settings.page_height
    ratio = min(width / units(page.width_mm), height / units(page.height_mm))
    picture_width, picture_height = round(units(page.width_mm) * ratio), round(units(page.height_mm) * ratio)
    for header in headers:
        pic = document.append_picture('page-background.jpg', raw, section_index=section_index, paragraph_index=0,
            width=picture_width, height=picture_height, media_type='image/jpeg', treat_as_char=False,
            text_wrap='BEHIND_TEXT', flow_with_text=False, allow_overlap=True, vert_rel_to='PAGE', horz_rel_to='COLUMN',
            # Header page/column coordinates start at its content frame.
            # Offset margins to cover the physical page, independent of text.
            vert_offset=(height - picture_height) // 2 - margins.get('top', 0),
            horz_offset=(width - picture_width) // 2 - margins.get('left', 0), shape_comment=MARKER)
        element = pic.element
        element.set('zOrder', '-1'); element.set('lock', '1')
        element.getparent().remove(element)
        attach_background(header, element)
    section.mark_modified()


def apply_backgrounds(document, pages: list[BackgroundPage], section_index: int | None = None):
    if len(pages) > 1:
        raise ValueError('현재 문서에는 배경 한 쪽을 선택해 적용해 주세요. 여러 쪽 양식은 새 문서로 만들 수 있습니다.')
    checked = checked_pages(pages)
    indices = range(len(document.sections)) if section_index is None else [section_index]
    for index in indices:
        if checked:
            replace_background(document, index, checked[0][0], checked[0][1])
        else:
            replace_background(document, index, None)


def make_background_document(layout, pages: list[BackgroundPage]):
    from jakal_hwpx import HwpxDocument
    from .page_layout import apply_layout
    document = HwpxDocument.blank()
    checked = checked_pages(pages)
    # Clone blank sections before inserting any pictures to avoid copying a
    # previous page's background and IDs into every subsequent section.
    for _ in checked[1:]:
        document.add_section(text='')
    for index, (page, raw) in enumerate(checked):
        page_layout = layout.model_copy(update={'width_mm': page.width_mm, 'height_mm': page.height_mm})
        page_layout = type(layout).model_validate(page_layout.model_dump())
        apply_layout(document, page_layout, index)
        replace_background(document, index, page, raw)
    return document
