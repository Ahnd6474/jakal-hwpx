"""Bounded, stateless extraction of page forms. Uploaded code is never run."""
from __future__ import annotations

import base64
import io
import json
import math
import os
import re
import subprocess
import sys
import time
import warnings
from pathlib import Path

from PIL import Image, ImageOps
from .files import atomic_json
from .page_background import MAX_IMAGES, MAX_PAGES, MAX_PIXELS, PREFIX, checked_pages
from .page_layout import read_layouts
from .schemas import BackgroundPage, PageLayout

REPO = Path(__file__).resolve().parents[1]
MAX_RESULT = 20 * 1024 * 1024


class ExtractionError(ValueError):
    pass


def encoded(image: Image.Image, quality=93) -> str:
    data = io.BytesIO()
    image.save(data, format='JPEG', quality=quality, optimize=True, subsampling=0)
    if data.tell() > 4 * 1024 * 1024:
        reduced = image.copy()
        try:
            reduced.thumbnail((2500, 2500), Image.Resampling.LANCZOS)
            data = io.BytesIO(); reduced.save(data, format='JPEG', quality=88, optimize=True)
        finally:
            reduced.close()
    return PREFIX + base64.b64encode(data.getvalue()).decode('ascii')


def thumbnail(image: Image.Image) -> str:
    small = image.copy()
    try:
        small.thumbnail((180, 240))
        return encoded(small, 75)
    finally:
        small.close()


def page_layout(width, height):
    return PageLayout(width_mm=round(width, 2), height_mm=round(height, 2), header_mm=0, footer_mm=0).model_dump()


def raster_result(name, pages, thumbs, source):
    if not pages or len(pages) > MAX_PAGES or sum(len(p['image']) * 3 // 4 for p in pages) > MAX_IMAGES:
        raise ExtractionError('배경 양식은 최대 10쪽, 이미지 합계 12MB 이하여야 합니다. 파일을 나누어 가져와 주세요.')
    return {'name': name, 'source': source, 'choices': [
        {'label': f'{i + 1}쪽', 'layout': page_layout(page['width_mm'], page['height_mm']), 'notes': []} for i, page in enumerate(pages)],
        'background_pages': pages, 'thumbnails': thumbs,
        'notes': ['원본의 모든 글자·표·그림이 페이지 배경에 포함됩니다. 배경 글자는 개별 편집되지 않으며, 그 위에 새 내용을 작성합니다.']}


def extract_pdf(path: Path, name: str):
    import pypdfium2 as pdfium
    try:
        pdf = pdfium.PdfDocument(path)
    except pdfium.PdfiumError:
        raise ExtractionError('PDF를 열 수 없습니다. 암호를 해제하고 올바른 PDF인지 확인해 주세요.') from None
    pages, thumbs = [], []
    try:
        if not 1 <= len(pdf) <= MAX_PAGES:
            raise ExtractionError('PDF 양식은 한 번에 최대 10쪽까지 가져올 수 있습니다.')
        for index in range(len(pdf)):
            page = pdf[index]
            try:
                width, height = page.get_size()
                if not all(math.isfinite(v) and v > 0 for v in (width, height)):
                    raise ExtractionError('PDF의 용지 크기가 올바르지 않습니다.')
                physical = page_layout(width * 25.4 / 72, height * 25.4 / 72)
                scale = min(200 / 72, math.sqrt(8_000_000 / (width * height)), 3508 / max(width, height))
                bitmap = page.render(scale=scale)
                try:
                    image = bitmap.to_pil().convert('RGB')
                    try:
                        pages.append({'width_mm': physical['width_mm'], 'height_mm': physical['height_mm'], 'image': encoded(image)})
                        thumbs.append(thumbnail(image))
                    finally:
                        image.close()
                finally:
                    bitmap.close()
            finally:
                page.close()
    finally:
        pdf.close()
    return raster_result(name, pages, thumbs, 'pdf')


def extract_images(paths: list[Path], name: str):
    pages, thumbs = [], []
    for path in paths:
        with warnings.catch_warnings():
            warnings.simplefilter('error', Image.DecompressionBombWarning)
            try:
                with Image.open(path) as original:
                    if original.format not in {'JPEG', 'PNG'} or original.width * original.height > MAX_PIXELS:
                        raise ExtractionError('JPG·PNG 이미지는 한 장당 2,500만 화소 이하여야 합니다.')
                    original.load()
                    oriented = ImageOps.exif_transpose(original)
                    if oriented.mode in {'RGBA', 'LA'} or 'transparency' in oriented.info:
                        rgba = oriented.convert('RGBA')
                        image = Image.new('RGB', oriented.size, 'white')
                        image.paste(rgba, mask=rgba.getchannel('A')); rgba.close()
                    else:
                        image = oriented.convert('RGB')
                    if oriented is not original:
                        oriented.close()
            except (OSError, Image.DecompressionBombError, Image.DecompressionBombWarning):
                raise ExtractionError('이미지를 읽지 못했습니다. 올바른 JPG·PNG 파일인지 확인해 주세요.') from None
        try:
            landscape = image.width > image.height
            width, height = (297, 210) if landscape else (210, 297)
            size = (2339, 1654) if landscape else (1654, 2339)
            # A4 with white letterboxing preserves photo aspect ratio. EXIF
            # orientation is applied before fitting; no document OCR is used.
            fitted = ImageOps.contain(image, size, Image.Resampling.LANCZOS)
            image.close(); image = fitted
            canvas = Image.new('RGB', size, 'white')
            try:
                canvas.paste(image, ((size[0] - image.width) // 2, (size[1] - image.height) // 2))
                pages.append({'width_mm': width, 'height_mm': height, 'image': encoded(canvas)})
                thumbs.append(thumbnail(canvas))
            finally:
                canvas.close()
        finally:
            image.close()
    result = raster_result(name, pages, thumbs, 'images')
    result['notes'].append('JPG·PNG는 가로·세로 방향에 맞는 A4 용지에 비율을 유지해 배치했습니다. 선택한 파일 순서를 쪽 순서로 사용합니다.')
    return result


def extract_hwpx(path: Path, name: str):
    from jakal_hwpx import HwpxDocument
    from jakal_hwpx.namespaces import NS
    from .engine import inspect_archive
    try:
        inspect_archive(path)
        document = HwpxDocument.open(path, validate=False)
        if document.is_distribution_protected:
            raise ExtractionError('보호된 HWPX는 쪽 양식을 추출할 수 없습니다.')
        if not 1 <= len(document.sections) <= 20:
            raise ExtractionError('한 번에 최대 20개 구역의 쪽 양식을 가져올 수 있습니다.')
        choices = []
        for section in read_layouts(document):
            notes = []
            layout = section['layout']
            for key in ('header_text', 'footer_text'):
                text = ' '.join(section[key].split())
                layout[key] = text[:200]
                if len(text) > 200:
                    notes.append('머리말·꼬리말 텍스트는 200자까지만 가져왔습니다.')
            layout['page_numbers'] = {'NONE': 'none', 'BOTTOM_CENTER': 'bottom-center', 'BOTTOM_RIGHT': 'bottom-right'}.get(section['page_number_position'], 'keep')
            if section['column_regions'] > 1:
                notes.append('구역 안에 여러 단 설정이 있습니다. 첫 단 설정을 가져왔습니다.')
            root = document.sections[section['section']].root_element
            if root.xpath('.//hp:header//hp:tbl | .//hp:header//hp:pic | .//hp:footer//hp:tbl | .//hp:footer//hp:pic', namespaces=NS):
                notes.append('머리말·꼬리말의 표·로고·그림 서식은 포함하지 않고 텍스트와 쪽 설정을 가져왔습니다.')
            choices.append({'label': f"구역 {section['section'] + 1}", 'layout': layout, 'notes': notes})
        return {'name': name, 'source': 'hwpx', 'choices': choices, 'background_pages': [], 'thumbnails': [],
            'notes': ['HWPX의 용지·여백·단·구분선과 머리말·꼬리말 텍스트를 가져왔습니다. 본문 내용은 포함하지 않습니다.']}
    except ExtractionError:
        raise
    except Exception:
        raise ExtractionError('HWPX의 문서 구조를 읽지 못했습니다. 정상적인 HWPX 파일인지 확인해 주세요.') from None


def extract_json(path: Path, name: str):
    value = json.loads(path.read_text(encoding='utf-8-sig'))
    if value.get('version') not in {1, 2} or not isinstance(value.get('name'), str) or not 1 <= len(value['name'].strip()) <= 60:
        raise ExtractionError('올바른 쪽 양식 JSON 파일이 아닙니다.')
    layout = PageLayout.model_validate(value['layout']).model_dump()
    pages = [BackgroundPage.model_validate(p) for p in value.get('background_pages', [])]
    if pages:
        from .page_background import decode_image
        checked = checked_pages(pages)
        normalized, thumbs = [], []
        for page, raw in checked:
            image = decode_image(page)
            try:
                normalized.append({**page.model_dump(), 'image': PREFIX + base64.b64encode(raw).decode('ascii')})
                thumbs.append(thumbnail(image))
            finally:
                image.close()
        result = raster_result(value['name'], normalized, thumbs, 'json')
        result['choices'][0]['layout'] = layout
        return result
    return {'name': value['name'], 'source': 'json', 'choices': [{'label': name, 'layout': layout, 'notes': []}], 'background_pages': [], 'thumbnails': [], 'notes': []}


def extract(directory: Path, inputs: list[dict]):
    if not 1 <= len(inputs) <= MAX_PAGES:
        raise ExtractionError('이미지 모음은 최대 10장까지 가져올 수 있습니다.')
    paths = []
    for item in inputs:
        if not re.fullmatch(r'input-\d{2}\.(?:hwpx|pdf|jpg|jpeg|png|json)', item['file']):
            raise ExtractionError('올바르지 않은 가져오기 요청입니다.')
        paths.append(directory / item['file'])
    suffixes = {path.suffix for path in paths}
    name = Path(inputs[0]['name']).stem[:60] or '가져온 쪽 양식'
    if suffixes <= {'.jpg', '.jpeg', '.png'}:
        return extract_images(paths, name)
    if len(paths) != 1:
        raise ExtractionError('HWPX·PDF·JSON은 한 파일씩, JPG·PNG는 여러 장을 함께 선택해 주세요.')
    if paths[0].suffix == '.hwpx':
        return extract_hwpx(paths[0], name)
    if paths[0].suffix == '.pdf':
        return extract_pdf(paths[0], name)
    return extract_json(paths[0], name)


def run_extraction(directory: Path, inputs: list[dict], *, timeout=60):
    """The API only streams files; all parsers run with time/RSS limits."""
    atomic_json(directory / 'request.json', inputs)
    from .process_environment import parser_environment
    env = parser_environment(REPO)
    import psutil
    with subprocess.Popen([sys.executable, '-m', 'server.template_extract', str(directory)], cwd=REPO, env=env,
            stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0) as process:
        began = time.monotonic()
        try:
            while process.poll() is None:
                if time.monotonic() - began > timeout:
                    raise ExtractionError('양식 추출 제한 시간(60초)을 넘었습니다. 파일을 나누어 가져와 주세요.')
                try:
                    parent = psutil.Process(process.pid)
                    if sum(p.memory_info().rss for p in [parent, *parent.children(recursive=True)]) > 768 * 1024 * 1024:
                        raise ExtractionError('양식 추출 메모리 제한(768MB)을 넘었습니다. 파일을 나누어 가져와 주세요.')
                except psutil.NoSuchProcess:
                    pass
                time.sleep(.05)
            result = directory / 'result.json'
            if process.returncode or not result.exists() or result.stat().st_size > MAX_RESULT:
                raise ExtractionError('양식을 추출하지 못했습니다. 파일 형식과 크기를 확인해 주세요.')
            value = json.loads(result.read_text(encoding='utf-8'))
            if 'error' in value:
                raise ExtractionError(value['error'])
            return value['result']
        finally:
            if process.poll() is None:
                process.kill()
            process.wait(timeout=5)


if __name__ == '__main__':
    directory = Path(sys.argv[1])
    try:
        output = {'result': extract(directory, json.loads((directory / 'request.json').read_text(encoding='utf-8')))}
    except ExtractionError as error:
        output = {'error': str(error)}
    except ImportError:
        output = {'error': '이미지·PDF 처리 모듈이 없습니다. 서버의 web 의존성을 설치해 주세요.'}
    except Exception:
        output = {'error': '양식 파일을 읽지 못했습니다. 파일이 손상되었거나 지원 범위를 벗어났습니다.'}
    atomic_json(directory / 'result.json', output)
