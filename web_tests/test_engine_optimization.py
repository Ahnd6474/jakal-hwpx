import struct
from zipfile import BadZipFile, ZIP_STORED, ZipFile

from lxml import etree
import pytest

from jakal_hwpx.namespaces import NS
from server.engine import inspect_archive, paragraph_text
from web_tests.test_api import raw_document


def test_compiled_paragraph_extraction_keeps_own_inline_text_and_excludes_nested_paragraphs():
    node = etree.fromstring(f'''<hp:p xmlns:hp="{NS['hp']}"><hp:run><hp:t>가<hp:tab/>나<hp:nbSpace/>다<hp:fwSpace/>라<hp:hyphen/>마<hp:lineBreak/>😀</hp:t><hp:tbl><hp:subList><hp:p><hp:run><hp:t>셀 텍스트</hp:t></hp:run></hp:p></hp:subList></hp:tbl><hp:t>끝</hp:t></hp:run></hp:p>''')
    assert paragraph_text(node) == '가\t나\u00a0다\u3000라\u00ad마\n😀끝'
    nested = node.find('.//hp:subList/hp:p', NS)
    assert paragraph_text(nested) == '셀 텍스트'
    orphan = etree.fromstring(f'<hp:run xmlns:hp="{NS["hp"]}"><hp:t>not a paragraph</hp:t></hp:run>')
    assert paragraph_text(orphan) == ''


def test_streamed_attachment_preflight_still_reads_all_bytes_and_checks_crc(tmp_path, raw_document):
    path = tmp_path / 'attachment.hwpx'
    path.write_bytes(raw_document)
    with ZipFile(path, 'a') as archive:
        archive.writestr('BinData/large.bin', b'x' * 3_000_000, compress_type=ZIP_STORED)
    inspect_archive(path)
    with ZipFile(path) as archive:
        info = archive.getinfo('BinData/large.bin')
    raw = bytearray(path.read_bytes())
    name_length, extra_length = struct.unpack_from('<HH', raw, info.header_offset + 26)
    offset = info.header_offset + 30 + name_length + extra_length
    raw[offset + 2_500_000] ^= 0xFF
    path.write_bytes(raw)
    with pytest.raises(BadZipFile):
        inspect_archive(path)
