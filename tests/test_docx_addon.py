"""The bundled DOCX addon preserves legacy APIs and optional dependencies."""
from pathlib import Path
import subprocess
import sys

from jakal_hwpx import Document, WordDocument
from jakal_hwpx.addons.docx import WordDocument as AddonWordDocument
from jakal_hwpx.word_document import WordDocument as LegacyWordDocument


def test_addon_and_legacy_imports_use_the_same_document_class(tmp_path: Path):
    assert AddonWordDocument is WordDocument is LegacyWordDocument
    assert issubclass(AddonWordDocument, Document)
    assert AddonWordDocument.__module__ == "jakal_hwpx.addons.docx.document"
    word = AddonWordDocument.blank()
    word.append_paragraph("Addon document")
    original = word.save(tmp_path / "input.docx")
    source = Document.open(original)
    assert type(source) is AddonWordDocument
    assert source.save(tmp_path / "copy.docx").read_bytes() == original.read_bytes()
    middle = source.save(tmp_path / "middle.hwpx")
    restored = Document.open(middle).to_docx_document()
    assert type(restored) is AddonWordDocument
    assert restored.word_document.paragraphs[0].text == "Addon document"


def test_core_and_addon_imports_work_without_python_docx():
    src = Path(__file__).resolve().parents[1] / "src"
    code = f"""
import importlib.abc
import sys
sys.path.insert(0, {str(src)!r})
class BlockDocx(importlib.abc.MetaPathFinder):
    def find_spec(self, fullname, path=None, target=None):
        if fullname == 'docx' or fullname.startswith('docx.'):
            raise ModuleNotFoundError('No module named docx', name='docx')
sys.meta_path.insert(0, BlockDocx())
from jakal_hwpx import Document, HwpxDocument
assert not any(name.startswith('jakal_hwpx.addons.docx') for name in sys.modules)
document = HwpxDocument.blank()
document.append_paragraph('Core still works')
assert 'Core still works' in document.get_document_text()
from jakal_hwpx.addons.docx import WordDocument
assert 'docx' not in sys.modules
for operation in [WordDocument.blank, lambda: Document.read_docx('missing.docx'),
                  lambda: Document.blank().to_docx_document()]:
    try:
        operation()
    except ModuleNotFoundError as error:
        assert 'jakal-hwpx[docx]' in str(error)
    else:
        raise AssertionError('A Word operation must require the optional dependency')
assert Document.blank().source_format is None
"""
    subprocess.run([sys.executable, "-I", "-c", code], check=True, capture_output=True, text=True)
