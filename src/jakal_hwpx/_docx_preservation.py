"""Compatibility imports; implementation lives in ``jakal_hwpx.addons.docx``."""
from .addons.docx.preservation import DocxSource, capture_docx_source, replace_word_match

__all__ = ['DocxSource', 'capture_docx_source', 'replace_word_match']


def __getattr__(name: str):
    from .addons.docx import preservation

    return getattr(preservation, name)
