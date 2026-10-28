"""Compatibility imports; implementation lives in ``jakal_hwpx.addons.docx``."""
from .addons.docx.reader import populate_from_docx

__all__ = ['populate_from_docx']


def __getattr__(name: str):
    from .addons.docx import reader

    return getattr(reader, name)
