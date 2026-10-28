"""Compatibility imports; implementation lives in ``jakal_hwpx.addons.docx``."""
from .addons.docx.writer import build_docx

__all__ = ['build_docx']


def __getattr__(name: str):
    from .addons.docx import writer

    return getattr(writer, name)
