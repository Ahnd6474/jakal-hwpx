"""DOCX addon bundled with jakal-hwpx. Install with ``jakal-hwpx[docx]``.

WordDocument shares the core Document model. Reading, writing, package
preservation and equation mapping live here; python-docx loads only on use.
"""
from .document import WordDocument

__all__ = ["WordDocument"]
