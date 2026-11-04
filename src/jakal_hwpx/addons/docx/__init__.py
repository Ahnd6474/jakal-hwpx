"""DOCX addon bundled with jakal-hwpx. Install with ``jakal-hwpx[docx]``.

WordDocument shares the core Document model. Reading, writing, package
preservation and equation mapping live here; python-docx loads only on use.
"""
from .document import WordDocument
from .editing import WordEditor, WordContainer, WordParagraph, WordTable, WordPicture
from .editing_stories import WordSection, WordNote
from .editing_fields import WordField, WordList
from .editing_ranges import WordRange
from .editing_navigation import WordLink, WordBookmark
from .editing_math import WordEquation
from .editing_content import WordContentControl
from .endpoint import WordEndpoint, VersionConflict, ObjectNotFound, make_http_server

__all__ = ["WordDocument", "WordEditor", "WordContainer", "WordParagraph", "WordTable", "WordPicture", "WordSection", "WordNote", "WordField", "WordList", "WordRange", "WordLink", "WordBookmark", "WordEquation", "WordContentControl", "WordEndpoint", "VersionConflict", "ObjectNotFound", "make_http_server"]
