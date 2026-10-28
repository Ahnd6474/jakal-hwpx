"""Load the optional Word dependency at the boundary where it is needed."""
from typing import Any, Callable


def document_factory() -> Callable[..., Any]:
    try:
        from docx import Document
    except ModuleNotFoundError as error:
        if error.name != "docx":
            raise
        raise ModuleNotFoundError(
            'The DOCX addon requires python-docx. Install it with: pip install "jakal-hwpx[docx]"'
        ) from error
    return Document
