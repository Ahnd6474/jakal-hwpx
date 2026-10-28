"""Compatibility imports; implementation lives in ``jakal_hwpx.addons.docx``."""
from .addons.docx.equations import hwp_to_omml, omml_to_hwp

__all__ = ['hwp_to_omml', 'omml_to_hwp']


def __getattr__(name: str):
    from .addons.docx import equations

    return getattr(equations, name)
