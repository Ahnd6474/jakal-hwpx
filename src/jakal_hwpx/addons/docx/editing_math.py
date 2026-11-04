"""Editable, native OMML equations; no rasterization or cached formula bodies."""
from __future__ import annotations

from .editing import _Block
from .equations import omml_to_hwp, hwp_to_omml, capture_properties, restore_properties


class WordEquation(_Block):
    def delete(self):
        self._check()
        self._node.getparent().remove(self._node)

    @property
    def script(self):
        self._check()
        return omml_to_hwp(self._node)

    def set_script(self, script):
        from .identity import ID
        self._check()
        replacement = restore_properties(hwp_to_omml(script), capture_properties(self._node))
        if self._node.get(ID):
            replacement.set(ID, self._node.get(ID))
        self._node.getparent().replace(self._node, replacement)
        self._node = replacement
        return self
