from __future__ import annotations

from copy import deepcopy
from typing import Any

from lxml import etree

from .namespaces import NS, local_name, resolve_tag


class HwpxXmlNode:
    def __init__(self, element: etree._Element, part: Any):
        self._element = element
        self._part = part

    @property
    def element(self) -> etree._Element:
        self._part.mark_modified()
        return self._element

    @property
    def tag(self) -> str:
        return self._element.tag

    @property
    def local_name(self) -> str:
        return local_name(self._element.tag)

    @property
    def text(self) -> str | None:
        return self._element.text

    @text.setter
    def text(self, value: str | None) -> None:
        self._part.mark_modified()
        self._element.text = value

    @property
    def tail(self) -> str | None:
        return self._element.tail

    @tail.setter
    def tail(self, value: str | None) -> None:
        self._part.mark_modified()
        self._element.tail = value

    @property
    def attributes(self) -> dict[str, str]:
        return dict(self._element.attrib)

    @property
    def children(self) -> list["HwpxXmlNode"]:
        return [HwpxXmlNode(child, self._part) for child in self._element]

    @property
    def parent(self) -> "HwpxXmlNode | None":
        parent = self._element.getparent()
        if parent is None:
            return None
        return HwpxXmlNode(parent, self._part)

    def get_attr(self, name: str, default: str | None = None) -> str | None:
        return self._element.get(name, default)

    def set_attr(self, name: str, value: Any) -> "HwpxXmlNode":
        self._part.mark_modified()
        self._element.set(name, "" if value is None else str(value))
        return self

    def remove_attr(self, name: str) -> "HwpxXmlNode":
        if name in self._element.attrib:
            self._part.mark_modified()
            del self._element.attrib[name]
        return self

    def xpath(self, expression: str, namespaces: dict[str, str] | None = None) -> list[Any]:
        values = self._element.xpath(expression, namespaces=namespaces or NS)
        if not isinstance(values, list):
            values = [values]
        wrapped: list[Any] = []
        for value in values:
            if isinstance(value, etree._Element):
                wrapped.append(HwpxXmlNode(value, self._part))
            else:
                wrapped.append(value)
        return wrapped

    def find(self, expression: str, namespaces: dict[str, str] | None = None) -> "HwpxXmlNode | None":
        result = self._element.xpath(expression, namespaces=namespaces or NS)
        for item in result:
            if isinstance(item, etree._Element):
                return HwpxXmlNode(item, self._part)
        return None

    def findall(self, expression: str, namespaces: dict[str, str] | None = None) -> list["HwpxXmlNode"]:
        return [
            HwpxXmlNode(item, self._part)
            for item in self._element.xpath(expression, namespaces=namespaces or NS)
            if isinstance(item, etree._Element)
        ]

    def iter(self, expression: str | None = None, namespaces: dict[str, str] | None = None) -> list["HwpxXmlNode"]:
        if expression is None:
            return [HwpxXmlNode(item, self._part) for item in self._element.iter()]
        return self.findall(expression, namespaces=namespaces)

    def append_child(
        self,
        tag: str,
        text: str | None = None,
        attributes: dict[str, Any] | None = None,
    ) -> "HwpxXmlNode":
        self._part.mark_modified()
        child = etree.Element(resolve_tag(tag))
        if text is not None:
            child.text = text
        if attributes:
            for key, value in attributes.items():
                child.set(key, str(value))
        self._element.append(child)
        return HwpxXmlNode(child, self._part)

    def insert_child(
        self,
        index: int,
        tag: str,
        text: str | None = None,
        attributes: dict[str, Any] | None = None,
    ) -> "HwpxXmlNode":
        self._part.mark_modified()
        child = etree.Element(resolve_tag(tag))
        if text is not None:
            child.text = text
        if attributes:
            for key, value in attributes.items():
                child.set(key, str(value))
        self._element.insert(index, child)
        return HwpxXmlNode(child, self._part)

    def ensure_child(self, tag: str, attributes: dict[str, Any] | None = None) -> "HwpxXmlNode":
        resolved = resolve_tag(tag)
        for child in self._element:
            if child.tag == resolved:
                return HwpxXmlNode(child, self._part)
        return self.append_child(tag, attributes=attributes)

    def remove(self) -> None:
        parent = self._element.getparent()
        if parent is None:
            raise ValueError("Root node cannot be removed.")
        self._part.mark_modified()
        parent.remove(self._element)

    def clear(self, keep_attributes: bool = True) -> "HwpxXmlNode":
        self._part.mark_modified()
        attributes = dict(self._element.attrib) if keep_attributes else {}
        self._element.clear()
        if keep_attributes:
            self._element.attrib.update(attributes)
        return self

    def clone(self) -> "HwpxXmlNode":
        return HwpxXmlNode(deepcopy(self._element), self._part)

    def add_existing_child(self, node: "HwpxXmlNode") -> "HwpxXmlNode":
        self._part.mark_modified()
        self._element.append(deepcopy(node._element))
        return HwpxXmlNode(self._element[-1], self._part)

    def to_xml(self, pretty: bool = False) -> str:
        return etree.tostring(self._element, encoding="unicode", pretty_print=pretty)

    def __iter__(self):
        return iter(self.children)

    def __repr__(self) -> str:
        return f"HwpxXmlNode(tag={self.tag!r})"
