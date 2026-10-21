from __future__ import annotations

from lxml import etree

from .elements import BookmarkXml, FieldXml
from .namespaces import qname


class HwpxDocumentFieldMixin:
    def append_bookmark(
        self,
        name: str,
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> BookmarkXml:
        self._ensure_editable_sections()
        paragraph = self._resolve_paragraph_for_insert(section_index=section_index, paragraph_index=paragraph_index)
        run = self._append_run(paragraph, char_pr_id=char_pr_id)
        ctrl = etree.SubElement(run, qname("hp", "ctrl"))
        bookmark = etree.SubElement(ctrl, qname("hp", "bookmark"))
        bookmark.set("name", name)
        etree.SubElement(run, qname("hp", "t")).text = ""
        section = self.sections[section_index]
        section.mark_modified()
        return BookmarkXml(self, section, bookmark)

    def append_field(
        self,
        *,
        field_type: str,
        display_text: str | None = None,
        name: str | None = None,
        parameters: dict[str, int | str] | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
        editable: bool = False,
        dirty: bool = False,
    ) -> FieldXml:
        self._ensure_editable_sections()
        paragraph = self._resolve_paragraph_for_insert(section_index=section_index, paragraph_index=paragraph_index)
        begin_id = self._next_control_number(".//hp:fieldBegin/@id")
        field_id = self._next_control_number(".//hp:fieldBegin/@fieldid | .//hp:fieldEnd/@fieldid")

        begin_run = self._append_run(paragraph, char_pr_id=char_pr_id)
        begin_ctrl = etree.SubElement(begin_run, qname("hp", "ctrl"))
        field_begin = etree.SubElement(begin_ctrl, qname("hp", "fieldBegin"))
        field_begin.set("id", begin_id)
        field_begin.set("type", field_type)
        field_begin.set("name", name or "")
        field_begin.set("editable", "1" if editable else "0")
        field_begin.set("dirty", "1" if dirty else "0")
        field_begin.set("zorder", "0")
        field_begin.set("fieldid", field_id)

        if parameters:
            params = etree.SubElement(field_begin, qname("hp", "parameters"))
            params.set("cnt", str(len(parameters)))
            params.set("name", "")
            for key, value in parameters.items():
                tag_name = "integerParam" if isinstance(value, int) else "stringParam"
                parameter = etree.SubElement(params, qname("hp", tag_name))
                parameter.set("name", key)
                parameter.text = str(value)

        if display_text is not None:
            display_run = self._append_run(paragraph, char_pr_id=char_pr_id)
            etree.SubElement(display_run, qname("hp", "t")).text = display_text

        end_run = self._append_run(paragraph, char_pr_id=char_pr_id)
        end_ctrl = etree.SubElement(end_run, qname("hp", "ctrl"))
        field_end = etree.SubElement(end_ctrl, qname("hp", "fieldEnd"))
        field_end.set("beginIDRef", begin_id)
        field_end.set("fieldid", field_id)

        section = self.sections[section_index]
        section.mark_modified()
        return FieldXml(self, section, field_begin)

    def append_hyperlink(
        self,
        target: str,
        *,
        display_text: str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> FieldXml:
        parameters: dict[str, int | str] = {"Command": target, "Path": target}
        if target.startswith("mailto:"):
            parameters["Category"] = "HWPHYPERLINK_TYPE_EMAIL"
        elif target.startswith(("http://", "https://")):
            parameters["Category"] = "HWPHYPERLINK_TYPE_WEB"
        else:
            parameters["Category"] = "HWPHYPERLINK_TYPE_PATH"
        return self.append_field(
            field_type="HYPERLINK",
            display_text=display_text,
            parameters=parameters,
            section_index=section_index,
            paragraph_index=paragraph_index,
            char_pr_id=char_pr_id,
        )

    def append_mail_merge_field(
        self,
        field_name: str,
        *,
        display_text: str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> FieldXml:
        field = self.append_field(
            field_type="MAILMERGE",
            display_text=display_text,
            name=field_name,
            parameters={"FieldName": field_name, "MergeField": field_name},
            section_index=section_index,
            paragraph_index=paragraph_index,
            char_pr_id=char_pr_id,
        )
        field.configure_mail_merge(field_name, display_text=display_text)
        return field

    def append_calculation_field(
        self,
        expression: str,
        *,
        display_text: str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> FieldXml:
        field = self.append_field(
            field_type="FORMULA",
            display_text=display_text,
            parameters={"Expression": expression, "Command": expression},
            section_index=section_index,
            paragraph_index=paragraph_index,
            char_pr_id=char_pr_id,
        )
        field.configure_calculation(expression, display_text=display_text)
        return field

    def append_cross_reference(
        self,
        bookmark_name: str,
        *,
        display_text: str | None = None,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
    ) -> FieldXml:
        field = self.append_field(
            field_type="CROSSREF",
            display_text=display_text,
            name=bookmark_name,
            parameters={"BookmarkName": bookmark_name, "Path": bookmark_name},
            section_index=section_index,
            paragraph_index=paragraph_index,
            char_pr_id=char_pr_id,
        )
        field.configure_cross_reference(bookmark_name, display_text=display_text)
        return field


for _compat_method in HwpxDocumentFieldMixin.__dict__.values():
    if callable(_compat_method):
        _compat_method.__module__ = "jakal_hwpx.document"

