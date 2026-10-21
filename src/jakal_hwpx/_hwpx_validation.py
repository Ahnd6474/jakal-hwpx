from __future__ import annotations

import json
import os
import zipfile
from pathlib import Path
from typing import TYPE_CHECKING, Any

from lxml import etree

from ._hwpx_common import (
    _FORM_METADATA_COMMAND_PREFIX,
    _FORM_TAGS,
    _MEMO_CARRIER_FIELD_TYPE,
    _bool_like_literal_is_valid,
    _expected_binary_media_type,
    _graphic_attribute_xpath,
    _issue,
    _new_temp_dir,
    _normalize_path,
)
from .exceptions import HwpxValidationError, ValidationIssue
from .namespaces import NS
from .parts import BinaryDataPart, ContainerPart, ContentHpfPart, HeaderPart, MimetypePart, XmlPart


class HwpxValidationMixin:
    """Package validation and strict-lint behavior shared by ``HwpxDocument``."""

    if TYPE_CHECKING:
        _parts: dict[str, Any]
        _zip_infos: dict[str, zipfile.ZipInfo]
        duplicate_entries: list[str]
        sections: list[Any]
        is_distribution_protected: bool
        content_hpf: ContentHpfPart
        header: HeaderPart

        @classmethod
        def open(cls, path: str | os.PathLike[str], *, validate: bool = True) -> Any: ...

        def save(self, path: str | os.PathLike[str], validate: bool = True, *, auto_repair: bool = True) -> Path: ...
        def _cleanup_temp_dir(self, path: Path) -> None: ...
        def control_preservation_validation_errors(self) -> list[ValidationIssue]: ...
        def stale_paragraph_layout_validation_errors(self) -> list[ValidationIssue]: ...
        def list_part_paths(self) -> list[str]: ...
        def tables(self, *, section_index: int | None = None) -> list[Any]: ...
        def styles(self) -> list[Any]: ...
        def paragraph_styles(self) -> list[Any]: ...
        def character_styles(self) -> list[Any]: ...
        def memo_shapes(self) -> list[Any]: ...
        def bookmarks(self) -> list[Any]: ...
        def fields(self) -> list[Any]: ...
        def pictures(self, *, section_index: int | None = None) -> list[Any]: ...
        def shapes(self, *, section_index: int | None = None) -> list[Any]: ...
        def charts(self, *, section_index: int | None = None) -> list[Any]: ...
        def cross_references(self) -> list[Any]: ...

    def validation_errors(self) -> list[ValidationIssue]:
        errors: list[ValidationIssue] = []
        content_hpf = self._parts.get("Contents/content.hpf")
        is_distribution_protected = isinstance(content_hpf, ContentHpfPart) and content_hpf.is_distribution_protected
        required = [
            "mimetype",
            "version.xml",
            "Contents/content.hpf",
            "META-INF/container.xml",
        ]
        if not is_distribution_protected:
            required.append("Contents/header.xml")
        for path in required:
            if path not in self._parts:
                errors.append(_issue("missing_part", f"Missing required part: {path}", part_path=path))

        if self.duplicate_entries:
            errors.append(
                _issue(
                    "duplicate_zip_entry",
                    f"Duplicate zip entries found: {', '.join(sorted(set(self.duplicate_entries)))}",
                )
            )

        mimetype = self._parts.get("mimetype")
        if isinstance(mimetype, MimetypePart) and mimetype.mime.strip() != "application/hwp+zip":
            errors.append(
                _issue("mimetype", "mimetype must be application/hwp+zip", part_path="mimetype")
            )

        xml_parts = [part for part in self._parts.values() if isinstance(part, XmlPart)]
        for part in xml_parts:
            try:
                part.to_bytes()
            except Exception as exc:  # noqa: BLE001
                errors.append(
                    _issue("xml_serialization", f"Invalid XML part serialization: {exc}", part_path=part.path)
                )

        container = self._parts.get("META-INF/container.xml")
        if isinstance(container, ContainerPart):
            for rootfile_path in container.rootfile_paths():
                normalized = _normalize_path(rootfile_path)
                if normalized not in self._parts:
                    if normalized == "Preview/PrvText.txt":
                        continue
                    errors.append(
                        _issue(
                            "container_reference",
                            f"container.xml references missing rootfile: {normalized}",
                            part_path="META-INF/container.xml",
                        )
                    )

        if isinstance(content_hpf, ContentHpfPart):
            manifest_ids: set[str] = set()
            manifest_hrefs: dict[str, str] = {}
            for item in content_hpf.manifest_items():
                item_id = item.get("id")
                href = item.get("href")
                if not item_id:
                    errors.append(
                        _issue(
                            "manifest_item",
                            "content.hpf manifest item is missing id",
                            part_path="Contents/content.hpf",
                        )
                    )
                elif item_id in manifest_ids:
                    errors.append(
                        _issue(
                            "manifest_item",
                            f"Duplicate content.hpf manifest id: {item_id}",
                            part_path="Contents/content.hpf",
                        )
                    )
                else:
                    manifest_ids.add(item_id)
                if not href:
                    errors.append(
                        _issue(
                            "manifest_item",
                            "content.hpf manifest item is missing href",
                            part_path="Contents/content.hpf",
                        )
                    )
                    continue
                normalized_href = _normalize_path(href)
                if normalized_href in manifest_hrefs:
                    errors.append(
                        _issue(
                            "manifest_item",
                            f"Duplicate content.hpf manifest href: {normalized_href}",
                            part_path="Contents/content.hpf",
                            context=f"ids={manifest_hrefs[normalized_href]},{item_id or ''}",
                        )
                    )
                elif item_id:
                    manifest_hrefs[normalized_href] = item_id
                if normalized_href not in self._parts:
                    errors.append(
                        _issue(
                            "manifest_reference",
                            f"content.hpf manifest references missing part: {href}",
                            part_path="Contents/content.hpf",
                        )
                    )

            for itemref in content_hpf.spine_itemrefs():
                item_id = itemref.get("idref")
                if not item_id:
                    errors.append(
                        _issue(
                            "spine_itemref",
                            "content.hpf spine itemref is missing idref",
                            part_path="Contents/content.hpf",
                        )
                    )
                elif item_id not in manifest_ids:
                    errors.append(
                        _issue(
                            "spine_itemref",
                            f"content.hpf spine references unknown manifest id: {item_id}",
                            part_path="Contents/content.hpf",
                        )
                    )

        header = self._parts.get("Contents/header.xml")
        if isinstance(header, HeaderPart):
            if header.section_count and header.section_count != len(self.sections):
                errors.append(
                    _issue(
                        "header_section_count",
                        f"header.xml secCnt={header.section_count} does not match section count={len(self.sections)}",
                        part_path="Contents/header.xml",
                    )
                )

        if not is_distribution_protected and not self.sections:
            errors.append(_issue("document_structure", "Document must contain at least one section part"))

        errors.extend(self.control_preservation_validation_errors())
        errors.extend(self.stale_paragraph_layout_validation_errors())
        return errors

    def validate(self) -> None:
        errors = self.validation_errors()
        if errors:
            raise HwpxValidationError(errors)

    def roundtrip_validate(self) -> None:
        temp_dir = _new_temp_dir("jakal_hwpx_")
        try:
            temp_path = temp_dir / "roundtrip.hwpx"
            self.save(temp_path)
            reopened = type(self).open(temp_path)
            reopened.validate()
        finally:
            self._cleanup_temp_dir(temp_dir)

    def xml_validation_errors(self) -> list[ValidationIssue]:
        errors: list[ValidationIssue] = []
        expected_roots = {
            "version.xml": "HCFVersion",
            "Contents/content.hpf": "package",
            "Contents/header.xml": "head",
            "settings.xml": "HWPApplicationSetting",
            "META-INF/container.xml": "container",
            "META-INF/manifest.xml": "manifest",
            "META-INF/container.rdf": "RDF",
        }
        for path, expected_local_name in expected_roots.items():
            part = self._parts.get(path)
            if part is None:
                continue
            if not isinstance(part, XmlPart):
                errors.append(
                    _issue(
                        "xml_part_type",
                        f"{path} is expected to be XML but is loaded as {type(part).__name__}",
                        part_path=path,
                    )
                )
                continue
            actual = part.root.local_name
            if actual != expected_local_name:
                errors.append(
                    _issue(
                        "xml_root",
                        f"{path} root element must be {expected_local_name}, got {actual}",
                        part_path=path,
                    )
                )
        for section_index, section in enumerate(self.sections):
            if section.root.local_name != "sec":
                errors.append(
                    _issue(
                        "xml_root",
                        f"{section.path} root element must be sec, got {section.root.local_name}",
                        part_path=section.path,
                        section_index=section_index,
                    )
                )
            if not section.readonly_root_element.xpath("./hp:p", namespaces=NS):
                errors.append(
                    _issue(
                        "section_structure",
                        f"{section.path} must contain at least one paragraph",
                        part_path=section.path,
                        section_index=section_index,
                    )
                )
            for table in self.tables(section_index=section_index):
                if table.row_count < 1 or table.column_count < 1:
                    errors.append(
                        _issue(
                            "table_structure",
                            f"{section.path} contains a table with invalid row/column count",
                            part_path=section.path,
                            section_index=section_index,
                        )
                    )
                for cell in table.cells():
                    if cell.row_span < 1 or cell.col_span < 1:
                        errors.append(
                            _issue(
                                "table_cell_span",
                                f"{section.path} table cell has invalid span",
                                part_path=section.path,
                                section_index=section_index,
                                cell_row=cell.row,
                                cell_column=cell.column,
                            )
                        )
        return errors

    def schema_validation_errors(self, schema_map: dict[str, str | os.PathLike[str]]) -> list[ValidationIssue]:
        validators: dict[Path, etree.XMLSchema] = {}
        errors: list[ValidationIssue] = []
        for part_path, schema_path in schema_map.items():
            part = self._parts.get(_normalize_path(part_path))
            if part is None:
                errors.append(
                    _issue("schema_validation", f"Schema configured for missing part: {part_path}", part_path=part_path)
                )
                continue
            if not isinstance(part, XmlPart):
                errors.append(
                    _issue("schema_validation", f"Schema validation requires XML part: {part_path}", part_path=part_path)
                )
                continue
            resolved_schema_path = Path(schema_path)
            if resolved_schema_path not in validators:
                schema_doc = etree.parse(str(resolved_schema_path))
                validators[resolved_schema_path] = etree.XMLSchema(schema_doc)
            validator = validators[resolved_schema_path]
            if not validator.validate(part._root):
                for entry in validator.error_log:
                    errors.append(
                        _issue(
                            "schema_validation",
                            entry.message,
                            part_path=part_path,
                            line=entry.line,
                            column=entry.column,
                        )
                    )
        return errors

    def reference_validation_errors(self) -> list[ValidationIssue]:
        errors = self.validation_errors()
        style_ids = {style.style_id for style in self.styles() if style.style_id is not None} if not self.is_distribution_protected else set()
        para_pr_ids = {
            style.style_id for style in self.paragraph_styles() if style.style_id is not None
        } if not self.is_distribution_protected else set()
        char_pr_ids = {
            style.style_id for style in self.character_styles() if style.style_id is not None
        } if not self.is_distribution_protected else set()
        manifest_ids = {item.get("id") for item in self.content_hpf.manifest_items() if item.get("id")}
        bookmark_names: set[str] = set()
        field_begin_ids: dict[str, str] = {}
        field_id_map: dict[str, str] = {}

        for bookmark in self.bookmarks():
            if not bookmark.name:
                errors.append(_issue("bookmark", "bookmark is missing name"))
                continue
            if bookmark.name in bookmark_names:
                errors.append(_issue("bookmark", f"duplicate bookmark name: {bookmark.name}"))
            bookmark_names.add(bookmark.name)

        for field in self.fields():
            if field.control_id:
                if field.control_id in field_begin_ids:
                    errors.append(_issue("field_reference", f"duplicate fieldBegin id: {field.control_id}"))
                field_begin_ids[field.control_id] = field.field_id or ""
            if field.field_id:
                if field.field_id in field_id_map:
                    errors.append(_issue("field_reference", f"duplicate fieldid: {field.field_id}"))
                field_id_map[field.field_id] = field.control_id or ""
            if (field.field_type or "").upper() == _MEMO_CARRIER_FIELD_TYPE:
                parameters = field.parameter_map()
                if not _bool_like_literal_is_valid(parameters.get("Visible")):
                    errors.append(
                        _issue(
                            "metadata_schema",
                            "JAKAL_MEMO carrier Visible parameter must be a bool-like literal.",
                            part_path=getattr(field.section, "path", None),
                            xpath=".//hp:fieldBegin[@type='JAKAL_MEMO']",
                        )
                    )
                order_value = parameters.get("Order")
                if order_value not in (None, ""):
                    try:
                        int(order_value)
                    except (TypeError, ValueError):
                        errors.append(
                            _issue(
                                "metadata_schema",
                                "JAKAL_MEMO carrier Order parameter must be an integer.",
                                part_path=getattr(field.section, "path", None),
                                xpath=".//hp:fieldBegin[@type='JAKAL_MEMO']",
                            )
                        )

        for section_index, section in enumerate(self.sections):
            for paragraph_index, paragraph in enumerate(
                section.readonly_root_element.xpath("./hp:p", namespaces=NS)
            ):
                style_id = paragraph.get("styleIDRef")
                para_pr_id = paragraph.get("paraPrIDRef")
                if style_id and style_ids and style_id not in style_ids:
                    errors.append(
                        _issue(
                            "style_reference",
                            f"paragraph references unknown styleIDRef={style_id}",
                            part_path=section.path,
                            section_index=section_index,
                            paragraph_index=paragraph_index,
                        )
                    )
                if para_pr_id and para_pr_ids and para_pr_id not in para_pr_ids:
                    errors.append(
                        _issue(
                            "style_reference",
                            f"paragraph references unknown paraPrIDRef={para_pr_id}",
                            part_path=section.path,
                            section_index=section_index,
                            paragraph_index=paragraph_index,
                        )
                    )
                for run in paragraph.xpath("./hp:run", namespaces=NS):
                    char_pr_id = run.get("charPrIDRef")
                    if char_pr_id and char_pr_ids and char_pr_id not in char_pr_ids:
                        errors.append(
                            _issue(
                                "style_reference",
                                f"run references unknown charPrIDRef={char_pr_id}",
                                part_path=section.path,
                                section_index=section_index,
                                paragraph_index=paragraph_index,
                            )
                        )
            for picture in self.pictures(section_index=section_index):
                if picture.binary_item_id and picture.binary_item_id not in manifest_ids:
                    errors.append(
                        _issue(
                            "binary_reference",
                            f"picture references unknown manifest id {picture.binary_item_id}",
                            part_path=section.path,
                            section_index=section_index,
                        )
                    )
            for shape in self.shapes(section_index=section_index):
                binary_refs = shape.element.xpath("./@binaryItemIDRef", namespaces=NS)
                for binary_ref in binary_refs:
                    if binary_ref not in manifest_ids:
                        errors.append(
                            _issue(
                                "binary_reference",
                                f"shape references unknown manifest id {binary_ref}",
                                part_path=section.path,
                                section_index=section_index,
                            )
                        )
            for binary_ref in section.readonly_root_element.xpath(".//hp:ole/@binaryItemIDRef", namespaces=NS):
                if binary_ref not in manifest_ids:
                    errors.append(
                        _issue(
                            "binary_reference",
                            f"ole references unknown manifest id {binary_ref}",
                            part_path=section.path,
                            section_index=section_index,
                        )
                    )
            for chart in self.charts(section_index=section_index):
                chart_ref = chart.chart_part_path
                if not chart_ref:
                    errors.append(
                        _issue(
                            "chart_reference",
                            "chart is missing chartIDRef",
                            part_path=section.path,
                            section_index=section_index,
                        )
                    )
                    continue
                if _normalize_path(chart_ref) not in self._parts:
                    errors.append(
                        _issue(
                            "chart_reference",
                            f"chart references missing chart part: {chart_ref}",
                            part_path=section.path,
                            section_index=section_index,
                        )
                    )
            for field_end in section.readonly_root_element.xpath(".//hp:fieldEnd", namespaces=NS):
                begin_id = field_end.get("beginIDRef")
                field_id = field_end.get("fieldid")
                if not begin_id:
                    errors.append(
                        _issue(
                            "field_reference",
                            "fieldEnd is missing beginIDRef",
                            part_path=section.path,
                            section_index=section_index,
                            xpath=".//hp:fieldEnd",
                        )
                    )
                    continue
                if begin_id not in field_begin_ids:
                    errors.append(
                        _issue(
                            "field_reference",
                            f"fieldEnd references unknown fieldBegin id {begin_id}",
                            part_path=section.path,
                            section_index=section_index,
                            xpath=".//hp:fieldEnd",
                        )
                    )
                    continue
                expected_field_id = field_begin_ids[begin_id]
                if field_id and expected_field_id and field_id != expected_field_id:
                    errors.append(
                        _issue(
                            "field_reference",
                            f"fieldEnd fieldid={field_id} does not match fieldBegin fieldid={expected_field_id}",
                            part_path=section.path,
                            section_index=section_index,
                            xpath=".//hp:fieldEnd",
                        )
                    )
        for field in self.cross_references():
            target = field.get_parameter("BookmarkName") or field.get_parameter("Path") or field.name
            if target and target not in bookmark_names:
                errors.append(_issue("cross_reference", f"cross reference points to unknown bookmark: {target}"))
        return errors

    def save_reopen_validation_errors(self) -> list[ValidationIssue]:
        temp_dir = _new_temp_dir("jakal_hwpx_reopen_")
        try:
            temp_path = temp_dir / "validation.hwpx"
            self.save(temp_path)
            reopened = type(self).open(temp_path)
            return reopened.validation_errors()
        except Exception as exc:  # noqa: BLE001
            if isinstance(exc, HwpxValidationError):
                return exc.errors
            return [_issue("save_reopen_validation", str(exc))]
        finally:
            self._cleanup_temp_dir(temp_dir)

    def strict_lint_errors(self) -> list[ValidationIssue]:
        errors = list(self.reference_validation_errors())

        ordered_paths = self.list_part_paths()
        if ordered_paths and ordered_paths[0] != "mimetype":
            errors.append(
                _issue(
                    "mimetype_packaging",
                    "mimetype must be the first package entry",
                    part_path="mimetype",
                )
            )

        mimetype_info = self._zip_infos.get("mimetype")
        if mimetype_info is not None and mimetype_info.compress_type != zipfile.ZIP_STORED:
            errors.append(
                _issue(
                    "mimetype_packaging",
                    "mimetype must be stored without compression",
                    part_path="mimetype",
                )
            )

        if not self.is_distribution_protected:
            style_ids = {style.style_id for style in self.styles() if style.style_id is not None}
            para_pr_ids = {style.style_id for style in self.paragraph_styles() if style.style_id is not None}
            char_pr_ids = {style.style_id for style in self.character_styles() if style.style_id is not None}
            memo_shape_ids = {memo_shape.memo_shape_id for memo_shape in self.memo_shapes() if memo_shape.memo_shape_id is not None}
            memo_properties_nodes = self.header.readonly_root_element.xpath(".//hh:memoProperties", namespaces=NS)
            for memo_properties in memo_properties_nodes:
                memo_pr_nodes = memo_properties.xpath("./hh:memoPr", namespaces=NS)
                expected_item_count = memo_properties.get("itemCnt")
                if expected_item_count not in (None, "", str(len(memo_pr_nodes))):
                    errors.append(
                        _issue(
                            "memo_reference",
                            f"memoProperties itemCnt={expected_item_count} does not match memoPr count={len(memo_pr_nodes)}",
                            part_path="Contents/header.xml",
                        )
                    )
                seen_memo_shape_ids: set[str] = set()
                for memo_pr in memo_pr_nodes:
                    memo_shape_id = memo_pr.get("id")
                    if not memo_shape_id:
                        continue
                    if memo_shape_id in seen_memo_shape_ids:
                        errors.append(
                            _issue(
                                "memo_reference",
                                f"Duplicate hh:memoPr id: {memo_shape_id}",
                                part_path="Contents/header.xml",
                            )
                        )
                    seen_memo_shape_ids.add(memo_shape_id)
            for section_index, section in enumerate(self.sections):
                sec_pr_nodes = section.readonly_root_element.xpath("./hp:p[1]//hp:secPr[1]", namespaces=NS)
                if sec_pr_nodes:
                    memo_shape_id = sec_pr_nodes[0].get("memoShapeIDRef")
                    if memo_shape_id not in (None, "", "0") and memo_shape_id not in memo_shape_ids:
                        errors.append(
                            _issue(
                                "memo_reference",
                                f"memoShapeIDRef points to unknown hh:memoPr id {memo_shape_id}",
                                part_path=section.path,
                                section_index=section_index,
                            )
                        )
                paragraphs = section.readonly_root_element.xpath(".//hp:p", namespaces=NS)
                for paragraph_index, paragraph in enumerate(paragraphs):
                    if style_ids and not paragraph.get("styleIDRef"):
                        errors.append(
                            _issue(
                                "style_binding",
                                "paragraph is missing styleIDRef",
                                part_path=section.path,
                                section_index=section_index,
                                paragraph_index=paragraph_index,
                            )
                        )
                    if para_pr_ids and not paragraph.get("paraPrIDRef"):
                        errors.append(
                            _issue(
                                "style_binding",
                                "paragraph is missing paraPrIDRef",
                                part_path=section.path,
                                section_index=section_index,
                                paragraph_index=paragraph_index,
                            )
                        )
                    for run_index, run in enumerate(paragraph.xpath("./hp:run", namespaces=NS)):
                        if char_pr_ids and not run.get("charPrIDRef"):
                            errors.append(
                                _issue(
                                    "style_binding",
                                    "run is missing charPrIDRef",
                                    part_path=section.path,
                                    section_index=section_index,
                                    paragraph_index=paragraph_index,
                                    context=f"run[{run_index}]",
                                )
                            )
                    form_nodes = paragraph.xpath(" | ".join(f".//hp:{tag}" for tag in _FORM_TAGS), namespaces=NS)
                    for form_index, form in enumerate(form_nodes):
                        kind = etree.QName(form).localname
                        command = form.get("command", "")
                        if command.startswith(_FORM_METADATA_COMMAND_PREFIX):
                            payload_text = command[len(_FORM_METADATA_COMMAND_PREFIX) :]
                            try:
                                payload = json.loads(payload_text)
                            except json.JSONDecodeError as exc:
                                errors.append(
                                    _issue(
                                        "metadata_schema",
                                        f"{kind} command metadata must be valid JSON: {exc}",
                                        part_path=section.path,
                                        section_index=section_index,
                                        paragraph_index=paragraph_index,
                                        context=f"form[{form_index}]",
                                    )
                                )
                                payload = None
                            if payload is not None and not isinstance(payload, dict):
                                errors.append(
                                    _issue(
                                        "metadata_schema",
                                        f"{kind} command metadata must decode to a JSON object.",
                                        part_path=section.path,
                                        section_index=section_index,
                                        paragraph_index=paragraph_index,
                                        context=f"form[{form_index}]",
                                    )
                                )
                            elif isinstance(payload, dict):
                                if "placeholder" in payload and not isinstance(payload["placeholder"], str):
                                    errors.append(
                                        _issue(
                                            "metadata_schema",
                                            f"{kind} placeholder metadata must be a string.",
                                            part_path=section.path,
                                            section_index=section_index,
                                            paragraph_index=paragraph_index,
                                            context=f"form[{form_index}]",
                                        )
                                    )
                                if "items" in payload and not isinstance(payload["items"], list):
                                    errors.append(
                                        _issue(
                                            "metadata_schema",
                                            f"{kind} items metadata must be a list.",
                                            part_path=section.path,
                                            section_index=section_index,
                                            paragraph_index=paragraph_index,
                                            context=f"form[{form_index}]",
                                        )
                                    )
                                elif isinstance(payload.get("items"), list) and any(
                                    not isinstance(item, str) for item in payload["items"]
                                ):
                                    errors.append(
                                        _issue(
                                            "metadata_schema",
                                            f"{kind} items metadata entries must be strings.",
                                            part_path=section.path,
                                            section_index=section_index,
                                            paragraph_index=paragraph_index,
                                            context=f"form[{form_index}]",
                                        )
                                    )
                                if "label" in payload and not isinstance(payload["label"], str):
                                    errors.append(
                                        _issue(
                                            "metadata_schema",
                                            f"{kind} label metadata must be a string.",
                                            part_path=section.path,
                                            section_index=section_index,
                                            paragraph_index=paragraph_index,
                                            context=f"form[{form_index}]",
                                        )
                                    )
                        for required_xpath, label in (
                            ("./hp:sz", "hp:sz"),
                            ("./hp:pos", "hp:pos"),
                            ("./hp:outMargin", "hp:outMargin"),
                        ):
                            if not form.xpath(required_xpath, namespaces=NS):
                                errors.append(
                                    _issue(
                                        "control_subtree",
                                        f"{kind} is missing {label}",
                                        part_path=section.path,
                                        section_index=section_index,
                                        paragraph_index=paragraph_index,
                                        context=f"form[{form_index}]",
                                    )
                                )
                        if form.xpath("./hp:caption", namespaces=NS) and not form.xpath("./hp:caption/hp:subList", namespaces=NS):
                            errors.append(
                                _issue(
                                    "control_subtree",
                                    f"{kind} caption is missing hp:subList",
                                    part_path=section.path,
                                    section_index=section_index,
                                    paragraph_index=paragraph_index,
                                    context=f"form[{form_index}]",
                                )
                            )
                        if kind != "scrollBar" and not form.xpath("./hp:formCharPr", namespaces=NS):
                            errors.append(
                                _issue(
                                    "control_subtree",
                                    f"{kind} is missing hp:formCharPr",
                                    part_path=section.path,
                                    section_index=section_index,
                                    paragraph_index=paragraph_index,
                                    context=f"form[{form_index}]",
                                )
                            )
                        if kind == "edit" and not form.xpath("./hp:text", namespaces=NS):
                            errors.append(
                                _issue(
                                    "control_subtree",
                                    "edit is missing hp:text",
                                    part_path=section.path,
                                    section_index=section_index,
                                    paragraph_index=paragraph_index,
                                    context=f"form[{form_index}]",
                                )
                            )
                    memo_nodes = paragraph.xpath(".//hp:hiddenComment", namespaces=NS)
                    for memo_index, memo in enumerate(memo_nodes):
                        if not memo.xpath("./hp:subList", namespaces=NS):
                            errors.append(
                                _issue(
                                    "control_subtree",
                                    "hiddenComment is missing hp:subList",
                                    part_path=section.path,
                                    section_index=section_index,
                                    paragraph_index=paragraph_index,
                                    context=f"memo[{memo_index}]",
                                )
                            )
                    chart_nodes = paragraph.xpath(".//hp:chart", namespaces=NS)
                    for chart_index, chart in enumerate(chart_nodes):
                        for required_xpath, label in (
                            ("./hp:sz", "hp:sz"),
                            ("./hp:pos", "hp:pos"),
                            ("./hp:outMargin", "hp:outMargin"),
                        ):
                            if not chart.xpath(required_xpath, namespaces=NS):
                                errors.append(
                                    _issue(
                                        "control_subtree",
                                        f"chart is missing {label}",
                                        part_path=section.path,
                                        section_index=section_index,
                                        paragraph_index=paragraph_index,
                                        context=f"chart[{chart_index}]",
                                    )
                                )
                        case = chart.getparent()
                        switch = case.getparent() if case is not None else None
                        if (
                            case is None
                            or etree.QName(case).localname != "case"
                            or switch is None
                            or etree.QName(switch).localname != "switch"
                        ):
                            errors.append(
                                _issue(
                                    "chart_closure",
                                    "native hp:chart control must be wrapped by hp:switch/hp:case",
                                    part_path=section.path,
                                    section_index=section_index,
                                    paragraph_index=paragraph_index,
                                    context=f"chart[{chart_index}]",
                                )
                            )
                        elif not switch.xpath("./hp:default/hp:ole", namespaces=NS):
                            errors.append(
                                _issue(
                                    "chart_closure",
                                    "native hp:chart control is missing hp:default/hp:ole fallback",
                                    part_path=section.path,
                                    section_index=section_index,
                                    paragraph_index=paragraph_index,
                                    context=f"chart[{chart_index}]",
                                )
                            )
                        elif not switch.xpath("./hp:default/hp:ole/hp:rotationInfo", namespaces=NS):
                            errors.append(
                                _issue(
                                    "chart_closure",
                                    "native hp:chart fallback hp:ole is missing hp:rotationInfo",
                                    part_path=section.path,
                                    section_index=section_index,
                                    paragraph_index=paragraph_index,
                                    context=f"chart[{chart_index}]",
                                )
                            )

        binary_part_paths = {
            path
            for path, part in self._parts.items()
            if path.startswith("BinData/") and isinstance(part, BinaryDataPart)
        }
        manifest_binary_items: dict[str, etree._Element] = {}
        manifest_binary_hrefs: set[str] = set()
        for item in self.content_hpf.manifest_items():
            item_id = item.get("id")
            href = item.get("href")
            if not item_id or not href:
                continue
            normalized_href = _normalize_path(href)
            if not normalized_href.startswith("BinData/"):
                continue
            manifest_binary_items[item_id] = item
            manifest_binary_hrefs.add(normalized_href)
            expected_media_type = _expected_binary_media_type(normalized_href)
            actual_media_type = item.get("media-type")
            if expected_media_type and actual_media_type and actual_media_type.lower() != expected_media_type.lower():
                errors.append(
                    _issue(
                        "binary_media_type",
                        f"manifest media-type {actual_media_type} does not match expected {expected_media_type} for {normalized_href}",
                        part_path="Contents/content.hpf",
                        context=f"manifest id={item_id}",
                    )
                )

        for part_path in sorted(binary_part_paths - manifest_binary_hrefs):
            errors.append(
                _issue(
                    "binary_closure",
                    f"BinData part is not referenced by content.hpf manifest: {part_path}",
                    part_path=part_path,
                )
            )

        referenced_binary_item_ids: set[str] = set()
        for section in self.sections:
            values = section.readonly_root_element.xpath(
                ".//hc:img/@binaryItemIDRef | " + _graphic_attribute_xpath("binaryItemIDRef"),
                namespaces=NS,
            )
            referenced_binary_item_ids.update(str(value) for value in values if str(value).strip())

        for item_id, item in manifest_binary_items.items():
            href = _normalize_path(item.get("href") or "")
            if item_id not in referenced_binary_item_ids:
                errors.append(
                    _issue(
                        "binary_closure",
                        f"manifest binary item is not referenced by any control: {item_id}",
                        part_path="Contents/content.hpf",
                        context=f"href={href}",
                    )
                )

        for item_id in sorted(referenced_binary_item_ids):
            item = manifest_binary_items.get(item_id)
            if item is None:
                continue
            href = _normalize_path(item.get("href") or "")
            if not href.startswith("BinData/"):
                errors.append(
                    _issue(
                        "binary_closure",
                        f"binaryItemIDRef {item_id} must resolve to a BinData part, got {href}",
                        part_path="Contents/content.hpf",
                        context=f"manifest id={item_id}",
                    )
                )

        referenced_chart_parts = {
            _normalize_path(chart.chart_part_path)
            for chart in self.charts()
            if chart.chart_part_path
        }
        chart_part_paths = {
            path
            for path, part in self._parts.items()
            if path.startswith("Chart/") and path.endswith(".xml")
        }
        manifest_hrefs = {
            _normalize_path(item.get("href") or "")
            for item in self.content_hpf.manifest_items()
            if item.get("href")
        }
        for chart_path in sorted(chart_part_paths - referenced_chart_parts):
            errors.append(
                _issue(
                    "chart_closure",
                    f"Chart part is not referenced by any native hp:chart control: {chart_path}",
                    part_path=chart_path,
                )
            )
        for chart_path in sorted(referenced_chart_parts):
            part = self._parts.get(chart_path)
            if part is None:
                continue
            if not isinstance(part, XmlPart):
                errors.append(
                    _issue(
                        "chart_closure",
                        f"Chart part must be XML: {chart_path}",
                        part_path=chart_path,
                    )
                )
            if chart_path not in manifest_hrefs:
                errors.append(
                    _issue(
                        "chart_closure",
                        f"Chart part is not declared in content.hpf manifest: {chart_path}",
                        part_path=chart_path,
                    )
                )

        return errors

    def strict_validate(self) -> None:
        errors = self.strict_lint_errors()
        if errors:
            raise HwpxValidationError(errors)

    def strict_lint_report(self, *, include_none: bool = False) -> list[dict[str, object]]:
        return [issue.to_dict(include_none=include_none) for issue in self.strict_lint_errors()]

    def format_strict_lint_errors(self) -> str:
        issues = self.strict_lint_errors()
        if not issues:
            return ""
        lines: list[str] = []
        for index, issue in enumerate(issues, start=1):
            location = issue.location() or "document"
            lines.append(f"{index}. [{issue.code or issue.kind}] {location}: {issue.message}")
            if issue.hint:
                lines.append(f"   Hint: {issue.hint}")
        return "\n".join(lines)

