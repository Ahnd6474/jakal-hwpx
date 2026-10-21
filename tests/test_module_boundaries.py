from __future__ import annotations

import pickle

from jakal_hwpx import (
    _hancom_models,
    _hancom_styles,
    _hancom_from_hwp,
    _hancom_from_hwpx,
    _hancom_to_hwp,
    _hancom_to_hwpx,
    _hwp_objects,
    _hwp_graphic_payloads,
    _hwp_object_payloads,
    _hwp_page_payloads,
    _hwp_payload_codecs,
    _hwp_record_tree,
    _hwp_shape_objects,
    _hwp_table_objects,
    _hwpx2py_authoring,
    _hwpx2py_macro,
    _hwpx2py_semantic,
    _hwpx_chart_xml,
    _hwpx_common,
    _hwpx_document_controls,
    _hwpx_document_fields,
    _hwpx_document_graphics,
    _hwpx_element_helpers,
    _hwpx_factory,
    _hwpx_validation,
    document,
    elements,
    hancom_document,
    hwp_binary,
    hwp_document,
    hwpx2py,
)


def test_public_facades_reexport_extracted_implementations() -> None:
    assert hancom_document.Paragraph is _hancom_models.Paragraph
    assert hancom_document.Paragraph.__module__ == "jakal_hwpx.hancom_document"
    assert hancom_document._build_hwp_paragraph_style_payload is _hancom_styles._build_hwp_paragraph_style_payload
    assert hwp_binary._parse_character_shape_payload is _hwp_payload_codecs._parse_character_shape_payload
    assert hwp_binary._build_shape_specific_payload is _hwp_graphic_payloads._build_shape_specific_payload
    assert hwp_binary._parse_page_def_payload is _hwp_page_payloads._parse_page_def_payload
    assert hwp_document._parse_shape_fill_payload is _hwp_object_payloads._parse_shape_fill_payload
    assert hwp_document.HwpSection is _hwp_objects.HwpSection
    assert hwp_document.HwpSection.__module__ == "jakal_hwpx.hwp_document"
    assert hwp_document.HwpTableObject is _hwp_table_objects.HwpTableObject
    assert hwp_document.HwpShapeObject is _hwp_shape_objects.HwpShapeObject
    assert hwp_binary.DocInfoModel is _hwp_record_tree.DocInfoModel
    assert hancom_document._extract_hwp_control is _hancom_from_hwp._extract_hwp_control
    assert hancom_document._extract_hwpx_paragraph_blocks is _hancom_from_hwpx._extract_hwpx_paragraph_blocks
    assert hancom_document._append_hancom_block_to_hwp is _hancom_to_hwp._append_hancom_block_to_hwp
    assert hancom_document._write_section_to_hwpx is _hancom_to_hwpx._write_section_to_hwpx
    assert document._build_blank_document_state is _hwpx_factory._build_blank_document_state
    assert document._normalize_path is _hwpx_common._normalize_path
    assert document._canonical_append_block_type is _hwpx_document_controls._canonical_append_block_type
    assert document.HwpxDocument.append_table is _hwpx_document_controls.HwpxDocumentControlMixin.append_table
    assert document.HwpxDocument.append_table is _hwpx_document_graphics.HwpxDocumentGraphicMixin.append_table
    assert document.HwpxDocument.append_field is _hwpx_document_fields.HwpxDocumentFieldMixin.append_field
    assert document.HwpxDocument.validation_errors is _hwpx_validation.HwpxValidationMixin.validation_errors
    assert issubclass(document.HwpxDocument, _hwpx_validation.HwpxValidationMixin)
    assert elements._build_chart_part_root is _hwpx_chart_xml._build_chart_part_root
    assert elements._set_graphic_layout is _hwpx_element_helpers._set_graphic_layout
    assert hwpx2py._emit_block is _hwpx2py_semantic._emit_block
    assert hwpx2py._macro_expr_for_block is _hwpx2py_macro._macro_expr_for_block
    assert hwpx2py._emit_authoring_block is _hwpx2py_authoring._emit_authoring_block

    paragraph = hancom_document.Paragraph("stable")
    assert pickle.loads(pickle.dumps(paragraph)) == paragraph


def test_extracted_factories_still_build_public_document_models() -> None:
    parts, part_order, zip_infos, duplicates = _hwpx_factory._build_blank_document_state()

    assert part_order[0] == "mimetype"
    assert set(part_order) == set(parts)
    assert zip_infos == {}
    assert duplicates == []


def test_hwpx_compatibility_aliases_are_public() -> None:
    import jakal_hwpx

    assert jakal_hwpx.HeaderFooterBlock is elements.HeaderFooterXml
    assert jakal_hwpx.ShapeObject is elements.ShapeXml
    assert jakal_hwpx.TableCell is elements.TableCellXml
