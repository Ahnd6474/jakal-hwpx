from __future__ import annotations

import mimetypes
import os
from dataclasses import dataclass, field
from pathlib import PurePosixPath
from typing import Mapping, Sequence

from lxml import etree

from .document import HwpxDocument
from .namespaces import NS, qname
from .pdf import PdfDocument, PdfMetadata, _wrap_text_for_page


_HWP_UNITS_PER_POINT = 100.0
_DEFAULT_MARKER_PREFIX = "[NON_TEXT]"
_DEFAULT_PAGE_MARGIN = 36.0


@dataclass
class BridgeTextBlock:
    text: str
    left: float
    top: float
    width: float = 0.0
    height: float = 0.0


@dataclass
class BridgePageFeatures:
    image_count: int = 0
    vector_operation_count: int = 0
    annotation_count: int = 0
    table_count: int = 0
    shape_count: int = 0
    equation_count: int = 0
    notes: list[str] = field(default_factory=list)

    def has_non_text(self) -> bool:
        return any(
            (
                self.image_count,
                self.vector_operation_count,
                self.annotation_count,
                self.table_count,
                self.shape_count,
                self.equation_count,
                self.notes,
            )
        )

    def to_marker_text(self, prefix: str = _DEFAULT_MARKER_PREFIX) -> str:
        parts = [
            f"images={self.image_count}",
            f"vector_ops={self.vector_operation_count}",
            f"annotations={self.annotation_count}",
        ]
        if self.table_count:
            parts.append(f"tables={self.table_count}")
        if self.shape_count:
            parts.append(f"shapes={self.shape_count}")
        if self.equation_count:
            parts.append(f"equations={self.equation_count}")
        parts.extend(self.notes)
        return f"{prefix} " + " ".join(parts)


@dataclass
class BridgePage:
    width_points: float
    height_points: float
    text: str = ""
    rotation: int = 0
    text_blocks: list[BridgeTextBlock] = field(default_factory=list)
    features: BridgePageFeatures = field(default_factory=BridgePageFeatures)


@dataclass
class DocumentBridge:
    source_format: str
    pages: list[BridgePage]
    metadata: PdfMetadata = field(default_factory=PdfMetadata)


def _coerce_ocr_text(ocr_text_by_page: Sequence[str] | Mapping[int, str] | None, index: int) -> str:
    if ocr_text_by_page is None:
        return ""
    if isinstance(ocr_text_by_page, Mapping):
        return ocr_text_by_page.get(index, "")
    if index < len(ocr_text_by_page):
        return ocr_text_by_page[index]
    return ""


def _coerce_block(value: BridgeTextBlock | Mapping[str, object]) -> BridgeTextBlock:
    if isinstance(value, BridgeTextBlock):
        return value
    return BridgeTextBlock(
        text=str(value.get("text", "")),
        left=float(value.get("left", 0.0)),
        top=float(value.get("top", 0.0)),
        width=float(value.get("width", 0.0)),
        height=float(value.get("height", 0.0)),
    )


def _coerce_ocr_blocks(
    ocr_blocks_by_page: Sequence[Sequence[BridgeTextBlock | Mapping[str, object]]] | Mapping[int, Sequence[BridgeTextBlock | Mapping[str, object]]] | None,
    index: int,
) -> list[BridgeTextBlock]:
    if ocr_blocks_by_page is None:
        return []
    raw_page = ocr_blocks_by_page.get(index, []) if isinstance(ocr_blocks_by_page, Mapping) else (
        ocr_blocks_by_page[index] if index < len(ocr_blocks_by_page) else []
    )
    return [_coerce_block(block) for block in raw_page]


def _ordered_blocks(blocks: Sequence[BridgeTextBlock]) -> list[BridgeTextBlock]:
    return sorted(blocks, key=lambda item: (round(item.top, 3), round(item.left, 3)))


def _point_to_hwp_units(value: float) -> int:
    return max(int(round(value * _HWP_UNITS_PER_POINT)), 1)


def _hwp_units_to_points(value: int | None) -> float:
    if value is None:
        return 0.0
    return float(value) / _HWP_UNITS_PER_POINT


def _bridge_preview_text(bridge: DocumentBridge) -> str:
    chunks = [page.text.strip() for page in bridge.pages if page.text.strip()]
    return "\n\n".join(chunks)[:4096]


def _block_text(page: BridgePage) -> str:
    if page.text_blocks:
        return "\n".join(block.text for block in _ordered_blocks(page.text_blocks) if block.text.strip())
    return page.text


def _set_section_text(document: HwpxDocument, section_index: int, page: BridgePage, marker: str | None = None) -> None:
    lines: list[str]
    if page.text_blocks:
        lines = [block.text for block in _ordered_blocks(page.text_blocks) if block.text]
    else:
        lines = page.text.splitlines() if page.text else []
    if not lines:
        lines = [""]
    if marker:
        lines.append(marker)

    document.set_paragraph_text(section_index, 0, lines[0])
    for line in lines[1:]:
        document.append_paragraph(line, section_index=section_index)


def _ensure_pdf_document(value: PdfDocument | str | os.PathLike[str]) -> PdfDocument:
    if isinstance(value, PdfDocument):
        return value
    return PdfDocument.open(value)


def _ensure_hwpx_document(value: HwpxDocument | str | os.PathLike[str]) -> HwpxDocument:
    if isinstance(value, HwpxDocument):
        return value
    return HwpxDocument.open(value)


def _append_picture_to_hwpx(
    document: HwpxDocument,
    *,
    section_index: int,
    image_bytes: bytes,
    image_name: str,
    width_points: float,
    height_points: float,
    comment: str,
) -> None:
    if not image_bytes:
        return

    manifest_id = document.content_hpf.next_manifest_id("image")
    unique_name = f"page{section_index + 1}_{manifest_id}_{PurePosixPath(image_name).name}"
    media_type = mimetypes.guess_type(unique_name)[0] or "application/octet-stream"
    document.add_or_replace_binary(unique_name, image_bytes, media_type=media_type, manifest_id=manifest_id)

    paragraph = document.append_paragraph("", section_index=section_index).element
    run = paragraph.xpath("./hp:run[1]", namespaces=NS)[0]
    width_units = _point_to_hwp_units(max(width_points, 1.0))
    height_units = _point_to_hwp_units(max(height_points, 1.0))

    picture = etree.Element(qname("hp", "pic"))
    picture.set("id", document._next_control_number(".//@id"))
    picture.set("zOrder", "1")
    picture.set("numberingType", "PICTURE")
    picture.set("textWrap", "TOP_AND_BOTTOM")
    picture.set("textFlow", "BOTH_SIDES")
    picture.set("lock", "0")
    picture.set("dropcapstyle", "None")
    picture.set("href", "")
    picture.set("groupLevel", "0")
    picture.set("instid", document._next_control_number(".//@instid"))
    picture.set("reverse", "0")

    offset = etree.SubElement(picture, qname("hp", "offset"))
    offset.set("x", "0")
    offset.set("y", "0")
    for tag in ("orgSz", "curSz"):
        size = etree.SubElement(picture, qname("hp", tag))
        size.set("width", str(width_units))
        size.set("height", str(height_units))
    flip = etree.SubElement(picture, qname("hp", "flip"))
    flip.set("horizontal", "0")
    flip.set("vertical", "0")
    rotation = etree.SubElement(picture, qname("hp", "rotationInfo"))
    rotation.set("angle", "0")
    rotation.set("centerX", str(width_units // 2))
    rotation.set("centerY", str(height_units // 2))
    rotation.set("rotateimage", "1")
    rendering = etree.SubElement(picture, qname("hp", "renderingInfo"))
    for matrix_tag in ("transMatrix", "scaMatrix", "rotMatrix"):
        matrix = etree.SubElement(rendering, qname("hc", matrix_tag))
        matrix.set("e1", "1")
        matrix.set("e2", "0")
        matrix.set("e3", "0")
        matrix.set("e4", "0")
        matrix.set("e5", "1")
        matrix.set("e6", "0")
    image_node = etree.SubElement(picture, qname("hc", "img"))
    image_node.set("binaryItemIDRef", manifest_id)
    image_node.set("bright", "0")
    image_node.set("contrast", "0")
    image_node.set("effect", "REAL_PIC")
    image_node.set("alpha", "0")
    img_rect = etree.SubElement(picture, qname("hp", "imgRect"))
    for point_name, x_value, y_value in (
        ("pt0", 0, 0),
        ("pt1", width_units, 0),
        ("pt2", width_units, height_units),
        ("pt3", 0, height_units),
    ):
        point = etree.SubElement(img_rect, qname("hc", point_name))
        point.set("x", str(x_value))
        point.set("y", str(y_value))
    clip = etree.SubElement(picture, qname("hp", "imgClip"))
    clip.set("left", "0")
    clip.set("right", str(width_units))
    clip.set("top", "0")
    clip.set("bottom", str(height_units))
    for tag in ("inMargin", "outMargin"):
        margin = etree.SubElement(picture, qname("hp", tag))
        for key in ("left", "right", "top", "bottom"):
            margin.set(key, "0")
    dim = etree.SubElement(picture, qname("hp", "imgDim"))
    dim.set("dimwidth", str(width_units))
    dim.set("dimheight", str(height_units))
    etree.SubElement(picture, qname("hp", "effects"))
    size = etree.SubElement(picture, qname("hp", "sz"))
    size.set("width", str(width_units))
    size.set("widthRelTo", "ABSOLUTE")
    size.set("height", str(height_units))
    size.set("heightRelTo", "ABSOLUTE")
    size.set("protect", "0")
    pos = etree.SubElement(picture, qname("hp", "pos"))
    pos.set("treatAsChar", "1")
    pos.set("affectLSpacing", "0")
    pos.set("flowWithText", "1")
    pos.set("allowOverlap", "0")
    pos.set("holdAnchorAndSO", "0")
    pos.set("vertRelTo", "PARA")
    pos.set("horzRelTo", "COLUMN")
    pos.set("vertAlign", "TOP")
    pos.set("horzAlign", "LEFT")
    pos.set("vertOffset", "0")
    pos.set("horzOffset", "0")
    shape_comment = etree.SubElement(picture, qname("hp", "shapeComment"))
    shape_comment.text = comment

    text_node = run.xpath("./hp:t", namespaces=NS)
    if text_node:
        run.insert(run.index(text_node[0]), picture)
        text_node[0].text = ""
    else:
        run.append(picture)
        etree.SubElement(run, qname("hp", "t"))
    document.sections[section_index].mark_modified()


def _render_picture_to_pdf(pdf_page, data: bytes, width_points: float, height_points: float, x: float, y: float, fallback_text: str) -> None:
    try:
        pdf_page.add_image(data, x=x, y=y, width=width_points, height=height_points)
    except Exception:  # noqa: BLE001
        pdf_page.draw_rectangle(x, y, width_points, height_points)
        pdf_page.add_text(fallback_text, x=x + 4, y=y + height_points - 14, font_size=10)


def _render_table_to_pdf(pdf_page, table, *, page_width: float, top_y: float) -> float:
    col_count = max(table.column_count, 1)
    row_count = max(table.row_count, 1)
    col_widths = [72.0] * col_count
    row_heights = [24.0] * row_count

    for cell in table.cells():
        width_values = cell.element.xpath("./hp:cellSz/@width", namespaces=NS)
        height_values = cell.element.xpath("./hp:cellSz/@height", namespaces=NS)
        width = _hwp_units_to_points(int(width_values[0])) if width_values else 72.0
        height = _hwp_units_to_points(int(height_values[0])) if height_values else 24.0
        span_width = max(cell.col_span, 1)
        span_height = max(cell.row_span, 1)
        per_col = width / span_width
        per_row = height / span_height
        for index in range(cell.column, min(cell.column + span_width, col_count)):
            col_widths[index] = max(col_widths[index], per_col)
        for index in range(cell.row, min(cell.row + span_height, row_count)):
            row_heights[index] = max(row_heights[index], per_row)

    usable_width = max(page_width - (_DEFAULT_PAGE_MARGIN * 2), 1.0)
    total_width = sum(col_widths)
    scale = min(1.0, usable_width / total_width) if total_width else 1.0
    col_widths = [value * scale for value in col_widths]
    row_heights = [value * scale for value in row_heights]
    total_height = sum(row_heights)
    left = _DEFAULT_PAGE_MARGIN
    bottom = top_y - total_height

    for cell in table.cells():
        x = left + sum(col_widths[: cell.column])
        y_top = top_y - sum(row_heights[: cell.row])
        width = sum(col_widths[cell.column : cell.column + max(cell.col_span, 1)])
        height = sum(row_heights[cell.row : cell.row + max(cell.row_span, 1)])
        y = y_top - height
        pdf_page.draw_rectangle(x, y, width, height, line_width=0.8)
        if cell.text.strip():
            lines = _wrap_text_for_page(cell.text.strip(), usable_width=max(width - 8.0, 1.0), font_size=9.0)
            pdf_page.add_text("\n".join(lines), x=x + 4, y=y + height - 12, font_size=9.0, leading=11.0)

    return bottom - 18.0


def _shape_size_points(shape) -> tuple[float, float]:
    width_values = shape.element.xpath("./hp:sz/@width", namespaces=NS)
    height_values = shape.element.xpath("./hp:sz/@height", namespaces=NS)
    if width_values and height_values:
        return _hwp_units_to_points(int(width_values[0])), _hwp_units_to_points(int(height_values[0]))
    pt0 = shape.element.xpath("./hc:pt0", namespaces=NS)
    pt1 = shape.element.xpath("./hc:pt1", namespaces=NS)
    if pt0 and pt1:
        width = abs(int(pt1[0].get("x", "0")) - int(pt0[0].get("x", "0")))
        height = abs(int(pt1[0].get("y", "0")) - int(pt0[0].get("y", "0")))
        return _hwp_units_to_points(width), _hwp_units_to_points(height)
    return 120.0, 36.0


def _render_shape_to_pdf(pdf_page, shape, *, page_width: float, top_y: float) -> float:
    width, height = _shape_size_points(shape)
    width = min(width, page_width - (_DEFAULT_PAGE_MARGIN * 2))
    height = max(height, 18.0)
    x = _DEFAULT_PAGE_MARGIN
    y = top_y - height

    if shape.kind in {"rect", "textart"}:
        pdf_page.draw_rectangle(x, y, width, height, line_width=1.0)
        if shape.text.strip():
            lines = _wrap_text_for_page(shape.text.strip(), usable_width=max(width - 8.0, 1.0), font_size=10.0)
            pdf_page.add_text("\n".join(lines), x=x + 4, y=y + height - 14, font_size=10.0)
        return y - 18.0

    if shape.kind == "line":
        pt0 = shape.element.xpath("./hc:pt0", namespaces=NS)
        pt1 = shape.element.xpath("./hc:pt1", namespaces=NS)
        if pt0 and pt1:
            x1 = x
            y1 = top_y
            x2 = x + _hwp_units_to_points(int(pt1[0].get("x", "0")) - int(pt0[0].get("x", "0")))
            y2 = top_y - _hwp_units_to_points(int(pt1[0].get("y", "0")) - int(pt0[0].get("y", "0")))
        else:
            x1, y1, x2, y2 = x, top_y, x + width, y
        pdf_page.draw_line(x1, y1, x2, y2, line_width=1.0)
        return min(y1, y2) - 18.0

    pdf_page.draw_rectangle(x, y, width, height, line_width=1.0)
    if shape.text.strip():
        pdf_page.add_text(shape.text.strip(), x=x + 4, y=y + height - 14, font_size=10.0)
    return y - 18.0


def _estimate_text_bottom(page: BridgePage, *, font_size: float) -> float:
    block_text = _block_text(page).strip()
    if not block_text:
        return page.height_points - _DEFAULT_PAGE_MARGIN
    lines = _wrap_text_for_page(
        block_text,
        usable_width=max(page.width_points - (_DEFAULT_PAGE_MARGIN * 2), 1.0),
        font_size=font_size,
    )
    text_height = len(lines) * font_size * 1.4
    return page.height_points - _DEFAULT_PAGE_MARGIN - text_height - 18.0


def _render_hwpx_non_text_features_to_pdf(
    hwpx_document: HwpxDocument,
    pdf_document: PdfDocument,
    bridge: DocumentBridge,
    *,
    font_size: float,
) -> None:
    for index, _section in enumerate(hwpx_document.sections):
        pdf_page = pdf_document.pages[index]
        cursor_y = _estimate_text_bottom(bridge.pages[index], font_size=font_size)

        for table in hwpx_document.tables(section_index=index):
            cursor_y = _render_table_to_pdf(pdf_page, table, page_width=bridge.pages[index].width_points, top_y=cursor_y)

        for picture in hwpx_document.pictures(section_index=index):
            size_values = picture.element.xpath("./hp:sz/@width | ./hp:sz/@height", namespaces=NS)
            if len(size_values) >= 2:
                width_points = _hwp_units_to_points(int(size_values[0]))
                height_points = _hwp_units_to_points(int(size_values[1]))
            else:
                width_points, height_points = 120.0, 120.0
            y = cursor_y - height_points
            fallback = picture.shape_comment.strip() or "embedded image"
            _render_picture_to_pdf(
                pdf_page,
                picture.binary_data(),
                width_points=width_points,
                height_points=height_points,
                x=_DEFAULT_PAGE_MARGIN,
                y=y,
                fallback_text=fallback,
            )
            cursor_y = y - 18.0

        for shape in hwpx_document.shapes(section_index=index):
            cursor_y = _render_shape_to_pdf(pdf_page, shape, page_width=bridge.pages[index].width_points, top_y=cursor_y)


def pdf_to_bridge(
    pdf: PdfDocument | str | os.PathLike[str],
    *,
    ocr_text_by_page: Sequence[str] | Mapping[int, str] | None = None,
    ocr_blocks_by_page: Sequence[Sequence[BridgeTextBlock | Mapping[str, object]]] | Mapping[int, Sequence[BridgeTextBlock | Mapping[str, object]]] | None = None,
) -> DocumentBridge:
    pdf_document = _ensure_pdf_document(pdf)
    pages: list[BridgePage] = []

    for index, page in enumerate(pdf_document.pages):
        analysis = page.analyze_non_text()
        blocks = _coerce_ocr_blocks(ocr_blocks_by_page, index)
        ocr_text = _coerce_ocr_text(ocr_text_by_page, index)
        notes = []
        if analysis.xobject_names:
            notes.append("xobjects=" + ",".join(analysis.xobject_names))
        pages.append(
            BridgePage(
                width_points=page.width,
                height_points=page.height,
                rotation=page.rotation,
                text=ocr_text or "\n".join(block.text for block in _ordered_blocks(blocks)),
                text_blocks=blocks,
                features=BridgePageFeatures(
                    image_count=analysis.image_count,
                    vector_operation_count=analysis.vector_summary.path_operator_count
                    + analysis.vector_summary.paint_operator_count,
                    annotation_count=analysis.annotation_count,
                    notes=notes,
                ),
            )
        )

    return DocumentBridge(source_format="pdf", pages=pages, metadata=pdf_document.metadata())


def hwpx_to_bridge(hwpx: HwpxDocument | str | os.PathLike[str]) -> DocumentBridge:
    hwpx_document = _ensure_hwpx_document(hwpx)
    metadata = hwpx_document.metadata()
    pages: list[BridgePage] = []

    for index, section in enumerate(hwpx_document.sections):
        settings = hwpx_document.section_settings(index)
        pages.append(
            BridgePage(
                width_points=_hwp_units_to_points(settings.page_width),
                height_points=_hwp_units_to_points(settings.page_height),
                text=section.extract_text(paragraph_separator="\n").strip(),
                features=BridgePageFeatures(
                    image_count=len(hwpx_document.pictures(section_index=index)),
                    table_count=len(hwpx_document.tables(section_index=index)),
                    shape_count=len(hwpx_document.shapes(section_index=index)),
                    equation_count=len(hwpx_document.equations(section_index=index)),
                ),
            )
        )

    return DocumentBridge(
        source_format="hwpx",
        pages=pages,
        metadata=PdfMetadata(
            title=metadata.title,
            author=metadata.creator,
            subject=metadata.subject,
            creator=metadata.creator,
            keywords=metadata.keyword,
        ),
    )


def bridge_to_hwpx(
    bridge: DocumentBridge,
    *,
    include_non_text_markers: bool = True,
) -> HwpxDocument:
    document = HwpxDocument.blank()
    page_count = max(len(bridge.pages), 1)
    while len(document.sections) < page_count:
        document.add_section(text="")

    document.set_metadata(
        title=bridge.metadata.title,
        creator=bridge.metadata.author or bridge.metadata.creator,
        subject=bridge.metadata.subject,
        keyword=bridge.metadata.keywords,
    )

    for index, page in enumerate(bridge.pages or [BridgePage(width_points=595.0, height_points=842.0)]):
        settings = document.section_settings(index)
        if page.width_points > 0:
            settings.set_page_size(width=_point_to_hwp_units(page.width_points))
        if page.height_points > 0:
            settings.set_page_size(height=_point_to_hwp_units(page.height_points))
        marker = page.features.to_marker_text() if include_non_text_markers and page.features.has_non_text() else None
        _set_section_text(document, index, page, marker=marker)

    preview_text = _bridge_preview_text(bridge)
    if preview_text:
        document.set_preview_text(preview_text)
    return document


def bridge_to_pdf(
    bridge: DocumentBridge,
    *,
    include_non_text_markers: bool = True,
    font_size: float = 12.0,
) -> PdfDocument:
    document = PdfDocument.blank()
    document.set_metadata(
        title=bridge.metadata.title,
        author=bridge.metadata.author,
        subject=bridge.metadata.subject,
        creator=bridge.metadata.creator,
        keywords=bridge.metadata.keywords,
    )

    pages = bridge.pages or [BridgePage(width_points=595.0, height_points=842.0)]
    for page in pages:
        pdf_page = document.add_page(
            width=page.width_points or 595.0,
            height=page.height_points or 842.0,
        )
        if page.text_blocks:
            for block in _ordered_blocks(page.text_blocks):
                baseline_y = (page.height_points or 842.0) - block.top - max(block.height, font_size)
                pdf_page.add_text(block.text, x=block.left, y=baseline_y, font_size=font_size)
        else:
            content = page.text
            if include_non_text_markers and page.features.has_non_text():
                marker = page.features.to_marker_text()
                content = f"{content}\n\n{marker}".strip()
            if content:
                pdf_page.add_wrapped_text(content, font_size=font_size, margin_left=_DEFAULT_PAGE_MARGIN, margin_right=_DEFAULT_PAGE_MARGIN, margin_top=_DEFAULT_PAGE_MARGIN)
    return document


def pdf_to_hwpx(
    pdf: PdfDocument | str | os.PathLike[str],
    *,
    ocr_text_by_page: Sequence[str] | Mapping[int, str] | None = None,
    ocr_blocks_by_page: Sequence[Sequence[BridgeTextBlock | Mapping[str, object]]] | Mapping[int, Sequence[BridgeTextBlock | Mapping[str, object]]] | None = None,
    include_non_text_markers: bool = True,
) -> HwpxDocument:
    pdf_document = _ensure_pdf_document(pdf)
    bridge = pdf_to_bridge(pdf_document, ocr_text_by_page=ocr_text_by_page, ocr_blocks_by_page=ocr_blocks_by_page)
    document = bridge_to_hwpx(bridge, include_non_text_markers=include_non_text_markers)

    for index, page in enumerate(pdf_document.pages):
        for placement_index, placement in enumerate(page.image_placements()):
            comment = (
                f"Imported from PDF image {placement.name}\n"
                f"position=({placement.x:.2f},{placement.y:.2f}) size=({placement.width:.2f},{placement.height:.2f})"
            )
            _append_picture_to_hwpx(
                document,
                section_index=index,
                image_bytes=placement.data,
                image_name=f"{placement_index + 1}_{placement.name}",
                width_points=placement.width,
                height_points=placement.height,
                comment=comment,
            )
    return document


def hwpx_to_pdf(
    hwpx: HwpxDocument | str | os.PathLike[str],
    *,
    include_non_text_markers: bool = True,
    font_size: float = 12.0,
) -> PdfDocument:
    hwpx_document = _ensure_hwpx_document(hwpx)
    bridge = hwpx_to_bridge(hwpx_document)
    document = bridge_to_pdf(bridge, include_non_text_markers=include_non_text_markers, font_size=font_size)
    _render_hwpx_non_text_features_to_pdf(hwpx_document, document, bridge, font_size=font_size)
    return document
