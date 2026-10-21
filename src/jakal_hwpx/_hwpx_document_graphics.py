from __future__ import annotations

from pathlib import PurePosixPath

from lxml import etree

from ._hwpx_common import _graphic_attribute_xpath
from ._ole_payloads import normalize_hwpx_ole_payload
from .elements import OleXml, PictureXml, ShapeXml, TableXml, _set_graphic_layout, _set_margin_values
from .namespaces import qname


class HwpxDocumentGraphicMixin:
    def append_table(
        self,
        rows: int,
        columns: int,
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
        cell_texts: list[list[str]] | None = None,
        cell_width: int = 16000,
        row_height: int = 1800,
        table_width: int | None = None,
        treat_as_char: bool = True,
        text_wrap: str = "TOP_AND_BOTTOM",
        text_flow: str = "BOTH_SIDES",
        affect_line_spacing: bool = False,
        flow_with_text: bool = True,
        allow_overlap: bool = False,
        hold_anchor_and_so: bool = False,
        vert_rel_to: str = "PARA",
        horz_rel_to: str = "COLUMN",
        vert_align: str = "TOP",
        horz_align: str = "LEFT",
        vert_offset: int = 0,
        horz_offset: int = 0,
        out_margin_left: int = 0,
        out_margin_right: int = 0,
        out_margin_top: int = 0,
        out_margin_bottom: int = 0,
    ) -> TableXml:
        self._ensure_editable_sections()
        if rows < 1 or columns < 1:
            raise ValueError("rows and columns must be at least 1.")
        if cell_width < 1 or row_height < 1:
            raise ValueError("cell_width and row_height must be positive.")

        width = table_width if table_width is not None else cell_width * columns
        if width < 1:
            raise ValueError("table_width must be positive.")

        table = etree.Element(qname("hp", "tbl"))
        table.set("id", self._next_control_number(_graphic_attribute_xpath("id")))
        table.set("zOrder", self._next_control_number(_graphic_attribute_xpath("zOrder")))
        table.set("numberingType", "TABLE")
        table.set("textWrap", "TOP_AND_BOTTOM")
        table.set("textFlow", "BOTH_SIDES")
        table.set("lock", "0")
        table.set("dropcapstyle", "None")
        table.set("pageBreak", "CELL")
        table.set("repeatHeader", "1")
        table.set("rowCnt", str(rows))
        table.set("colCnt", str(columns))
        table.set("cellSpacing", "0")
        table.set("borderFillIDRef", "0")
        table.set("noAdjust", "0")

        size = etree.SubElement(table, qname("hp", "sz"))
        size.set("width", str(width))
        size.set("widthRelTo", "ABSOLUTE")
        size.set("height", str(row_height * rows))
        size.set("heightRelTo", "ABSOLUTE")
        size.set("protect", "0")

        position = etree.SubElement(table, qname("hp", "pos"))
        position.set("treatAsChar", "1" if treat_as_char else "0")
        position.set("affectLSpacing", "0")
        position.set("flowWithText", "1")
        position.set("allowOverlap", "0")
        position.set("holdAnchorAndSO", "0")
        position.set("vertRelTo", "PARA")
        position.set("horzRelTo", "COLUMN")
        position.set("vertAlign", "TOP")
        position.set("horzAlign", "LEFT")
        position.set("vertOffset", "0")
        position.set("horzOffset", "0")

        out_margin = etree.SubElement(table, qname("hp", "outMargin"))
        for key in ("left", "right", "top", "bottom"):
            out_margin.set(key, "0")
        _set_graphic_layout(
            table,
            text_wrap=text_wrap,
            text_flow=text_flow,
            treat_as_char=treat_as_char,
            affect_line_spacing=affect_line_spacing,
            flow_with_text=flow_with_text,
            allow_overlap=allow_overlap,
            hold_anchor_and_so=hold_anchor_and_so,
            vert_rel_to=vert_rel_to,
            horz_rel_to=horz_rel_to,
            vert_align=vert_align,
            horz_align=horz_align,
            vert_offset=vert_offset,
            horz_offset=horz_offset,
        )
        _set_margin_values(
            table,
            "./hp:outMargin",
            left=out_margin_left,
            right=out_margin_right,
            top=out_margin_top,
            bottom=out_margin_bottom,
        )

        in_margin = etree.SubElement(table, qname("hp", "inMargin"))
        in_margin.set("left", "141")
        in_margin.set("right", "141")
        in_margin.set("top", "141")
        in_margin.set("bottom", "141")

        nested_paragraph_id = self._next_control_number(".//hp:p/@id")
        for row_index in range(rows):
            row = etree.SubElement(table, qname("hp", "tr"))
            for column_index in range(columns):
                cell = etree.SubElement(row, qname("hp", "tc"))
                cell.set("name", "")
                cell.set("header", "0")
                cell.set("hasMargin", "0")
                cell.set("protect", "0")
                cell.set("editable", "0")
                cell.set("dirty", "0")
                cell.set("borderFillIDRef", "0")

                sub_list = etree.SubElement(cell, qname("hp", "subList"))
                sub_list.set("id", "")
                sub_list.set("textDirection", "HORIZONTAL")
                sub_list.set("lineWrap", "BREAK")
                sub_list.set("vertAlign", "CENTER")
                sub_list.set("linkListIDRef", "0")
                sub_list.set("linkListNextIDRef", "0")
                sub_list.set("textWidth", "0")
                sub_list.set("textHeight", "0")
                sub_list.set("hasTextRef", "0")
                sub_list.set("hasNumRef", "0")

                paragraph = etree.SubElement(sub_list, qname("hp", "p"))
                paragraph.set("id", nested_paragraph_id)
                nested_paragraph_id = str(int(nested_paragraph_id) + 1)
                paragraph.set("paraPrIDRef", "0")
                paragraph.set("styleIDRef", "0")
                paragraph.set("pageBreak", "0")
                paragraph.set("columnBreak", "0")
                paragraph.set("merged", "0")

                run = etree.SubElement(paragraph, qname("hp", "run"))
                run.set("charPrIDRef", char_pr_id or "0")
                text_node = etree.SubElement(run, qname("hp", "t"))
                value = ""
                if cell_texts is not None and row_index < len(cell_texts) and column_index < len(cell_texts[row_index]):
                    value = cell_texts[row_index][column_index]
                text_node.text = value

                cell_addr = etree.SubElement(cell, qname("hp", "cellAddr"))
                cell_addr.set("colAddr", str(column_index))
                cell_addr.set("rowAddr", str(row_index))

                cell_span = etree.SubElement(cell, qname("hp", "cellSpan"))
                cell_span.set("colSpan", "1")
                cell_span.set("rowSpan", "1")

                cell_size = etree.SubElement(cell, qname("hp", "cellSz"))
                cell_size.set("width", str(cell_width))
                cell_size.set("height", str(row_height))

                cell_margin = etree.SubElement(cell, qname("hp", "cellMargin"))
                cell_margin.set("left", "141")
                cell_margin.set("right", "141")
                cell_margin.set("top", "141")
                cell_margin.set("bottom", "141")

        section = self._append_run_child(section_index, paragraph_index, table, char_pr_id=char_pr_id)
        return TableXml(self, section, table)

    def append_picture(
        self,
        name: str,
        data: bytes,
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
        media_type: str | None = None,
        manifest_id: str | None = None,
        width: int = 7200,
        height: int = 7200,
        original_width: int | None = None,
        original_height: int | None = None,
        current_width: int | None = None,
        current_height: int | None = None,
        treat_as_char: bool = True,
        shape_comment: str | None = None,
        text_wrap: str = "TOP_AND_BOTTOM",
        text_flow: str = "BOTH_SIDES",
        affect_line_spacing: bool = False,
        flow_with_text: bool = True,
        allow_overlap: bool = False,
        hold_anchor_and_so: bool = False,
        vert_rel_to: str = "PARA",
        horz_rel_to: str = "COLUMN",
        vert_align: str = "TOP",
        horz_align: str = "LEFT",
        vert_offset: int = 0,
        horz_offset: int = 0,
        out_margin_left: int = 0,
        out_margin_right: int = 0,
        out_margin_top: int = 0,
        out_margin_bottom: int = 0,
    ) -> PictureXml:
        self._ensure_editable_sections()
        if width < 1 or height < 1:
            raise ValueError("width and height must be positive.")

        binary_name = name if manifest_id is not None else self._unused_binary_name(name)
        resolved_manifest_id = manifest_id or self.content_hpf.next_manifest_id(PurePosixPath(binary_name).stem or "image")
        self.add_or_replace_binary(binary_name, data, media_type=media_type, manifest_id=resolved_manifest_id)

        image_width = original_width or width
        image_height = original_height or height
        display_width = current_width if current_width is not None else width
        display_height = current_height if current_height is not None else height
        picture = etree.Element(qname("hp", "pic"))
        picture.set("id", self._next_control_number(_graphic_attribute_xpath("id")))
        picture.set("zOrder", self._next_control_number(_graphic_attribute_xpath("zOrder")))
        picture.set("numberingType", "PICTURE")
        picture.set("textWrap", "TOP_AND_BOTTOM")
        picture.set("textFlow", "BOTH_SIDES")
        picture.set("lock", "0")
        picture.set("dropcapstyle", "None")
        picture.set("href", "")
        picture.set("groupLevel", "0")
        picture.set("instid", self._next_control_number(_graphic_attribute_xpath("instid")))
        picture.set("reverse", "0")

        offset = etree.SubElement(picture, qname("hp", "offset"))
        offset.set("x", "0")
        offset.set("y", "0")

        original_size = etree.SubElement(picture, qname("hp", "orgSz"))
        original_size.set("width", str(image_width))
        original_size.set("height", str(image_height))

        current_size = etree.SubElement(picture, qname("hp", "curSz"))
        current_size.set("width", str(display_width))
        current_size.set("height", str(display_height))

        flip = etree.SubElement(picture, qname("hp", "flip"))
        flip.set("horizontal", "0")
        flip.set("vertical", "0")

        rotation = etree.SubElement(picture, qname("hp", "rotationInfo"))
        rotation.set("angle", "0")
        rotation.set("centerX", str(width // 2))
        rotation.set("centerY", str(height // 2))
        rotation.set("rotateimage", "1")

        rendering = etree.SubElement(picture, qname("hp", "renderingInfo"))
        self._append_identity_matrix(rendering, "transMatrix")
        scale_x = width / image_width if image_width else 1
        scale_y = height / image_height if image_height else 1
        self._append_matrix(rendering, "scaMatrix", e1=f"{scale_x:.6f}", e5=f"{scale_y:.6f}")
        self._append_identity_matrix(rendering, "rotMatrix")

        image = etree.SubElement(picture, qname("hc", "img"))
        image.set("binaryItemIDRef", resolved_manifest_id)
        image.set("bright", "0")
        image.set("contrast", "0")
        image.set("effect", "REAL_PIC")
        image.set("alpha", "0")

        image_rect = etree.SubElement(picture, qname("hp", "imgRect"))
        for index, (x, y) in enumerate(((0, 0), (image_width, 0), (image_width, image_height), (0, image_height))):
            point = etree.SubElement(image_rect, qname("hc", f"pt{index}"))
            point.set("x", str(x))
            point.set("y", str(y))

        image_clip = etree.SubElement(picture, qname("hp", "imgClip"))
        image_clip.set("left", "0")
        image_clip.set("right", str(image_width))
        image_clip.set("top", "0")
        image_clip.set("bottom", str(image_height))

        in_margin = etree.SubElement(picture, qname("hp", "inMargin"))
        for key in ("left", "right", "top", "bottom"):
            in_margin.set(key, "0")

        image_dim = etree.SubElement(picture, qname("hp", "imgDim"))
        image_dim.set("dimwidth", str(image_width))
        image_dim.set("dimheight", str(image_height))
        etree.SubElement(picture, qname("hp", "effects"))

        size = etree.SubElement(picture, qname("hp", "sz"))
        size.set("width", str(width))
        size.set("widthRelTo", "ABSOLUTE")
        size.set("height", str(height))
        size.set("heightRelTo", "ABSOLUTE")
        size.set("protect", "0")

        position = etree.SubElement(picture, qname("hp", "pos"))
        position.set("treatAsChar", "1" if treat_as_char else "0")
        position.set("affectLSpacing", "0")
        position.set("flowWithText", "1")
        position.set("allowOverlap", "0")
        position.set("holdAnchorAndSO", "0")
        position.set("vertRelTo", "PARA")
        position.set("horzRelTo", "COLUMN")
        position.set("vertAlign", "TOP")
        position.set("horzAlign", "LEFT")
        position.set("vertOffset", "0")
        position.set("horzOffset", "0")

        out_margin = etree.SubElement(picture, qname("hp", "outMargin"))
        for key in ("left", "right", "top", "bottom"):
            out_margin.set(key, "0")
        _set_graphic_layout(
            picture,
            text_wrap=text_wrap,
            text_flow=text_flow,
            treat_as_char=treat_as_char,
            affect_line_spacing=affect_line_spacing,
            flow_with_text=flow_with_text,
            allow_overlap=allow_overlap,
            hold_anchor_and_so=hold_anchor_and_so,
            vert_rel_to=vert_rel_to,
            horz_rel_to=horz_rel_to,
            vert_align=vert_align,
            horz_align=horz_align,
            vert_offset=vert_offset,
            horz_offset=horz_offset,
        )
        _set_margin_values(
            picture,
            "./hp:outMargin",
            left=out_margin_left,
            right=out_margin_right,
            top=out_margin_top,
            bottom=out_margin_bottom,
        )

        if shape_comment:
            comment = etree.SubElement(picture, qname("hp", "shapeComment"))
            comment.text = shape_comment

        section = self._append_run_child(section_index, paragraph_index, picture, char_pr_id=char_pr_id)
        return PictureXml(self, section, picture)

    def append_shape(
        self,
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
        kind: str = "rect",
        text: str = "",
        width: int = 12000,
        height: int = 3200,
        treat_as_char: bool = True,
        shape_comment: str | None = None,
        fill_color: str = "#FFFFFF",
        line_color: str = "#000000",
        line_width: int = 33,
        original_width: int | None = None,
        original_height: int | None = None,
        current_width: int | None = None,
        current_height: int | None = None,
        text_wrap: str = "TOP_AND_BOTTOM",
        text_flow: str = "BOTH_SIDES",
        affect_line_spacing: bool = False,
        flow_with_text: bool = True,
        allow_overlap: bool = False,
        hold_anchor_and_so: bool = False,
        vert_rel_to: str = "PARA",
        horz_rel_to: str = "COLUMN",
        vert_align: str = "TOP",
        horz_align: str = "LEFT",
        vert_offset: int = 0,
        horz_offset: int = 0,
        out_margin_left: int = 0,
        out_margin_right: int = 0,
        out_margin_top: int = 0,
        out_margin_bottom: int = 0,
    ) -> ShapeXml:
        self._ensure_editable_sections()
        supported_kinds = {"rect", "ellipse", "arc", "polygon", "curve", "connectLine", "line", "textart", "container"}
        if kind not in supported_kinds:
            supported = ", ".join(sorted(supported_kinds))
            raise ValueError(f"append_shape() supports kind values: {supported}")
        if width < 1 or height < 1:
            raise ValueError("width and height must be positive.")

        shape = etree.Element(qname("hp", kind))
        shape.set("id", self._next_control_number(_graphic_attribute_xpath("id")))
        shape.set("zOrder", self._next_control_number(_graphic_attribute_xpath("zOrder")))
        shape.set("numberingType", "PICTURE")
        shape.set("textWrap", "TOP_AND_BOTTOM")
        shape.set("textFlow", "BOTH_SIDES")
        shape.set("lock", "0")
        shape.set("dropcapstyle", "None")
        shape.set("href", "")
        shape.set("groupLevel", "0")
        shape.set("instid", self._next_control_number(_graphic_attribute_xpath("instid")))
        if kind == "textart":
            shape.set("text", text)
        elif kind == "ellipse":
            shape.set("intervalDirty", "0")
            shape.set("hasArcPr", "0")
            shape.set("arcType", "NORMAL")
        elif kind == "arc":
            shape.set("type", "NORMAL")
        elif kind == "connectLine":
            shape.set("type", "STRAIGHT_NOARROW")
        if kind not in {"line", "connectLine"}:
            shape.set("ratio", "0")

        offset = etree.SubElement(shape, qname("hp", "offset"))
        offset.set("x", "0")
        offset.set("y", "0")

        display_width = current_width if current_width is not None else width
        display_height = current_height if current_height is not None else height
        image_width = original_width or width
        image_height = original_height or height

        original_size = etree.SubElement(shape, qname("hp", "orgSz"))
        original_size.set("width", str(image_width))
        original_size.set("height", str(image_height))

        current_size = etree.SubElement(shape, qname("hp", "curSz"))
        current_size.set("width", str(display_width))
        current_size.set("height", str(display_height))

        flip = etree.SubElement(shape, qname("hp", "flip"))
        flip.set("horizontal", "0")
        flip.set("vertical", "0")

        rotation = etree.SubElement(shape, qname("hp", "rotationInfo"))
        rotation.set("angle", "0")
        rotation.set("centerX", str(width // 2))
        rotation.set("centerY", str(height // 2))
        rotation.set("rotateimage", "1")

        rendering = etree.SubElement(shape, qname("hp", "renderingInfo"))
        self._append_identity_matrix(rendering, "transMatrix")
        self._append_identity_matrix(rendering, "scaMatrix")
        self._append_identity_matrix(rendering, "rotMatrix")

        if kind != "container":
            line_shape = etree.SubElement(shape, qname("hp", "lineShape"))
            line_shape.set("color", line_color)
            line_shape.set("width", str(line_width))
            line_shape.set("style", "SOLID" if line_width > 0 else "NONE")
            line_shape.set("endCap", "FLAT")
            line_shape.set("headStyle", "NORMAL")
            line_shape.set("tailStyle", "NORMAL")
            line_shape.set("headfill", "1")
            line_shape.set("tailfill", "1")
            line_shape.set("headSz", "MEDIUM_MEDIUM")
            line_shape.set("tailSz", "MEDIUM_MEDIUM")
            line_shape.set("outlineStyle", "NORMAL")
            line_shape.set("alpha", "0")

            if kind not in {"line", "connectLine"}:
                fill_brush = etree.SubElement(shape, qname("hc", "fillBrush"))
                win_brush = etree.SubElement(fill_brush, qname("hc", "winBrush"))
                win_brush.set("faceColor", fill_color)
                win_brush.set("hatchColor", line_color)
                win_brush.set("alpha", "0")

            shadow = etree.SubElement(shape, qname("hp", "shadow"))
            shadow.set("type", "NONE")
            shadow.set("color", "#B2B2B2")
            shadow.set("offsetX", "0")
            shadow.set("offsetY", "0")
            shadow.set("alpha", "0")

        if kind in {"rect", "ellipse", "arc", "polygon", "curve", "container"}:
            draw_text = etree.SubElement(shape, qname("hp", "drawText"))
            draw_text.set("lastWidth", str(width))
            draw_text.set("name", "")
            draw_text.set("editable", "0")

            sub_list = etree.SubElement(draw_text, qname("hp", "subList"))
            sub_list.set("id", "")
            sub_list.set("textDirection", "HORIZONTAL")
            sub_list.set("lineWrap", "BREAK")
            sub_list.set("vertAlign", "CENTER")
            sub_list.set("linkListIDRef", "0")
            sub_list.set("linkListNextIDRef", "0")
            sub_list.set("textWidth", "0")
            sub_list.set("textHeight", "0")
            sub_list.set("hasTextRef", "0")
            sub_list.set("hasNumRef", "0")

            paragraph = etree.SubElement(sub_list, qname("hp", "p"))
            paragraph.set("id", self._next_control_number(".//hp:p/@id"))
            paragraph.set("paraPrIDRef", "0")
            paragraph.set("styleIDRef", "0")
            paragraph.set("pageBreak", "0")
            paragraph.set("columnBreak", "0")
            paragraph.set("merged", "0")

            run = etree.SubElement(paragraph, qname("hp", "run"))
            run.set("charPrIDRef", char_pr_id or "0")
            text_node = etree.SubElement(run, qname("hp", "t"))
            text_node.text = text

            text_margin = etree.SubElement(draw_text, qname("hp", "textMargin"))
            text_margin.set("left", "283")
            text_margin.set("right", "283")
            text_margin.set("top", "283")
            text_margin.set("bottom", "283")
        elif kind == "textart":
            textart_pr = etree.SubElement(shape, qname("hp", "textartPr"))
            textart_pr.set("fontName", "HY태백B")
            textart_pr.set("fontStyle", "보통")
            textart_pr.set("fontType", "TTF")
            textart_pr.set("textShape", "PLAIN_TEXT")
            textart_pr.set("lineSpacing", "120")
            textart_pr.set("charSpacing", "100")
            textart_pr.set("align", "LEFT")
            textart_shadow = etree.SubElement(textart_pr, qname("hp", "shadow"))
            textart_shadow.set("type", "NONE")
            textart_shadow.set("color", "#000000")
            textart_shadow.set("offsetX", "0")
            textart_shadow.set("offsetY", "0")
            textart_shadow.set("alpha", "0")

        if kind == "line":
            points = (("startPt", 0, 0), ("endPt", width, height))
        elif kind == "connectLine":
            points = (("startPt", 0, 0), ("endPt", width, height))
        elif kind in {"rect", "textart"}:
            points = tuple(
                (f"pt{index}", x, y)
                for index, (x, y) in enumerate(((0, 0), (width, 0), (width, height), (0, height)))
            )
        elif kind == "ellipse":
            points = (
                ("center", width // 2, height // 2),
                ("ax1", width, height // 2),
                ("ax2", width // 2, height),
                ("start1", width, height // 2),
                ("start2", width, height // 2),
                ("end1", width, height // 2),
                ("end2", width, height // 2),
            )
        elif kind == "arc":
            points = (
                ("center", width // 2, height // 2),
                ("ax1", width, height // 2),
                ("ax2", width // 2, height),
            )
        elif kind == "polygon":
            points = tuple(("pt", x, y) for x, y in ((0, 0), (width, 0), (width, height), (0, height)))
        else:
            points = ()

        for point_name, x, y in points:
            point = etree.SubElement(shape, qname("hc", point_name))
            point.set("x", str(x))
            point.set("y", str(y))
            if kind == "connectLine":
                point.set("subjectIDRef", "0")
                point.set("subjectIdx", "0")

        if kind == "curve":
            segment = etree.SubElement(shape, qname("hp", "seg"))
            segment.set("type", "LINE")
            segment.set("x1", "0")
            segment.set("y1", "0")
            segment.set("x2", str(width))
            segment.set("y2", str(height))

        size = etree.SubElement(shape, qname("hp", "sz"))
        size.set("width", str(width))
        size.set("widthRelTo", "ABSOLUTE")
        size.set("height", str(height))
        size.set("heightRelTo", "ABSOLUTE")
        size.set("protect", "0")

        position = etree.SubElement(shape, qname("hp", "pos"))
        position.set("treatAsChar", "1" if treat_as_char else "0")
        position.set("affectLSpacing", "0")
        position.set("flowWithText", "1")
        position.set("allowOverlap", "0")
        position.set("holdAnchorAndSO", "0")
        position.set("vertRelTo", "PARA")
        position.set("horzRelTo", "COLUMN")
        position.set("vertAlign", "TOP")
        position.set("horzAlign", "LEFT")
        position.set("vertOffset", "0")
        position.set("horzOffset", "0")

        out_margin = etree.SubElement(shape, qname("hp", "outMargin"))
        for key in ("left", "right", "top", "bottom"):
            out_margin.set(key, "0")
        _set_graphic_layout(
            shape,
            text_wrap=text_wrap,
            text_flow=text_flow,
            treat_as_char=treat_as_char,
            affect_line_spacing=affect_line_spacing,
            flow_with_text=flow_with_text,
            allow_overlap=allow_overlap,
            hold_anchor_and_so=hold_anchor_and_so,
            vert_rel_to=vert_rel_to,
            horz_rel_to=horz_rel_to,
            vert_align=vert_align,
            horz_align=horz_align,
            vert_offset=vert_offset,
            horz_offset=horz_offset,
        )
        _set_margin_values(
            shape,
            "./hp:outMargin",
            left=out_margin_left,
            right=out_margin_right,
            top=out_margin_top,
            bottom=out_margin_bottom,
        )

        if shape_comment:
            comment = etree.SubElement(shape, qname("hp", "shapeComment"))
            comment.text = shape_comment

        section = self._append_run_child(section_index, paragraph_index, shape, char_pr_id=char_pr_id)
        return ShapeXml(self, section, shape)

    def append_ole(
        self,
        name: str,
        data: bytes,
        *,
        section_index: int = 0,
        paragraph_index: int | None = None,
        char_pr_id: str | None = None,
        media_type: str = "application/ole",
        manifest_id: str | None = None,
        width: int = 42001,
        height: int = 13501,
        original_width: int | None = None,
        original_height: int | None = None,
        current_width: int = 0,
        current_height: int = 0,
        treat_as_char: bool = False,
        shape_comment: str | None = None,
        object_type: str = "EMBEDDED",
        has_moniker: bool = False,
        draw_aspect: str = "CONTENT",
        eq_baseline: int = 0,
        text_wrap: str = "SQUARE",
        text_flow: str = "BOTH_SIDES",
        affect_line_spacing: bool = False,
        flow_with_text: bool = True,
        allow_overlap: bool = False,
        hold_anchor_and_so: bool = False,
        vert_rel_to: str = "PARA",
        horz_rel_to: str = "COLUMN",
        vert_align: str = "TOP",
        horz_align: str = "LEFT",
        vert_offset: int = 0,
        horz_offset: int = 0,
        out_margin_left: int = 0,
        out_margin_right: int = 0,
        out_margin_top: int = 0,
        out_margin_bottom: int = 0,
    ) -> OleXml:
        self._ensure_editable_sections()
        if width < 1 or height < 1:
            raise ValueError("width and height must be positive.")

        binary_name = name if manifest_id is not None else self._unused_binary_name(name)
        resolved_manifest_id = manifest_id or self.content_hpf.next_manifest_id(PurePosixPath(binary_name).stem or "ole")
        self.add_or_replace_binary(
            binary_name,
            normalize_hwpx_ole_payload(data),
            media_type=media_type,
            manifest_id=resolved_manifest_id,
            is_embedded=False,
        )

        resolved_original_width = original_width or width
        resolved_original_height = original_height or height

        ole = etree.Element(qname("hp", "ole"))
        ole.set("id", self._next_control_number(_graphic_attribute_xpath("id")))
        ole.set("zOrder", self._next_control_number(_graphic_attribute_xpath("zOrder")))
        ole.set("numberingType", "PICTURE")
        ole.set("textWrap", text_wrap)
        ole.set("textFlow", text_flow)
        ole.set("lock", "0")
        ole.set("dropcapstyle", "None")
        ole.set("href", "")
        ole.set("groupLevel", "0")
        ole.set("instid", self._next_control_number(_graphic_attribute_xpath("instid")))
        ole.set("objectType", object_type)
        ole.set("binaryItemIDRef", resolved_manifest_id)
        ole.set("hasMoniker", "1" if has_moniker else "0")
        ole.set("drawAspect", draw_aspect)
        ole.set("eqBaseLine", str(eq_baseline))

        offset = etree.SubElement(ole, qname("hp", "offset"))
        offset.set("x", "0")
        offset.set("y", "0")

        original_size = etree.SubElement(ole, qname("hp", "orgSz"))
        original_size.set("width", str(resolved_original_width))
        original_size.set("height", str(resolved_original_height))

        current_size = etree.SubElement(ole, qname("hp", "curSz"))
        current_size.set("width", str(current_width))
        current_size.set("height", str(current_height))

        flip = etree.SubElement(ole, qname("hp", "flip"))
        flip.set("horizontal", "0")
        flip.set("vertical", "0")

        rotation = etree.SubElement(ole, qname("hp", "rotationInfo"))
        rotation.set("angle", "0")
        rotation.set("centerX", str(width // 2))
        rotation.set("centerY", str(height // 2))
        rotation.set("rotateimage", "1")

        rendering = etree.SubElement(ole, qname("hp", "renderingInfo"))
        self._append_identity_matrix(rendering, "transMatrix")
        self._append_identity_matrix(rendering, "scaMatrix")
        self._append_identity_matrix(rendering, "rotMatrix")

        extent = etree.SubElement(ole, qname("hc", "extent"))
        extent.set("x", str(width))
        extent.set("y", str(height))

        line_shape = etree.SubElement(ole, qname("hp", "lineShape"))
        line_shape.set("color", "#000000")
        line_shape.set("width", "0")
        line_shape.set("style", "NONE")
        line_shape.set("endCap", "ROUND")
        line_shape.set("headStyle", "NORMAL")
        line_shape.set("tailStyle", "NORMAL")
        line_shape.set("headfill", "0")
        line_shape.set("tailfill", "0")
        line_shape.set("headSz", "SMALL_SMALL")
        line_shape.set("tailSz", "SMALL_SMALL")
        line_shape.set("outlineStyle", "OUTER")
        line_shape.set("alpha", "0")

        size = etree.SubElement(ole, qname("hp", "sz"))
        size.set("width", str(width))
        size.set("widthRelTo", "ABSOLUTE")
        size.set("height", str(height))
        size.set("heightRelTo", "ABSOLUTE")
        size.set("protect", "0")

        position = etree.SubElement(ole, qname("hp", "pos"))
        position.set("treatAsChar", "1" if treat_as_char else "0")
        position.set("affectLSpacing", "0")
        position.set("flowWithText", "1")
        position.set("allowOverlap", "0")
        position.set("holdAnchorAndSO", "0")
        position.set("vertRelTo", "PARA")
        position.set("horzRelTo", "COLUMN")
        position.set("vertAlign", "TOP")
        position.set("horzAlign", "LEFT")
        position.set("vertOffset", "0")
        position.set("horzOffset", "0")

        out_margin = etree.SubElement(ole, qname("hp", "outMargin"))
        for key in ("left", "right", "top", "bottom"):
            out_margin.set(key, "0")

        _set_graphic_layout(
            ole,
            text_wrap=text_wrap,
            text_flow=text_flow,
            treat_as_char=treat_as_char,
            affect_line_spacing=affect_line_spacing,
            flow_with_text=flow_with_text,
            allow_overlap=allow_overlap,
            hold_anchor_and_so=hold_anchor_and_so,
            vert_rel_to=vert_rel_to,
            horz_rel_to=horz_rel_to,
            vert_align=vert_align,
            horz_align=horz_align,
            vert_offset=vert_offset,
            horz_offset=horz_offset,
        )
        _set_margin_values(
            ole,
            "./hp:outMargin",
            left=out_margin_left,
            right=out_margin_right,
            top=out_margin_top,
            bottom=out_margin_bottom,
        )

        if shape_comment:
            comment = etree.SubElement(ole, qname("hp", "shapeComment"))
            comment.text = shape_comment

        section = self._append_run_child(section_index, paragraph_index, ole, char_pr_id=char_pr_id)
        return OleXml(self, section, ole)


for _compat_method in HwpxDocumentGraphicMixin.__dict__.values():
    if callable(_compat_method):
        _compat_method.__module__ = "jakal_hwpx.document"

