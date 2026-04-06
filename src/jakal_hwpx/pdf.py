from __future__ import annotations

import io
import math
import os
import textwrap
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from pypdf import PdfReader, PdfWriter, Transformation
from pypdf.generic import ArrayObject, DecodedStreamObject, DictionaryObject, NameObject

try:
    from PIL import Image
except ImportError:  # pragma: no cover - optional runtime dependency
    Image = None


_STANDARD_FONT_RESOURCE = "/F1"
_STANDARD_FONT_NAME = "/Helvetica"
_TEXT_LEADING_RATIO = 1.4
_DEFAULT_PAGE_MARGIN = 72.0
_VECTOR_PATH_OPERATORS = {b"m", b"l", b"c", b"v", b"y", b"re"}
_VECTOR_PAINT_OPERATORS = {b"S", b"s", b"f", b"F", b"f*", b"B", b"B*", b"b", b"b*", b"n"}
_TEXT_OPERATORS = {b"BT", b"ET", b"Tj", b"TJ", b"'", b'"'}
_IMAGE_OPERATORS = {b"Do"}
_GRAPHICS_STATE_PUSH = b"q"
_GRAPHICS_STATE_POP = b"Q"
_GRAPHICS_STATE_CONCAT = b"cm"


def _resolve_object(value: Any) -> Any:
    return value.get_object() if hasattr(value, "get_object") else value


def _escape_pdf_text(value: str) -> str:
    return value.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def _wrap_text_for_page(text: str, *, usable_width: float, font_size: float) -> list[str]:
    if not text:
        return [""]

    average_character_width = max(font_size * 0.55, 1.0)
    max_chars = max(int(usable_width / average_character_width), 1)
    lines: list[str] = []
    for paragraph in text.splitlines() or [""]:
        if not paragraph:
            lines.append("")
            continue
        lines.extend(textwrap.wrap(paragraph, width=max_chars, break_long_words=True, break_on_hyphens=False))
    return lines or [""]


def _matrix_multiply(left: tuple[float, float, float, float, float, float], right: tuple[float, float, float, float, float, float]) -> tuple[float, float, float, float, float, float]:
    a1, b1, c1, d1, e1, f1 = left
    a2, b2, c2, d2, e2, f2 = right
    return (
        a1 * a2 + c1 * b2,
        b1 * a2 + d1 * b2,
        a1 * c2 + c1 * d2,
        b1 * c2 + d1 * d2,
        a1 * e2 + c1 * f2 + e1,
        b1 * e2 + d1 * f2 + f1,
    )


@dataclass
class PdfMetadata:
    title: str | None = None
    author: str | None = None
    subject: str | None = None
    creator: str | None = None
    producer: str | None = None
    keywords: str | None = None
    extra: dict[str, str] = field(default_factory=dict)


@dataclass
class PdfPageImage:
    name: str
    width: int | None = None
    height: int | None = None
    bits_per_component: int | None = None
    color_space: str | None = None
    data: bytes = b""


@dataclass
class PdfImagePlacement:
    name: str
    data: bytes
    x: float
    y: float
    width: float
    height: float


@dataclass
class PdfAnnotation:
    subtype: str | None = None
    contents: str | None = None
    rect: tuple[float, float, float, float] | None = None


@dataclass
class PdfVectorSummary:
    path_operator_count: int = 0
    paint_operator_count: int = 0
    text_operator_count: int = 0
    image_operator_count: int = 0


@dataclass
class PdfPageAnalysis:
    image_count: int = 0
    annotation_count: int = 0
    xobject_names: list[str] = field(default_factory=list)
    vector_summary: PdfVectorSummary = field(default_factory=PdfVectorSummary)


class PdfPage:
    def __init__(self, document: "PdfDocument", index: int):
        self.document = document
        self.index = index

    @property
    def number(self) -> int:
        return self.index + 1

    @property
    def width(self) -> float:
        page = self.document._page_object(self.index)
        return float(page.mediabox.right) - float(page.mediabox.left)

    @property
    def height(self) -> float:
        page = self.document._page_object(self.index)
        return float(page.mediabox.top) - float(page.mediabox.bottom)

    @property
    def rotation(self) -> int:
        page = self.document._page_object(self.index)
        return int(page.rotation or 0)

    def extract_text(self) -> str:
        page = self.document._page_object(self.index)
        return page.extract_text() or ""

    def images(self) -> list[PdfPageImage]:
        page = self.document._page_object(self.index)
        images: list[PdfPageImage] = []
        for image in getattr(page, "images", []):
            descriptor = _resolve_object(getattr(image, "indirect_reference", None)) or {}
            images.append(
                PdfPageImage(
                    name=getattr(image, "name", "image"),
                    width=descriptor.get("/Width"),
                    height=descriptor.get("/Height"),
                    bits_per_component=descriptor.get("/BitsPerComponent"),
                    color_space=str(descriptor.get("/ColorSpace")) if descriptor.get("/ColorSpace") is not None else None,
                    data=getattr(image, "data", b""),
                )
            )
        return images

    def image_placements(self) -> list[PdfImagePlacement]:
        page = self.document._page_object(self.index)
        resources = _resolve_object(page.get("/Resources")) or {}
        xobjects = _resolve_object(resources.get("/XObject")) or {}

        image_map: dict[str, PdfPageImage] = {}
        for image in self.images():
            base_name = Path(image.name).stem
            image_map[image.name] = image
            image_map[base_name] = image
            image_map[f"/{base_name}"] = image

        placements: list[PdfImagePlacement] = []
        stack: list[tuple[float, float, float, float, float, float]] = []
        current = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
        contents = page.get_contents()
        operations = getattr(contents, "operations", []) if contents is not None else []

        for operands, operator in operations:
            if operator == _GRAPHICS_STATE_PUSH:
                stack.append(current)
                continue
            if operator == _GRAPHICS_STATE_POP:
                current = stack.pop() if stack else (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
                continue
            if operator == _GRAPHICS_STATE_CONCAT and len(operands) == 6:
                matrix = tuple(float(value) for value in operands)
                current = _matrix_multiply(current, matrix)
                continue
            if operator != b"Do" or not operands:
                continue

            raw_name = operands[0]
            name = str(raw_name)
            xobject = _resolve_object(xobjects.get(NameObject(name))) if NameObject(name) in xobjects else None
            if xobject is None or xobject.get("/Subtype") != "/Image":
                continue

            image = image_map.get(name) or image_map.get(name.lstrip("/"))
            if image is None:
                continue

            a, b, c, d, e, f = current
            width = math.hypot(a, b)
            height = math.hypot(c, d)
            placements.append(
                PdfImagePlacement(
                    name=image.name,
                    data=image.data,
                    x=e,
                    y=f,
                    width=width,
                    height=height,
                )
            )
        return placements

    def annotations(self) -> list[PdfAnnotation]:
        page = self.document._page_object(self.index)
        values = _resolve_object(page.get("/Annots")) or []
        annotations: list[PdfAnnotation] = []
        for annotation_ref in values:
            annotation = _resolve_object(annotation_ref) or {}
            rect = annotation.get("/Rect")
            resolved_rect: tuple[float, float, float, float] | None = None
            if rect is not None and len(rect) == 4:
                resolved_rect = tuple(float(value) for value in rect)
            annotations.append(
                PdfAnnotation(
                    subtype=str(annotation.get("/Subtype")) if annotation.get("/Subtype") is not None else None,
                    contents=annotation.get("/Contents"),
                    rect=resolved_rect,
                )
            )
        return annotations

    def analyze_non_text(self) -> PdfPageAnalysis:
        page = self.document._page_object(self.index)
        resources = _resolve_object(page.get("/Resources")) or {}
        xobjects = _resolve_object(resources.get("/XObject")) or {}
        summary = PdfVectorSummary()
        contents = page.get_contents()
        operations = getattr(contents, "operations", []) if contents is not None else []
        for _operands, operator in operations:
            if operator in _VECTOR_PATH_OPERATORS:
                summary.path_operator_count += 1
            elif operator in _VECTOR_PAINT_OPERATORS:
                summary.paint_operator_count += 1
            elif operator in _TEXT_OPERATORS:
                summary.text_operator_count += 1
            elif operator in _IMAGE_OPERATORS:
                summary.image_operator_count += 1

        return PdfPageAnalysis(
            image_count=len(self.images()),
            annotation_count=len(self.annotations()),
            xobject_names=[str(name) for name in xobjects.keys()],
            vector_summary=summary,
        )

    def add_text(
        self,
        text: str,
        *,
        x: float = _DEFAULT_PAGE_MARGIN,
        y: float | None = None,
        font_size: float = 12.0,
        leading: float | None = None,
    ) -> None:
        writer_page = self.document._writer_page_object(self.index)
        self.document._ensure_standard_font(writer_page)

        resolved_y = self.height - _DEFAULT_PAGE_MARGIN if y is None else y
        resolved_leading = font_size * _TEXT_LEADING_RATIO if leading is None else leading
        lines = text.splitlines() or [""]

        commands = [
            "BT",
            f"{_STANDARD_FONT_RESOURCE} {font_size:.2f} Tf",
            f"1 0 0 1 {x:.2f} {resolved_y:.2f} Tm",
        ]
        first_line = True
        for line in lines:
            if not first_line:
                commands.append(f"0 {-resolved_leading:.2f} Td")
            commands.append(f"({_escape_pdf_text(line)}) Tj")
            first_line = False
        commands.append("ET")
        self.document._append_content_stream(writer_page, ("\n".join(commands) + "\n").encode("utf-8"))

    def add_wrapped_text(
        self,
        text: str,
        *,
        margin_left: float = _DEFAULT_PAGE_MARGIN,
        margin_right: float = _DEFAULT_PAGE_MARGIN,
        margin_top: float = _DEFAULT_PAGE_MARGIN,
        font_size: float = 12.0,
        leading: float | None = None,
    ) -> None:
        usable_width = max(self.width - margin_left - margin_right, 1.0)
        lines = _wrap_text_for_page(text, usable_width=usable_width, font_size=font_size)
        self.add_text(
            "\n".join(lines),
            x=margin_left,
            y=self.height - margin_top,
            font_size=font_size,
            leading=leading,
        )

    def draw_line(
        self,
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        *,
        line_width: float = 1.0,
    ) -> None:
        commands = [
            f"{line_width:.2f} w",
            f"{x1:.2f} {y1:.2f} m",
            f"{x2:.2f} {y2:.2f} l",
            "S",
        ]
        self.document._append_content_stream(self.document._writer_page_object(self.index), ("\n".join(commands) + "\n").encode("utf-8"))

    def draw_rectangle(
        self,
        x: float,
        y: float,
        width: float,
        height: float,
        *,
        line_width: float = 1.0,
        fill_gray: float | None = None,
    ) -> None:
        commands = [f"{line_width:.2f} w"]
        if fill_gray is not None:
            commands.append(f"{fill_gray:.3f} g")
        commands.append(f"{x:.2f} {y:.2f} {width:.2f} {height:.2f} re")
        commands.append("B" if fill_gray is not None else "S")
        self.document._append_content_stream(self.document._writer_page_object(self.index), ("\n".join(commands) + "\n").encode("utf-8"))

    def add_image(
        self,
        data: bytes,
        *,
        x: float,
        y: float,
        width: float,
        height: float,
    ) -> None:
        if Image is None:
            raise RuntimeError("Pillow is required to place raster images into generated PDFs.")

        with Image.open(io.BytesIO(data)) as image:
            pdf_bytes = io.BytesIO()
            image.save(pdf_bytes, format="PDF")

        source_page = PdfReader(io.BytesIO(pdf_bytes.getvalue())).pages[0]
        source_width = float(source_page.mediabox.right) - float(source_page.mediabox.left)
        source_height = float(source_page.mediabox.top) - float(source_page.mediabox.bottom)
        scale_x = width / source_width if source_width else 1.0
        scale_y = height / source_height if source_height else 1.0
        self.document._writer_page_object(self.index).merge_transformed_page(
            source_page,
            Transformation().scale(scale_x, scale_y).translate(x, y),
        )


class PdfDocument:
    def __init__(
        self,
        *,
        reader: PdfReader | None = None,
        writer: PdfWriter | None = None,
        source_path: Path | None = None,
        original_bytes: bytes | None = None,
    ):
        self._reader = reader
        self._writer = writer
        self.source_path = source_path
        self._original_bytes = original_bytes

    @classmethod
    def blank(cls) -> "PdfDocument":
        return cls(writer=PdfWriter())

    @classmethod
    def open(cls, path: str | os.PathLike[str]) -> "PdfDocument":
        input_path = Path(path)
        raw_bytes = input_path.read_bytes()
        return cls.from_bytes(raw_bytes, source_path=input_path)

    @classmethod
    def from_bytes(cls, raw_bytes: bytes, source_path: Path | None = None) -> "PdfDocument":
        reader = PdfReader(io.BytesIO(raw_bytes))
        return cls(reader=reader, source_path=source_path, original_bytes=raw_bytes)

    @property
    def pages(self) -> list[PdfPage]:
        page_count = len(self._writer.pages) if self._writer is not None else len(self._reader.pages)
        return [PdfPage(self, index) for index in range(page_count)]

    def metadata(self) -> PdfMetadata:
        source = self._writer.metadata if self._writer is not None else self._reader.metadata
        metadata = PdfMetadata()
        raw_items = dict(source or {})
        metadata.title = raw_items.pop("/Title", None)
        metadata.author = raw_items.pop("/Author", None)
        metadata.subject = raw_items.pop("/Subject", None)
        metadata.creator = raw_items.pop("/Creator", None)
        metadata.producer = raw_items.pop("/Producer", None)
        metadata.keywords = raw_items.pop("/Keywords", None)
        metadata.extra = {str(key): str(value) for key, value in raw_items.items()}
        return metadata

    def set_metadata(
        self,
        *,
        title: str | None = None,
        author: str | None = None,
        subject: str | None = None,
        creator: str | None = None,
        producer: str | None = None,
        keywords: str | None = None,
        extra: dict[str, str] | None = None,
    ) -> None:
        self._ensure_writer()
        info: dict[str, str] = {}
        if title is not None:
            info["/Title"] = title
        if author is not None:
            info["/Author"] = author
        if subject is not None:
            info["/Subject"] = subject
        if creator is not None:
            info["/Creator"] = creator
        if producer is not None:
            info["/Producer"] = producer
        if keywords is not None:
            info["/Keywords"] = keywords
        if extra:
            info.update(extra)
        self._writer.add_metadata(info)

    def add_page(self, *, width: float = 595.0, height: float = 842.0) -> PdfPage:
        self._ensure_writer()
        self._writer.add_blank_page(width=width, height=height)
        return PdfPage(self, len(self._writer.pages) - 1)

    def get_document_text(self, page_separator: str = "\n\n") -> str:
        values = [page.extract_text().strip() for page in self.pages]
        return page_separator.join(value for value in values if value)

    def compile(self) -> bytes:
        if self._writer is None:
            return self._original_bytes or b""

        buffer = io.BytesIO()
        self._writer.write(buffer)
        return buffer.getvalue()

    def save(self, path: str | os.PathLike[str]) -> Path:
        output_path = Path(path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_bytes(self.compile())
        return output_path

    def _ensure_writer(self) -> None:
        if self._writer is not None:
            return
        assert self._reader is not None
        writer = PdfWriter()
        writer.clone_document_from_reader(self._reader)
        self._writer = writer

    def _page_object(self, index: int):
        if self._writer is not None:
            return self._writer.pages[index]
        assert self._reader is not None
        return self._reader.pages[index]

    def _writer_page_object(self, index: int):
        self._ensure_writer()
        return self._writer.pages[index]

    def _ensure_standard_font(self, page: Any) -> None:
        resources = _resolve_object(page.get("/Resources"))
        if resources is None:
            resources = DictionaryObject()
            page[NameObject("/Resources")] = resources
        fonts = _resolve_object(resources.get("/Font"))
        if fonts is None:
            fonts = DictionaryObject()
            resources[NameObject("/Font")] = fonts
        if NameObject(_STANDARD_FONT_RESOURCE) in fonts:
            return

        font = DictionaryObject(
            {
                NameObject("/Type"): NameObject("/Font"),
                NameObject("/Subtype"): NameObject("/Type1"),
                NameObject("/BaseFont"): NameObject(_STANDARD_FONT_NAME),
            }
        )
        fonts[NameObject(_STANDARD_FONT_RESOURCE)] = self._writer._add_object(font)

    def _append_content_stream(self, page: Any, data: bytes) -> None:
        self._ensure_writer()
        stream = DecodedStreamObject()
        stream.set_data(data)
        stream_ref = self._writer._add_object(stream)
        contents = page.get("/Contents")
        if contents is None:
            page[NameObject("/Contents")] = stream_ref
            return
        if isinstance(contents, ArrayObject):
            contents.append(stream_ref)
            return
        page[NameObject("/Contents")] = ArrayObject([contents, stream_ref])
