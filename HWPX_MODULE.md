# HWPX Module Guide

This document describes the Python package in `src/jakal_hwpx`.

If you only need repository setup, sample locations, or test commands, start with [README.md](./README.md).

## Package Scope

`jakal_hwpx` currently has three layers:

1. `HwpxDocument`: the core HWPX package reader/editor/writer
2. `PdfDocument`: a lightweight PDF reader/generator built on `pypdf`
3. bridge helpers: `pdf_to_hwpx()`, `hwpx_to_pdf()`, `pdf_to_bridge()`, `hwpx_to_bridge()`

The HWPX layer is the stable center of the project. The PDF layer and bridge are intentionally conservative and focus on moving text, page size, images, and non-text summaries safely rather than claiming layout-perfect conversion.

## Main Entry Points

### `HwpxDocument`

Use this for `.hwpx` package editing.

Common workflows:

- open an existing HWPX file
- edit paragraphs, tables, pictures, styles, headers, fields, notes, equations, and shapes
- validate package structure before writing
- save or compile back to bytes

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("examples/samples/hwpx/AI와_특이점_보고서.hwpx")
doc.set_metadata(title="Edited title", creator="jakal-hwpx")
doc.append_paragraph("Added from Python.", section_index=0)
doc.save("build/edited.hwpx")
```

### `PdfDocument`

Use this for basic PDF inspection and generation.

Current responsibilities:

- open existing PDFs
- inspect page text, image resources, annotations, and vector activity
- create simple PDFs with text, lines, rectangles, and raster images

```python
from jakal_hwpx import PdfDocument

pdf = PdfDocument.blank()
page = pdf.add_page(width=595, height=842)
page.add_text("Hello PDF", x=72, y=760)
pdf.save("build/hello.pdf")
```

### Bridge APIs

Use these when you want format conversion with an explicit intermediate model.

- `pdf_to_bridge()`
- `hwpx_to_bridge()`
- `bridge_to_hwpx()`
- `bridge_to_pdf()`
- `pdf_to_hwpx()`
- `hwpx_to_pdf()`

```python
from jakal_hwpx import BridgeTextBlock, pdf_to_hwpx

hwpx = pdf_to_hwpx(
    "examples/samples/pdf/평가원 수학 양식 (2)-jakal_hwpx_수정.pdf",
    ocr_blocks_by_page={
        0: [
            BridgeTextBlock(text="OCR line 1", left=36, top=48, width=220, height=12),
            BridgeTextBlock(text="OCR line 2", left=36, top=66, width=220, height=12),
        ]
    },
)
hwpx.save("build/from-pdf.hwpx")
```

## HWPX Layer

### Core Constructors

- `HwpxDocument()`
- `HwpxDocument.blank()`
- `HwpxDocument.open(path)`
- `HwpxDocument.from_bytes(raw_bytes)`

### Common Accessors

- `metadata()`
- `sections`
- `headers()`
- `footers()`
- `tables()`
- `pictures()`
- `notes()`
- `bookmarks()`
- `fields()`
- `equations()`
- `shapes()`
- `styles()`
- `paragraph_styles()`
- `character_styles()`

### Common Editing Helpers

- `set_metadata()`
- `replace_text()`
- `append_paragraph()`
- `insert_paragraph()`
- `set_paragraph_text()`
- `delete_paragraph()`
- `apply_style_to_paragraph()`
- `apply_style_batch()`
- `add_section()`
- `remove_section()`
- `set_preview_text()`
- `add_or_replace_binary()`
- `append_bookmark()`
- `append_hyperlink()`
- `append_mail_merge_field()`
- `append_calculation_field()`
- `append_cross_reference()`
- `compile()`
- `save()`

### Validation Helpers

- `validation_errors()`
- `validate()`
- `roundtrip_validate()`
- `xml_validation_errors()`
- `schema_validation_errors()`
- `reference_validation_errors()`
- `save_reopen_validation_errors()`

### Low-Level Escape Hatch

`HwpxXmlNode` is still available when the typed wrappers do not cover a case yet.

## PDF Layer

### Main Types

- `PdfDocument`
- `PdfPage`
- `PdfMetadata`
- `PdfPageImage`
- `PdfImagePlacement`
- `PdfAnnotation`
- `PdfPageAnalysis`
- `PdfVectorSummary`

### What `PdfDocument` Supports Today

- open from disk or bytes
- inspect metadata
- add blank pages
- write simple text content
- draw lines and rectangles
- merge raster images into generated pages
- analyze non-text content in existing pages

### Non-Text Analysis

For each page you can inspect:

- image count
- annotation count
- XObject names
- vector path operation count
- vector paint operation count
- text operation count
- image draw operation count
- image placements extracted from `Do` operators and graphics-state transforms

This is the main basis for the current `pdf -> hwpx` bridge.

## Bridge Model

### Main Types

- `DocumentBridge`
- `BridgePage`
- `BridgePageFeatures`
- `BridgeTextBlock`

### `pdf -> hwpx`

Current behavior:

- page size becomes section page size
- OCR text can be supplied as page-level strings
- OCR text can also be supplied as positioned `BridgeTextBlock` objects
- raster images referenced on a PDF page are imported as HWPX `hp:pic`
- non-text metadata that cannot yet be faithfully recreated is preserved as marker text like `[NON_TEXT] images=... vector_ops=...`

Important limitation:

- OCR is treated as an upstream dependency
- the bridge does not run OCR itself

### `hwpx -> pdf`

Current behavior:

- section text becomes readable PDF text
- page size becomes PDF page size
- HWPX tables are rendered as simple stroked rectangles with cell text
- HWPX pictures are rendered as raster images when binary data is available
- basic shapes such as rectangles, textart-like boxes, and lines are rendered as simple PDF drawing operations

Important limitation:

- this is a structural bridge, not a layout-faithful renderer
- complex HWPX typography, exact positioning, equation rendering, and advanced drawing semantics are not fully reproduced yet

## Examples

### Open, Edit, Save HWPX

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("examples/samples/hwpx/social_security_application_example.hwpx")
doc.tables()[0].set_cell_text(0, 0, "Edited")
doc.save("build/edited-form.hwpx")
```

### Build a PDF From HWPX

```python
from jakal_hwpx import hwpx_to_pdf

pdf = hwpx_to_pdf("examples/samples/hwpx/AI와_특이점_보고서.hwpx")
pdf.save("build/from-hwpx.pdf")
```

### Build an HWPX From PDF

```python
from jakal_hwpx import pdf_to_hwpx

doc = pdf_to_hwpx(
    "examples/samples/pdf/평가원 수학 양식 (2)-jakal_hwpx_수정.pdf",
    ocr_text_by_page=["OCR output for page 1"],
)
doc.save("build/from-pdf.hwpx")
```

## Sample Files and Test Fixtures

Repository sample files are organized as:

- `examples/samples/hwpx/`
- `examples/samples/hwp/`
- `examples/samples/pdf/`
- `examples/output_smoke/`
- `examples/output/`

Tests use `examples/output_smoke/` first for a stable committed corpus, then fall back to other configured sample roots.

## Maintainer Notes

- `.hwp -> .hwpx` conversion relies on the bundled Java tooling in `tools/`
- generated validation artifacts belong in `build/validation/`
- the HWPX layer should stay the most stable part of the package
- bridge work should prefer adding new modules over destabilizing core HWPX editing behavior
