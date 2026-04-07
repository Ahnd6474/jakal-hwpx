# jakal-hwpx

`jakal-hwpx` is a Python library for opening, editing, validating, and saving ZIP-based HWPX documents.

## Repository Layout

- `src/jakal_hwpx`: package source
- `tests`: test suite
- `examples/samples/hwpx`: editable HWPX sample documents
- `examples/output_smoke`: small smoke-test corpus
- `examples/output`: generated showcase outputs

## Requirements

- Python 3.11 or newer

## Install

From PyPI:

```bash
python -m pip install --upgrade pip
python -m pip install jakal-hwpx
```

From a local checkout:

```bash
python -m pip install --upgrade pip
python -m pip install .
```

For development:

```bash
python -m pip install -e .[dev]
```

Package name: `jakal-hwpx`  
Import path: `jakal_hwpx`

## Quick Start

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.blank()
doc.set_metadata(title="Example", creator="jakal-hwpx")
doc.set_paragraph_text(0, 0, "Hello HWPX")
doc.save("build/hello.hwpx")
```

The package works with ZIP-based HWPX packages. It does not ship a bundled `.hwp` to `.hwpx` converter.

## Main API

Most usage starts with `HwpxDocument`.

- `HwpxDocument.open(path)`: open an existing HWPX file
- `HwpxDocument.blank()`: create a minimal new document
- `metadata()` / `set_metadata()`: read or update document metadata
- `get_document_text()`: extract body text
- `set_paragraph_text()`: replace paragraph text
- `append_paragraph()`: append a paragraph to a section
- `replace_text()`: replace text across the document
- `section_settings()`: inspect or change page settings
- `tables()`, `pictures()`, `notes()`, `fields()`: access structured document elements
- `validation_errors()`: validate document structure
- `save(path)`: write the package back to disk

Useful helper types exposed by the package include:

- `SectionSettings`
- `Table`, `TableCell`
- `HeaderFooterBlock`
- `Bookmark`, `Field`, `Note`, `Equation`, `ShapeObject`

See [HWPX_MODULE.md](./HWPX_MODULE.md) for the package structure and API notes.

## Examples

Open and validate a document:

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("input.hwpx")

print(doc.metadata())
print(doc.get_document_text())
print(doc.validation_errors())
print(doc.reference_validation_errors())
```

Edit metadata and body text:

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("input.hwpx")
doc.set_metadata(title="Edited Title", creator="Docs Team", keyword="example")
doc.replace_text("Draft", "Final")
doc.append_paragraph("Additional paragraph", section_index=0)
doc.save("build/edited.hwpx")
```

Update page settings and table contents:

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("input.hwpx")

settings = doc.section_settings(0)
settings.set_page_size(width=60000, height=85000)
settings.set_margins(left=7000, right=7000, top=5000, bottom=5000)

table = doc.tables()[0]
table.set_cell_text(0, 0, "Updated")
table.append_row()[0].set_text("New row")

doc.save("build/layout-updated.hwpx")
```

Add bookmarks, hyperlinks, and calculated fields:

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.blank()
bookmark = doc.append_bookmark("summary_anchor")
doc.append_hyperlink("https://example.com", display_text="Example")
doc.append_calculation_field("40+2", display_text="42")
doc.append_cross_reference(bookmark.name or "summary_anchor", display_text="Jump to summary")
doc.save("build/fields.hwpx")
```

## Tests

```bash
python -m pip install -e .[dev]
python -m pytest -q
```

The test suite looks for sample HWPX files in this order:

1. `JAKAL_HWPX_SAMPLE_DIR`
2. `all_hwpx_flat/`
3. `examples/output_smoke/`
4. `examples/output/`
5. `examples/samples/hwpx/`

## Additional Documents

- [HWPX_MODULE.md](./HWPX_MODULE.md): package structure and API details
- [examples/SHOWCASE.md](./examples/SHOWCASE.md): showcase generation flow
- [RELEASING.md](./RELEASING.md): release checklist
- [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md): notes for sample documents and redistribution

## License

Project-authored source code is licensed under the [MIT License](./LICENSE).

Sample documents and committed generated outputs may carry separate rights. Review [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md) before redistributing repository contents.
