# jakal-hwpx

Python tools for reading, editing, validating, and writing `HWPX` documents.

Korean overview: [README.ko.md](./README.ko.md)

## What This Repo Contains

- `src/jakal_hwpx`: the importable Python package
- `examples/samples`: sample `hwpx` and `hwp` documents used for local experimentation
- `examples/output_smoke`: committed smoke corpus for tests
- `examples/output`: showcase outputs generated from the smoke corpus
- `tools`: bundled Java-based `.hwp -> .hwpx` converter assets used by maintainer workflows

This README is intentionally repo-level. Detailed module and API notes live in [HWPX_MODULE.md](./HWPX_MODULE.md).

## Installation

You need Python 3.11 or newer.

Install from PyPI:

```bash
python -m pip install --upgrade pip
python -m pip install jakal-hwpx
```

Install from a local checkout:

```bash
python -m pip install --upgrade pip
python -m pip install .
```

For editable development mode:

```bash
python -m pip install -e .[dev]
```

The package name is `jakal-hwpx` and the import path is `jakal_hwpx`.

## Quick Start

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.blank()
doc.set_metadata(title="Example", creator="jakal-hwpx")
doc.set_paragraph_text(0, 0, "Hello HWPX")
doc.save("build/hello.hwpx")
```

For module-level usage, document parts, and editing helpers, see [HWPX_MODULE.md](./HWPX_MODULE.md).

## Core API

Most users only need `HwpxDocument` and a few wrapper types returned from it.

- `HwpxDocument`: open, create, edit, validate, compile, and save HWPX packages
- `SectionSettings`: read or update page size and margins for a section
- `Table` and `TableCell`: inspect and edit table content
- `HeaderFooterBlock`: read or replace header and footer text
- `Bookmark`, `Field`, `Note`, `Equation`, `ShapeObject`: inspect and update richer document features
- `HwpxPart` and related part classes: low-level access when you need to work with package internals directly

### `HwpxDocument` at a glance

| Method | Purpose |
| --- | --- |
| `open(path)` | Open an existing HWPX package from disk |
| `blank()` | Create a new in-memory document with default parts |
| `metadata()` / `set_metadata()` | Read or update document metadata |
| `get_document_text()` | Extract body text across sections |
| `set_paragraph_text()` | Replace one paragraph's text |
| `append_paragraph()` | Add a paragraph to a section |
| `replace_text()` | Replace text across the document |
| `section_settings()` | Access page size and margins for a section |
| `tables()`, `pictures()`, `notes()`, `fields()` | Access richer document objects |
| `validation_errors()` | Check package-level validity before save |
| `save(path)` | Write the package back to disk |

### Wrapper types you will see most often

| Type | Typical use |
| --- | --- |
| `SectionSettings` | Change page size, margins, and orientation |
| `Table` / `TableCell` | Update table text and append or merge rows |
| `HeaderFooterBlock` | Replace header or footer text |
| `Field` | Read or update hyperlink, mail merge, calculation, and cross-reference fields |
| `Picture` | Replace or inspect embedded binary image data |
| `Note` | Update footnotes and endnotes |
| `ShapeObject` | Read or edit text-bearing shapes |

## Example Workflows

### Open, inspect, and validate

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("input.hwpx")

print(doc.metadata())
print(doc.get_document_text())
print(doc.validation_errors())
print(doc.reference_validation_errors())
```

### Edit metadata and body text

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("input.hwpx")
doc.set_metadata(title="Edited title", creator="Docs Team", keyword="example")
doc.replace_text("draft", "final")
doc.append_paragraph("Appended paragraph", section_index=0)
doc.save("build/edited.hwpx")
```

### Change page settings and table contents

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

### Create hyperlinks, bookmarks, and calculated fields

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.blank()
bookmark = doc.append_bookmark("summary_anchor")
doc.append_hyperlink("https://example.com", display_text="Example")
doc.append_calculation_field("40+2", display_text="42")
doc.append_cross_reference(bookmark.name or "summary_anchor", display_text="Jump to summary")
doc.save("build/fields.hwpx")
```

## Repository Layout

```text
src/jakal_hwpx/          Python package
examples/samples/hwpx/   sample HWPX documents
examples/samples/hwp/    sample HWP documents
examples/output_smoke/   committed smoke corpus used by tests
examples/output/         generated showcase outputs
scripts/                 maintainer scripts
tools/                   bundled HWP conversion tooling
```

## Testing

Run the default test suite:

```bash
python -m pip install -e .[dev]
python -m pytest -q
```

The tests first look for samples in:

1. `JAKAL_HWPX_SAMPLE_DIR`
2. `all_hwpx_flat/`
3. `examples/output_smoke/`
4. `examples/output/`
5. `examples/samples/hwpx/`

## Sample Files

Repository sample inputs now live under `examples/samples/` instead of the repo root:

- `examples/samples/hwpx/`
- `examples/samples/hwp/`

Generated validation outputs belong under `build/validation/`.

## More Docs

- [HWPX_MODULE.md](./HWPX_MODULE.md): package structure, module roles, and API usage
- [examples/SHOWCASE.md](./examples/SHOWCASE.md): showcase generation workflow
- [RELEASING.md](./RELEASING.md): release checklist
- [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md): scope notes for sample files, bundled tools, and HWPX-related naming

## License

Original project-authored source code in this repository is available under the [MIT License](./LICENSE).

Sample documents, committed outputs, and bundled toolchain artifacts under `tools/` can be subject to separate rights or upstream licenses. See [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md) before redistributing those assets.
