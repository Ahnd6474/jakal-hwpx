# `jakal_hwpx` Module Notes

This document describes the Python package under `src/jakal_hwpx/`.

## Main Entry Point

### `HwpxDocument`

`HwpxDocument` is the primary public API.

It supports:

- opening an existing `.hwpx`
- creating a blank document
- reading and updating package metadata
- editing paragraph text
- appending paragraphs
- adding or replacing package parts
- validating package structure and references
- saving and round-tripping the package

Example:

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("example.hwpx")
doc.replace_text("before", "after")
doc.save("example-edited.hwpx")
```

### `HwpxDocument` API Reference

| Method | Returns | What it is for |
| --- | --- | --- |
| `open(path)` | `HwpxDocument` | Open an existing HWPX package |
| `blank()` | `HwpxDocument` | Build a new blank package in memory |
| `metadata()` | `DocumentMetadata` | Read title, creator, subject, keywords, and related metadata |
| `set_metadata(**values)` | `None` | Update package metadata fields |
| `get_document_text()` | `str` | Extract combined section text |
| `set_paragraph_text(section_index, paragraph_index, text)` | wrapper for the edited paragraph | Replace one paragraph while preserving surrounding structure |
| `append_paragraph(text, section_index=0)` | paragraph wrapper | Add a new paragraph to a section |
| `replace_text(old, new, count=-1)` | `int` | Replace text across sections and, by default, headers |
| `add_section(text=...)` | `SectionPart` | Append a new section to the document |
| `section_settings(index)` | `SectionSettings` | Access page size and margin settings |
| `tables()`, `pictures()`, `notes()`, `bookmarks()`, `fields()` | list of wrappers | Traverse rich content in the document |
| `styles()`, `paragraph_styles()`, `character_styles()` | list of wrappers | Inspect or update style definitions |
| `apply_style_to_paragraph(...)` | `None` | Bind a style to a single paragraph |
| `apply_style_batch(...)` | `int` | Apply styles to paragraphs matching a text filter |
| `append_bookmark(...)`, `append_hyperlink(...)`, `append_mail_merge_field(...)`, `append_calculation_field(...)`, `append_cross_reference(...)` | wrapper object | Create richer field-like content |
| `add_or_replace_binary(...)` | `BinaryDataPart` | Add or replace an embedded package binary |
| `compile(validate=True)` | `bytes` | Serialize the package in memory |
| `save(path, validate=True)` | `Path` | Save the package to disk |
| `validation_errors()` | `list[str]` | Check package-level validity |
| `xml_validation_errors()` | `list[str]` | Check XML roots and structural XML expectations |
| `reference_validation_errors()` | `list[str]` | Check references between styles, fields, bookmarks, and manifest items |
| `save_reopen_validation_errors()` | `list[str]` | Save and reopen as a practical round-trip check |

## Module Overview

The package is small enough that most work starts in `document.py`, then moves into wrapper types from `elements.py` if you need more control.

| Module | What it does | When to use it |
| --- | --- | --- |
| `document.py` | High-level document container and editing API | Start here for almost every workflow |
| `elements.py` | Wrappers for tables, pictures, notes, fields, shapes, and styles | Use when `HwpxDocument` returns a richer object you want to inspect or mutate |
| `parts.py` | Low-level package part classes for XML, text, binary, preview, and manifest content | Use when you need direct access to package internals |
| `xmlnode.py` | Small XML helper abstraction used across the package | Mostly internal, but useful when extending behavior |
| `namespaces.py` | Namespace map, QName helpers, and section matching constants | Useful when writing custom XPath or XML manipulation |
| `exceptions.py` | Package-specific error types | Catch these when handling invalid or malformed HWPX files |

## Package Layout

### `document.py`

Core document container and editing logic.

Key responsibilities:

- open, build, compile, and save HWPX zip packages
- expose high-level editing helpers
- preserve and update internal package parts
- perform structural validation

Common entry points:

- `HwpxDocument.open()` and `HwpxDocument.blank()`
- `set_metadata()`, `metadata()`, `get_document_text()`
- `set_paragraph_text()`, `append_paragraph()`, `replace_text()`
- `section_settings()`, `tables()`, `pictures()`, `notes()`, `fields()`
- `validation_errors()`, `reference_validation_errors()`, `save_reopen_validation_errors()`
- `compile()` and `save()`

### `parts.py`

Low-level package part model.

Key responsibilities:

- represent content parts such as `header.xml`, `section*.xml`, and `content.hpf`
- distinguish XML, text, binary, and preview parts
- expose metadata helpers for `content.hpf`

### `elements.py`

Wrappers over frequently edited XML structures.

Examples:

- paragraph and character styles
- tables and table cells
- pictures
- notes
- equations
- bookmarks and fields
- headers and footers

This is where most feature-specific editing lives. For example:

- `Table.set_cell_text()` and `Table.append_row()`
- `HeaderFooterBlock.set_text()`
- `Field.set_display_text()` and `Field.set_hyperlink_target()`
- `SectionSettings.set_page_size()` and `SectionSettings.set_margins()`
- `CharacterStyle.set_text_color()` and `ParagraphStyle.set_alignment()`

### Common wrapper types

| Type | Key methods / properties | Notes |
| --- | --- | --- |
| `HeaderFooterBlock` | `text`, `set_text()`, `replace_text()` | Returned by `headers()` and `footers()` |
| `Table` | `row_count`, `column_count`, `cells()`, `cell()`, `set_cell_text()`, `append_row()`, `merge_cells()` | Main entry point for table editing |
| `TableCell` | `row`, `column`, `text`, `row_span`, `col_span`, `set_text()` | Returned from `Table` helpers |
| `Picture` | `binary_item_id`, `shape_comment`, `binary_data()`, `replace_binary()` | Lets you inspect or replace embedded image data |
| `SectionSettings` | `page_width`, `page_height`, `landscape`, `margins()`, `set_page_size()`, `set_margins()` | Page layout per section |
| `StyleDefinition` | `style_id`, `name`, `set_name()`, `bind_refs()` | Top-level style object |
| `ParagraphStyle` | `alignment_horizontal`, `line_spacing`, `set_alignment()`, `set_line_spacing()` | Paragraph-level formatting |
| `CharacterStyle` | `text_color`, `height`, `set_text_color()`, `set_height()` | Character-level formatting |
| `Note` | `kind`, `number`, `text`, `set_text()` | Footnote and endnote wrapper |
| `Bookmark` | `name`, `rename()` | Bookmark editing |
| `Field` | `field_type`, `field_id`, `parameter_map()`, `set_parameter()`, `set_display_text()` | Covers hyperlink, mail merge, calculation, and cross-reference fields |
| `Equation` | `script`, `shape_comment` | Exposes equation script text |
| `ShapeObject` | `kind`, `shape_comment`, `text`, `set_text()` | Covers text art and other text-bearing shapes |

### `xmlnode.py`

Small XML convenience wrapper used by other modules.

### `namespaces.py`

Namespace constants, QName helpers, and section path matching.

### `exceptions.py`

Custom exception types for invalid packages and validation failures.

## Validation Surfaces

The package exposes several complementary validation helpers:

- `validation_errors()`
- `xml_validation_errors()`
- `reference_validation_errors()`
- `save_reopen_validation_errors()`
- `roundtrip_validate()`

These checks catch different classes of failure: malformed XML, broken references, missing manifest entries, and packages that cannot be reopened after save.

In practice:

- `validation_errors()` catches package-level structure problems
- `xml_validation_errors()` catches malformed or unexpected XML roots
- `reference_validation_errors()` catches broken style, field, bookmark, and manifest references
- `save_reopen_validation_errors()` is a pragmatic round-trip smoke test

## Public Types Re-Exported from `jakal_hwpx`

The package root re-exports:

- `HwpxDocument`
- `DocumentMetadata`
- `HwpxPart`
- `XmlPart`
- `SectionPart`
- `HeaderPart`
- `ContentHpfPart`
- `SettingsPart`
- `VersionPart`
- `MimetypePart`
- `ContainerPart`
- `ContainerRdfPart`
- `ManifestPart`
- `BinaryDataPart`
- `GenericBinaryPart`
- `GenericTextPart`
- `GenericXmlPart`
- `PreviewImagePart`
- `PreviewTextPart`
- `ScriptPart`
- `Picture`
- `Table`
- `TableCell`
- `Note`
- `Equation`
- `Bookmark`
- `Field`
- `AutoNumber`
- `HeaderFooterBlock`
- `SectionSettings`
- `StyleDefinition`
- `ParagraphStyle`
- `CharacterStyle`
- `ShapeObject`
- `HwpxError`
- `HwpxValidationError`
- `InvalidHwpxFileError`

### Error types

| Type | When you will see it |
| --- | --- |
| `InvalidHwpxFileError` | Input file is not a valid zip-based HWPX package |
| `HwpxValidationError` | Validation failed while opening, compiling, or saving |
| `HwpxError` | Base error type for package-specific failures |

## Round-Trip and Editing Notes

### Safe Editing Model

The package is designed around preserving package structure as much as possible:

- existing parts are loaded and kept in memory
- unknown parts are preserved by type inference
- zip entry metadata is cloned on compile when possible
- the document can be reopened and validated after save

### Distribution-Protected Packages

If `content.hpf` marks the package as distribution-protected, editable section XML may not be present.

In that case:

- the package can still be opened
- structural validation still works
- high-level editing APIs can be unavailable

### Why the Package Uses Element Wrappers

HWPX editing often requires working across:

- document metadata
- package manifests
- section XML
- style tables
- embedded binaries

The wrappers in `elements.py` let the rest of the code expose domain-specific operations without pushing raw XPath and XML bookkeeping into every call site.

## Typical Usage Patterns

### Inspect and validate a document

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("example.hwpx")

print(doc.metadata())
print(doc.validation_errors())
print(doc.reference_validation_errors())
```

### Update text and save

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("example.hwpx")
doc.replace_text("draft", "final")
doc.save("example-final.hwpx")
```

### Generate a blank document

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.blank()
doc.set_metadata(title="Generated")
doc.set_paragraph_text(0, 0, "Hello")
doc.save("build/blank.hwpx")
```

### Edit page settings and styles

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("example.hwpx")

settings = doc.section_settings(0)
settings.set_page_size(width=60000, height=85000)
settings.set_margins(left=7000, right=7000)

style = doc.styles()[0]
para_style = doc.paragraph_styles()[0]
char_style = doc.character_styles()[0]

style.set_name("Body Center")
para_style.set_alignment(horizontal="CENTER")
char_style.set_text_color("#112233")

doc.apply_style_to_paragraph(
    0,
    0,
    style_id=style.style_id,
    para_pr_id=para_style.style_id,
    char_pr_id=char_style.style_id,
)
doc.save("build/styled.hwpx")
```

### Work with tables, headers, and footers

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("example.hwpx")

if doc.headers():
    doc.headers()[0].set_text("Edited header")
if doc.footers():
    doc.footers()[0].set_text("Edited footer")

table = doc.tables()[0]
table.set_cell_text(0, 0, "Updated value")
table.append_row()[0].set_text("Appended row")

doc.save("build/structured-edit.hwpx")
```

### Create bookmarks and dynamic fields

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.blank()

bookmark = doc.append_bookmark("summary_anchor")
doc.append_hyperlink("https://example.com", display_text="Open Example")
doc.append_mail_merge_field("customer_name", display_text="CUSTOMER_NAME")
doc.append_calculation_field("40+2", display_text="42")
doc.append_cross_reference(bookmark.name or "summary_anchor", display_text="Go to summary")

assert doc.reference_validation_errors() == []
doc.save("build/fields.hwpx")
```

### Add binary data and keep package validation intact

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.blank()
doc.add_or_replace_binary(
    "assets/custom.bin",
    b"abc123",
    media_type="application/octet-stream",
    manifest_id="custom_asset",
)

assert doc.validation_errors() == []
doc.save("build/with-binary.hwpx")
```

## Choosing an API Level

Use the highest-level API that fits the job.

- Start with `HwpxDocument` when you need to open, edit, validate, or save a document.
- Use wrapper objects from `elements.py` when the document contains tables, fields, notes, pictures, or style objects you want to modify in place.
- Drop to `parts.py` only when you need direct package access, such as custom binaries, preview parts, or raw part inspection.

## Related Repo Assets

This document is about the Python package itself. The repository also includes:

- `examples/samples/`
- `examples/output_smoke/`
- `examples/output/`
- `tools/`

The Java-based `.hwp -> .hwpx` conversion helpers live under `tools/` and are not part of the importable Python package.
