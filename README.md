# jakal-hwpx

Round-trip-safe HWPX reading, editing, validation, and writing for Python.

Korean version: [README.ko.md](./README.ko.md)

`jakal_hwpx` is the Python package in `src/jakal_hwpx`. This README focuses on the library you import and use in code. Some broader repository workflows depend on a local HWPX corpus and Windows-only tooling that are not part of a normal checkout, so they are treated as optional maintainer workflows here instead of the main path.

## Table of contents

- [Installation](#installation)
- [Quick start](#quick-start)
- [What is jakal-hwpx?](#what-is-jakal-hwpx)
- [Why use it?](#why-use-it)
- [API](#api)
  - [`HwpxDocument`](#hwpxdocument)
  - [`DocumentMetadata`](#documentmetadata)
  - [Element wrappers](#element-wrappers)
  - [Part classes](#part-classes)
  - [`HwpxXmlNode`](#hwpxxmlnode)
  - [Exceptions](#exceptions)
- [Examples](#examples)
- [Maintainer workflows](#maintainer-workflows)
- [Further reading](#further-reading)
- [License](#license)

## Installation

You need Python 3.11 or newer.

Install the package from a local checkout:

```bash
python -m pip install --upgrade pip
python -m pip install -e .
```

The import path is `jakal_hwpx`. The project name in `pyproject.toml` is `jakal-hwpx`.

If you want to run the test suite as well:

```bash
python -m pip install pytest
```

If you want to use the Windows COM-based Hancom verification helpers, install `pywin32` too:

```bash
python -m pip install pywin32
```

## Quick start

```python
from pathlib import Path

from jakal_hwpx import HwpxDocument

source = Path("example.hwpx")
target = Path("example-edited.hwpx")

doc = HwpxDocument.open(source)
doc.set_metadata(title="Edited title", creator="jakal-hwpx")
doc.replace_text("old text", "new text", count=1)
doc.append_paragraph("Added from Python.", section_index=0)

if doc.headers():
    doc.headers()[0].set_text("Edited header")

doc.save(target)
print(doc.validation_errors())
# []
```

## What is jakal-hwpx?

`jakal_hwpx` is a Python library for editing zip-based HWPX packages without tearing them apart by hand. It opens an existing document, exposes typed wrappers for the structures you usually want to change, and writes the package back out while preserving untouched parts as original bytes.

The package covers metadata, paragraphs, styles, page settings, tables, images, headers and footers, notes, bookmarks, fields, equations, shapes, binary assets, and low-level XML access.

## Why use it?

HWPX is not just one XML file. A real document is a package of XML parts, manifest entries, binary assets, references, and layout data that all have to stay in sync.

`jakal_hwpx` tries to make those edits safer:

- It keeps untouched parts intact instead of regenerating everything.
- It gives you typed helpers for common edits such as paragraph text, header/footer text, tables, fields, and styles.
- It validates package structure and cross-part references before writing.
- It still leaves an escape hatch through `HwpxXmlNode` when you need to edit something the typed wrappers do not cover yet.
- It can preserve distribution-protected packages even when high-level editing is unavailable for encrypted sections.

If you only need raw XML access, `lxml` may be enough. If you need programmatic edits to existing HWPX files, this library gives you a better starting point.

## API

### `HwpxDocument`

`HwpxDocument` is the main entry point.

#### Constructors

| Signature | Returns | Description |
|-----------|---------|-------------|
| `HwpxDocument.open(path)` | `HwpxDocument` | Open a `.hwpx` file from disk. |
| `HwpxDocument.from_bytes(raw_bytes)` | `HwpxDocument` | Open a package from in-memory bytes. |

`HwpxDocument.open()` raises `FileNotFoundError` when the path does not exist.

#### Core properties and part access

| Member | Returns | Description |
|--------|---------|-------------|
| `mimetype` | `MimetypePart` | The raw `mimetype` entry. |
| `content_hpf` | `ContentHpfPart` | `Contents/content.hpf`, including metadata and manifest helpers. |
| `header` | `HeaderPart` | `Contents/header.xml`. |
| `container` | `ContainerPart` | `META-INF/container.xml`. |
| `preview_text` | `PreviewTextPart \| None` | `Preview/PrvText.txt` if present. |
| `is_distribution_protected` | `bool` | `True` when the package uses protected encrypted parts. |
| `sections` | `list[SectionPart]` | Sorted section XML parts. |
| `get_part(path, expected_type=None)` | `HwpxPart` | Fetch any package part by path. |
| `list_part_paths()` | `list[str]` | Current part order used when writing the zip. |
| `add_part(path, raw_bytes)` | `HwpxPart` | Add or replace a package part directly. |
| `remove_part(path)` | `None` | Remove a package part. |

#### Feature accessors

| Method | Returns | Description |
|--------|---------|-------------|
| `metadata()` | `DocumentMetadata` | Read document metadata. |
| `headers(section_index=None)` | `list[HeaderFooterBlock]` | Header blocks across the document or in one section. |
| `footers(section_index=None)` | `list[HeaderFooterBlock]` | Footer blocks across the document or in one section. |
| `tables(section_index=None)` | `list[Table]` | Tables in the document. |
| `pictures(section_index=None)` | `list[Picture]` | Picture objects in the document. |
| `section_settings(section_index=0)` | `SectionSettings` | Page size and margin settings for a section. |
| `notes(section_index=None)` | `list[Note]` | Footnotes and endnotes. |
| `bookmarks(section_index=None)` | `list[Bookmark]` | Bookmark controls. |
| `fields(section_index=None)` | `list[Field]` | All field controls. |
| `hyperlinks(section_index=None)` | `list[Field]` | Hyperlink fields only. |
| `mail_merge_fields(section_index=None)` | `list[Field]` | Mail merge fields only. |
| `calculation_fields(section_index=None)` | `list[Field]` | Formula and calculation fields only. |
| `cross_references(section_index=None)` | `list[Field]` | Cross-reference fields only. |
| `auto_numbers(section_index=None)` | `list[AutoNumber]` | Automatic numbering controls. |
| `equations(section_index=None)` | `list[Equation]` | Equation objects. |
| `shapes(section_index=None)` | `list[ShapeObject]` | Generic shape objects, including textart. |
| `styles()` | `list[StyleDefinition]` | Document styles from `header.xml`. |
| `paragraph_styles()` | `list[ParagraphStyle]` | Paragraph style records. |
| `character_styles()` | `list[CharacterStyle]` | Character style records. |
| `get_style(style_id)` | `StyleDefinition` | Look up a style by id. |
| `get_paragraph_style(style_id)` | `ParagraphStyle` | Look up a paragraph style by id. |
| `get_character_style(style_id)` | `CharacterStyle` | Look up a character style by id. |

#### Editing and output

| Method | Returns | Description |
|--------|---------|-------------|
| `set_metadata(**values)` | `None` | Update title, language, creator, subject, description, dates, keyword, and related metadata. |
| `get_document_text(section_separator="\n\n")` | `str` | Flatten section text for inspection. |
| `replace_text(old, new, count=-1, include_header=True)` | `int` | Replace text across editable XML parts. |
| `append_paragraph(text, section_index=0, ...)` | `HwpxXmlNode` | Add a paragraph to a section. |
| `insert_paragraph(section_index, paragraph_index, text, ...)` | `HwpxXmlNode` | Insert a paragraph at a specific paragraph index. |
| `set_paragraph_text(section_index, paragraph_index, text, ...)` | `HwpxXmlNode` | Replace the text of one paragraph. |
| `delete_paragraph(section_index, paragraph_index)` | `None` | Remove a paragraph. |
| `apply_style_to_paragraph(section_index, paragraph_index, ...)` | `None` | Attach style ids to one paragraph. |
| `apply_style_batch(section_index=..., text_contains=..., regex=..., ...)` | `int` | Apply style ids to matching paragraphs. |
| `add_section(clone_from=0, text=None)` | `SectionPart` | Create a new section. |
| `remove_section(section_index)` | `None` | Delete a section and update package metadata. |
| `set_preview_text(text)` | `PreviewTextPart` | Create or replace preview text. |
| `add_or_replace_binary(name, data, media_type=None, manifest_id=None)` | `BinaryDataPart` | Store binary data and update the manifest. |
| `append_bookmark(name, ...)` | `Bookmark` | Add a bookmark control. |
| `append_field(field_type, ...)` | `Field` | Add a generic field control. |
| `append_hyperlink(target, display_text, ...)` | `Field` | Add a hyperlink field. |
| `append_mail_merge_field(name, display_text, ...)` | `Field` | Add a mail merge field. |
| `append_calculation_field(expression, display_text, ...)` | `Field` | Add a formula field. |
| `append_cross_reference(bookmark_name, display_text, ...)` | `Field` | Add a cross-reference field. |
| `compile(validate=True)` | `bytes` | Build the in-memory package into `.hwpx` bytes. |
| `save(path, validate=True)` | `Path` | Write the package to disk. |

#### Validation and Hancom verification

| Method | Returns | Description |
|--------|---------|-------------|
| `validation_errors()` | `list[str]` | Structural package checks for required parts, manifest entries, duplicate zip paths, and section counts. |
| `validate()` | `None` | Raise `HwpxValidationError` if `validation_errors()` is not empty. |
| `roundtrip_validate()` | `None` | Save to a temp directory, reopen, and validate again. |
| `xml_validation_errors()` | `list[str]` | XML root-name and section/table sanity checks. |
| `schema_validation_errors(schema_map)` | `list[str]` | Optional XSD validation for specific XML parts. |
| `reference_validation_errors()` | `list[str]` | Style, manifest, bookmark, and field reference checks. |
| `save_reopen_validation_errors()` | `list[str]` | Save, reopen, and report structural errors. |
| `discover_hancom_executable()` | `Path \| None` | Try to find an installed Hancom executable. |
| `hancom_open_validation_errors(executable_path=None, timeout_seconds=15)` | `list[str]` | Save a temp file, try to open it in Hancom, and report failures. |
| `open_in_hancom(executable_path=None, timeout_seconds=15)` | `None` | Open a temp copy in Hancom and raise if launch/open fails. |

If `is_distribution_protected` is `True`, high-level section editing is intentionally blocked because the encrypted section data is not editable through the public helpers.

### `DocumentMetadata`

`DocumentMetadata` is a small dataclass returned by `HwpxDocument.metadata()`.

```python
DocumentMetadata(
    title: str | None = None,
    language: str | None = None,
    creator: str | None = None,
    subject: str | None = None,
    description: str | None = None,
    lastsaveby: str | None = None,
    created: str | None = None,
    modified: str | None = None,
    date: str | None = None,
    keyword: str | None = None,
    extra: dict[str, str] = {},
)
```

### Element wrappers

These wrappers sit on top of XML nodes and cover the common editing paths.

| Class | Use for | Key members |
|-------|---------|-------------|
| `HeaderFooterBlock` | Header and footer text | `kind`, `text`, `replace_text()`, `set_text()` |
| `TableCell` | One table cell | `row`, `column`, `text`, `row_span`, `col_span`, `set_text()` |
| `Table` | Table editing | `row_count`, `column_count`, `cells()`, `rows()`, `cell()`, `set_cell_text()`, `append_row()`, `merge_cells()` |
| `Picture` | Image metadata and binary binding | `binary_item_id`, `shape_comment`, `binary_part_path()`, `binary_data()`, `replace_binary()`, `bind_binary_item()` |
| `StyleDefinition` | Named document styles | `style_id`, `name`, `english_name`, `para_pr_id`, `char_pr_id`, `set_name()`, `set_english_name()`, `bind_refs()` |
| `ParagraphStyle` | Paragraph formatting | `style_id`, `alignment_horizontal`, `line_spacing`, `set_alignment()`, `set_line_spacing()` |
| `CharacterStyle` | Character formatting | `style_id`, `text_color`, `height`, `set_text_color()`, `set_height()` |
| `SectionSettings` | Page setup | `page_width`, `page_height`, `landscape`, `margins()`, `set_page_size()`, `set_margins()` |
| `Note` | Footnotes and endnotes | `kind`, `number`, `text`, `set_text()` |
| `Bookmark` | Bookmark controls | `name`, `rename()` |
| `Field` | Hyperlinks, mail merge, formulas, and cross references | `field_type`, `field_id`, `control_id`, `name`, `parameter_map()`, `get_parameter()`, `set_parameter()`, `display_text`, `set_display_text()`, `set_hyperlink_target()`, `configure_mail_merge()`, `configure_calculation()`, `configure_cross_reference()` |
| `AutoNumber` | Automatic numbering controls | `kind`, `number`, `number_type`, `set_number()`, `set_number_type()` |
| `Equation` | Equation editing | `script`, `shape_comment` |
| `ShapeObject` | Shape comments and textart text | `kind`, `shape_comment`, `text`, `set_text()` |

### Part classes

The package also exports the lower-level part model.

| Class | Represents |
|-------|------------|
| `HwpxPart` | Base class for every package part |
| `GenericBinaryPart` | Arbitrary binary payload |
| `MimetypePart` | `mimetype` |
| `GenericTextPart` | Arbitrary UTF-8 text part |
| `PreviewTextPart` | `Preview/PrvText.txt` |
| `PreviewImagePart` | Preview image binary |
| `BinaryDataPart` | `BinData/*` entries |
| `ScriptPart` | Script text part |
| `XmlPart` | Base class for XML-backed parts |
| `GenericXmlPart` | Arbitrary XML part |
| `VersionPart` | `version.xml` |
| `ContainerPart` | `META-INF/container.xml` |
| `ManifestPart` | `META-INF/manifest.xml` |
| `ContainerRdfPart` | `META-INF/container.rdf` |
| `ContentHpfPart` | `Contents/content.hpf` |
| `HeaderPart` | `Contents/header.xml` |
| `SettingsPart` | `settings.xml` |
| `SectionPart` | `Contents/sectionN.xml` |

### `HwpxXmlNode`

`HwpxXmlNode` is the low-level escape hatch for namespace-aware XML editing.

Use it when the typed wrappers do not cover what you need yet.

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("example.hwpx")
section = doc.sections[0]
paragraph = section.paragraphs()[0]

for node in paragraph.xpath(".//hp:t"):
    if node.text:
        node.text = node.text.replace("before", "after")

doc.save("example-edited.hwpx")
```

### Exceptions

| Exception | Raised when |
|-----------|-------------|
| `HwpxError` | Base package exception. |
| `InvalidHwpxFileError` | The input exists but is not a valid zip-based HWPX package. |
| `HwpxValidationError` | `validate()` finds one or more structural problems. |

## Examples

### Feature-level editing

```python
doc = HwpxDocument.open("example.hwpx")

doc.headers()[0].set_text("Edited header")
doc.footers()[0].set_text("Edited footer")

section = doc.section_settings(0)
section.set_page_size(width=60000, height=84000)
section.set_margins(left=8500, right=8500)

table = doc.tables()[0]
table.set_cell_text(0, 0, "Edited cell")
table.append_row()

doc.append_bookmark("python_anchor")
doc.append_hyperlink("https://example.com", display_text="Example")
doc.append_mail_merge_field("customer_name", display_text="CUSTOMER_NAME")
doc.append_calculation_field("40+2", display_text="42")
doc.append_cross_reference("python_anchor", display_text="Anchor Ref")

doc.save("example-edited.hwpx")
```

### Validation helpers

```python
doc = HwpxDocument.open("example.hwpx")

print(doc.xml_validation_errors())
print(doc.reference_validation_errors())
print(doc.save_reopen_validation_errors())
```

If Hancom is installed on Windows, you can also use `open_in_hancom()` or `hancom_open_validation_errors()`. Set `HWPX_HANCOM_EXE` only when auto-discovery is not enough.

## Maintainer workflows

The repository includes broader validation and showcase scripts, but those are not required to use the library itself.

Two constraints matter:

- Most repository tests expect a local HWPX corpus that is not committed with the package sources.
- Hancom model verification is Windows-only and requires a local Hancom installation.

If you maintain such a setup, the showcase script can be pointed at your own paths explicitly:

```bash
python examples/build_showcase_bundle.py --corpus-dir <path-to-hwpx-corpus> --output-dir <path-to-output> --skip-hancom
```

For test execution, check `tests/conftest.py` first. The current suite assumes a maintainer-provided local corpus layout.

## Further reading

- [`HWPX_MODULE.md`](./HWPX_MODULE.md) for a shorter module overview.
- [`examples/SHOWCASE.md`](./examples/SHOWCASE.md) for showcase generation details.

## License

This repository does not currently include a top-level license file. Add one before publishing or redistributing the project.
