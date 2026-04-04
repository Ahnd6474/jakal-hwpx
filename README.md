# jakal-hwpx

Round-trip-safe HWPX reading, editing, validation, and writing for Python.

This repository is centered on the `jakal_hwpx` package under `src/jakal_hwpx`. It also keeps a local HWPX corpus, showcase generators, validation tests, and Windows tooling used to check that edits still behave the way Hancom expects.

## Table of contents

- [Installation](#installation)
- [Quick start](#quick-start)
- [What is jakal-hwpx?](#what-is-jakal-hwpx)
- [Why this repo exists](#why-this-repo-exists)
- [API](#api)
  - [`HwpxDocument`](#hwpxdocument)
  - [`DocumentMetadata`](#documentmetadata)
  - [Element wrappers](#element-wrappers)
  - [Part classes](#part-classes)
  - [`HwpxXmlNode`](#hwpxxmlnode)
  - [Exceptions](#exceptions)
- [Examples](#examples)
- [Further reading](#further-reading)
- [License](#license)
- [Contributing](#contributing)

## Installation

You need Python 3.11 or newer.

Install the package from a local checkout:

```bash
python -m venv .venv
python -m pip install --upgrade pip
python -m pip install -e .
```

The import path is `jakal_hwpx`. The project name in `pyproject.toml` is `jakal-hwpx`.

`lxml` is the only runtime dependency right now, and it is installed automatically by `pip install -e .`.

If you want to run the repo's tests as well:

```bash
python -m pip install pytest
```

If you want to run the COM-based Hancom model verification tests on Windows, install `pywin32` too:

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

`jakal_hwpx` is a Python library for working with zip-based HWPX packages. It opens an existing document, exposes typed wrappers for common editable structures, and writes the package back out without rebuilding every file from scratch.

The package covers document metadata, paragraphs, styles, tables, images, headers and footers, notes, bookmarks, fields, equations, shapes, binary attachments, and low-level XML access. The repo around it adds a real sample corpus in `all_hwpx_flat`, showcase generation in `examples/`, and validation tests that compare saved output against actual Hancom behavior.

## Why this repo exists

HWPX files are not just "some XML in a zip". They are a package of linked XML parts, manifest entries, binary assets, and references that are easy to break with unzip-edit-rezip scripts.

This repo exists to make those edits safer:

- Untouched parts stay as their original bytes unless you modify them.
- The library validates package structure, manifest references, style references, and field links before writing.
- The tests do more than parse XML. They reopen saved files, and on Windows they can check real Hancom load behavior as well.
- Distribution-protected packages can still be preserved and re-saved, even when their encrypted section XML is not editable through the high-level API.

Use this if you need scripted edits to existing HWPX documents. If all you want is raw XML traversal, `lxml` alone may be enough. If you need HWP to HWPX conversion, the PowerShell scripts and bundled Java tooling in this repo are the relevant pieces, not the `jakal_hwpx` Python API.

## API

### `HwpxDocument`

`HwpxDocument` is the main entry point. It loads a package, exposes typed helpers for common structures, and handles validation plus save/compile workflows.

#### Constructors

| Signature | Returns | Description |
|-----------|---------|-------------|
| `HwpxDocument.open(path)` | `HwpxDocument` | Open a `.hwpx` file from disk. |
| `HwpxDocument.from_bytes(raw_bytes)` | `HwpxDocument` | Open a package from in-memory bytes. |

`HwpxDocument.open()` raises the built-in `FileNotFoundError` when the path does not exist.

#### Core properties and part access

| Member | Returns | Description |
|--------|---------|-------------|
| `mimetype` | `MimetypePart` | The raw `mimetype` package entry. |
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
| `metadata()` | `DocumentMetadata` | Read document metadata from `content.hpf`. |
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
| `auto_numbers(section_index=None)` | `list[AutoNumber]` | Automatic numbering controls such as `hp:newNum`. |
| `equations(section_index=None)` | `list[Equation]` | Equation objects. |
| `shapes(section_index=None)` | `list[ShapeObject]` | Generic shape objects, including textart. |
| `styles()` | `list[StyleDefinition]` | Document styles from `header.xml`. |
| `paragraph_styles()` | `list[ParagraphStyle]` | Paragraph style records. |
| `character_styles()` | `list[CharacterStyle]` | Character style records. |
| `get_style(style_id)` | `StyleDefinition` | Look up a style by id. |
| `get_paragraph_style(style_id)` | `ParagraphStyle` | Look up a paragraph style by id. |
| `get_character_style(style_id)` | `CharacterStyle` | Look up a character style by id. |

#### Text, structure, and package editing

| Method | Returns | Description |
|--------|---------|-------------|
| `set_metadata(**values)` | `None` | Update title, language, creator, subject, description, dates, keyword, and similar metadata. |
| `get_document_text(section_separator="\n\n")` | `str` | Flatten section text for inspection. |
| `replace_text(old, new, count=-1, include_header=True)` | `int` | Replace text across editable XML parts. |
| `append_paragraph(text, section_index=0, ...)` | `HwpxXmlNode` | Add a paragraph to a section. |
| `insert_paragraph(section_index, paragraph_index, text, ...)` | `HwpxXmlNode` | Insert a paragraph at a specific paragraph index in a section. |
| `set_paragraph_text(section_index, paragraph_index, text, ...)` | `HwpxXmlNode` | Replace the text of one paragraph. |
| `delete_paragraph(section_index, paragraph_index)` | `None` | Remove a paragraph. |
| `apply_style_to_paragraph(section_index, paragraph_index, ...)` | `None` | Attach style, paragraph style, or character style ids to one paragraph. |
| `apply_style_batch(section_index, text_contains, ...)` | `int` | Apply style ids to paragraphs whose text matches a substring. |
| `add_section(clone_from=0, text=None)` | `SectionPart` | Create a new section, optionally cloning from an existing one. |
| `remove_section(section_index)` | `None` | Delete a section and update package metadata. |
| `set_preview_text(text)` | `PreviewTextPart` | Create or replace preview text. |
| `add_or_replace_binary(name, data, media_type=None, manifest_id=None)` | `BinaryDataPart` | Store binary data in the package and update the manifest. |
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
| `reference_validation_errors()` | `list[str]` | Style id, manifest id, bookmark, and field reference checks. |
| `save_reopen_validation_errors()` | `list[str]` | Save, reopen, and report any structural errors. |
| `discover_hancom_executable()` | `Path \| None` | Find `Hanword.exe`, `Hword.exe`, or `Hwp.exe` from `PATH`, `HWPX_HANCOM_EXE`, or common install folders. |
| `hancom_open_validation_errors(executable_path=None, timeout_seconds=15)` | `list[str]` | Save the document to a temp file, try to open it in Hancom, and report launch/open failures. |
| `open_in_hancom(executable_path=None, timeout_seconds=15)` | `None` | Open a temp copy in Hancom and raise if launch/open fails. |

If `is_distribution_protected` is `True`, high-level section editing is intentionally blocked because the section and header data are encrypted. The package can still be preserved and re-saved.

### `DocumentMetadata`

`DocumentMetadata` is a simple dataclass returned by `HwpxDocument.metadata()`.

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

These wrappers sit on top of XML nodes and give you safer, feature-level edits.

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
| `Field` | Hyperlinks, mail merge fields, formulas, and cross references | `field_type`, `field_id`, `control_id`, `name`, `parameter_map()`, `get_parameter()`, `set_parameter()`, `display_text`, `set_display_text()`, `set_hyperlink_target()`, `configure_mail_merge()`, `configure_calculation()`, `configure_cross_reference()` |
| `AutoNumber` | Automatic numbering controls | `kind`, `number`, `number_type`, `set_number()`, `set_number_type()` |
| `Equation` | Equation script editing | `script`, `shape_comment` |
| `ShapeObject` | Shape comments and textart text | `kind`, `shape_comment`, `text`, `set_text()` |

### Part classes

The package also exports the low-level part model used internally.

| Class | Represents | Notes |
|-------|------------|-------|
| `HwpxPart` | Base class for every package part | Holds `path`, `raw_bytes`, `modified`, and `to_bytes()`. |
| `GenericBinaryPart` | Arbitrary binary payload | Raw bytes via `.data`. |
| `MimetypePart` | `mimetype` | Exposes `.mime`. |
| `GenericTextPart` | Arbitrary UTF-8 text part | Exposes `.text`. |
| `PreviewTextPart` | `Preview/PrvText.txt` | Subclass of `GenericTextPart`. |
| `PreviewImagePart` | Preview image binary | Subclass of `GenericBinaryPart`. |
| `BinaryDataPart` | `BinData/*` entries | Subclass of `GenericBinaryPart`. |
| `ScriptPart` | Script text part | Subclass of `GenericTextPart`. |
| `XmlPart` | Base class for XML-backed parts | Gives you `root`, `root_element`, `xpath()`, `find()`, `findall()`, `extract_hp_text()`, and `replace_hp_text()`. |
| `GenericXmlPart` | Arbitrary XML part | Subclass of `XmlPart`. |
| `VersionPart` | `version.xml` | Version XML wrapper. |
| `ContainerPart` | `META-INF/container.xml` | Includes `rootfile_paths()` and `ensure_rootfile()`. |
| `ManifestPart` | `META-INF/manifest.xml` | XML wrapper for package manifest data. |
| `ContainerRdfPart` | `META-INF/container.rdf` | XML wrapper for RDF container data. |
| `ContentHpfPart` | `Contents/content.hpf` | Metadata, manifest, and spine helpers such as `metadata()`, `set_metadata()`, `manifest_items()`, `ensure_manifest_item()`, and `ensure_spine_itemref()`. |
| `HeaderPart` | `Contents/header.xml` | Includes `section_count` and `set_section_count()`. |
| `SettingsPart` | `settings.xml` | XML wrapper for application settings. |
| `SectionPart` | `Contents/sectionN.xml` | Paragraph and text helpers such as `section_index()`, `paragraphs()`, `text_fragments()`, `extract_text()`, `append_paragraph()`, `insert_paragraph()`, `set_paragraph_text()`, and `delete_paragraph()`. |

### `HwpxXmlNode`

`HwpxXmlNode` is the low-level escape hatch for namespace-aware XML editing without dropping all the way down to raw `lxml` elements.

Use it when the typed wrappers do not cover the element you need yet.

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

`HwpxXmlNode` exposes attribute helpers, XPath helpers, child insertion, cloning, removal, and XML serialization through methods such as `get_attr()`, `set_attr()`, `remove_attr()`, `xpath()`, `find()`, `findall()`, `append_child()`, `insert_child()`, `ensure_child()`, `remove()`, `clear()`, `clone()`, `add_existing_child()`, and `to_xml()`.

### Exceptions

| Exception | Raised when |
|-----------|-------------|
| `HwpxError` | Base package exception. |
| `InvalidHwpxFileError` | The input exists but is not a valid zip-based HWPX package. |
| `HwpxValidationError` | `validate()` finds one or more structural problems. |

## Examples

### Run the test suite

The tests use the checked-in HWPX corpus under `all_hwpx_flat` and skip invalid zero-byte files automatically.

```bash
python -m pytest -q
```

### Build the showcase bundle

This generates example output documents under `examples/output/` and writes a manifest plus a Markdown report.

```bash
python examples/build_showcase_bundle.py
```

If you want faster generation without launching Hancom:

```bash
python examples/build_showcase_bundle.py --skip-hancom
```

See `examples/SHOWCASE.md` for the generated file list and what each one demonstrates.

### Run Hancom open validation on Windows

`HwpxDocument.open_in_hancom()` and `hancom_open_validation_errors()` can launch a local Hancom installation and verify that a saved temp file still opens.

In PowerShell:

```powershell
$env:HWPX_HANCOM_EXE = "C:\Program Files\Hancom\Office\HOffice130\Bin\Hwp.exe"
python -m pytest -q tests/test_hancom_model_verification.py
```

The repo also includes `tests/test_hancom_model_verification.py`, which goes further than a plain launch test. It uses `win32com.client` to export the loaded document as `HWPML2X` and check where edited content actually landed in Hancom's document model.

### Use the repo's supporting tooling

The Python package is the main product here, but the repo also carries the support files used to build and validate against a real corpus:

- `scripts/flatten-hwpx.ps1` collects converted files into `all_hwpx_flat`.
- `scripts/build-hwp-converter.ps1` rebuilds `tools/hwp-batch-converter.jar`.
- `scripts/process-hwp-chunk.ps1` and `scripts/collect-hwp-docs.ps1` are Windows-side helpers for batch conversion work.
- `vendor/hwp2hwpx-src/` keeps the upstream Java conversion source used by the build script.

## Further reading

- [`HWPX_MODULE.md`](./HWPX_MODULE.md) for a shorter module-level overview.
- [`examples/SHOWCASE.md`](./examples/SHOWCASE.md) for showcase generation details.
- [`vendor/hwp2hwpx-src/README.md`](./vendor/hwp2hwpx-src/README.md) for the vendored Java converter source.

## License

This repository does not currently include a top-level license file. If you plan to publish or redistribute it, add one first.

## Contributing

Before you change the API, run the checks that exercise the current contract:

```bash
python -m pytest -q
python examples/build_showcase_bundle.py --skip-hancom
```

If you are changing Hancom-facing behavior on Windows, run the extra verification path too:

```bash
python -m pytest -q tests/test_hancom_model_verification.py
```

The sample corpus is large and mixed quality. The fixtures only use files that are valid zip-based `.hwpx` packages, so do not be surprised when `all_hwpx_flat` also contains zero-byte placeholders or failed conversions.
