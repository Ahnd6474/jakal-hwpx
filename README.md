# jakal-hwpx

`jakal-hwpx` is a Python library for reading, editing, and writing `HWPX` and `HWP` documents,
with an optional beta addon for `DOCX` (Microsoft Word).

The package exposes a shared document API and format-specific implementations:

- `Document`: the shared document API (`HancomDocument` remains a compatibility alias)
- Beta DOCX addon `WordDocument(Document)`: retains Word details and dispatches shared edits to a Word backend
- `HwpxDocument`: direct HWPX package and XML editing
- `HwpDocument`: direct native HWP object and binary editing

![Module layers](https://raw.githubusercontent.com/Ahnd6474/jakal-hwpx/main/docs/images/module-layers.svg)

## Installation

Requires Python 3.11 or newer.

```bash
python -m pip install --upgrade pip
python -m pip install jakal-hwpx
```

For local development:

```bash
python -m pip install -e ".[dev]"
```

Package name on PyPI is `jakal-hwpx`. Import path is `jakal_hwpx`.

## Quick start

For most application code, start with `Document`.

```python
from jakal_hwpx import Document

doc = Document.blank()
doc.metadata.title = "Quarterly report"
doc.append_paragraph("Sales summary")
doc.append_table(
    rows=2,
    cols=2,
    cell_texts=[["Item", "Value"], ["Q1", "120"]],
)

doc.write_to_hwpx("build/report.hwpx")
doc.write_to_hwp("build/report.hwp")
```

Reading an existing document uses the same model:

```python
from jakal_hwpx import Document

doc = Document.open("input.hwp")
doc.append_paragraph("Review complete")
doc.save("build/output.hwpx")
```

### DOCX addon (beta)

Install the current beta and its DOCX dependencies:

```bash
python -m pip install --upgrade --pre "jakal-hwpx[docx]"
```

Or install this checkout from the repository root:

```bash
python -m pip install -e ".[docx]"
```

The bundled addon lives in `jakal_hwpx.addons.docx` and shares the core library's
release. The default HWP/HWPX installation does not require python-docx; the addon
and its backend load only when requested. DOCX is a beta API in `1.1.0b1`;
behavior and fidelity may change before the stable `1.1.0` release.

#### Create a Word document

This example needs no input file. Output directories are created automatically.

```python
from jakal_hwpx.addons.docx import WordDocument

doc = WordDocument.blank()
doc.append_paragraph("Quarterly report")
doc.append_paragraph("Status: draft")
doc.append_table(
    rows=2,
    cols=2,
    cell_texts=[["Item", "Value"], ["Q1", "120"]],
)
doc.save("build/docx-example/original.docx")
```

#### Read, edit, and preserve Word details

Run this after the creation example, or replace the path with an existing DOCX.
`Document.open()` selects `WordDocument` automatically. An untouched DOCX save
preserves the original file byte for byte.

```python
from jakal_hwpx import Document
from jakal_hwpx.addons.docx import WordDocument
from docx.shared import Pt

doc = Document.open("build/docx-example/original.docx")
assert isinstance(doc, WordDocument)
doc.save("build/docx-example/untouched-copy.docx")

count = doc.replace_text("Status: draft", "Status: reviewed")
doc.append_paragraph("Review complete")

with doc.edit_word() as word:
    # `word` is the python-docx backend; changes refresh the common model on exit.
    for run in word.paragraphs[0].runs:
        run.bold = True
        run.font.size = Pt(18)

# Save after leaving edit_word() so native formatting changes are included.
doc.save("build/docx-example/edited.docx")
print(f"Replaced {count} occurrence(s)")
```

#### Convert between DOCX, HWPX, and HWP

Run this after the editing example to convert DOCX to HWPX and back:

```python
from jakal_hwpx import Document

Document.open("build/docx-example/edited.docx").save("build/docx-example/converted.hwpx")
Document.open("build/docx-example/converted.hwpx").save("build/docx-example/restored.docx")
```

For your own documents, choose the output format by its extension:

```python
from jakal_hwpx import Document

Document.open("input.docx").save("build/from_word.hwpx")
Document.open("input.docx").save("build/from_word.hwp")
Document.open("input.hwpx").save("build/from_hwpx.docx")
Document.open("input.hwp").save("build/from_hwp.docx")

# Obtain an editable WordDocument before saving, if needed.
word = Document.open("input.hwpx").to_docx_document()
with word.edit_word() as backend:
    backend.add_paragraph("Added in Word")
word.save("build/from_hwpx_edited.docx")
```

#### Run the complete example

From the repository root, after installing `.[docx]`:

```bash
python examples/docx_addon.py
# Optional: choose a different output directory.
python examples/docx_addon.py --output-dir build/my-word-demo
```

The [runnable example](examples/docx_addon.py) creates `original.docx`,
`untouched-copy.docx`, `edited.docx`, `converted.hwpx`, and `restored.docx` in
`build/docx-example/` by default. It checks the untouched copy's bytes and prints
each output path. Microsoft Word and Hancom Office are not required to run it.

Shared DOCX edits preserve unrelated package parts. Cross-format conversion maps
supported content and formatting through the common model; it does not guarantee
identical layout or preservation of every Word feature. Unsupported Word features
produce `UserWarning` messages during conversion. The addon preserves independent
body reference and footnote/endnote marker formatting and maps supported first-page
headers/footers, lists, and fields to HWPX equivalents or substitute representations.
HWP has additional limits; an external converter remains opt-in.

The existing `from jakal_hwpx import WordDocument` import remains supported.
The low-level `HwpxDocument` and `HwpDocument` classes also provide
`save_as_docx()` and `to_docx_document()`. Equation conversion supports a subset of
OMML/Hancom math; other structures and embedded objects have limits.
See the [DOCX addon guide](docs/docx-addon.md),
[WordDocument API and export limits](docs/word-document.md), and
[formatting and foreground-render verification](docs/docx-fidelity-benchmark.md).

## Direct HWPX authoring

Use `HwpxDocument` when you want explicit control over paragraphs, runs, controls, package parts, or validation.

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.blank()
doc.append_paragraph("Inline math:")
doc.append_inline("equ", "x+y", width=3200, height=1800)
doc.append_inline("text", " = z")
doc.append_block("equ", "a+b", width=2800, height=1700)

doc.strict_validate()
doc.save("build/direct.hwpx")
```

There are now two authoring styles for controls:

- `append_block(kind=..., content=..., **kwargs)`: insert a block-level control
- `append_inline(kind=..., content=..., **kwargs)`: reuse the target paragraph instead of creating a new one

Examples:

```python
doc.append("equ", "x+y")
doc.append_inline("equ", "x+y")
doc.append_inline("text", " + z")
doc.append_block("table", [["A", "B"], ["1", "2"]])
```

Supported type aliases include:

- `text`, `paragraph`
- `eq`, `equ`, `equation`
- `pic`, `image`, `picture`
- `table`, `tbl`
- `bookmark`, `field`, `hyperlink`
- `note`, `footnote`, `endnote`
- `form`, `memo`, `chart`, `ole`, `shape`
- `autonum`, `newnum`
- `header`, `footer`

The older explicit APIs still exist and remain the lowest-level stable surface:

- `append_equation()`
- `append_inline_equation()`
- `append_picture()`
- `append_shape()`
- `append_ole()`
- `append_table()`

## Choosing the right layer

| Layer | Use it for |
|---|---|
| `Document` | format-neutral authoring, conversion, block-level editing |
| `WordDocument` (DOCX addon, beta) | Word package preservation, shared edits, and python-docx formatting |
| `HwpxDocument` | direct HWPX package editing, XML placement, validation, control surgery |
| `HwpDocument` | direct native HWP object editing and low-level HWP inspection |
| `HwpBinaryDocument` | record tree, streams, DocInfo, and section model inspection |

Additional public API documentation lives in [HWPX_MODULE.md](https://github.com/Ahnd6474/jakal-hwpx/blob/main/HWPX_MODULE.md).

## Validation and release checks

Basic test run:

```bash
python -m pytest tests/test_document_model.py tests/test_hancom_document.py -q
```

Packaging and release validation:

```bash
python -m build
python -m twine check dist/*
python scripts/check_release.py --profile ci
```

On the Windows release machine, run the full Hancom gate as well:

```powershell
python scripts/check_release.py --profile release
```

## Documentation

- [HWPX_MODULE.md](https://github.com/Ahnd6474/jakal-hwpx/blob/main/HWPX_MODULE.md): public module and API index
- [docs/README.md](https://github.com/Ahnd6474/jakal-hwpx/blob/main/docs/README.md): docs directory index
- [docs/hancom-document.md](https://github.com/Ahnd6474/jakal-hwpx/blob/main/docs/hancom-document.md): `Document`
- [docs/docx-addon.md](docs/docx-addon.md): DOCX addon installation and architecture
- [docs/word-document.md](docs/word-document.md): `WordDocument`, Word editing, and conversion limits
- [examples/docx_addon.py](examples/docx_addon.py): runnable DOCX creation, editing, and HWPX round trip
- [docs/hwpx-document.md](https://github.com/Ahnd6474/jakal-hwpx/blob/main/docs/hwpx-document.md): `HwpxDocument`
- [docs/hwp-document.md](https://github.com/Ahnd6474/jakal-hwpx/blob/main/docs/hwp-document.md): `HwpDocument`
- [docs/bridge-and-binary.md](https://github.com/Ahnd6474/jakal-hwpx/blob/main/docs/bridge-and-binary.md): bridge and binary internals
- [RELEASING.md](https://github.com/Ahnd6474/jakal-hwpx/blob/main/RELEASING.md): release checklist

## License

Released under the [MIT License](https://github.com/Ahnd6474/jakal-hwpx/blob/main/LICENSE).
