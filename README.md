# jakal-hwpx

`jakal-hwpx` is a Python library for reading, editing, and writing both `HWPX` and `HWP` documents.

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
python -m pip install -e .[dev]
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

Install the bundled DOCX addon with `pip install "jakal-hwpx[docx]"`.
Its implementation lives in `jakal_hwpx.addons.docx` and shares the same release
as the core library. The default HWP/HWPX installation does not require python-docx;
the addon and its backend load only when requested.
The DOCX surface is released as a beta API in `1.1.0b1`; behavior and fidelity may
change based on feedback before the stable `1.1.0` release.
`Document.open("input.docx")` returns a `WordDocument`, which retains the complete
original DOCX package alongside its common model view.

```python
from jakal_hwpx import Document
from jakal_hwpx.addons.docx import WordDocument

doc = Document.open("input.docx")
assert isinstance(doc, WordDocument)
doc.append_paragraph("Review complete")
doc.replace_text("Old text", "New text")
doc.save("build/output.docx")
doc.save("build/output.hwpx")  # Python conversion through the common model
doc.save("build/output.hwp")   # Python conversion through the common model

with doc.edit_word() as word:
    word.paragraphs[0].runs[0].bold = True
```

Untouched DOCX output is byte-identical. Shared edits preserve unrelated parts,
and Word-specific details can be edited through the python-docx backend.
HWP/HWPX output is generated in Python without Hancom Office. The common model
retains supported content and formatting; unsupported Word features produce
`UserWarning` messages during conversion. An external converter remains opt-in.
The existing `from jakal_hwpx import WordDocument` import remains supported.
See the [DOCX addon](docs/docx-addon.md) and [WordDocument](docs/word-document.md)
for supported edits, requirements, and visual limits.

HWP and HWPX can also be exported to editable Word documents in Python:

```python
Document.open("input.hwpx").save("build/from_hwpx.docx")
Document.open("input.hwp").save("build/from_hwp.docx")
word = Document.open("input.hwpx").to_docx_document()  # WordDocument
```

The low-level `HwpxDocument` and `HwpDocument` classes also provide
`save_as_docx()` and `to_docx_document()`. This maps supported shared content and
formatting; native equations and unsupported embedded objects require an explicit
external converter. See the [DOCX export limits](docs/word-document.md#hwphwpx에서-docx-내보내기).

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
- [docs/hwpx-document.md](https://github.com/Ahnd6474/jakal-hwpx/blob/main/docs/hwpx-document.md): `HwpxDocument`
- [docs/hwp-document.md](https://github.com/Ahnd6474/jakal-hwpx/blob/main/docs/hwp-document.md): `HwpDocument`
- [docs/bridge-and-binary.md](https://github.com/Ahnd6474/jakal-hwpx/blob/main/docs/bridge-and-binary.md): bridge and binary internals
- [RELEASING.md](https://github.com/Ahnd6474/jakal-hwpx/blob/main/RELEASING.md): release checklist

## License

Released under the [MIT License](https://github.com/Ahnd6474/jakal-hwpx/blob/main/LICENSE).
