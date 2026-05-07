# jakal-hwpx

`jakal-hwpx` is a Python library for reading, editing, and writing both `HWPX` and `HWP` documents.

The package exposes three main layers:

- `HancomDocument`: a format-neutral document model for most authoring and conversion work
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

For most application code, start with `HancomDocument`.

```python
from jakal_hwpx import HancomDocument

doc = HancomDocument.blank()
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
from jakal_hwpx import HancomDocument

doc = HancomDocument.read_hwp("input.hwp")
doc.append_paragraph("Review complete")
doc.write_to_hwpx("build/output.hwpx")
```

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

- `append_block(type=..., content=..., **kwargs)`: insert a block-level control
- `append_inline(type=..., content=..., **kwargs)`: reuse the target paragraph instead of creating a new one

Examples:

```python
doc.append_block("equ", "x+y")
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
| `HancomDocument` | format-neutral authoring, conversion, block-level editing |
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
- [docs/hancom-document.md](https://github.com/Ahnd6474/jakal-hwpx/blob/main/docs/hancom-document.md): `HancomDocument`
- [docs/hwpx-document.md](https://github.com/Ahnd6474/jakal-hwpx/blob/main/docs/hwpx-document.md): `HwpxDocument`
- [docs/hwp-document.md](https://github.com/Ahnd6474/jakal-hwpx/blob/main/docs/hwp-document.md): `HwpDocument`
- [docs/bridge-and-binary.md](https://github.com/Ahnd6474/jakal-hwpx/blob/main/docs/bridge-and-binary.md): bridge and binary internals
- [RELEASING.md](https://github.com/Ahnd6474/jakal-hwpx/blob/main/RELEASING.md): release checklist

## License

Released under the [MIT License](https://github.com/Ahnd6474/jakal-hwpx/blob/main/LICENSE).
