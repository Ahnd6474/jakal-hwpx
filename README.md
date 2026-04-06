# jakal-hwpx

Python tools for reading, editing, validating, and bridging `HWPX`, plus early `PDF <-> HWPX` conversion helpers.

Korean overview: [README.ko.md](./README.ko.md)

## What This Repo Contains

- `src/jakal_hwpx`: the importable Python package
- `examples/samples`: sample `hwpx`, `hwp`, and `pdf` documents used for local experimentation
- `examples/output_smoke`: committed smoke corpus for tests
- `examples/output`: showcase outputs generated from the smoke corpus
- `tools`: bundled Java-based `.hwp -> .hwpx` converter assets used by maintainer workflows

This README is intentionally repo-level. Detailed module and API notes live in [HWPX_MODULE.md](./HWPX_MODULE.md).

## Installation

You need Python 3.11 or newer.

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

doc = HwpxDocument.open("examples/samples/hwpx/AI와_특이점_보고서.hwpx")
doc.replace_text("before", "after", count=1)
doc.save("build/edited.hwpx")
```

For module-level usage, `PdfDocument`, or bridge APIs such as `pdf_to_hwpx()` and `hwpx_to_pdf()`, see [HWPX_MODULE.md](./HWPX_MODULE.md).

## Repository Layout

```text
src/jakal_hwpx/          Python package
examples/samples/hwpx/   sample HWPX documents
examples/samples/hwp/    sample HWP documents
examples/samples/pdf/    sample PDF documents
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
- `examples/samples/pdf/`

Generated validation outputs belong under `build/validation/`.

## More Docs

- [HWPX_MODULE.md](./HWPX_MODULE.md): package structure, module roles, and API usage
- [examples/SHOWCASE.md](./examples/SHOWCASE.md): showcase generation workflow
- [RELEASING.md](./RELEASING.md): release checklist
- [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md): scope notes for sample files, bundled tools, and HWPX-related naming

## License

Original project-authored source code in this repository is available under the [MIT License](./LICENSE).

Sample documents, committed outputs, and bundled toolchain artifacts under `tools/` can be subject to separate rights or upstream licenses. See [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md) before redistributing those assets.
