# jakal-hwpx

Read, edit, validate, and convert HWPX and native HWP documents in Python.

`jakal-hwpx` gives you three working surfaces:

- `HwpxDocument` for direct HWPX authoring and editing
- `HwpDocument` for native HWP read/write and object-level editing
- `HancomDocument` for a common HWP/HWPX bridge model

Normal use is pure Python. An installed Hancom application is optional and only used for smoke validation scripts.

## Table of contents

- [Installation](#installation)
- [Quick start](#quick-start)
- [What it is](#what-it-is)
- [Why use it](#why-use-it)
- [Main entry points](#main-entry-points)
- [Support matrix](#support-matrix)
- [API overview](#api-overview)
- [Examples](#examples)
- [Testing and validation](#testing-and-validation)
- [Further reading](#further-reading)
- [License](#license)

## Installation

Requirements:

- Python 3.11 or newer
- No Git LFS setup

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

Install with development dependencies:

```bash
python -m pip install -e .[dev]
```

Package name: `jakal-hwpx`  
Import path: `jakal_hwpx`

## Quick start

Use `HancomDocument` when you want one editing model that can write both formats:

```python
from jakal_hwpx import HancomDocument

doc = HancomDocument.blank()
doc.metadata.title = "Quarterly report"

doc.append_header("Internal")
doc.append_paragraph("Revenue summary")
doc.append_table(rows=2, cols=2, cell_texts=[["Q1", "Q2"], ["120", "135"]])
doc.append_equation("x+y", width=3200, height=1800)

doc.write_to_hwpx("build/report.hwpx")
doc.write_to_hwp("build/report.hwp")
```

If you are only working with HWPX, `HwpxDocument` is the simplest path:

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.blank()
doc.set_metadata(title="Hello")
doc.append_paragraph("Hello HWPX")
doc.save("build/hello.hwpx")
```

## What it is

`jakal-hwpx` is a document editing library for Hancom formats. It covers direct HWPX editing, native HWP editing, and pure-Python bridge workflows between the two.

The project is not trying to mimic Hancom's GUI. The focus is programmatic editing: generating documents, changing structured content, preserving supported controls through round trips, and validating the result before you save it.

The stable compatibility target is the top-level `jakal_hwpx` import surface. Internal module paths such as `jakal_hwpx.document` or `jakal_hwpx.parts` are implementation details and may move.

## Why use it

- You can edit HWPX directly without unpacking ZIP parts by hand.
- You can read and write native HWP without depending on COM automation.
- You can move between HWP and HWPX through `HancomDocument` when you want one editing model.
- You get validation on both sides: strict HWPX package checks and strict HWP structure checks.
- You can drop to lower-level APIs when you need typed record access instead of high-level editing helpers.

## Main entry points

| Entry point | Best for | Notes |
|---|---|---|
| `HwpxDocument` | Direct HWPX authoring and editing | Fastest path when your input and output are both HWPX |
| `HwpDocument` | Native HWP editing | Object wrappers for paragraphs, tables, fields, shapes, notes, section settings, and more |
| `HancomDocument` | HWP/HWPX bridge workflows | Common IR for `read_hwp()`, `read_hwpx()`, `write_to_hwp()`, and `write_to_hwpx()` |
| `HwpBinaryDocument` | Low-level native HWP work | Typed record access, section/docinfo models, stream-level reencode work |

## Support matrix

### Workflow support

| Flow | Current level | Best fit | Notes |
|---|---|---|---|
| `HWPX -> HWPX` | Excellent | Direct editing and generation | Broadest authoring surface |
| `HWP -> HWP` | Strong | Native HWP editing and reencode | Most conservative path for existing HWP files |
| `HWP -> HWPX` | Strong | Lifting native HWP into richer XML | Good semantic coverage with less normalization pressure |
| `HWPX -> HWP` | Good | Generating native HWP from supported HWPX content | Most normalization happens here |

### Control-family support

Status legend:

- `Full`: high-level edit surface is in place and covered in roundtrip tests
- `Strong`: supported and exercised, but with more native-format constraints
- `Partial`: typed access or wrappers exist, but the edit surface is narrower

| Control family | `HWPX -> HWPX` | `HWP -> HWP` | `HWP -> HWPX` | `HWPX -> HWP` |
|---|---|---:|---:|---:|
| Paragraphs and styles | Full | Strong | Strong | Strong |
| Header and footer | Full | Strong | Strong | Strong |
| Fields, bookmarks, hyperlinks | Full | Strong | Strong | Strong |
| Notes, auto numbers, page numbers | Full | Strong | Strong | Strong |
| Section settings and page border fill | Full | Strong | Strong | Strong |
| Tables | Full | Strong | Strong | Strong |
| Pictures | Full | Strong | Strong | Strong |
| Shapes and `connectLine` | Full | Strong | Strong | Strong |
| Equations | Full | Strong | Strong | Strong |
| OLE | Full | Strong | Strong | Strong |
| Charts | Partial | Partial | Partial | Partial |
| Form objects | Partial | Partial | Partial | Partial |
| Memo and comment controls | Partial | Partial | Partial | Partial |

### What "good" means here

This library is already useful for document automation. It is not a drop-in replacement for Hancom's GUI editor.

If your job is:

- generate templates
- fill fields
- edit paragraphs
- update tables
- change section settings
- preserve supported controls through save/reopen cycles

then `jakal-hwpx` is in good shape.

If your job is:

- arbitrary WYSIWYG editing
- every edge-case object Hancom can open
- pixel-perfect behavior across rare vendor-specific controls

Hancom still has more freedom.

## API overview

### `HwpxDocument`

Main direct-edit entry point for HWPX.

Common operations:

- `HwpxDocument.open(path)`
- `HwpxDocument.blank()`
- `append_paragraph()`
- `append_header()`, `append_footer()`
- `append_table()`, `append_picture()`, `append_shape()`, `append_equation()`, `append_ole()`
- `append_field()`, `append_hyperlink()`, `append_bookmark()`, `append_note()`
- `section_settings()`
- `strict_lint_errors()`, `strict_validate()`
- `save(path)`

### `HwpDocument`

Main native HWP editing surface.

Common operations:

- `HwpDocument.open(path)`
- `HwpDocument.blank()`
- `append_paragraph()`
- `append_table()`, `append_picture()`, `append_shape()`, `append_equation()`, `append_ole()`
- `append_field()`, `append_hyperlink()`, `append_bookmark()`, `append_note()`
- `append_header()`, `append_footer()`, `append_auto_number()`
- `section(index)` and object wrappers such as `tables()`, `pictures()`, `fields()`, `notes()`
- `strict_lint_errors()`, `strict_validate()`
- `save(path)` for both `.hwp` and `.hwpx`

### `HancomDocument`

Common IR for HWP/HWPX bridge workflows.

Common operations:

- `HancomDocument.blank()`
- `HancomDocument.read_hwpx(path)`
- `HancomDocument.read_hwp(path)`
- `append_paragraph()`
- `append_table()`, `append_picture()`, `append_shape()`, `append_equation()`, `append_ole()`
- `append_field()`, `append_hyperlink()`, `append_bookmark()`, `append_note()`
- `append_header()`, `append_footer()`
- `write_to_hwpx(path)`
- `write_to_hwp(path)`

### `HwpBinaryDocument`

Low-level native HWP entry point.

Use it when you need typed record trees, `DocInfoModel`, `SectionModel`, or stream-level reencode work instead of high-level document editing.

## Examples

### Edit an existing HWPX file

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.open("input.hwpx")
doc.replace_text("Draft", "Final")
doc.append_paragraph("Approved for distribution")
doc.strict_validate()
doc.save("build/edited.hwpx")
```

### Edit an existing HWP file

```python
from jakal_hwpx import HwpDocument

doc = HwpDocument.open("input.hwp")
doc.append_paragraph("Bridge paragraph")
doc.append_hyperlink("https://example.com", text="Example")
doc.strict_validate()
doc.save("build/edited.hwp")
```

### Convert HWP to HWPX

```python
from jakal_hwpx import HancomDocument

doc = HancomDocument.read_hwp("input.hwp")
doc.write_to_hwpx("build/exported.hwpx")
```

### Convert HWPX to HWP

```python
from jakal_hwpx import HancomDocument

doc = HancomDocument.read_hwpx("input.hwpx")
doc.write_to_hwp("build/exported.hwp")
```

### Inspect native HWP records

```python
from jakal_hwpx import HwpBinaryDocument

doc = HwpBinaryDocument.open("input.hwp")
print(doc.file_header().version)
print(doc.docinfo_model().id_mappings_record().named_counts())
print(doc.section_model(0).controls())
```

## Testing and validation

Run the core test suite:

```bash
python -m pip install -e .[dev]
python -m pytest -q
```

Run the stability and release checks:

```bash
python scripts/check_release.py
```

Optional Hancom-based smoke validation on Windows:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/setup_hancom_security_module.ps1 -DownloadIfMissing
powershell -ExecutionPolicy Bypass -File scripts/run_hancom_smoke_validation.ps1 -InputPath examples\samples\hwpx\AI와_특이점_보고서.hwpx -OutputPath .codex-temp\hancom-smoke\sample.roundtrip.hwpx
powershell -ExecutionPolicy Bypass -File scripts/run_hancom_corpus_smoke_validation.ps1
```

## Further reading

- [HWPX_MODULE.md](./HWPX_MODULE.md) for package structure and API notes
- [STABILITY_CONTRACT.md](./STABILITY_CONTRACT.md) for release criteria and non-goals
- [scripts/check_release.py](./scripts/check_release.py) for the release gate
- [scripts/audit_hwp_lossless_roundtrip.py](./scripts/audit_hwp_lossless_roundtrip.py) for HWP reencode audits
- [scripts/run_bridge_stability_lab.py](./scripts/run_bridge_stability_lab.py) for bridge matrix runs
- [examples/SHOWCASE.md](./examples/SHOWCASE.md) for generated document examples
- [RELEASING.md](./RELEASING.md) for packaging and release steps
- [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md) for sample document notices

Advanced tools are also exported at the top level:

- `build_hwp_pure_profile()`
- `append_feature_from_profile()`
- `run_template_lab()`
- donor-scanning helpers from `hwp_collection`

Those are useful when you need template-backed native HWP authoring or corpus analysis, but they are not the main path for day-to-day document editing.

## License

The project source code is licensed under [MIT](./LICENSE).

Sample documents and generated outputs may have separate rights. Check [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md) before redistributing them.
