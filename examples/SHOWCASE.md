# Showcase Examples

`build_showcase_bundle.py` creates promotion-ready HWPX outputs from the local sample corpus.

## Run

```powershell
python examples/build_showcase_bundle.py
```

## Output

By default the script writes to `examples/output/`.

Generated files:

- `showcase_layout_headers.hwpx`
- `showcase_table_picture.hwpx`
- `showcase_fields_references.hwpx`
- `showcase_notes.hwpx`
- `showcase_numbering.hwpx`
- `showcase_equation.hwpx`
- `showcase_shapes.hwpx`
- `showcase_manifest.json`
- `showcase_report.md`

## What Each File Demonstrates

- `showcase_layout_headers.hwpx`: header/footer editing, section/page setup, paragraph insertion, style batch application
- `showcase_table_picture.hwpx`: table text changes, merged row editing, picture comment editing
- `showcase_fields_references.hwpx`: bookmark, hyperlink, mail merge, formula, cross-reference generation
- `showcase_notes.hwpx`: footnote/endnote content editing
- `showcase_numbering.hwpx`: body-level automatic numbering update using `hp:newNum`
- `showcase_equation.hwpx`: equation script editing
- `showcase_shapes.hwpx`: shape comment editing and textart update when available

## Validation

Each generated file is checked with:

- `xml_validation_errors()`
- `reference_validation_errors()`
- `save_reopen_validation_errors()`

`showcase_manifest.json` stores the validation result for every output file.
