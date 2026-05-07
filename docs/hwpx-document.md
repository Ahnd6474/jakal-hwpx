# `HwpxDocument`

`HwpxDocument` is the direct HWPX package editing layer.

Use it when you need:

- exact paragraph or control placement
- direct XML access
- package part inspection or mutation
- HWPX validation and strict linting

```python
from jakal_hwpx import HwpxDocument
```

## Create, open, and save

| API | Purpose |
|---|---|
| `HwpxDocument.blank()` | build a new blank HWPX document |
| `HwpxDocument.open(path, validate=True)` | open an existing HWPX package |
| `HwpxDocument.from_bytes(raw_bytes)` | open from bytes |
| `compile(validate=True, auto_repair=True)` | compile the current package to bytes |
| `save(path, validate=True, auto_repair=True)` | save to `.hwpx` |
| `to_hancom_document()` | convert to `HancomDocument` |
| `to_hwp_document()` | materialize native `HwpDocument` |
| `save_as_hwp(path, *, converter=None, force_refresh=True)` | write native HWP |

## Package and XML access

| API | Purpose |
|---|---|
| `list_part_paths()` | list package parts |
| `get_part(path, expected_type=None)` | return a part wrapper |
| `add_part(path, raw_bytes)` | add a package part |
| `remove_part(path)` | remove a package part |
| `add_or_replace_binary(name, data, *, media_type=None, manifest_id=None, is_embedded=True)` | manage binary parts |
| `xml_part(path)` | return a generic XML wrapper |
| `section_xml(section_index=0)` | return a section XML wrapper |
| `header_xml()` | return the header XML wrapper |
| `content_hpf_xml()` | return content metadata XML |
| `settings_xml()` | return settings XML |

## Paragraph editing

| API | Purpose |
|---|---|
| `append_paragraph(text, ...)` | add a paragraph |
| `insert_paragraph(...)` | insert a paragraph at an index |
| `set_paragraph_text(section_index, paragraph_index, text)` | replace paragraph text while preserving structural controls |
| `delete_paragraph(section_index, paragraph_index)` | delete a paragraph |
| `paragraph_count(section_index=0)` | count paragraphs |
| `paragraph_xml(section_index, paragraph_index, pretty_print=False)` | inspect a paragraph as XML |
| `insert_paragraph_xml(...)` | insert raw `hp:p` XML |
| `replace_paragraph_xml(...)` | replace a paragraph with raw `hp:p` XML |
| `move_paragraph(...)` | move a paragraph |
| `copy_paragraph(...)` | copy a paragraph |

## Generic authoring dispatchers

Two generic dispatchers are now available on top of the existing explicit `append_*` APIs.

### `append_block(type, content=None, **kwargs)`

Insert a block-level control or paragraph.

```python
doc.append_block("equ", "x+y")
doc.append_block("table", [["A", "B"], ["1", "2"]])
doc.append_block("bookmark", "anchor_1")
```

### `append_inline(type, content=None, **kwargs)`

Reuse the target paragraph instead of creating a new one.

```python
doc.append_paragraph("Inline math:")
doc.append_inline("equ", "x+y")
doc.append_inline("text", " = z")
```

If `paragraph_index` is omitted, `append_inline()` uses the current section's last paragraph.

### Supported aliases

| Alias examples | Canonical target |
|---|---|
| `text`, `para`, `paragraph` | paragraph |
| `eq`, `equ`, `equation` | equation |
| `pic`, `image`, `picture` | picture |
| `table`, `tbl` | table |
| `bookmark` | bookmark |
| `field` | field |
| `link`, `hyperlink` | hyperlink |
| `note`, `footnote`, `endnote` | note |
| `form` | form |
| `memo`, `comment` | memo |
| `chart` | chart |
| `autonum`, `newnum` | auto number |
| `header`, `footer` | header or footer |
| `shape`, `ole` | same name |

Notes:

- `append_inline("header", ...)` and `append_inline("footer", ...)` are rejected.
- `append_block("picture", ...)` and `append_block("ole", ...)` require `data=...`.
- `append_block("table", content=[[...]])` will infer `rows`, `columns`, and `cell_texts`.

## Explicit append APIs

The lower-level explicit APIs remain available and are still the most stable surface when you want precise control over arguments.

| API | Return |
|---|---|
| `append_header(...)` | `HeaderFooterXml` |
| `append_footer(...)` | `HeaderFooterXml` |
| `append_note(...)` | `NoteXml` |
| `append_footnote(...)` | `NoteXml` |
| `append_endnote(...)` | `NoteXml` |
| `append_form(...)` | `FormXml` |
| `append_memo(...)` | `MemoXml` |
| `append_chart(...)` | `ChartXml` |
| `append_auto_number(...)` | `AutoNumberXml` |
| `append_new_number(...)` | `AutoNumberXml` |
| `append_equation(...)` | `EquationXml` |
| `append_inline_equation(...)` | `EquationXml` |
| `append_table(...)` | `TableXml` |
| `append_picture(...)` | `PictureXml` |
| `append_shape(...)` | `ShapeXml` |
| `append_ole(...)` | `OleXml` |
| `append_bookmark(...)` | `BookmarkXml` |
| `append_field(...)` | `FieldXml` |
| `append_hyperlink(...)` | `FieldXml` |
| `append_mail_merge_field(...)` | `FieldXml` |
| `append_calculation_field(...)` | `FieldXml` |
| `append_cross_reference(...)` | `FieldXml` |

## Control-level XML editing

| API | Purpose |
|---|---|
| `append_control_xml(xml, *, section_index=0, paragraph_index=None, char_pr_id=None)` | append control XML |
| `append_run_xml(xml, *, section_index=0, paragraph_index=None, char_pr_id=None)` | append run XML |
| `control_count(section_index, paragraph_index)` | count top-level controls in a paragraph |
| `insert_control_xml_at(...)` | insert control XML at an exact index |
| `move_control(...)` | move a top-level control |
| `copy_control(...)` | copy a top-level control |
| `delete_control(...)` | delete a top-level control |

These APIs are close to the raw HWPX structure. Use them when the higher-level append APIs are too restrictive.

## Selectors

Selectors return wrapper lists.

| API | Return |
|---|---|
| `headers()`, `footers()` | `list[HeaderFooterXml]` |
| `tables()` | `list[TableXml]` |
| `pictures()` | `list[PictureXml]` |
| `oles()` | `list[OleXml]` |
| `notes()` | `list[NoteXml]` |
| `memos()` | `list[MemoXml]` |
| `forms()` | `list[FormXml]` |
| `charts()` | `list[ChartXml]` |
| `bookmarks()` | `list[BookmarkXml]` |
| `fields()` | `list[FieldXml]` |
| `hyperlinks()` | `list[FieldXml]` |
| `auto_numbers()` | `list[AutoNumberXml]` |
| `equations()` | `list[EquationXml]` |
| `shapes()` | `list[ShapeXml]` |
| `styles()`, `paragraph_styles()`, `character_styles()` | style wrappers |
| `memo_shapes()` | `list[MemoShapeXml]` |

## Validation and repair

| API | Purpose |
|---|---|
| `validation_errors()` | basic validation issues |
| `validate()` | raise on basic validation failure |
| `reference_validation_errors()` | broken part or relationship references |
| `roundtrip_validate()` | compile and reopen validation |
| `strict_lint_errors()` | stricter semantic lint issues |
| `strict_validate()` | raise on strict lint failure |
| `strict_lint_report()` | structured lint report |
| `format_strict_lint_errors()` | human-readable lint report |
| `repair_stale_paragraph_layout()` | clear stale paragraph `linesegarray` cache patterns |
| `HwpxDocument.repair(path, output_path=None)` | repair a file on disk |

`save()` and `compile()` run with `auto_repair=True` by default.

## Example

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.blank()
doc.append_paragraph("Answer:")
doc.append_inline("equ", "x+y")
doc.append_inline("text", " = z")
doc.append_block("note", "Check derivation", kind="footNote", number=1)

doc.strict_validate()
doc.save("build/output.hwpx")
```

## Related docs

- [README.md](../README.md)
- [HWPX_MODULE.md](../HWPX_MODULE.md)
- [hancom-document.md](./hancom-document.md)
- [hwp-document.md](./hwp-document.md)
