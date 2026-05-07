# `jakal_hwpx` module guide

This file is the public API index for `jakal_hwpx`.

If you are starting fresh, read this in order:

1. [README.md](./README.md)
2. [docs/hancom-document.md](./docs/hancom-document.md)
3. [docs/hwpx-document.md](./docs/hwpx-document.md)
4. [docs/hwp-document.md](./docs/hwp-document.md)

## Public import rule

Prefer root-package imports:

```python
from jakal_hwpx import HancomDocument, HwpxDocument, HwpDocument
```

Direct submodule imports still work, but the root import path is the documented public entry point.

## Main layers

| Layer | Primary objects | Use it for |
|---|---|---|
| Format-neutral model | `HancomDocument`, `HancomSection`, block dataclasses | authoring, conversion, roundtrip-safe document workflows |
| HWPX package model | `HwpxDocument`, `*Xml` wrappers, `HwpxXmlNode` | direct HWPX editing, XML placement, validation, package parts |
| Native HWP model | `HwpDocument`, `Hwp*Object` wrappers | direct native HWP editing and inspection |
| Binary inspection | `HwpBinaryDocument`, `RecordNode`, `SectionModel` | streams, records, DocInfo, low-level debugging |
| Code generation | `generate_hwpx_script()`, `generate_hwp_script()` | recreate documents as Python scripts |

## Common public objects

### Document models

- `HancomDocument`
- `HwpxDocument`
- `HwpDocument`
- `HwpHwpxBridge`
- `HwpBinaryDocument`

### Hancom block types

- `Paragraph`
- `Table`
- `Picture`
- `Shape`
- `Equation`
- `Ole`
- `Field`
- `Hyperlink`
- `Bookmark`
- `AutoNumber`
- `Note`
- `HeaderFooter`
- `Memo`
- `Form`
- `Chart`

### HWPX XML wrappers

- `HeaderFooterXml`
- `TableXml`, `TableCellXml`
- `PictureXml`
- `ShapeXml`
- `EquationXml`
- `OleXml`
- `FieldXml`
- `BookmarkXml`
- `AutoNumberXml`
- `NoteXml`
- `MemoXml`
- `FormXml`
- `ChartXml`
- `ParagraphStyleXml`
- `CharacterStyleXml`
- `StyleDefinitionXml`
- `SectionSettingsXml`
- `HwpxXmlNode`

### Native HWP wrappers

- `HwpParagraphObject`
- `HwpTableObject`, `HwpTableCellObject`
- `HwpPictureObject`
- `HwpShapeObject`
- `HwpEquationObject`
- `HwpOleObject`
- `HwpFieldObject`
- `HwpHyperlinkObject`
- `HwpBookmarkObject`
- `HwpHeaderFooterObject`
- `HwpNoteObject`
- `HwpPageNumObject`
- `HwpFormObject`
- `HwpMemoObject`
- `HwpChartObject`

### Exceptions

- `HwpxError`
- `InvalidHwpxFileError`
- `HwpxValidationError`
- `InvalidHwpFileError`
- `HwpBinaryEditError`
- `HancomInteropError`
- `ValidationIssue`

## Recommended entry points

### Start with `HancomDocument`

Use this when the caller does not need to care whether the source is HWP or HWPX.

```python
from jakal_hwpx import HancomDocument

doc = HancomDocument.read_hwpx("input.hwpx")
doc.append_paragraph("Done")
doc.write_to_hwp("build/output.hwp")
```

### Use `HwpxDocument` for exact placement

This layer now has both explicit append APIs and generic dispatchers.

```python
from jakal_hwpx import HwpxDocument

doc = HwpxDocument.blank()
doc.append_paragraph("Inline math:")
doc.append_inline("equ", "x+y")
doc.append_inline("text", " = z")
doc.append_block("equ", "a+b")
```

### Use `HwpDocument` for native HWP work

```python
from jakal_hwpx import HwpDocument

doc = HwpDocument.open("input.hwp")
doc.append_paragraph("Added in native HWP mode")
doc.save("build/output.hwp")
```

## Detailed docs

| File | Scope |
|---|---|
| [docs/hancom-document.md](./docs/hancom-document.md) | format-neutral document model |
| [docs/hwpx-document.md](./docs/hwpx-document.md) | HWPX package editing and authoring |
| [docs/hwp-document.md](./docs/hwp-document.md) | native HWP object model |
| [docs/bridge-and-binary.md](./docs/bridge-and-binary.md) | bridge and binary internals |
| [docs/hwpx2py.md](./docs/hwpx2py.md) | HWPX to Python script generation |
| [docs/hwp2py.md](./docs/hwp2py.md) | HWP to Python script generation |

## Release and packaging

Release checks and packaging steps are documented in [RELEASING.md](./RELEASING.md).
