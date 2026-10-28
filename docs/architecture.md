# Internal architecture

`jakal_hwpx` keeps the historical public modules as compatibility façades. New implementation code should be placed in
the smallest internal module that owns the relevant responsibility instead of extending an existing façade indefinitely.

## Public façades

| Module | Responsibility |
|---|---|
| `document.py` | Public `HwpxDocument` editing API and validation façade |
| `hwp_document.py` | Public native HWP document façade |
| `shared_document.py`, `hancom_document.py` | Public `Document` base and legacy `HancomDocument` compatibility alias |
| `word_document.py` | `WordDocument(Document)`: Word backend, native edits, DOCX preservation, pure Python HWP/HWPX output and optional converter dispatch |
| `hancom_converter.py` | Isolated optional Hancom Office converter |
| `hwp_binary.py` | Public record model and `HwpBinaryDocument` API |

Existing imports from these modules remain supported. Internal implementations imported by a façade are deliberately
re-exported where older callers may already depend on them.

## Internal boundaries

| Module | Owned responsibility |
|---|---|
| `_hwpx_archive.py` | HWPX ZIP entry reading, ordering, metadata cloning, and writing |
| `_hwpx_common.py` | Shared HWPX path, metadata, validation-issue, and temporary-directory helpers |
| `_hwpx_factory.py` | Construction of a blank HWPX package and its XML parts |
| `_hwpx_validation.py` | HWPX package validation, reference checks, and strict linting |
| `_hancom_models.py` | Format-neutral dataclasses and semantic field behavior |
| `_hancom_styles.py` | HWP/HWPX style and DocInfo payload conversion |
| `_hancom_from_hwp.py`, `_hancom_to_hwp.py` | Native HWP conversion dispatch by direction |
| `_hancom_from_hwpx.py`, `_hancom_to_hwpx.py` | HWPX conversion dispatch by direction |
| `_hancom_from_docx.py` | Word content projection into shared model views |
| `_hancom_to_docx.py` | Portable Word export from HWP/HWPX/shared models, with grouped runs and rich HWPX table cells |
| `_equation_conversion.py` | Structured OMML ↔ HWP equation mapping for fractions, scripts and radicals; explicit rejection of unsupported math |
| `_docx_preservation.py` | Original Word package, surgical common edits, native backend changes and untouched-part preservation |
| `_hwp_constants.py` | Native HWP tags, flags, templates, and binary constants |
| `_hwp_payload_codecs.py` | Pure native HWP payload decoders |
| `_hwp_graphic_payloads.py` | Native graphic layout, line, fill, crop, and OLE payload builders |
| `_hwp_page_payloads.py` | Native section, page, numbering, border, and note payload codecs |
| `_hwp_object_payloads.py` | Native object metadata and geometry payload helpers |
| `_hwp_objects.py` | Core native HWP section/control wrappers and compatibility re-exports |
| `_hwp_table_objects.py`, `_hwp_shape_objects.py` | Native table and shape wrapper families |
| `_hwp_record_tree.py` | Typed record specialization and DocInfo/section tree models |
| `_hwpx_chart_xml.py` | Chart XML construction and metadata codecs |
| `_hwpx_element_helpers.py` | Shared HWPX geometry, text editing, paragraph, and form XML helpers |
| `_hwpx_document_controls.py` | Header, note, form, chart, equation, and style creation mixin |
| `_hwpx_document_graphics.py` | Table, picture, shape, and OLE creation mixin |
| `_hwpx_document_fields.py` | Bookmark, field, hyperlink, and cross-reference creation mixin |
| `_hwpx2py_semantic.py`, `_hwpx2py_macro.py`, `_hwpx2py_authoring.py` | Python code emitters by output mode |

## Dependency direction

Low-level codec and model modules do not import the public façade that consumes them at runtime. Type-only references use
`TYPE_CHECKING`. Directional conversion and object-family adapters are loaded only after their façade has defined the
shared compatibility helpers they consume. Public modules may import and re-export internal names, while new low-level
modules should depend only on constants, parts, record types, and models. This keeps the module initialization graph
acyclic without breaking historical imports.

`Document.open()` selects `WordDocument` for DOCX. Shared operations dispatch to the Word
implementation, which updates its python-docx backend and refreshes shared views. Word-specific
details remain in the source package rather than being reconstructed from the common projection.
DOCX output writes this package; HWP/HWPX output uses the common model and native
Python writers by default. An explicitly supplied converter can handle features
outside the common model. `to_common_document()` exposes the portable projection.

When adding functionality:

1. Put byte-layout logic in a payload codec module.
2. Put format-neutral state and behavior in `_hancom_models.py`.
3. Keep orchestration and public compatibility methods in the façade.
4. Add the new internal module to the release package check; add it to the staged mypy boundary once its dynamic API is typed.
