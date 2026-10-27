# Third-Party Notices

The top-level [LICENSE](./LICENSE) applies to original source code in this repository unless a file, directory, or notice says otherwise.

## HWP and HWPX Compatibility Notice

This project is an independent compatibility tool for working with `HWP` and `HWPX` documents.

`Hancom`, `Hangul`, `HWP`, and `HWPX` are referenced only to describe file-format compatibility. No affiliation with, sponsorship by, or endorsement from the respective rights holders is claimed.

## Sample and Fixture Documents

Sample or fixture documents in this repository, including files under `examples/`, committed generated outputs, and any repository-root `*.hwp`, `*.hwpx`, or `*.pdf` files, may carry separate copyright, trademark, privacy, publicity, or template rights.

Unless a file explicitly says otherwise, treat those documents as test or reference assets only. They are not automatically relicensed under the project's MIT license.

## Redistribution Guidance

The optional web workspace uses `@rhwp/core` (MIT) for read-only HWP/HWPX
rendering and DOMPurify (Apache-2.0 OR MPL-2.0) to sanitize SVG. Their notices
are distributed under `web/public/licenses/`. The UI fonts use Spoqa Han Sans
Neo under SIL OFL 1.1; see `web/public/fonts/LICENSE.txt`.

The optional page-template importer uses Pillow (MIT-CMU) for image processing
and pypdfium2 (Apache-2.0 OR BSD-3-Clause) for PDF rendering. pypdfium2 wheels
bundle PDFium, which has its own BSD-style license and third-party notices.
When packaging the server runtime, retain the license files distributed with
Pillow, pypdfium2 and its PDFium binary, including the PDFium dependency notices.
See the [Pillow license](https://github.com/python-pillow/Pillow/blob/main/LICENSE)
and [pypdfium2 licensing documentation](https://pypdfium2.readthedocs.io/en/stable/readme.html#licensing).

If you redistribute this repository or a derivative:

- keep the top-level `LICENSE` with project-authored code
- confirm that you have the right to redistribute included sample documents
- avoid using third-party product or trademark names in a way that suggests endorsement
