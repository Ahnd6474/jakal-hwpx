from __future__ import annotations

import mimetypes
import os
import tempfile
from pathlib import Path, PurePosixPath
from uuid import uuid4

from ._hwpx_archive import normalize_package_path
from .exceptions import ValidationIssue

_GRAPHIC_CONTROL_TAGS = (
    "tbl",
    "pic",
    "rect",
    "line",
    "ellipse",
    "arc",
    "polygon",
    "curve",
    "connectLine",
    "textart",
    "container",
    "ole",
)

_FORM_TAG_BY_TYPE = {
    "CHECKBOX": "checkBtn",
    "RADIOBUTTON": "radioBtn",
    "BUTTON": "btn",
    "INPUT": "edit",
    "EDIT": "edit",
    "COMBOBOX": "comboBox",
    "LISTBOX": "listBox",
    "SCROLLBAR": "scrollBar",
}
_FORM_TAGS = tuple(_FORM_TAG_BY_TYPE.values())
_FORM_METADATA_COMMAND_PREFIX = "JAKAL_FORM_META:"
_MEMO_CARRIER_FIELD_TYPE = "JAKAL_MEMO"
_BUTTON_FORM_TAGS = {"checkBtn", "radioBtn", "btn"}


def _normalize_path(path: str) -> str:
    return normalize_package_path(path)


def _bool_like_literal_is_valid(value: str | None) -> bool:
    if value is None:
        return True
    return value.strip().lower() in {"0", "1", "false", "true", "no", "yes", "n", "y", "off", "on"}


def _issue(
    kind: str,
    message: str,
    *,
    part_path: str | None = None,
    section_index: int | None = None,
    paragraph_index: int | None = None,
    cell_row: int | None = None,
    cell_column: int | None = None,
    xpath: str | None = None,
    line: int | None = None,
    column: int | None = None,
    context: str | None = None,
) -> ValidationIssue:
    return ValidationIssue(
        kind=kind,
        message=message,
        part_path=part_path,
        section_index=section_index,
        paragraph_index=paragraph_index,
        cell_row=cell_row,
        cell_column=cell_column,
        xpath=xpath,
        line=line,
        column=column,
        context=context,
    )


def _graphic_attribute_xpath(attribute: str) -> str:
    return " | ".join(f".//hp:{tag}/@{attribute}" for tag in _GRAPHIC_CONTROL_TAGS)


def _expected_binary_media_type(path: str) -> str | None:
    suffix = PurePosixPath(path).suffix.lower()
    if suffix == ".ole":
        return "application/ole"
    return mimetypes.guess_type(PurePosixPath(path).name)[0]


def _new_temp_dir(prefix: str) -> Path:
    root = Path(
        os.environ.get("JAKAL_HWPX_TEMP_ROOT")
        or os.environ.get("TMPDIR")
        or os.environ.get("TEMP")
        or tempfile.gettempdir()
    )
    root.mkdir(parents=True, exist_ok=True)
    for _ in range(32):
        candidate = root / f"{prefix}{uuid4().hex}"
        try:
            candidate.mkdir(parents=True, exist_ok=False)
            return candidate
        except FileExistsError:
            continue
    return Path(tempfile.mkdtemp(prefix=prefix, dir=str(root)))
