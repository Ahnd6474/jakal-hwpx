from __future__ import annotations

from dataclasses import dataclass


class HwpxError(Exception):
    pass


class InvalidHwpxFileError(HwpxError):
    pass


class InvalidHwpFileError(HwpxError):
    pass


class HwpBinaryEditError(HwpxError):
    pass


class HancomInteropError(HwpxError):
    pass


@dataclass(frozen=True)
class ValidationIssue:
    kind: str
    message: str
    code: str | None = None
    part_path: str | None = None
    section_index: int | None = None
    paragraph_index: int | None = None
    cell_row: int | None = None
    cell_column: int | None = None
    xpath: str | None = None
    line: int | None = None
    column: int | None = None
    context: str | None = None

    def __post_init__(self) -> None:
        if self.code is None:
            object.__setattr__(self, "code", self.kind)

    def location(self) -> str:
        parts: list[str] = []
        if self.part_path:
            parts.append(self.part_path)
        elif self.section_index is not None:
            parts.append(f"section[{self.section_index}]")

        if self.context:
            parts.append(self.context)
        if self.paragraph_index is not None and (not self.context or "paragraph[" not in self.context):
            parts.append(f"paragraph[{self.paragraph_index}]")
        if (
            self.cell_row is not None
            and self.cell_column is not None
            and (not self.context or "cell(" not in self.context)
        ):
            parts.append(f"cell(row={self.cell_row}, column={self.cell_column})")

        if self.xpath:
            parts.append(f"xpath={self.xpath}")
        if self.line is not None:
            if self.column is not None:
                parts.append(f"line={self.line}, column={self.column}")
            else:
                parts.append(f"line={self.line}")
        return " ".join(parts)

    def to_dict(self, *, include_none: bool = False) -> dict[str, object]:
        payload: dict[str, object] = {
            "code": self.code or self.kind,
            "kind": self.kind,
            "message": self.message,
            "location": self.location(),
            "part_path": self.part_path,
            "section_index": self.section_index,
            "paragraph_index": self.paragraph_index,
            "cell_row": self.cell_row,
            "cell_column": self.cell_column,
            "xpath": self.xpath,
            "line": self.line,
            "column": self.column,
            "context": self.context,
        }
        if include_none:
            return payload
        return {key: value for key, value in payload.items() if value is not None and value != ""}

    def __str__(self) -> str:
        prefix = f"[{self.code or self.kind}] " if (self.code or self.kind) else ""
        location = self.location()
        if location:
            return f"{prefix}{location}: {self.message}"
        return f"{prefix}{self.message}"

    def __repr__(self) -> str:
        return str(self)


class HwpxValidationError(HwpxError):
    def __init__(self, errors: list[ValidationIssue]):
        super().__init__("\n".join(str(error) for error in errors))
        self.errors = errors

    def to_dicts(self, *, include_none: bool = False) -> list[dict[str, object]]:
        return [error.to_dict(include_none=include_none) for error in self.errors]
