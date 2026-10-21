from typing import Annotated, Literal

from pydantic import BaseModel, ConfigDict, Field


class TemplateRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    template: Literal["blank", "report", "meeting", "table", "exam", "worksheet"]
    title: str = Field(min_length=1, max_length=100)
    question_count: int = Field(default=3, ge=1, le=20)


class ParagraphEdit(BaseModel):
    model_config = ConfigDict(extra="forbid")
    type: Literal["paragraph"]
    section: int = Field(ge=0)
    index: int = Field(ge=0)
    text: str = Field(max_length=100_000)


class CellEdit(BaseModel):
    model_config = ConfigDict(extra="forbid")
    type: Literal["cell"]
    section: int = Field(ge=0)
    table: int = Field(ge=0)
    row: int = Field(ge=0)
    column: int = Field(ge=0)
    text: str = Field(max_length=100_000)


class RevisionRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    revision: int = Field(ge=0)


class CommandRevisionRequest(RevisionRequest):
    settings_version: int | None = Field(default=None, ge=0)


class EditRequest(CommandRevisionRequest):
    operations: list[Annotated[ParagraphEdit | CellEdit, Field(discriminator="type")]] = Field(min_length=1, max_length=500)


class RenameRequest(RevisionRequest):
    expected_name: str = Field(min_length=1, max_length=180)
    name: str = Field(min_length=1, max_length=160)


class CopyRequest(RevisionRequest):
    name: str = Field(min_length=1, max_length=160)


class RestoreRequest(CommandRevisionRequest):
    source_revision: int = Field(ge=0)


class AdvancedSettings(BaseModel):
    model_config = ConfigDict(extra="forbid")
    instructions: str = Field(default="", max_length=8_000)
    protected_terms: list[Annotated[str, Field(min_length=1, max_length=200)]] = Field(default_factory=list, max_length=100)
    strict_validation: bool = False


class SettingsRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    version: int = Field(ge=0)
    settings: AdvancedSettings


