from typing import Annotated, Literal

from pydantic import BaseModel, ConfigDict, Field


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


class EditRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    revision: int = Field(ge=0)
    operations: list[Annotated[ParagraphEdit | CellEdit, Field(discriminator="type")]] = Field(min_length=1, max_length=500)


class RevisionRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    revision: int = Field(ge=0)


class RestoreRequest(RevisionRequest):
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


