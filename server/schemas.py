from typing import Annotated, Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator


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


class PageLayout(BaseModel):
    model_config = ConfigDict(extra="forbid", allow_inf_nan=False)
    width_mm: float = Field(default=210, ge=80, le=600)
    height_mm: float = Field(default=297, ge=80, le=800)
    left_mm: float = Field(default=20, ge=0, le=100)
    right_mm: float = Field(default=20, ge=0, le=100)
    top_mm: float = Field(default=20, ge=0, le=100)
    bottom_mm: float = Field(default=20, ge=0, le=100)
    header_mm: float = Field(default=10, ge=0, le=60)
    footer_mm: float = Field(default=10, ge=0, le=60)
    columns: Literal[1, 2, 3] = 1
    gap_mm: float = Field(default=8, ge=0, le=40)
    separator: bool = False
    header_text: str | None = Field(default=None, max_length=200)
    footer_text: str | None = Field(default=None, max_length=200)
    page_numbers: Literal["keep", "none", "bottom-center", "bottom-right"] = "keep"

    @model_validator(mode="after")
    def usable_area(self):
        if (self.width_mm - self.left_mm - self.right_mm - self.gap_mm * (self.columns - 1)) / self.columns < 20:
            raise ValueError("각 단의 본문 너비는 20mm 이상이어야 합니다.")
        if self.height_mm - self.top_mm - self.bottom_mm - self.header_mm - self.footer_mm < 40:
            raise ValueError("본문 높이는 40mm 이상이어야 합니다.")
        if self.header_mm > self.top_mm or self.footer_mm > self.bottom_mm:
            raise ValueError("머리말·꼬리말 위치는 위·아래 여백 이내여야 합니다.")
        for text in (self.header_text, self.footer_text):
            if text and any(ord(c) < 32 or 0xD800 <= ord(c) <= 0xDFFF for c in text):
                raise ValueError("머리말·꼬리말에는 줄바꿈이나 제어 문자를 사용할 수 없습니다.")
        return self


class PageLayoutRequest(CommandRevisionRequest):
    section: int | None = Field(default=None, ge=0)
    layout: PageLayout
    background_pages: list['BackgroundPage'] | None = Field(default=None, max_length=10)


class PageLayoutDocumentRequest(BaseModel):
    model_config = ConfigDict(extra='forbid')
    title: str = Field(min_length=1, max_length=100)
    layout: PageLayout
    background_pages: list['BackgroundPage'] | None = Field(default=None, max_length=10)


class BackgroundPage(BaseModel):
    model_config = ConfigDict(extra='forbid', allow_inf_nan=False)
    width_mm: float = Field(ge=80, le=600)
    height_mm: float = Field(ge=80, le=800)
    image: str = Field(min_length=30, max_length=6_000_000)


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


