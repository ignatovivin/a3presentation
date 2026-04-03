from __future__ import annotations

from pydantic import BaseModel, Field

from a3presentation.domain.chart import ChartabilityAssessment
from a3presentation.domain.presentation import TableBlock
from a3presentation.domain.template import TemplateManifest


class TemplateSummary(BaseModel):
    template_id: str
    display_name: str
    description: str | None = None


class TextPlanRequest(BaseModel):
    template_id: str
    raw_text: str = Field(min_length=1)
    title: str | None = None
    tables: list[TableBlock] = Field(default_factory=list)
    blocks: list["DocumentBlock"] = Field(default_factory=list)


class GeneratePresentationResponse(BaseModel):
    output_path: str
    file_name: str
    download_url: str


class UploadTemplateResponse(BaseModel):
    template_id: str
    manifest_path: str
    template_path: str


class TemplateDetailsResponse(BaseModel):
    manifest: TemplateManifest
    has_template_file: bool


class AnalyzeTemplateResponse(BaseModel):
    template_id: str
    manifest_path: str


class AutoUploadTemplateResponse(BaseModel):
    template_id: str
    manifest_path: str
    template_path: str
    analyzed: bool = True


class DocumentBlock(BaseModel):
    kind: str
    text: str | None = None
    level: int | None = None
    style_name: str | None = None
    style_id: str | None = None
    items: list[str] = Field(default_factory=list)
    table: TableBlock | None = None
    hyperlinks: list[str] = Field(default_factory=list)
    run_count: int | None = None
    image_name: str | None = None
    image_content_type: str | None = None
    image_base64: str | None = None


class ExtractTextResponse(BaseModel):
    file_name: str
    text: str
    tables: list[TableBlock] = Field(default_factory=list)
    blocks: list[DocumentBlock] = Field(default_factory=list)
    chart_assessments: list[ChartabilityAssessment] = Field(default_factory=list)
