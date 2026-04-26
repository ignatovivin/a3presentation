from __future__ import annotations

from pydantic import BaseModel, Field

from a3presentation.domain.chart import ChartabilityAssessment, ChartSpec
from a3presentation.domain.presentation import PresentationPlan, TableBlock
from a3presentation.domain.template import ComponentGeometry, TemplateManifest


class ChartOverride(BaseModel):
    table_id: str
    mode: str = Field(pattern="^(table|chart)$")
    selected_chart: ChartSpec | None = None


class TemplateSummary(BaseModel):
    template_id: str
    display_name: str
    description: str | None = None


class InventoryTargetSummary(BaseModel):
    key: str
    name: str
    source: str = Field(pattern="^(layout|prototype|direct_shape_binding)$")
    source_label: str | None = None
    supported_slide_kinds: list[str] = Field(default_factory=list)
    representation_hints: list[str] = Field(default_factory=list)
    editable_slot_count: int = 0
    editable_roles: list[str] = Field(default_factory=list)


class TemplateInventorySummary(BaseModel):
    generation_mode: str = Field(pattern="^(layout|prototype)$")
    usability_status: str = Field(pattern="^(usable|usable_with_degradation|not_safely_editable)$", default="not_safely_editable")
    has_usable_layout_inventory: bool = False
    has_prototype_inventory: bool = False
    degradation_mode: str | None = None
    warnings: list[str] = Field(default_factory=list)
    layout_target_count: int = 0
    prototype_target_count: int = 0
    direct_target_count: int = 0
    targets: list[InventoryTargetSummary] = Field(default_factory=list)


class EditableTargetSummary(BaseModel):
    key: str
    name: str
    source: str = Field(pattern="^(layout|prototype|direct_shape_binding)$")
    source_label: str | None = None
    runtime_profile_key: str | None = None
    supported_slide_kinds: list[str] = Field(default_factory=list)
    representation_hints: list[str] = Field(default_factory=list)
    editable_slot_count: int = 0
    editable_roles: list[str] = Field(default_factory=list)


class DetectedComponentSummary(BaseModel):
    component_id: str
    source_kind: str = Field(pattern="^(layout|slide)$")
    source_index: int
    source_name: str | None = None
    shape_name: str | None = None
    component_type: str
    role: str
    binding: str | None = None
    confidence: str = Field(pattern="^(high|medium|low)$")
    editability: str = Field(pattern="^(editable|semi_editable|decorative|unsafe)$")
    capabilities: list[str] = Field(default_factory=list)
    geometry: ComponentGeometry = Field(default_factory=ComponentGeometry)
    text_excerpt: str | None = None
    child_component_ids: list[str] = Field(default_factory=list)


class TextPlanRequest(BaseModel):
    template_id: str = Field(min_length=1)
    raw_text: str = Field(min_length=1)
    title: str | None = None
    tables: list[TableBlock] = Field(default_factory=list)
    blocks: list["DocumentBlock"] = Field(default_factory=list)
    chart_overrides: list[ChartOverride] = Field(default_factory=list)


class PlanWithTemplateResponse(BaseModel):
    plan: "PresentationPlan"
    manifest: TemplateManifest
    inventory_summary: TemplateInventorySummary
    editable_targets: list[EditableTargetSummary] = Field(default_factory=list)
    detected_components: list[DetectedComponentSummary] = Field(default_factory=list)
    slide_layout_reviews: list["SlideLayoutReview"] = Field(default_factory=list)


class SlideLayoutOption(BaseModel):
    key: str
    name: str
    source: str = Field(pattern="^(layout|prototype|direct_shape_binding)$")
    source_label: str | None = None
    runtime_profile_key: str | None = None
    supported_slide_kinds: list[str] = Field(default_factory=list)
    representation_hints: list[str] = Field(default_factory=list)
    editable_slot_count: int = 0
    editable_roles: list[str] = Field(default_factory=list)
    supports_current_slide_kind: bool = False
    estimated_text_capacity_chars: int | None = None
    match_summary: str | None = None
    recommendation_label: str | None = None
    recommendation_reasons: list[str] = Field(default_factory=list)


class SlideLayoutReview(BaseModel):
    slide_index: int
    current_layout_key: str | None = None
    current_target_key: str | None = None
    current_target_type: str | None = Field(default=None, pattern="^(layout|prototype|direct_shape_binding|auto_layout)$")
    current_target_source: str | None = None
    current_target_explanation: str | None = None
    current_target_confidence: str | None = None
    current_target_degradation_reasons: list[str] = Field(default_factory=list)
    current_runtime_profile_key: str | None = None
    available_layouts: list[SlideLayoutOption] = Field(default_factory=list)


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
    inventory_summary: TemplateInventorySummary
    editable_targets: list[EditableTargetSummary] = Field(default_factory=list)
    detected_components: list[DetectedComponentSummary] = Field(default_factory=list)


class AnalyzeTemplateResponse(BaseModel):
    template_id: str
    manifest_path: str
    inventory_summary: TemplateInventorySummary
    editable_targets: list[EditableTargetSummary] = Field(default_factory=list)
    detected_components: list[DetectedComponentSummary] = Field(default_factory=list)


class AutoUploadTemplateResponse(BaseModel):
    template_id: str
    manifest_path: str
    template_path: str
    analyzed: bool = True
    inventory_summary: TemplateInventorySummary
    editable_targets: list[EditableTargetSummary] = Field(default_factory=list)
    detected_components: list[DetectedComponentSummary] = Field(default_factory=list)


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
