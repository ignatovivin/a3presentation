from __future__ import annotations

from enum import Enum

from pydantic import BaseModel, Field


class GenerationMode(str, Enum):
    LAYOUT = "layout"
    PROTOTYPE = "prototype"


class PlaceholderKind(str, Enum):
    TITLE = "title"
    SUBTITLE = "subtitle"
    BODY = "body"
    IMAGE = "image"
    TABLE = "table"
    CHART = "chart"
    FOOTER = "footer"
    UNKNOWN = "unknown"


class PlaceholderSpec(BaseModel):
    name: str
    kind: PlaceholderKind = PlaceholderKind.UNKNOWN
    idx: int | None = None
    shape_name: str | None = None
    binding: str | None = None
    max_chars: int | None = None
    left_emu: int | None = None
    top_emu: int | None = None
    width_emu: int | None = None
    height_emu: int | None = None
    margin_left_emu: int | None = None
    margin_right_emu: int | None = None
    margin_top_emu: int | None = None
    margin_bottom_emu: int | None = None


class LayoutSpec(BaseModel):
    key: str
    name: str
    slide_master_index: int = 0
    slide_layout_index: int
    preview_path: str | None = None
    supported_slide_kinds: list[str] = Field(default_factory=list)
    placeholders: list[PlaceholderSpec] = Field(default_factory=list)


class PrototypeTokenSpec(BaseModel):
    token: str
    binding: str
    shape_name: str | None = None
    left_emu: int | None = None
    top_emu: int | None = None
    width_emu: int | None = None
    height_emu: int | None = None
    margin_left_emu: int | None = None
    margin_right_emu: int | None = None
    margin_top_emu: int | None = None
    margin_bottom_emu: int | None = None


class PrototypeSlideSpec(BaseModel):
    key: str
    name: str
    source_slide_index: int
    supported_slide_kinds: list[str] = Field(default_factory=list)
    tokens: list[PrototypeTokenSpec] = Field(default_factory=list)


class TemplateManifest(BaseModel):
    template_id: str
    display_name: str
    source_pptx: str = "template.pptx"
    description: str | None = None
    generation_mode: GenerationMode = GenerationMode.LAYOUT
    default_layout_key: str | None = None
    layouts: list[LayoutSpec] = Field(default_factory=list)
    prototype_slides: list[PrototypeSlideSpec] = Field(default_factory=list)
