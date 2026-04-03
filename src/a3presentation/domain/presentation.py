from __future__ import annotations

from enum import Enum

from pydantic import BaseModel, Field

from a3presentation.domain.chart import ChartSpec


class SlideKind(str, Enum):
    TITLE = "title"
    BULLETS = "bullets"
    TEXT = "text"
    TWO_COLUMN = "two_column"
    TABLE = "table"
    CHART = "chart"
    IMAGE = "image"


class BulletBlock(BaseModel):
    heading: str | None = None
    items: list[str] = Field(default_factory=list)


class TableBlock(BaseModel):
    headers: list[str] = Field(default_factory=list)
    rows: list[list[str]] = Field(default_factory=list)


class SlideSpec(BaseModel):
    kind: SlideKind
    title: str | None = None
    subtitle: str | None = None
    text: str | None = None
    bullets: list[str] = Field(default_factory=list)
    left_bullets: list[str] = Field(default_factory=list)
    right_bullets: list[str] = Field(default_factory=list)
    table: TableBlock | None = None
    chart: ChartSpec | None = None
    source_table_id: str | None = None
    notes: str | None = None
    preferred_layout_key: str | None = None
    image_base64: str | None = None
    image_content_type: str | None = None


class PresentationPlan(BaseModel):
    template_id: str
    title: str
    author: str | None = None
    subject: str | None = None
    slides: list[SlideSpec] = Field(default_factory=list)
