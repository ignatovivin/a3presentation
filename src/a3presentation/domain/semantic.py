from __future__ import annotations

from enum import Enum

from pydantic import BaseModel, Field

from a3presentation.domain.presentation import TableBlock


class DocumentKind(str, Enum):
    REPORT = "report"
    FORM = "form"
    TABLE_HEAVY = "table_heavy"
    RESUME = "resume"
    MIXED = "mixed"
    UNKNOWN = "unknown"


class DocumentStats(BaseModel):
    paragraph_count: int = 0
    heading_count: int = 0
    list_count: int = 0
    table_count: int = 0
    image_count: int = 0
    character_count: int = 0
    fact_count: int = 0
    contact_count: int = 0
    date_count: int = 0
    signature_count: int = 0


class SemanticFact(BaseModel):
    label: str
    value: str
    confidence: float = 0.5
    source_text: str | None = None


class SemanticImage(BaseModel):
    name: str | None = None
    alt_text: str | None = None
    content_type: str | None = None
    image_base64: str | None = None
    width_px: int | None = None
    height_px: int | None = None


class SemanticSection(BaseModel):
    title: str
    level: int = 1
    subtitle: str | None = None
    paragraphs: list[str] = Field(default_factory=list)
    bullets: list[str] = Field(default_factory=list)
    tables: list[TableBlock] = Field(default_factory=list)
    facts: list[SemanticFact] = Field(default_factory=list)
    images: list[SemanticImage] = Field(default_factory=list)


class SemanticDocument(BaseModel):
    title: str
    kind: DocumentKind = DocumentKind.UNKNOWN
    summary_lines: list[str] = Field(default_factory=list)
    facts: list[SemanticFact] = Field(default_factory=list)
    contacts: list[str] = Field(default_factory=list)
    dates: list[str] = Field(default_factory=list)
    signatures: list[str] = Field(default_factory=list)
    images: list[SemanticImage] = Field(default_factory=list)
    sections: list[SemanticSection] = Field(default_factory=list)
    loose_tables: list[TableBlock] = Field(default_factory=list)
    stats: DocumentStats = Field(default_factory=DocumentStats)
