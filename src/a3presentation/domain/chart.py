from __future__ import annotations

from enum import Enum

from pydantic import BaseModel, Field


class ChartValueType(str, Enum):
    TEXT = "text"
    NUMBER = "number"
    CURRENCY = "currency"
    PERCENT = "percent"
    DATE = "date"
    QUARTER = "quarter"
    YEAR = "year"
    EMPTY = "empty"
    MIXED = "mixed"


class ChartType(str, Enum):
    BAR = "bar"
    COLUMN = "column"
    LINE = "line"
    STACKED_BAR = "stacked_bar"
    STACKED_COLUMN = "stacked_column"
    COMBO = "combo"
    PIE = "pie"


class ChartConfidence(str, Enum):
    HIGH = "high"
    MEDIUM = "medium"
    LOW = "low"
    NONE = "none"


class ChartTableClassification(str, Enum):
    SINGLE_SERIES_CATEGORY = "single_series_category"
    MULTI_SERIES_CATEGORY = "multi_series_category"
    TIME_SERIES = "time_series"
    COMPOSITION = "composition"
    RANKING = "ranking"
    MATRIX_NUMERIC = "matrix_numeric"
    TEXT_DOMINANT = "text_dominant"
    MIXED_AMBIGUOUS = "mixed_ambiguous"
    NOT_CHARTABLE = "not_chartable"


class StructuredCell(BaseModel):
    text: str
    normalized_text: str
    value_type: ChartValueType
    parsed_value: float | None = None
    unit: str | None = None
    annotation: str | None = None
    is_header_like: bool = False


class StructuredTable(BaseModel):
    table_id: str
    header_rows: list[int] = Field(default_factory=list)
    label_columns: list[int] = Field(default_factory=list)
    numeric_columns: list[int] = Field(default_factory=list)
    time_columns: list[int] = Field(default_factory=list)
    data_start_row: int = 0
    cells: list[list[StructuredCell]] = Field(default_factory=list)
    summary_rows: list[int] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)


class ChartSeries(BaseModel):
    name: str
    values: list[float]
    unit: str | None = None
    axis: str = "primary"
    hidden: bool = False


class ChartSpec(BaseModel):
    chart_id: str
    source_table_id: str
    chart_type: ChartType
    title: str | None = None
    categories: list[str] = Field(default_factory=list)
    series: list[ChartSeries] = Field(default_factory=list)
    x_axis_title: str | None = None
    y_axis_title: str | None = None
    legend_visible: bool = True
    data_labels_visible: bool = False
    value_format: str = "number"
    stacking: str = "none"
    confidence: ChartConfidence = ChartConfidence.NONE
    warnings: list[str] = Field(default_factory=list)


class ChartabilityAssessment(BaseModel):
    table_id: str
    chartable: bool
    classification: ChartTableClassification
    confidence: ChartConfidence
    reasons: list[str] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)
    candidate_specs: list[ChartSpec] = Field(default_factory=list)
    structured_table: StructuredTable | None = None
