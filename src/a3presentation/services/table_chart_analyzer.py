from __future__ import annotations

import re
from dataclasses import dataclass

from a3presentation.domain.chart import (
    ChartConfidence,
    ChartSeries,
    ChartSpec,
    ChartTableClassification,
    ChartType,
    ChartValueType,
    ChartabilityAssessment,
    StructuredCell,
    StructuredTable,
)
from a3presentation.domain.presentation import TableBlock


@dataclass
class _ParsedValue:
    value_type: ChartValueType
    parsed_value: float | None = None
    unit: str | None = None
    annotation: str | None = None


class TableChartAnalyzer:
    SUMMARY_LABELS = {
        "итого",
        "всего",
        "итог",
        "год",
        "цель",
        "target",
        "total",
        "grand total",
    }
    QUARTER_PATTERN = re.compile(r"^q[1-4](?:\s*[\(\-–—].*)?(?:\s+\d{4})?$", re.IGNORECASE)
    YEAR_PATTERN = re.compile(r"^(19|20)\d{2}$")
    MONTH_NAMES = {
        "январь",
        "февраль",
        "март",
        "апрель",
        "май",
        "июнь",
        "июль",
        "август",
        "сентябрь",
        "октябрь",
        "ноябрь",
        "декабрь",
        "january",
        "february",
        "march",
        "april",
        "may",
        "june",
        "july",
        "august",
        "september",
        "october",
        "november",
        "december",
    }
    NUMBER_PATTERN = re.compile(r"[-+]?\d+(?:[.,]\d+)?")

    def analyze(self, table: TableBlock, *, table_id: str = "table_1") -> ChartabilityAssessment:
        structured = self._build_structured_table(table, table_id=table_id)
        if not structured.cells or len(structured.cells) < 3:
            return ChartabilityAssessment(
                table_id=table_id,
                chartable=False,
                classification=ChartTableClassification.NOT_CHARTABLE,
                confidence=ChartConfidence.NONE,
                reasons=["table has too few rows for chart analysis"],
                structured_table=structured,
            )

        data_rows = [row for idx, row in enumerate(structured.cells[structured.data_start_row :], start=structured.data_start_row) if idx not in structured.summary_rows]
        if not data_rows:
            return ChartabilityAssessment(
                table_id=table_id,
                chartable=False,
                classification=ChartTableClassification.NOT_CHARTABLE,
                confidence=ChartConfidence.NONE,
                reasons=["table has no data rows after summary filtering"],
                warnings=structured.warnings.copy(),
                structured_table=structured,
            )

        numeric_columns = structured.numeric_columns
        if not numeric_columns:
            classification = ChartTableClassification.TEXT_DOMINANT
            return ChartabilityAssessment(
                table_id=table_id,
                chartable=False,
                classification=classification,
                confidence=ChartConfidence.NONE,
                reasons=["table has no numeric columns"],
                warnings=structured.warnings.copy(),
                structured_table=structured,
            )

        label_column = structured.label_columns[0] if structured.label_columns else None
        if label_column is None:
            return ChartabilityAssessment(
                table_id=table_id,
                chartable=False,
                classification=ChartTableClassification.MIXED_AMBIGUOUS,
                confidence=ChartConfidence.LOW,
                reasons=["table has no clear category axis"],
                warnings=structured.warnings.copy(),
                structured_table=structured,
            )

        categories = [row[label_column].text.strip() for row in data_rows if label_column < len(row)]
        if len(categories) < 2:
            return ChartabilityAssessment(
                table_id=table_id,
                chartable=False,
                classification=ChartTableClassification.NOT_CHARTABLE,
                confidence=ChartConfidence.NONE,
                reasons=["table has too few categories"],
                warnings=structured.warnings.copy(),
                structured_table=structured,
            )

        classification = self._classify(structured, categories)
        candidates = self._build_candidates(structured, categories, data_rows, classification)
        chartable = bool(candidates)
        confidence = candidates[0].confidence if candidates else ChartConfidence.NONE
        reasons = self._build_reasons(classification, structured, categories, candidates)

        return ChartabilityAssessment(
            table_id=table_id,
            chartable=chartable,
            classification=classification if chartable else ChartTableClassification.NOT_CHARTABLE,
            confidence=confidence if chartable else ChartConfidence.NONE,
            reasons=reasons,
            warnings=structured.warnings.copy(),
            candidate_specs=candidates,
            structured_table=structured,
        )

    def _build_structured_table(self, table: TableBlock, *, table_id: str) -> StructuredTable:
        raw_rows: list[list[str]] = []
        if table.headers:
            raw_rows.append(table.headers)
        raw_rows.extend(table.rows)
        width = max((len(row) for row in raw_rows), default=0)
        padded_rows = [row + [""] * (width - len(row)) for row in raw_rows]
        cells = [
            [self._build_cell(cell_text, is_header_like=(row_index == 0)) for cell_text in row]
            for row_index, row in enumerate(padded_rows)
        ]

        numeric_columns = []
        time_columns = []
        label_columns = []
        summary_rows = []
        warnings: list[str] = []

        for column_index in range(width):
            column_cells = [row[column_index] for row in cells[1:]] if len(cells) > 1 else []
            numeric_count = sum(cell.parsed_value is not None for cell in column_cells)
            time_count = sum(cell.value_type in {ChartValueType.QUARTER, ChartValueType.YEAR, ChartValueType.DATE} for cell in column_cells)
            text_count = sum(cell.value_type == ChartValueType.TEXT for cell in column_cells)

            if numeric_count >= max(2, len(column_cells) // 2 if column_cells else 0):
                numeric_columns.append(column_index)
            elif time_count >= max(2, len(column_cells) // 2 if column_cells else 0):
                time_columns.append(column_index)
            elif text_count:
                label_columns.append(column_index)

        for row_index, row in enumerate(cells[1:], start=1):
            if row and row[0].normalized_text in self.SUMMARY_LABELS:
                summary_rows.append(row_index)
                warnings.append(f"summary row filtered: {row[0].text}")

        if not label_columns and width:
            label_columns = [0]

        return StructuredTable(
            table_id=table_id,
            header_rows=[0] if cells else [],
            label_columns=label_columns,
            numeric_columns=numeric_columns,
            time_columns=time_columns,
            data_start_row=1 if len(cells) > 1 else 0,
            cells=cells,
            summary_rows=summary_rows,
            warnings=warnings,
        )

    def _build_cell(self, text: str, *, is_header_like: bool) -> StructuredCell:
        normalized = " ".join(text.split()).strip()
        parsed = self._parse_value(normalized)
        return StructuredCell(
            text=text,
            normalized_text=normalized.lower(),
            value_type=parsed.value_type,
            parsed_value=parsed.parsed_value,
            unit=parsed.unit,
            annotation=parsed.annotation,
            is_header_like=is_header_like,
        )

    def _parse_value(self, text: str) -> _ParsedValue:
        if not text:
            return _ParsedValue(value_type=ChartValueType.EMPTY)

        lowered = text.lower().strip()
        if self.QUARTER_PATTERN.match(lowered):
            return _ParsedValue(value_type=ChartValueType.QUARTER)
        if self.YEAR_PATTERN.match(lowered):
            return _ParsedValue(value_type=ChartValueType.YEAR, parsed_value=float(lowered))
        if lowered in self.MONTH_NAMES:
            return _ParsedValue(value_type=ChartValueType.DATE)

        number_match = self.NUMBER_PATTERN.search(lowered.replace(" ", ""))
        if not number_match:
            return _ParsedValue(value_type=ChartValueType.TEXT)

        numeric_text = number_match.group(0).replace(",", ".")
        value = float(numeric_text)
        unit: str | None = None
        value_type = ChartValueType.NUMBER

        if "%" in lowered:
            value_type = ChartValueType.PERCENT
            unit = "%"
        elif "₽" in lowered or "руб" in lowered:
            value_type = ChartValueType.CURRENCY
            unit = "RUB"
            if "млрд" in lowered:
                value *= 1_000_000_000
            elif "млн" in lowered:
                value *= 1_000_000
            elif "тыс" in lowered:
                value *= 1_000
        elif "млрд" in lowered:
            value *= 1_000_000_000
        elif "млн" in lowered:
            value *= 1_000_000
        elif "тыс" in lowered:
            value *= 1_000

        annotation = None
        if "—" in text:
            annotation = text.split("—", 1)[1].strip()

        return _ParsedValue(value_type=value_type, parsed_value=value, unit=unit, annotation=annotation)

    def _classify(self, structured: StructuredTable, categories: list[str]) -> ChartTableClassification:
        if structured.time_columns and structured.label_columns and structured.label_columns[0] in structured.time_columns:
            return ChartTableClassification.TIME_SERIES
        if self._looks_like_time_axis(categories):
            return ChartTableClassification.TIME_SERIES
        if len(structured.numeric_columns) == 1:
            return ChartTableClassification.RANKING if len(categories) > 6 else ChartTableClassification.SINGLE_SERIES_CATEGORY
        if len(structured.numeric_columns) > 1:
            return ChartTableClassification.MULTI_SERIES_CATEGORY
        return ChartTableClassification.MIXED_AMBIGUOUS

    def _looks_like_time_axis(self, categories: list[str]) -> bool:
        matches = 0
        for category in categories:
            lowered = category.lower().strip()
            if (
                self.QUARTER_PATTERN.match(lowered)
                or self.YEAR_PATTERN.match(lowered)
                or lowered in self.MONTH_NAMES
                or any(month in lowered for month in self.MONTH_NAMES)
            ):
                matches += 1
        return matches >= max(2, len(categories) // 2)

    def _build_candidates(
        self,
        structured: StructuredTable,
        categories: list[str],
        data_rows: list[list[StructuredCell]],
        classification: ChartTableClassification,
    ) -> list[ChartSpec]:
        series: list[ChartSeries] = []
        units = set()
        for column_index in structured.numeric_columns:
            values: list[float] = []
            for row in data_rows:
                if column_index >= len(row) or row[column_index].parsed_value is None:
                    break
                values.append(row[column_index].parsed_value)
                if row[column_index].unit:
                    units.add(row[column_index].unit)
            else:
                if len(values) == len(categories):
                    series_name = self._series_name(structured, column_index)
                    series.append(ChartSeries(name=series_name, values=values, unit=self._series_unit(data_rows, column_index)))

        if not series:
            return []

        confidence = ChartConfidence.HIGH
        warnings: list[str] = []
        if len(categories) > 15:
            confidence = ChartConfidence.LOW
            warnings.append("too many categories for default chart suggestion")
        elif len(series) > 4:
            confidence = ChartConfidence.MEDIUM
            warnings.append("multiple series detected")

        if len(units) > 1:
            confidence = ChartConfidence.LOW
            warnings.append("mixed units detected across series")

        suggested_types: list[ChartType]
        if classification == ChartTableClassification.TIME_SERIES:
            suggested_types = [ChartType.LINE, ChartType.COLUMN]
        elif classification == ChartTableClassification.RANKING:
            suggested_types = [ChartType.BAR]
        elif classification == ChartTableClassification.SINGLE_SERIES_CATEGORY:
            suggested_types = [ChartType.COLUMN, ChartType.BAR]
        else:
            suggested_types = [ChartType.COLUMN, ChartType.LINE]

        value_format = "currency" if any(unit == "RUB" for unit in units) else "percent" if "%" in units else "number"
        y_axis_title = self._series_name(structured, structured.numeric_columns[0]) if len(series) == 1 else None

        return [
            ChartSpec(
                chart_id=f"{structured.table_id}_{chart_type.value}",
                source_table_id=structured.table_id,
                chart_type=chart_type,
                categories=categories,
                series=series,
                x_axis_title=self._series_name(structured, structured.label_columns[0]) if structured.label_columns else None,
                y_axis_title=y_axis_title,
                legend_visible=len(series) > 1,
                value_format=value_format,
                confidence=confidence,
                warnings=warnings.copy(),
            )
            for chart_type in suggested_types
        ]

    def _series_name(self, structured: StructuredTable, column_index: int) -> str:
        if structured.cells and structured.header_rows and column_index < len(structured.cells[0]):
            header = structured.cells[0][column_index].text.strip()
            if header:
                return header
        return f"Series {column_index + 1}"

    def _series_unit(self, rows: list[list[StructuredCell]], column_index: int) -> str | None:
        for row in rows:
            if column_index < len(row) and row[column_index].unit:
                return row[column_index].unit
        return None

    def _build_reasons(
        self,
        classification: ChartTableClassification,
        structured: StructuredTable,
        categories: list[str],
        candidates: list[ChartSpec],
    ) -> list[str]:
        reasons = [f"classified as {classification.value}"]
        reasons.append(f"detected {len(categories)} categories")
        reasons.append(f"detected {len(structured.numeric_columns)} numeric columns")
        if candidates:
            reasons.append(f"generated {len(candidates)} chart candidates")
        return reasons
