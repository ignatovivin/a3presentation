from __future__ import annotations

from a3presentation.domain.chart import ChartSeries, ChartSpec, ChartType

PRIMARY_AXIS = "primary"
SECONDARY_AXIS = "secondary"


def visible_chart_series(chart_spec: ChartSpec) -> list[ChartSeries]:
    return [series for series in chart_spec.series if not series.hidden]


def _unit_key(series: ChartSeries) -> str:
    return series.unit or "number"


def _resolved_series_axes(chart_spec: ChartSpec, visible_series: list[ChartSeries]) -> list[ChartSeries]:
    if chart_spec.chart_type != ChartType.COMBO:
        return [series.model_copy(update={"axis": PRIMARY_AXIS}) for series in visible_series]

    explicit_secondary = any(series.axis == SECONDARY_AXIS for series in visible_series)
    mixed_units = len({_unit_key(series) for series in visible_series}) > 1
    auto_secondary_index = len(visible_series) - 1 if mixed_units and len(visible_series) >= 2 else None

    resolved: list[ChartSeries] = []
    for index, series in enumerate(visible_series):
        axis = SECONDARY_AXIS if series.axis == SECONDARY_AXIS else PRIMARY_AXIS
        if not explicit_secondary and auto_secondary_index is not None and index == auto_secondary_index:
            axis = SECONDARY_AXIS
        elif axis != SECONDARY_AXIS:
            axis = PRIMARY_AXIS
        resolved.append(series.model_copy(update={"axis": axis}))
    return resolved


def combo_line_visible(chart_spec: ChartSpec, visible_series: list[ChartSeries] | None = None) -> bool:
    if chart_spec.chart_type != ChartType.COMBO:
        return False
    visible = visible_series if visible_series is not None else visible_chart_series(chart_spec)
    return bool(chart_spec.series and not chart_spec.series[-1].hidden and len(visible) >= 2)


def render_chart_type(chart_spec: ChartSpec) -> ChartType | None:
    visible = _resolved_series_axes(chart_spec, visible_chart_series(chart_spec))
    if not visible:
        return None
    if chart_spec.chart_type == ChartType.COMBO:
        return ChartType.COMBO if combo_line_visible(chart_spec, visible) else ChartType.COLUMN
    return chart_spec.chart_type


def render_chart_spec(chart_spec: ChartSpec | None) -> ChartSpec | None:
    if chart_spec is None:
        return None
    visible = _resolved_series_axes(chart_spec, visible_chart_series(chart_spec))
    if not visible:
        return None
    resolved_type = render_chart_type(chart_spec)
    if resolved_type is None:
        return None
    return chart_spec.model_copy(update={"series": visible, "chart_type": resolved_type}, deep=True)


def render_chart_series_count(chart_spec: ChartSpec | None) -> int | None:
    render_spec = render_chart_spec(chart_spec)
    if render_spec is None:
        return None
    return len(render_spec.series)


def chart_series_for_axis(chart_spec: ChartSpec | None, axis: str) -> list[ChartSeries]:
    render_spec = render_chart_spec(chart_spec)
    if render_spec is None:
        return []
    return [series for series in render_spec.series if series.axis == axis]


def uses_secondary_value_axis(chart_spec: ChartSpec | None) -> bool:
    render_spec = render_chart_spec(chart_spec)
    if render_spec is None or render_spec.chart_type != ChartType.COMBO:
        return False
    return any(series.axis == SECONDARY_AXIS for series in render_spec.series)


def can_transpose_chart_spec(chart_spec: ChartSpec | None) -> bool:
    if chart_spec is None:
        return False
    if chart_spec.chart_type == ChartType.PIE:
        return False
    if len(chart_spec.categories) < 2 or len(chart_spec.series) < 2:
        return False
    if not all(len(series.values) == len(chart_spec.categories) for series in chart_spec.series):
        return False
    unit_keys = {_unit_key(series) for series in chart_spec.series}
    return len(unit_keys) <= 1


def chart_axis_number_format_for_axis(chart_spec: ChartSpec | None, axis: str = PRIMARY_AXIS) -> str | None:
    render_spec = render_chart_spec(chart_spec)
    if render_spec is None or render_spec.chart_type == ChartType.PIE:
        return None

    axis_series = [series for series in render_spec.series if series.axis == axis]
    if not axis_series:
        return None

    if all(series.unit == "%" for series in axis_series):
        return '0"%"'

    max_value = max((abs(value) for series in axis_series for value in series.values), default=0.0)
    has_currency = any(series.unit == "RUB" for series in axis_series)
    if has_currency or render_spec.value_format == "currency":
        if max_value >= 1_000_000_000:
            return '0.0,,," млрд ₽"'
        if max_value >= 1_000_000:
            return '0.0,," млн ₽"'
        if max_value >= 1_000:
            return '0," тыс ₽"'
        return '#,##0" ₽"'

    if max_value >= 1_000_000_000:
        return '0.0,,," млрд"'
    if max_value >= 1_000_000:
        return '0.0,," млн"'
    if max_value >= 1_000:
        return '0," тыс"'
    return "#,##0"


def chart_axis_number_format(chart_spec: ChartSpec | None) -> str | None:
    return chart_axis_number_format_for_axis(chart_spec, PRIMARY_AXIS)
