import type { CSSProperties } from "react";
import type { ChartSeries, ChartSpec, ChartType } from "@/types";
import chartStyle from "@/chart-style.json";

type ChartPreviewProps = {
  spec: ChartSpec;
};

type AxisRole = "primary" | "secondary";

type RenderChartSpec = ChartSpec & {
  series: ChartSeries[];
  chart_type: ChartType;
};

type ScaleDomain = {
  min: number;
  max: number;
  span: number;
};

function visibleSeries(spec: ChartSpec): ChartSeries[] {
  return spec.series.filter((series) => !series.hidden);
}

function unitKey(series: ChartSeries): string {
  return series.unit ?? "number";
}

function resolveSeriesAxes(spec: ChartSpec, series: ChartSeries[]): ChartSeries[] {
  if (spec.chart_type !== "combo") {
    const explicitSecondary = series.some((item) => item.axis === "secondary");
    return series.map((item) => ({ ...item, axis: explicitSecondary && item.axis === "secondary" ? "secondary" : "primary" }));
  }

  const explicitSecondary = series.some((item) => item.axis === "secondary");
  const mixedUnits = new Set(series.map(unitKey)).size > 1;
  const autoSecondaryIndex = mixedUnits && series.length >= 2 ? series.length - 1 : null;

  return series.map((item, index) => {
    let axis: AxisRole = item.axis === "secondary" ? "secondary" : "primary";
    if (!explicitSecondary && autoSecondaryIndex !== null && index === autoSecondaryIndex) {
      axis = "secondary";
    } else if (axis !== "secondary") {
      axis = "primary";
    }
    return { ...item, axis };
  });
}

function comboLineVisible(spec: ChartSpec, series: ChartSeries[]): boolean {
  if (spec.chart_type !== "combo") {
    return false;
  }
  return Boolean(spec.series.length && !spec.series[spec.series.length - 1].hidden && series.length >= 2);
}

function renderChartSpec(spec: ChartSpec): RenderChartSpec | null {
  const visible = visibleSeries(spec);
  if (!visible.length) {
    return null;
  }

  const resolvedSeries = resolveSeriesAxes(spec, visible);
  const shouldRenderCombo = comboLineVisible(spec, resolvedSeries);
  const resolvedType = spec.chart_type === "combo" && !shouldRenderCombo ? "column" : spec.chart_type;
  const normalizedSeries = resolvedSeries;

  return {
    ...spec,
    chart_type: resolvedType,
    series: normalizedSeries,
  };
}

function seriesForAxis(spec: RenderChartSpec, axis: AxisRole): ChartSeries[] {
  return spec.series.filter((series) => series.axis === axis);
}

function usesSecondaryAxis(spec: RenderChartSpec): boolean {
  return spec.series.some((series) => series.axis === "secondary");
}

function scaleDomain(series: ChartSeries[]): ScaleDomain {
  const values = series.flatMap((item) => item.values);
  if (!values.length) {
    return { min: 0, max: 1, span: 1 };
  }
  const minValue = Math.min(...values, 0);
  const maxValue = Math.max(...values, 1);
  const span = Math.max(maxValue - minValue, 1);
  return { min: minValue, max: maxValue, span };
}

function tickValuesForDomain(domain: ScaleDomain, gridSteps: number): number[] {
  return Array.from({ length: gridSteps + 1 }, (_, index) => {
    const ratio = (gridSteps - index) / gridSteps;
    return domain.min + domain.span * ratio;
  });
}

function scaleDomainForStacked(series: ChartSeries[], categoryCount: number): ScaleDomain {
  const maxValue = Math.max(
    1,
    ...Array.from({ length: categoryCount }, (_, categoryIndex) =>
      series.reduce((sum, item) => sum + Math.max(0, item.values[categoryIndex] ?? 0), 0),
    ),
  );
  return { min: 0, max: maxValue, span: maxValue };
}

function positiveHeightPercent(value: number, domain: ScaleDomain, minimum = 6): number {
  if (domain.min < 0) {
    return Math.max(minimum, (Math.abs(value) / domain.span) * 100);
  }
  return Math.max(minimum, (value / Math.max(domain.max, 1)) * 100);
}

function pointBottomPercent(value: number, domain: ScaleDomain): number {
  return Math.max(0, Math.min(100, ((value - domain.min) / domain.span) * 100));
}

function pointXPercent(index: number, valueCount: number): number {
  return valueCount === 1 ? 50 : ((index + 0.5) / Math.max(valueCount, 1)) * 100;
}

function compactNumber(value: number): string {
  return new Intl.NumberFormat("ru-RU", {
    notation: "compact",
    maximumFractionDigits: Math.abs(value) >= 10 ? 0 : 1,
  }).format(value);
}

function formatSeriesValue(value: number, series: ChartSeries, fallbackFormat: string): string {
  const effectiveFormat = series.unit === "%" ? "percent" : series.unit === "RUB" ? "currency" : fallbackFormat;
  if (effectiveFormat === "percent") {
    return `${Number.isInteger(value) ? value.toFixed(0) : value.toFixed(value >= 10 ? 0 : 1)}%`;
  }
  if (effectiveFormat === "currency") {
    return `${compactNumber(value)} ₽`;
  }
  return compactNumber(value);
}

function axisKindForSeries(spec: RenderChartSpec, series: ChartSeries[]): "percent" | "currency" | "number" {
  if (series.length && series.every((item) => item.unit === "%")) {
    return "percent";
  }
  if (series.some((item) => item.unit === "RUB") || spec.value_format === "currency") {
    return "currency";
  }
  return "number";
}

function formatAxisValue(value: number, spec: RenderChartSpec, series: ChartSeries[]): string {
  const axisKind = axisKindForSeries(spec, series);
  if (axisKind === "percent") {
    return `${Number.isInteger(value) ? value.toFixed(0) : value.toFixed(value >= 10 ? 0 : 1)}%`;
  }

  const absValue = Math.abs(value);
  const suffix = axisKind === "currency" ? " ₽" : "";
  if (absValue >= 1_000_000_000) {
    return `${(value / 1_000_000_000).toFixed(absValue >= 10_000_000_000 ? 0 : 1)} млрд${suffix}`;
  }
  if (absValue >= 1_000_000) {
    return `${(value / 1_000_000).toFixed(absValue >= 10_000_000 ? 0 : 1)} млн${suffix}`;
  }
  if (absValue >= 1_000) {
    return `${(value / 1_000).toFixed(absValue >= 10_000 ? 0 : 1)} тыс${suffix}`;
  }
  if (axisKind === "currency") {
    return `${Math.round(value)} ₽`;
  }
  return `${Math.round(value)}`;
}

function shortenCategoryLabel(label: string, maxChars: number): string {
  if (label.length <= maxChars) {
    return label;
  }
  return `${label.slice(0, Math.max(0, maxChars - 1)).trimEnd()}…`;
}

function categoryTotal(series: ChartSeries[], categoryIndex: number): number {
  return series.reduce((sum, item) => sum + (item.values[categoryIndex] ?? 0), 0);
}

function compactSeriesMagnitude(series: ChartSeries): number {
  return Math.max(...series.values.map((value) => Math.abs(value)), 0);
}

export function ChartPreview({ spec }: ChartPreviewProps) {
  const renderSpec = renderChartSpec(spec);
  if (!renderSpec) {
    return <div className="chart-empty-state">Нет доступных рядов для визуализации.</div>;
  }

  const visible = renderSpec.series;
  const primarySeries = seriesForAxis(renderSpec, "primary");
  const secondarySeries = seriesForAxis(renderSpec, "secondary");
  const hasSecondaryAxis = usesSecondaryAxis(renderSpec) && secondarySeries.length > 0;
  const primaryDomain = scaleDomain(primarySeries.length ? primarySeries : visible);
  const secondaryDomain = scaleDomain(secondarySeries.length ? secondarySeries : visible);
  const pieSeries = renderSpec.chart_type === "pie" ? visible[0] ?? null : null;
  const pieTotal = pieSeries ? Math.max(pieSeries.values.reduce((sum, value) => sum + value, 0), 1) : 1;
  const categoryCount = renderSpec.categories.length;
  const seriesCount = visible.length;
  const gridSteps = 4;
  const isPieChart = renderSpec.chart_type === "pie";
  const isHorizontalBarChart = renderSpec.chart_type === "bar";
  const isHorizontalStackedChart = renderSpec.chart_type === "stacked_bar";
  const isStackedChart = renderSpec.chart_type === "stacked_bar" || renderSpec.chart_type === "stacked_column";
  const stackedDomain = scaleDomainForStacked(visible, categoryCount);
  const axisDomain = isStackedChart ? stackedDomain : primaryDomain;
  const primaryTickValues = tickValuesForDomain(axisDomain, gridSteps);
  const secondaryTickValues = tickValuesForDomain(secondaryDomain, gridSteps);
  const isLineChart = renderSpec.chart_type === "line";
  const isComboChart = renderSpec.chart_type === "combo";
  const isSecondaryColumnChart = renderSpec.chart_type === "column" && hasSecondaryAxis;
  const isDense = categoryCount >= chartStyle.denseCategoryThreshold;
  const isVeryDense = categoryCount >= chartStyle.veryDenseCategoryThreshold;
  const categorySlotWidth = isVeryDense
    ? chartStyle.categorySlotWidths.veryDense
    : isDense
      ? chartStyle.categorySlotWidths.dense
      : chartStyle.categorySlotWidths.default;
  const contentWidth = Math.max(categoryCount * categorySlotWidth, 420);
  const categoryLabelMaxChars = isVeryDense
    ? chartStyle.categoryLabelMaxChars.veryDense
    : isDense
      ? chartStyle.categoryLabelMaxChars.dense
      : chartStyle.categoryLabelMaxChars.default;
  const comboLineSeries = isComboChart ? secondarySeries[secondarySeries.length - 1] ?? null : null;
  const comboBarSeries = isComboChart ? primarySeries : visible;
  const comboBarDomain = scaleDomain(comboBarSeries.length ? comboBarSeries : visible);
  const labeledLineSeriesIndex =
    isLineChart && visible.length > 1
      ? visible.reduce((lowestIndex, series, index) => {
          return compactSeriesMagnitude(series) < compactSeriesMagnitude(visible[lowestIndex]) ? index : lowestIndex;
        }, 0)
      : 0;

  function buildLinePath(values: number[], domain: ScaleDomain): string {
    return values
      .map((value, index) => {
        const x = pointXPercent(index, values.length);
        const y = 100 - pointBottomPercent(value, domain);
        return `${index === 0 ? "M" : "L"} ${x} ${y}`;
      })
      .join(" ");
  }

  function seriesColor(index: number): string {
    return chartStyle.palette[index % chartStyle.palette.length];
  }

  function buildPieSegments(values: number[]): string[] {
    let currentAngle = -90;
    return values.map((value) => {
      const angle = (value / pieTotal) * 360;
      const startAngle = currentAngle;
      const endAngle = currentAngle + angle;
      currentAngle = endAngle;

      const x1 = 50 + 40 * Math.cos((Math.PI / 180) * startAngle);
      const y1 = 50 + 40 * Math.sin((Math.PI / 180) * startAngle);
      const x2 = 50 + 40 * Math.cos((Math.PI / 180) * endAngle);
      const y2 = 50 + 40 * Math.sin((Math.PI / 180) * endAngle);
      const largeArcFlag = angle > 180 ? 1 : 0;

      return `M 50 50 L ${x1} ${y1} A 40 40 0 ${largeArcFlag} 1 ${x2} ${y2} Z`;
    });
  }

  return (
    <div
      className={`chart-preview${isDense ? " is-dense" : ""}${isVeryDense ? " is-very-dense" : ""}`}
      style={
        {
          "--chart-content-width": `${contentWidth}px`,
          "--chart-category-count": categoryCount,
          "--chart-series-count": seriesCount,
        } as CSSProperties
      }
    >
      <div className="chart-preview-head">
        <div>
          <div className="chart-preview-title">{renderSpec.title || "Превью графика"}</div>
          <div className="chart-preview-meta">
            {categoryCount} категорий · {seriesCount} {seriesCount === 1 ? "ряд" : seriesCount < 5 ? "ряда" : "рядов"}
          </div>
        </div>
        {renderSpec.legend_visible ? (
          <div className="chart-preview-legend">
            {visible.map((series, index) => (
              <span className="chart-preview-legend-item" key={series.name}>
                <span className="chart-preview-swatch" style={{ background: seriesColor(index) }} />
                {series.name}
              </span>
            ))}
          </div>
        ) : null}
      </div>

      <div className={`chart-preview-frame${isPieChart ? " is-pie" : ""}${hasSecondaryAxis ? " has-secondary-axis" : ""}`}>
        {isPieChart ? (
          <>
            <div className="chart-preview-axis chart-preview-axis-pie">
              <div className="chart-preview-axis-tick">100%</div>
              <div className="chart-preview-axis-tick">50%</div>
              <div className="chart-preview-axis-tick">0%</div>
            </div>

            <div className="chart-preview-pie-layout">
              <div className="chart-preview-pie-card">
                <svg className="chart-preview-pie-svg" viewBox="0 0 100 100" aria-hidden="true">
                  {pieSeries
                    ? buildPieSegments(pieSeries.values).map((path, index) => (
                        <path key={`${renderSpec.chart_id}-pie-${index}`} d={path} fill={seriesColor(index)} />
                      ))
                    : null}
                  <circle cx="50" cy="50" r="16" fill="#ffffff" fillOpacity="0.94" />
                </svg>
                <div className="chart-preview-pie-total">
                  <span className="chart-preview-pie-total-label">Всего</span>
                  <span className="chart-preview-pie-total-value">
                    {formatSeriesValue(pieTotal, pieSeries ?? visible[0], renderSpec.value_format)}
                  </span>
                </div>
              </div>

              <div className="chart-preview-pie-breakdown">
                {renderSpec.categories.map((category, categoryIndex) => {
                  const value = pieSeries?.values[categoryIndex] ?? 0;
                  const percent = pieTotal > 0 ? (value / pieTotal) * 100 : 0;
                  return (
                    <div className="chart-preview-pie-item" key={`${renderSpec.chart_id}-pie-item-${categoryIndex}`}>
                      <div className="chart-preview-pie-item-head">
                        <span className="chart-preview-swatch" style={{ background: seriesColor(categoryIndex) }} />
                        <span className="chart-preview-pie-item-label" title={category}>
                          {shortenCategoryLabel(category, Math.max(categoryLabelMaxChars, 18))}
                        </span>
                      </div>
                      <div className="chart-preview-pie-item-metrics">
                        <span>{formatSeriesValue(value, pieSeries ?? visible[0], renderSpec.value_format)}</span>
                        <span>{percent.toFixed(percent >= 10 ? 0 : 1)}%</span>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          </>
        ) : (
          <>
            <div className="chart-preview-axis">
              {primaryTickValues.map((value, index) => (
                <div className="chart-preview-axis-tick" key={`${renderSpec.chart_id}-tick-primary-${index}`}>
                  {formatAxisValue(value, renderSpec, isStackedChart ? visible : primarySeries.length ? primarySeries : visible)}
                </div>
              ))}
            </div>

            <div className="chart-preview-stage-scroll" style={{ minWidth: 0 }}>
              <div className="chart-preview-stage" style={{ width: `${contentWidth}px` }}>
                <div className="chart-preview-gridlines">
                  {primaryTickValues.map((_, index) => (
                    <div className="chart-preview-gridline" key={`${renderSpec.chart_id}-grid-${index}`} />
                  ))}
                </div>

                {isLineChart ? (
                  <div className="chart-preview-line-stage">
                    <svg className="chart-preview-line-svg" viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true">
                      {visible.map((series, seriesIndex) => (
                        <path
                          key={series.name}
                          d={buildLinePath(series.values, series.axis === "secondary" ? secondaryDomain : primaryDomain)}
                          className="chart-preview-line"
                          style={{ stroke: seriesColor(seriesIndex) }}
                        />
                      ))}
                    </svg>

                    <div className="chart-preview-line-marker-layer">
                      {visible.flatMap((series, seriesIndex) =>
                        series.values.map((value, categoryIndex) => {
                          const x = pointXPercent(categoryIndex, series.values.length);
                          const bottom = pointBottomPercent(value, series.axis === "secondary" ? secondaryDomain : primaryDomain);
                          const shouldShowLabel = visible.length === 1 || seriesIndex === labeledLineSeriesIndex;
                          return (
                            <div
                              className="chart-preview-line-marker"
                              key={`${series.name}-${renderSpec.categories[categoryIndex] ?? categoryIndex}-marker`}
                              style={{ left: `${x}%`, bottom: `${bottom}%` }}
                              title={`${series.name}: ${formatSeriesValue(value, series, renderSpec.value_format)}`}
                            >
                              <span className="chart-preview-point" style={{ background: seriesColor(seriesIndex) }} />
                              {shouldShowLabel ? (
                                <span className="chart-preview-line-label">
                                  {formatSeriesValue(value, series, renderSpec.value_format)}
                                </span>
                              ) : null}
                            </div>
                          );
                        }),
                      )}
                    </div>

                    <div className="chart-preview-grid line-layout">
                      {renderSpec.categories.map((category) => (
                        <div className="chart-preview-group" key={`${renderSpec.chart_id}-${category}`}>
                          <div className="chart-preview-line-spacer" aria-hidden="true" />
                          <div className="chart-preview-category" title={category}>
                            {shortenCategoryLabel(category, categoryLabelMaxChars)}
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                ) : isHorizontalStackedChart ? (
                  <div className="chart-preview-horizontal-list">
                    {renderSpec.categories.map((category, categoryIndex) => (
                      <div className="chart-preview-horizontal-row" key={`${renderSpec.chart_id}-${category}`}>
                        <div className="chart-preview-horizontal-label" title={category}>
                          {shortenCategoryLabel(category, categoryLabelMaxChars)}
                        </div>
                        <div className="chart-preview-horizontal-track">
                          <div className="chart-preview-horizontal-stack">
                            {visible.map((series, seriesIndex) => {
                              const value = series.values[categoryIndex] ?? 0;
                              const widthPercent = Math.max(4, (value / stackedDomain.max) * 100);
                              return (
                                <div
                                  className="chart-preview-horizontal-segment"
                                  key={`${series.name}-${category}`}
                                  style={{ width: `${widthPercent}%`, background: seriesColor(seriesIndex) }}
                                  title={`${series.name}: ${formatSeriesValue(value, series, renderSpec.value_format)}`}
                                />
                              );
                            })}
                          </div>
                        </div>
                        <div className="chart-preview-horizontal-value">
                          {formatAxisValue(categoryTotal(visible, categoryIndex), renderSpec, visible)}
                        </div>
                      </div>
                    ))}
                  </div>
                ) : isHorizontalBarChart ? (
                  <div className="chart-preview-horizontal-list">
                    {renderSpec.categories.map((category, categoryIndex) => (
                      <div className="chart-preview-horizontal-row" key={`${renderSpec.chart_id}-${category}`}>
                        <div className="chart-preview-horizontal-label" title={category}>
                          {shortenCategoryLabel(category, categoryLabelMaxChars)}
                        </div>
                        <div className="chart-preview-horizontal-track">
                          {visible.map((series, seriesIndex) => {
                            const value = series.values[categoryIndex] ?? 0;
                            const widthPercent = positiveHeightPercent(value, primaryDomain, 4);
                            return (
                              <div className="chart-preview-horizontal-series" key={`${series.name}-${category}`}>
                                <div
                                  className="chart-preview-horizontal-bar"
                                  style={{ width: `${widthPercent}%`, background: seriesColor(seriesIndex) }}
                                />
                              </div>
                            );
                          })}
                        </div>
                        <div className="chart-preview-horizontal-value">
                          {formatAxisValue(categoryTotal(visible, categoryIndex), renderSpec, visible)}
                        </div>
                      </div>
                    ))}
                  </div>
                ) : isStackedChart ? (
                  <div className="chart-preview-grid">
                    {renderSpec.categories.map((category, categoryIndex) => (
                      <div className="chart-preview-group" key={`${renderSpec.chart_id}-${category}`}>
                        <div className="chart-preview-bars">
                          <div className="chart-preview-bar-wrap stacked-wrap">
                            <div className="chart-preview-bar-value">
                              {formatAxisValue(categoryTotal(visible, categoryIndex), renderSpec, visible)}
                            </div>
                            <div className="chart-preview-stack">
                              {visible.map((series, seriesIndex) => {
                                const value = series.values[categoryIndex] ?? 0;
                                const heightPercent = Math.max(6, (value / stackedDomain.max) * 100);
                                return (
                                  <div
                                    className="chart-preview-stack-segment"
                                    key={`${series.name}-${category}`}
                                    style={{ height: `${heightPercent}%`, background: seriesColor(seriesIndex) }}
                                    title={`${series.name}: ${formatSeriesValue(value, series, renderSpec.value_format)}`}
                                  />
                                );
                              })}
                            </div>
                          </div>
                        </div>
                        <div className="chart-preview-category" title={category}>
                          {shortenCategoryLabel(category, categoryLabelMaxChars)}
                        </div>
                      </div>
                    ))}
                  </div>
                ) : isComboChart ? (
                  <div className="chart-preview-line-stage">
                    {comboLineSeries ? (
                      <svg className="chart-preview-line-svg" viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true">
                        <path
                          d={buildLinePath(comboLineSeries.values, secondaryDomain)}
                          className="chart-preview-line"
                          style={{ stroke: seriesColor(visible.length - 1) }}
                        />
                      </svg>
                    ) : null}

                    {comboLineSeries ? (
                      <div className="chart-preview-line-marker-layer">
                        {comboLineSeries.values.map((value, categoryIndex) => {
                          const x = pointXPercent(categoryIndex, comboLineSeries.values.length);
                          const bottom = pointBottomPercent(value, secondaryDomain);
                          return (
                            <div
                              className="chart-preview-line-marker"
                              key={`${comboLineSeries.name}-${renderSpec.categories[categoryIndex] ?? categoryIndex}-marker`}
                              style={{ left: `${x}%`, bottom: `${bottom}%` }}
                              title={`${comboLineSeries.name}: ${formatSeriesValue(value, comboLineSeries, renderSpec.value_format)}`}
                            >
                              <span
                                className="chart-preview-point"
                                style={{ background: seriesColor(visible.length - 1) }}
                              />
                              <span className="chart-preview-line-label">
                                {formatSeriesValue(value, comboLineSeries, renderSpec.value_format)}
                              </span>
                            </div>
                          );
                        })}
                      </div>
                    ) : null}

                    <div className="chart-preview-grid line-layout">
                      {renderSpec.categories.map((category, categoryIndex) => (
                        <div className="chart-preview-group" key={`${renderSpec.chart_id}-${category}`}>
                          <div className="chart-preview-bars combo-bars">
                            {comboBarSeries.map((series, seriesIndex) => {
                              const value = series.values[categoryIndex] ?? 0;
                              const heightPercent = positiveHeightPercent(value, comboBarDomain);
                              return (
                                <div className="chart-preview-bar-wrap" key={`${series.name}-${category}`}>
                                  <div className="chart-preview-bar-value">
                                    {formatSeriesValue(value, series, renderSpec.value_format)}
                                  </div>
                                  <div
                                    className="chart-preview-bar"
                                    style={{ height: `${heightPercent}%`, background: seriesColor(seriesIndex) }}
                                  />
                                </div>
                              );
                            })}
                          </div>
                          <div className="chart-preview-category" title={category}>
                            {shortenCategoryLabel(category, categoryLabelMaxChars)}
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                ) : (
                  <div className="chart-preview-grid">
                    {renderSpec.categories.map((category, categoryIndex) => (
                      <div className="chart-preview-group" key={`${renderSpec.chart_id}-${category}`}>
                        <div className="chart-preview-bars">
                          {visible.map((series, seriesIndex) => {
                            const value = series.values[categoryIndex] ?? 0;
                            const heightPercent = positiveHeightPercent(
                              value,
                              series.axis === "secondary" && isSecondaryColumnChart ? secondaryDomain : primaryDomain,
                            );
                            return (
                              <div className="chart-preview-bar-wrap" key={`${series.name}-${category}`}>
                                <div className="chart-preview-bar-value">
                                  {formatSeriesValue(value, series, renderSpec.value_format)}
                                </div>
                                <div
                                  className="chart-preview-bar"
                                  style={{ height: `${heightPercent}%`, background: seriesColor(seriesIndex) }}
                                />
                              </div>
                            );
                          })}
                        </div>
                        <div className="chart-preview-category" title={category}>
                          {shortenCategoryLabel(category, categoryLabelMaxChars)}
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>

            {hasSecondaryAxis ? (
              <div className="chart-preview-axis" style={{ textAlign: "left" }}>
                {secondaryTickValues.map((value, index) => (
                  <div className="chart-preview-axis-tick" key={`${renderSpec.chart_id}-tick-secondary-${index}`}>
                    {formatAxisValue(value, renderSpec, secondarySeries)}
                  </div>
                ))}
              </div>
            ) : null}
          </>
        )}
      </div>
    </div>
  );
}
