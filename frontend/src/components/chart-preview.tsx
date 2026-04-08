import type { CSSProperties } from "react";
import type { ChartSeries, ChartSpec, ChartType } from "@/types";
import chartStyle from "@/chart-style.json";

type ChartPreviewProps = {
  spec: ChartSpec;
};

type RenderChartSpec = ChartSpec & {
  series: ChartSeries[];
  chart_type: ChartType;
};

type ScaleDomain = {
  min: number;
  max: number;
  span: number;
};

function formatCompactValue(value: number, valueFormat: string): string {
  const compact = new Intl.NumberFormat("ru-RU", {
    notation: "compact",
    maximumFractionDigits: 1,
  }).format(value);

  if (valueFormat === "currency") {
    return `${compact} ₽`;
  }
  if (valueFormat === "percent") {
    return `${value}%`;
  }
  return compact;
}

function shortenCategoryLabel(label: string, maxChars: number): string {
  if (label.length <= maxChars) {
    return label;
  }
  return `${label.slice(0, Math.max(0, maxChars - 1)).trimEnd()}…`;
}

function renderChartSpec(spec: ChartSpec): RenderChartSpec {
  const visibleSeries = spec.series.filter((series) => !series.hidden);
  const shouldRenderCombo =
    spec.chart_type === "combo" &&
    Boolean(spec.series.length && !spec.series[spec.series.length - 1].hidden && visibleSeries.length >= 2);

  return {
    ...spec,
    chart_type: spec.chart_type === "combo" && !shouldRenderCombo ? "column" : spec.chart_type,
    series: visibleSeries,
  };
}

function scaleDomain(series: ChartSeries[]): ScaleDomain {
  const values = series.flatMap((item) => item.values);
  const minValue = Math.min(...values, 0);
  const maxValue = Math.max(...values, 1);
  const span = Math.max(maxValue - minValue, 1);
  return { min: minValue, max: maxValue, span };
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

function categoryTotal(series: ChartSeries[], categoryIndex: number): number {
  return series.reduce((sum, item) => sum + (item.values[categoryIndex] ?? 0), 0);
}

function compactSeriesMagnitude(series: ChartSeries): number {
  return Math.max(...series.values.map((value) => Math.abs(value)), 0);
}

export function ChartPreview({ spec }: ChartPreviewProps) {
  const renderSpec = renderChartSpec(spec);
  const visibleSeries = renderSpec.series;
  const valueDomain = scaleDomain(visibleSeries);
  const pieSeries = renderSpec.chart_type === "pie" ? visibleSeries[0] ?? null : null;
  const pieTotal = pieSeries ? Math.max(pieSeries.values.reduce((sum, value) => sum + value, 0), 1) : 1;
  const stackedMaxValue = Math.max(
    ...renderSpec.categories.map((_, categoryIndex) => categoryTotal(visibleSeries, categoryIndex)),
    1,
  );
  const categoryCount = renderSpec.categories.length;
  const seriesCount = visibleSeries.length;
  const gridSteps = 4;
  const isPieChart = renderSpec.chart_type === "pie";
  const isHorizontalBarChart = renderSpec.chart_type === "bar";
  const isHorizontalStackedChart = renderSpec.chart_type === "stacked_bar";
  const isStackedChart = renderSpec.chart_type === "stacked_bar" || renderSpec.chart_type === "stacked_column";
  const tickValues = Array.from({ length: gridSteps + 1 }, (_, index) => {
    const ratio = (gridSteps - index) / gridSteps;
    if (isStackedChart) {
      return Math.max(0, stackedMaxValue * ratio);
    }
    return valueDomain.min + valueDomain.span * ratio;
  });
  const isLineChart = renderSpec.chart_type === "line";
  const isComboChart = renderSpec.chart_type === "combo";
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
  const comboLineSeries = isComboChart ? visibleSeries[visibleSeries.length - 1] ?? null : null;
  const comboBarSeries = isComboChart ? visibleSeries.slice(0, -1) : visibleSeries;
  const comboBarDomain = scaleDomain(comboBarSeries);
  const labeledLineSeriesIndex =
    isLineChart && visibleSeries.length > 1
      ? visibleSeries.reduce((lowestIndex, series, index) => {
          return compactSeriesMagnitude(series) < compactSeriesMagnitude(visibleSeries[lowestIndex]) ? index : lowestIndex;
        }, 0)
      : 0;

  if (!visibleSeries.length) {
    return <div className="chart-empty-state">Нет доступных рядов для визуализации.</div>;
  }

  function buildLinePath(values: number[]): string {
    return values
      .map((value, index) => {
        const x = pointXPercent(index, values.length);
        const y = 100 - pointBottomPercent(value, valueDomain);
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
            {visibleSeries.map((series, index) => (
              <span className="chart-preview-legend-item" key={series.name}>
                <span className="chart-preview-swatch" style={{ background: seriesColor(index) }} />
                {series.name}
              </span>
            ))}
          </div>
        ) : null}
      </div>

      <div className={`chart-preview-frame${isPieChart ? " is-pie" : ""}`}>
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
                    {formatCompactValue(pieTotal, renderSpec.value_format)}
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
                        <span>{formatCompactValue(value, renderSpec.value_format)}</span>
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
          {tickValues.map((value, index) => (
            <div className="chart-preview-axis-tick" key={`${renderSpec.chart_id}-tick-${index}`}>
              {formatCompactValue(value, renderSpec.value_format)}
            </div>
          ))}
        </div>

        <div className="chart-preview-stage-scroll">
          <div className="chart-preview-stage" style={{ width: `${contentWidth}px` }}>
            <div className="chart-preview-gridlines">
              {tickValues.map((_, index) => (
                <div className="chart-preview-gridline" key={`${renderSpec.chart_id}-grid-${index}`} />
              ))}
            </div>

            {isLineChart ? (
              <div className="chart-preview-line-stage">
                <svg className="chart-preview-line-svg" viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true">
                  {visibleSeries.map((series, seriesIndex) => (
                    <path
                      key={series.name}
                      d={buildLinePath(series.values)}
                      className="chart-preview-line"
                      style={{ stroke: seriesColor(seriesIndex) }}
                    />
                  ))}
                </svg>

                <div className="chart-preview-line-marker-layer">
                  {visibleSeries.flatMap((series, seriesIndex) =>
                    series.values.map((value, categoryIndex) => {
                      const x = pointXPercent(categoryIndex, series.values.length);
                      const bottom = pointBottomPercent(value, valueDomain);
                      const shouldShowLabel = visibleSeries.length === 1 || seriesIndex === labeledLineSeriesIndex;
                      return (
                        <div
                          className="chart-preview-line-marker"
                          key={`${series.name}-${renderSpec.categories[categoryIndex] ?? categoryIndex}-marker`}
                          style={{ left: `${x}%`, bottom: `${bottom}%` }}
                          title={`${series.name}: ${formatCompactValue(value, renderSpec.value_format)}`}
                        >
                          <span className="chart-preview-point" style={{ background: seriesColor(seriesIndex) }} />
                          {shouldShowLabel ? (
                            <span className="chart-preview-line-label">
                              {formatCompactValue(value, renderSpec.value_format)}
                            </span>
                          ) : null}
                        </div>
                      );
                    }),
                  )}
                </div>

                <div className="chart-preview-grid line-layout">
                  {renderSpec.categories.map((category, categoryIndex) => (
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
                        {visibleSeries.map((series, seriesIndex) => {
                          const value = series.values[categoryIndex] ?? 0;
                          const widthPercent = Math.max(4, (value / stackedMaxValue) * 100);
                          return (
                            <div
                              className="chart-preview-horizontal-segment"
                              key={`${series.name}-${category}`}
                              style={{ width: `${widthPercent}%`, background: seriesColor(seriesIndex) }}
                              title={`${series.name}: ${formatCompactValue(value, renderSpec.value_format)}`}
                            />
                          );
                        })}
                      </div>
                    </div>
                    <div className="chart-preview-horizontal-value">
                      {formatCompactValue(
                        categoryTotal(visibleSeries, categoryIndex),
                        renderSpec.value_format,
                      )}
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
                      {visibleSeries.map((series, seriesIndex) => {
                        const value = series.values[categoryIndex] ?? 0;
                        const widthPercent = positiveHeightPercent(value, valueDomain, 4);
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
                      {formatCompactValue(categoryTotal(visibleSeries, categoryIndex), renderSpec.value_format)}
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
                          {formatCompactValue(
                            categoryTotal(visibleSeries, categoryIndex),
                            renderSpec.value_format,
                          )}
                        </div>
                        <div className="chart-preview-stack">
                          {visibleSeries.map((series, seriesIndex) => {
                            const value = series.values[categoryIndex] ?? 0;
                            const heightPercent = Math.max(6, (value / stackedMaxValue) * 100);
                            return (
                              <div
                                className="chart-preview-stack-segment"
                                key={`${series.name}-${category}`}
                                style={{ height: `${heightPercent}%`, background: seriesColor(seriesIndex) }}
                                title={`${series.name}: ${formatCompactValue(value, renderSpec.value_format)}`}
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
                      d={buildLinePath(comboLineSeries.values)}
                      className="chart-preview-line"
                      style={{ stroke: seriesColor(visibleSeries.length - 1) }}
                    />
                  </svg>
                ) : null}

                {comboLineSeries ? (
                  <div className="chart-preview-line-marker-layer">
                    {comboLineSeries.values.map((value, categoryIndex) => {
                      const x = pointXPercent(categoryIndex, comboLineSeries.values.length);
                      const bottom = pointBottomPercent(value, valueDomain);
                      return (
                        <div
                          className="chart-preview-line-marker"
                          key={`${comboLineSeries.name}-${renderSpec.categories[categoryIndex] ?? categoryIndex}-marker`}
                          style={{ left: `${x}%`, bottom: `${bottom}%` }}
                          title={`${comboLineSeries.name}: ${formatCompactValue(value, renderSpec.value_format)}`}
                        >
                          <span
                            className="chart-preview-point"
                            style={{ background: seriesColor(visibleSeries.length - 1) }}
                          />
                          <span className="chart-preview-line-label">
                            {formatCompactValue(value, renderSpec.value_format)}
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
                              <div className="chart-preview-bar-value">{formatCompactValue(value, renderSpec.value_format)}</div>
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
                      {visibleSeries.map((series, seriesIndex) => {
                        const value = series.values[categoryIndex] ?? 0;
                        const heightPercent = positiveHeightPercent(value, valueDomain);
                        return (
                          <div className="chart-preview-bar-wrap" key={`${series.name}-${category}`}>
                            <div className="chart-preview-bar-value">{formatCompactValue(value, renderSpec.value_format)}</div>
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
          </>
        )}
      </div>
    </div>
  );
}
