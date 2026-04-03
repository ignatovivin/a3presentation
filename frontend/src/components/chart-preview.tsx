import type { CSSProperties } from "react";
import type { ChartSpec } from "@/types";
import chartStyle from "@/chart-style.json";

type ChartPreviewProps = {
  spec: ChartSpec;
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

export function ChartPreview({ spec }: ChartPreviewProps) {
  const visibleSeries = spec.series.filter((series) => !series.hidden);
  const maxValue = Math.max(...visibleSeries.flatMap((series) => series.values), 1);
  const minValue = Math.min(...visibleSeries.flatMap((series) => series.values), 0);
  const pieSeries = visibleSeries[0] ?? null;
  const pieTotal = pieSeries ? Math.max(pieSeries.values.reduce((sum, value) => sum + value, 0), 1) : 1;
  const stackedMaxValue = Math.max(
    ...spec.categories.map((_, categoryIndex) =>
      visibleSeries.reduce((sum, series) => sum + (series.values[categoryIndex] ?? 0), 0),
    ),
    1,
  );
  const categoryCount = spec.categories.length;
  const seriesCount = visibleSeries.length;
  const gridSteps = 4;
  const isPieChart = spec.chart_type === "pie";
  const isHorizontalBarChart = spec.chart_type === "bar";
  const isHorizontalStackedChart = spec.chart_type === "stacked_bar";
  const isStackedChart = spec.chart_type === "stacked_bar" || spec.chart_type === "stacked_column";
  const tickValues = Array.from({ length: gridSteps + 1 }, (_, index) => {
    const ratio = (gridSteps - index) / gridSteps;
    return Math.max(0, (isStackedChart ? stackedMaxValue : maxValue) * ratio);
  });
  const isLineChart = spec.chart_type === "line";
  const isComboChart = spec.chart_type === "combo";
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
  const comboBarMaxValue = Math.max(...comboBarSeries.flatMap((series) => series.values), 1);

  if (!visibleSeries.length) {
    return <div className="chart-empty-state">Нет доступных рядов для визуализации.</div>;
  }

  function buildLinePath(values: number[]): string {
    return values
      .map((value, index) => {
        const x = values.length === 1 ? 50 : (index / Math.max(values.length - 1, 1)) * 100;
        const y = 100 - Math.max(0, Math.min(100, (value / maxValue) * 100));
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
          <div className="chart-preview-title">{spec.title || "Превью графика"}</div>
          <div className="chart-preview-meta">
            {categoryCount} категорий · {seriesCount} {seriesCount === 1 ? "ряд" : seriesCount < 5 ? "ряда" : "рядов"}
          </div>
        </div>
        {spec.legend_visible ? (
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

      <div className="chart-preview-summary">
        <div className="chart-preview-stat">
          <span className="chart-preview-stat-label">Максимум</span>
          <span className="chart-preview-stat-value">{formatCompactValue(maxValue, spec.value_format)}</span>
        </div>
        <div className="chart-preview-stat">
          <span className="chart-preview-stat-label">Минимум</span>
          <span className="chart-preview-stat-value">{formatCompactValue(minValue, spec.value_format)}</span>
        </div>
        <div className="chart-preview-stat">
          <span className="chart-preview-stat-label">Тип</span>
          <span className="chart-preview-stat-value">
            {isPieChart ? "Структура" : isLineChart ? "Тренд" : isStackedChart ? "Накопление" : "Сравнение"}
          </span>
        </div>
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
                        <path key={`${spec.chart_id}-pie-${index}`} d={path} fill={seriesColor(index)} />
                      ))
                    : null}
                  <circle cx="50" cy="50" r="16" fill="#ffffff" fillOpacity="0.94" />
                </svg>
                <div className="chart-preview-pie-total">
                  <span className="chart-preview-pie-total-label">Всего</span>
                  <span className="chart-preview-pie-total-value">
                    {formatCompactValue(pieTotal, spec.value_format)}
                  </span>
                </div>
              </div>

              <div className="chart-preview-pie-breakdown">
                {spec.categories.map((category, categoryIndex) => {
                  const value = pieSeries?.values[categoryIndex] ?? 0;
                  const percent = pieTotal > 0 ? (value / pieTotal) * 100 : 0;
                  return (
                    <div className="chart-preview-pie-item" key={`${spec.chart_id}-pie-item-${categoryIndex}`}>
                      <div className="chart-preview-pie-item-head">
                        <span className="chart-preview-swatch" style={{ background: seriesColor(categoryIndex) }} />
                        <span className="chart-preview-pie-item-label" title={category}>
                          {shortenCategoryLabel(category, Math.max(categoryLabelMaxChars, 18))}
                        </span>
                      </div>
                      <div className="chart-preview-pie-item-metrics">
                        <span>{formatCompactValue(value, spec.value_format)}</span>
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
            <div className="chart-preview-axis-tick" key={`${spec.chart_id}-tick-${index}`}>
              {formatCompactValue(value, spec.value_format)}
            </div>
          ))}
        </div>

        <div className="chart-preview-stage-scroll">
          <div className="chart-preview-stage" style={{ width: `${contentWidth}px` }}>
            <div className="chart-preview-gridlines">
              {tickValues.map((_, index) => (
                <div className="chart-preview-gridline" key={`${spec.chart_id}-grid-${index}`} />
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

                <div className="chart-preview-grid line-layout">
                  {spec.categories.map((category, categoryIndex) => (
                    <div className="chart-preview-group" key={`${spec.chart_id}-${category}`}>
                      <div className="chart-preview-bars chart-preview-points">
                        {visibleSeries.map((series, seriesIndex) => {
                          const value = series.values[categoryIndex] ?? 0;
                          const pointBottom = Math.max(0, Math.min(100, (value / maxValue) * 100));
                          return (
                            <div className="chart-preview-bar-wrap point-wrap" key={`${series.name}-${category}`}>
                              <div className="chart-preview-bar-value">{formatCompactValue(value, spec.value_format)}</div>
                              <div
                                className="chart-preview-point"
                                style={{ bottom: `${pointBottom}%`, background: seriesColor(seriesIndex) }}
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
            ) : isHorizontalStackedChart ? (
              <div className="chart-preview-horizontal-list">
                {spec.categories.map((category, categoryIndex) => (
                  <div className="chart-preview-horizontal-row" key={`${spec.chart_id}-${category}`}>
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
                              title={`${series.name}: ${formatCompactValue(value, spec.value_format)}`}
                            />
                          );
                        })}
                      </div>
                    </div>
                    <div className="chart-preview-horizontal-value">
                      {formatCompactValue(
                        visibleSeries.reduce((sum, series) => sum + (series.values[categoryIndex] ?? 0), 0),
                        spec.value_format,
                      )}
                    </div>
                  </div>
                ))}
              </div>
            ) : isHorizontalBarChart ? (
              <div className="chart-preview-horizontal-list">
                {spec.categories.map((category, categoryIndex) => (
                  <div className="chart-preview-horizontal-row" key={`${spec.chart_id}-${category}`}>
                    <div className="chart-preview-horizontal-label" title={category}>
                      {shortenCategoryLabel(category, categoryLabelMaxChars)}
                    </div>
                    <div className="chart-preview-horizontal-track">
                      {visibleSeries.map((series, seriesIndex) => {
                        const value = series.values[categoryIndex] ?? 0;
                        const widthPercent = Math.max(4, (value / maxValue) * 100);
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
                      {formatCompactValue(visibleSeries[0]?.values[categoryIndex] ?? 0, spec.value_format)}
                    </div>
                  </div>
                ))}
              </div>
            ) : isStackedChart ? (
              <div className="chart-preview-grid">
                {spec.categories.map((category, categoryIndex) => (
                  <div className="chart-preview-group" key={`${spec.chart_id}-${category}`}>
                    <div className="chart-preview-bars">
                      <div className="chart-preview-bar-wrap stacked-wrap">
                        <div className="chart-preview-bar-value">
                          {formatCompactValue(
                            visibleSeries.reduce((sum, series) => sum + (series.values[categoryIndex] ?? 0), 0),
                            spec.value_format,
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
                                title={`${series.name}: ${formatCompactValue(value, spec.value_format)}`}
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

                <div className="chart-preview-grid line-layout">
                  {spec.categories.map((category, categoryIndex) => (
                    <div className="chart-preview-group" key={`${spec.chart_id}-${category}`}>
                      <div className="chart-preview-bars combo-bars">
                        {comboBarSeries.map((series, seriesIndex) => {
                          const value = series.values[categoryIndex] ?? 0;
                          const heightPercent = Math.max(6, (value / comboBarMaxValue) * 100);
                          return (
                            <div className="chart-preview-bar-wrap" key={`${series.name}-${category}`}>
                              <div className="chart-preview-bar-value">{formatCompactValue(value, spec.value_format)}</div>
                              <div
                                className="chart-preview-bar"
                                style={{ height: `${heightPercent}%`, background: seriesColor(seriesIndex) }}
                              />
                            </div>
                          );
                        })}
                        {comboLineSeries ? (
                          <div className="chart-preview-bar-wrap point-wrap combo-point-wrap">
                            <div className="chart-preview-bar-value">
                              {formatCompactValue(comboLineSeries.values[categoryIndex] ?? 0, spec.value_format)}
                            </div>
                            <div
                              className="chart-preview-point"
                              style={{
                                background: seriesColor(visibleSeries.length - 1),
                                bottom: `${Math.max(
                                  0,
                                  Math.min(100, ((comboLineSeries.values[categoryIndex] ?? 0) / maxValue) * 100),
                                )}%`,
                              }}
                            />
                          </div>
                        ) : null}
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
                {spec.categories.map((category, categoryIndex) => (
                  <div className="chart-preview-group" key={`${spec.chart_id}-${category}`}>
                    <div className="chart-preview-bars">
                      {visibleSeries.map((series, seriesIndex) => {
                        const value = series.values[categoryIndex] ?? 0;
                        const heightPercent = Math.max(6, (value / maxValue) * 100);
                        return (
                          <div className="chart-preview-bar-wrap" key={`${series.name}-${category}`}>
                            <div className="chart-preview-bar-value">{formatCompactValue(value, spec.value_format)}</div>
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
