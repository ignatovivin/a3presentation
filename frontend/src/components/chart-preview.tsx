import type { ChartSpec } from "@/types";

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

export function ChartPreview({ spec }: ChartPreviewProps) {
  const visibleSeries = spec.series.filter((series) => !series.hidden);
  const maxValue = Math.max(...visibleSeries.flatMap((series) => series.values), 1);

  if (!visibleSeries.length) {
    return <div className="chart-empty-state">Нет доступных рядов для визуализации.</div>;
  }

  return (
    <div className="chart-preview">
      <div className="chart-preview-head">
        <div>
          <div className="chart-preview-title">{spec.title || "Превью графика"}</div>
          <div className="chart-preview-meta">
            {spec.chart_type} · confidence: {spec.confidence}
          </div>
        </div>
        {spec.legend_visible ? (
          <div className="chart-preview-legend">
            {visibleSeries.map((series, index) => (
              <span className="chart-preview-legend-item" key={series.name}>
                <span className={`chart-preview-swatch chart-series-${index % 4}`} />
                {series.name}
              </span>
            ))}
          </div>
        ) : null}
      </div>

      <div className="chart-preview-grid">
        {spec.categories.map((category, categoryIndex) => (
          <div className="chart-preview-group" key={`${spec.chart_id}-${category}`}>
            <div className="chart-preview-bars">
              {visibleSeries.map((series, seriesIndex) => {
                const value = series.values[categoryIndex] ?? 0;
                const heightPercent = Math.max(10, (value / maxValue) * 100);
                return (
                  <div className="chart-preview-bar-wrap" key={`${series.name}-${category}`}>
                    <div className="chart-preview-bar-value">{formatCompactValue(value, spec.value_format)}</div>
                    <div
                      className={`chart-preview-bar chart-series-${seriesIndex % 4}`}
                      style={{ height: `${heightPercent}%` }}
                    />
                  </div>
                );
              })}
            </div>
            <div className="chart-preview-category">{category}</div>
          </div>
        ))}
      </div>
    </div>
  );
}
