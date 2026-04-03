import { ChangeEvent, useEffect, useState, useTransition } from "react";

import { buildDownloadUrl, buildPlan, extractTextFromDocument, fetchTemplates, generatePresentation } from "@/api";
import { ChartPreview } from "@/components/chart-preview";
import { StructureDrawer } from "@/components/structure-drawer";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Select } from "@/components/ui/select";
import type {
  ChartabilityAssessment,
  ChartOverride,
  ChartSpec,
  DocumentBlock,
  GeneratePresentationResponse,
  PresentationPlan,
  TableBlock,
  TemplateSummary,
} from "@/types";

const PRIMARY_TEMPLATE_ID = "corp_light_v1";

const initialText = `Вставьте текст или загрузите документ в формате docx для презентации`;

const chartTypeLabels: Record<string, string> = {
  bar: "Горизонтальные столбцы",
  column: "Вертикальные столбцы",
  line: "Линейный график",
  stacked_bar: "Горизонтальные столбцы с накоплением",
  stacked_column: "Вертикальные столбцы с накоплением",
  combo: "Комбинированный график",
  pie: "Круговая диаграмма",
};

const assessmentHintLabels: Record<string, string> = {
  "detected category axis with numeric series": "Найдены подписи категорий и числовые значения для построения графика.",
  "detected time-like axis with numeric series": "Найдена временная шкала и числовые значения для построения графика.",
  "detected one numeric series with categorical labels": "Таблицу можно показать как один ряд по категориям.",
  "detected multiple numeric series with categorical labels": "Таблицу можно показать как несколько рядов по общим категориям.",
  "table is mostly text with weak numeric signal": "В таблице слишком мало сопоставимых чисел для надёжного графика.",
  "table does not contain enough structured numeric data": "Числовых данных недостаточно для корректной визуализации.",
  "mixed units detected": "В таблице смешаны разные единицы измерения.",
  "too many categories": "Категорий слишком много, график будет перегружен.",
  "contains summary rows": "В таблице есть итоговые строки, которые могут искажать график.",
  "contains annotations in value cells": "В ячейках есть комментарии рядом с числами, поэтому данные требуют осторожной интерпретации.",
};

function isSystemHint(hint: string): boolean {
  return hint.startsWith("classified as ") || hint.startsWith("detected ") || hint.startsWith("generated ");
}

export function App() {
  const [templates, setTemplates] = useState<TemplateSummary[]>([]);
  const [selectedTemplateId, setSelectedTemplateId] = useState(PRIMARY_TEMPLATE_ID);
  const [rawText, setRawText] = useState("");
  const [documentTables, setDocumentTables] = useState<TableBlock[]>([]);
  const [documentBlocks, setDocumentBlocks] = useState<DocumentBlock[]>([]);
  const [chartAssessments, setChartAssessments] = useState<ChartabilityAssessment[]>([]);
  const [chartSelectionByTableId, setChartSelectionByTableId] = useState<Record<string, string>>({});
  const [chartModeByTableId, setChartModeByTableId] = useState<Record<string, "table" | "chart">>({});
  const [hiddenSeriesByTableId, setHiddenSeriesByTableId] = useState<Record<string, string[]>>({});
  const [chartOrientationByTableId, setChartOrientationByTableId] = useState<Record<string, "default" | "transposed">>({});
  const [savedChartSelectionByTableId, setSavedChartSelectionByTableId] = useState<Record<string, string>>({});
  const [savedChartModeByTableId, setSavedChartModeByTableId] = useState<Record<string, "table" | "chart">>({});
  const [savedHiddenSeriesByTableId, setSavedHiddenSeriesByTableId] = useState<Record<string, string[]>>({});
  const [savedChartOrientationByTableId, setSavedChartOrientationByTableId] = useState<
    Record<string, "default" | "transposed">
  >({});
  const [generationResult, setGenerationResult] = useState<GeneratePresentationResponse | null>(null);
  const [error, setError] = useState("");
  const [showLoadingNotice, setShowLoadingNotice] = useState(true);
  const [isStructureDrawerOpen, setIsStructureDrawerOpen] = useState(false);
  const [showAllTablesInDrawer, setShowAllTablesInDrawer] = useState(false);
  const [isPending, startTransition] = useTransition();

  useEffect(() => {
    startTransition(() => {
      fetchTemplates()
        .then((items) => {
          setTemplates(items);
          setShowLoadingNotice(false);
          if (items.some((item) => item.template_id === PRIMARY_TEMPLATE_ID)) {
            setSelectedTemplateId(PRIMARY_TEMPLATE_ID);
            return;
          }
          if (items[0]?.template_id) {
            setSelectedTemplateId(items[0].template_id);
          }
        })
        .catch((err: Error) => {
          setShowLoadingNotice(false);
          setError(err.message);
        });
    });
  }, []);

  function handleTextFileUpload(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    setError("");
    setGenerationResult(null);
    startTransition(() => {
      extractTextFromDocument(file)
        .then((result) => {
          const nextChartSelectionByTableId: Record<string, string> = Object.fromEntries(
            result.chart_assessments
              .filter((assessment) => assessment.candidate_specs.length > 0)
              .map((assessment) => [assessment.table_id, assessment.candidate_specs[0].chart_id]),
          );
          const nextChartModeByTableId: Record<string, "table" | "chart"> = Object.fromEntries(
            result.chart_assessments
              .filter((assessment) => assessment.chartable)
              .map((assessment) => [assessment.table_id, "table" satisfies "table" | "chart"]),
          );

          setRawText(result.text);
          setDocumentTables(result.tables);
          setDocumentBlocks(result.blocks);
          setChartAssessments(result.chart_assessments);
          setIsStructureDrawerOpen(false);
          setShowAllTablesInDrawer(false);
          setChartSelectionByTableId(nextChartSelectionByTableId);
          setChartModeByTableId(nextChartModeByTableId);
          setHiddenSeriesByTableId({});
          setChartOrientationByTableId({});
          setSavedChartSelectionByTableId(nextChartSelectionByTableId);
          setSavedChartModeByTableId(nextChartModeByTableId);
          setSavedHiddenSeriesByTableId({});
          setSavedChartOrientationByTableId({});
        })
        .catch((err: Error) => {
          setError(err.message);
        });
    });

    event.target.value = "";
  }

  function handleGenerate() {
    const normalizedText = rawText.trim();
    if (!normalizedText) {
      setError("Введите текст или загрузите документ.");
      return;
    }

    setError("");
    setGenerationResult(null);
    startTransition(() => {
      const effectiveChartSelectionByTableId = savedChartSelectionByTableId;
      const effectiveChartModeByTableId = savedChartModeByTableId;
      const effectiveHiddenSeriesByTableId = savedHiddenSeriesByTableId;
      const effectiveChartOrientationByTableId = savedChartOrientationByTableId;

      const chartOverrides: ChartOverride[] = chartAssessments.map((assessment) => {
        const mode = effectiveChartModeByTableId[assessment.table_id] ?? "table";
        const selectedChart =
          mode === "chart"
            ? selectedChartSpec(
                assessment,
                effectiveChartSelectionByTableId,
                effectiveHiddenSeriesByTableId,
                effectiveChartOrientationByTableId,
              )
            : null;
        return {
          table_id: assessment.table_id,
          mode,
          selected_chart: selectedChart,
        };
      });

      buildPlan({
        template_id: selectedTemplateId || PRIMARY_TEMPLATE_ID,
        raw_text: normalizedText,
        title: "A3 Presentation",
        tables: documentTables,
        blocks: documentBlocks,
        chart_overrides: chartOverrides,
      })
        .then((plan: PresentationPlan) => generatePresentation(plan))
        .then((result) => setGenerationResult(result))
        .catch((err: Error) => setError(err.message));
    });
  }

  function selectedChartSpec(
    assessment: ChartabilityAssessment,
    selectionByTableId = chartSelectionByTableId,
    hiddenByTableId = hiddenSeriesByTableId,
    orientationByTableId = chartOrientationByTableId,
  ): ChartSpec | null {
    const chartId = selectionByTableId[assessment.table_id];
    const baseSpec = assessment.candidate_specs.find((item) => item.chart_id === chartId) ?? assessment.candidate_specs[0] ?? null;
    if (!baseSpec) {
      return null;
    }

    const orientation = orientationByTableId[assessment.table_id] ?? "default";
    const orientedSpec = orientation === "transposed" ? transposeChartSpec(baseSpec) : baseSpec;
    const hiddenSeries = new Set(hiddenByTableId[assessment.table_id] ?? []);
    return {
      ...orientedSpec,
      chart_id: `${orientedSpec.chart_id}:${orientation}`,
      series: orientedSpec.series.map((series) => ({
        ...series,
        hidden: hiddenSeries.has(series.name),
      })),
    };
  }

  function canTransposeChart(spec: ChartSpec | null): boolean {
    if (!spec) {
      return false;
    }
    if (spec.categories.length < 2 || spec.series.length < 2) {
      return false;
    }
    return spec.series.every((series) => series.values.length === spec.categories.length);
  }

  function transposeChartSpec(spec: ChartSpec): ChartSpec {
    if (!canTransposeChart(spec)) {
      return spec;
    }

    return {
      ...spec,
      title: spec.title ? `${spec.title} · разворот` : spec.title,
      categories: spec.series.map((series) => series.name),
      series: spec.categories.map((category, categoryIndex) => ({
        name: category,
        values: spec.series.map((series) => series.values[categoryIndex] ?? 0),
        unit: spec.series[0]?.unit ?? null,
        axis: "primary",
        hidden: false,
      })),
    };
  }

  function toggleSeriesVisibility(tableId: string, seriesName: string) {
    setHiddenSeriesByTableId((current) => {
      const currentHidden = new Set(current[tableId] ?? []);
      if (currentHidden.has(seriesName)) {
        currentHidden.delete(seriesName);
      } else {
        currentHidden.add(seriesName);
      }
      return {
        ...current,
        [tableId]: Array.from(currentHidden),
      };
    });
  }

  const hasUnsavedStructureChanges =
    JSON.stringify(chartSelectionByTableId) !== JSON.stringify(savedChartSelectionByTableId) ||
    JSON.stringify(chartModeByTableId) !== JSON.stringify(savedChartModeByTableId) ||
    JSON.stringify(hiddenSeriesByTableId) !== JSON.stringify(savedHiddenSeriesByTableId) ||
    JSON.stringify(chartOrientationByTableId) !== JSON.stringify(savedChartOrientationByTableId);
  const chartableAssessments = chartAssessments.filter((assessment) => assessment.chartable);
  const visibleAssessments = showAllTablesInDrawer ? chartAssessments : chartableAssessments;

  function handleSaveStructureChoices() {
    setSavedChartSelectionByTableId(chartSelectionByTableId);
    setSavedChartModeByTableId(chartModeByTableId);
    setSavedHiddenSeriesByTableId(hiddenSeriesByTableId);
    setSavedChartOrientationByTableId(chartOrientationByTableId);
    setIsStructureDrawerOpen(false);
  }

  function renderTablePreview(assessment: ChartabilityAssessment) {
    const structuredTable = assessment.structured_table;
    if (!structuredTable?.cells.length) {
      return <div className="table-preview-empty">Нет структурированного preview таблицы.</div>;
    }

    return (
      <div className="table-preview">
        {structuredTable.cells.slice(0, 6).map((row, rowIndex) => (
          <div className="table-preview-row" key={`${assessment.table_id}-row-${rowIndex}`}>
            {row.map((cell, cellIndex) => (
              <div
                className={`table-preview-cell${rowIndex === 0 ? " is-header" : ""}`}
                key={`${assessment.table_id}-cell-${rowIndex}-${cellIndex}`}
              >
                {cell.text || "—"}
              </div>
            ))}
          </div>
        ))}
      </div>
    );
  }

  function chartTypeLabel(chartType: string): string {
    return chartTypeLabels[chartType] ?? chartType;
  }

  function assessmentHints(assessment: ChartabilityAssessment): string[] {
    const rawHints = [...assessment.reasons, ...assessment.warnings].filter((hint) => !isSystemHint(hint));
    if (rawHints.length === 0) {
      return assessment.chartable
        ? ["Эту таблицу можно использовать для графика."]
        : ["Для этой таблицы график лучше не использовать."];
    }

    return rawHints.map((hint) => assessmentHintLabels[hint] ?? hint);
  }

  return (
    <main className="app-shell">
      <section className="hero-block" data-node-id="634:1739">
        <h1 className="hero-title" data-node-id="633:1698">
          A3 Presentation
        </h1>
        <p className="hero-description" data-node-id="633:3191">
          Превращайте документ в готовую презентацию в корпоративном стиле
          <br />
          Сервис извлекает структуру документа, собирает план слайдов и генерирует без ручной верстки.
        </p>
      </section>

      <div className="composer-stack">
        {generationResult ? (
          <div className="status-panel status-success">
            <button type="button" className="status-close" onClick={() => setGenerationResult(null)} aria-label="Закрыть">
              ×
            </button>
            <div className="status-title">Презентация готова</div>
            <div className="status-text">{generationResult.file_name}</div>
            <button
              type="button"
              className="primary-button status-download"
              onClick={() => window.open(buildDownloadUrl(generationResult.download_url), "_blank", "noopener,noreferrer")}
            >
              Скачать презентацию
            </button>
          </div>
        ) : null}

        {error ? (
          <div className="status-panel status-error">
            <button type="button" className="status-close" onClick={() => setError("")} aria-label="Закрыть">
              ×
            </button>
            <div className="status-title">Ошибка</div>
            <div className="status-text">{error}</div>
          </div>
        ) : null}

        {templates.length === 0 && !error && showLoadingNotice ? (
          <div className="status-panel status-muted">
            <button type="button" className="status-close" onClick={() => setShowLoadingNotice(false)} aria-label="Скрыть">
              ×
            </button>
            Загрузка конфигурации шаблона...
          </div>
        ) : null}

        <section className="composer-card" data-node-id="633:1701">
          <div className="composer-inner" data-node-id="634:1782">
            <div className="textarea-wrap" data-node-id="633:3164">
              <textarea
                value={rawText}
                onChange={(event) => setRawText(event.target.value)}
                placeholder={initialText}
                className="main-textarea"
                rows={10}
              />
            </div>

            <div className="actions-row" data-node-id="634:1772">
              <label className="secondary-button" data-node-id="633:2618">
                <input
                  type="file"
                  accept=".txt,.md,.markdown,.pdf,.docx,text/plain,text/markdown,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                  className="sr-only"
                  onChange={handleTextFileUpload}
                />
                <span>Загрузить документ</span>
              </label>

              <div className="actions-group">
                {chartAssessments.length > 0 ? (
                  <button
                    type="button"
                    className="secondary-button"
                    onClick={() => setIsStructureDrawerOpen(true)}
                  >
                    Посмотреть структуру ({chartAssessments.length})
                  </button>
                ) : null}

                <button
                  type="button"
                  className="primary-button"
                  data-node-id="634:1743"
                  onClick={handleGenerate}
                  disabled={isPending}
                >
                  {isPending ? "Генерация..." : "Сгенерировать"}
                </button>
              </div>
            </div>
          </div>
        </section>

        {chartAssessments.length > 0 ? (
          <StructureDrawer
            open={isStructureDrawerOpen}
            onOpenChange={setIsStructureDrawerOpen}
            title="Таблицы документа"
            description={`Найдено таблиц: ${chartAssessments.length}. Для визуализации подходят: ${chartableAssessments.length}.`}
            footer={
              <div className="drawer-footer-actions">
                <div className="drawer-footer-note">
                  {hasUnsavedStructureChanges ? "Есть несохранённые изменения. Сохрани выбор, чтобы применить его при генерации." : "Выбор сохранён и будет использован при генерации."}
                </div>
                <button
                  type="button"
                  className="primary-button"
                  onClick={handleSaveStructureChoices}
                  disabled={!hasUnsavedStructureChanges}
                >
                  Сохранить выбор
                </button>
              </div>
            }
          >
            <div className="drawer-toolbar">
              <div className="drawer-toolbar-copy">
                {chartableAssessments.length > 0
                  ? "Сначала показаны только таблицы, которые можно превратить в графики."
                  : "В документе не найдено таблиц, которые система может уверенно превратить в графики."}
              </div>
              {chartAssessments.length > chartableAssessments.length ? (
                <label className="drawer-toggle">
                  <Input
                    type="checkbox"
                    className="drawer-toggle-input"
                    checked={showAllTablesInDrawer}
                    onChange={(event) => setShowAllTablesInDrawer(event.target.checked)}
                  />
                  <span>Показать все таблицы</span>
                </label>
              ) : null}
            </div>
            <section className="preview-gallery">
              {visibleAssessments.map((assessment) => {
                const selectedSpec = selectedChartSpec(assessment);
                const mode = chartModeByTableId[assessment.table_id] ?? "table";
                const selectedChartId =
                  chartSelectionByTableId[assessment.table_id] ?? assessment.candidate_specs[0]?.chart_id ?? "";

                return (
                  <Card className="preview-card" key={assessment.table_id}>
                    <CardHeader>
                      <CardTitle>
                        {assessment.chartable ? "Можно заменить на график" : "Оставить как таблицу"}
                      </CardTitle>
                      <CardDescription>
                        {assessment.chartable
                          ? "Выбери формат отображения и сохрани решение для итоговой презентации."
                          : "Для этой таблицы лучше оставить табличный вид, чтобы не исказить данные."}
                      </CardDescription>
                    </CardHeader>
                    <CardContent className="preview-card-content">
                      <div className="preview-toolbar">
                        {assessment.chartable ? (
                          <>
                            <button
                              type="button"
                              className={`secondary-button${mode === "table" ? " is-active" : ""}`}
                              onClick={() => setChartModeByTableId((current) => ({ ...current, [assessment.table_id]: "table" }))}
                            >
                              Таблица
                            </button>
                            <button
                              type="button"
                              className={`secondary-button${mode === "chart" ? " is-active" : ""}`}
                              onClick={() => setChartModeByTableId((current) => ({ ...current, [assessment.table_id]: "chart" }))}
                              disabled={!selectedSpec}
                            >
                              График
                            </button>
                          </>
                        ) : (
                          <div className="preview-toolbar-note">Для этой таблицы вариант с графиком не предлагается.</div>
                        )}
                        {assessment.chartable && mode === "chart" && assessment.candidate_specs.length > 1 ? (
                          <Select
                            className="chart-type-select"
                            value={selectedChartId}
                            onChange={(event) =>
                              setChartSelectionByTableId((current) => ({ ...current, [assessment.table_id]: event.target.value }))
                            }
                          >
                            {assessment.candidate_specs.map((spec) => (
                              <option key={spec.chart_id} value={spec.chart_id}>
                                {chartTypeLabel(spec.chart_type)}
                              </option>
                            ))}
                          </Select>
                        ) : null}
                        {assessment.chartable && mode === "chart" && canTransposeChart(selectedSpec) ? (
                          <Select
                            className="chart-type-select"
                            value={chartOrientationByTableId[assessment.table_id] ?? "default"}
                            onChange={(event) =>
                              setChartOrientationByTableId((current) => ({
                                ...current,
                                [assessment.table_id]: event.target.value as "default" | "transposed",
                              }))
                            }
                          >
                            <option value="default">Ряды по колонкам</option>
                            <option value="transposed">Ряды по строкам</option>
                          </Select>
                        ) : null}
                      </div>

                      {mode === "chart" && selectedSpec ? <ChartPreview spec={selectedSpec} /> : renderTablePreview(assessment)}

                      {mode === "chart" && selectedSpec && selectedSpec.series.length > 1 ? (
                        <div className="series-controls">
                          <div className="series-controls-title">Ряды графика</div>
                          <div className="series-controls-list">
                            {selectedSpec.series.map((series) => (
                              <label className="series-toggle" key={`${assessment.table_id}-${series.name}`}>
                                <Input
                                  type="checkbox"
                                  className="series-toggle-input"
                                  checked={!series.hidden}
                                  onChange={() => toggleSeriesVisibility(assessment.table_id, series.name)}
                                />
                                <span>{series.name}</span>
                              </label>
                            ))}
                          </div>
                        </div>
                      ) : null}

                      {assessmentHints(assessment).length > 0 ? (
                        <div className="preview-reasons">
                          {assessmentHints(assessment).map((reason) => (
                            <div className="preview-reason" key={reason}>
                              {reason}
                            </div>
                          ))}
                        </div>
                      ) : null}
                    </CardContent>
                  </Card>
                );
              })}
            </section>
          </StructureDrawer>
        ) : null}
      </div>
    </main>
  );
}
