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
  SlideSpec,
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

function chartVariantPriority(spec: ChartSpec): number {
  const label = spec.variant_label?.toLowerCase() ?? "";
  if (label.startsWith("сравнение")) {
    return 0;
  }
  if (spec.chart_type === "combo" || label.startsWith("комбинированный")) {
    return 0;
  }
  if (label.startsWith("единичный")) {
    return 1;
  }
  return 3;
}

function isSelectableChartType(chartType: string): boolean {
  return chartType !== "combo";
}

function chartControlType(spec: ChartSpec): string {
  return spec.chart_type === "combo" ? "column" : spec.chart_type;
}

export function App() {
  const [templates, setTemplates] = useState<TemplateSummary[]>([]);
  const [selectedTemplateId, setSelectedTemplateId] = useState(PRIMARY_TEMPLATE_ID);
  const [rawText, setRawText] = useState("");
  const [attachedDocumentName, setAttachedDocumentName] = useState("");
  const [attachedDocumentText, setAttachedDocumentText] = useState("");
  const [documentTables, setDocumentTables] = useState<TableBlock[]>([]);
  const [documentBlocks, setDocumentBlocks] = useState<DocumentBlock[]>([]);
  const [chartAssessments, setChartAssessments] = useState<ChartabilityAssessment[]>([]);
  const [chartSelectionByTableId, setChartSelectionByTableId] = useState<Record<string, string>>({});
  const [chartModeByTableId, setChartModeByTableId] = useState<Record<string, "table" | "chart">>({});
  const [hiddenSeriesByTableId, setHiddenSeriesByTableId] = useState<Record<string, string[]>>({});
  const [savedChartSelectionByTableId, setSavedChartSelectionByTableId] = useState<Record<string, string>>({});
  const [savedChartModeByTableId, setSavedChartModeByTableId] = useState<Record<string, "table" | "chart">>({});
  const [savedHiddenSeriesByTableId, setSavedHiddenSeriesByTableId] = useState<Record<string, string[]>>({});
  const [reviewPlan, setReviewPlan] = useState<PresentationPlan | null>(null);
  const [cardSlideIndexes, setCardSlideIndexes] = useState<number[]>([]);
  const [isPreparingReviewPlan, setIsPreparingReviewPlan] = useState(false);
  const [isGeneratingPresentation, setIsGeneratingPresentation] = useState(false);
  const [generationResult, setGenerationResult] = useState<GeneratePresentationResponse | null>(null);
  const [error, setError] = useState("");
  const [showLoadingNotice, setShowLoadingNotice] = useState(true);
  const [isStructureDrawerOpen, setIsStructureDrawerOpen] = useState(false);
  const [showAllTablesInDrawer, setShowAllTablesInDrawer] = useState(false);
  const [drawerTab, setDrawerTab] = useState<"charts" | "text">("charts");
  const [isPending, startTransition] = useTransition();

  type CardSlideFit = "high" | "medium";
  type CardSlideChoice = {
    index: number;
    slide: SlideSpec;
    items: string[];
    fit: CardSlideFit;
    reason: string;
  };

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
              .map((assessment) => [assessment.table_id, preferredChartSpec(assessment)?.chart_id ?? assessment.candidate_specs[0].chart_id]),
          );
          const nextChartModeByTableId: Record<string, "table" | "chart"> = Object.fromEntries(
            result.chart_assessments
              .filter((assessment) => assessment.chartable)
              .map((assessment) => [assessment.table_id, "table" satisfies "table" | "chart"]),
          );

          setRawText("");
          setAttachedDocumentName(result.file_name || file.name);
          setAttachedDocumentText(result.text);
          setReviewPlan(null);
          setCardSlideIndexes([]);
          setDocumentTables(result.tables);
          setDocumentBlocks(result.blocks);
          setChartAssessments(result.chart_assessments);
          setIsStructureDrawerOpen(false);
          setShowAllTablesInDrawer(false);
          setChartSelectionByTableId(nextChartSelectionByTableId);
          setChartModeByTableId(nextChartModeByTableId);
          setHiddenSeriesByTableId({});
          setSavedChartSelectionByTableId(nextChartSelectionByTableId);
          setSavedChartModeByTableId(nextChartModeByTableId);
          setSavedHiddenSeriesByTableId({});
        })
        .catch((err: Error) => {
          setAttachedDocumentName("");
          setAttachedDocumentText("");
          setError(err.message);
        });
    });

    event.target.value = "";
  }

  function handleGenerate() {
    if (isGeneratingPresentation || isPreparingReviewPlan) {
      return;
    }
    if (reviewPlan) {
      setError("");
      setGenerationResult(null);
      setIsGeneratingPresentation(true);
      startTransition(() => {
        generatePresentation(applyCardSlideChoices(reviewPlan, cardSlideIndexes))
          .then((result) => setGenerationResult(result))
          .catch((err: Error) => setError(err.message))
          .finally(() => setIsGeneratingPresentation(false));
      });
      return;
    }

    prepareReviewPlan({ generateAfter: true });
  }

  function prepareReviewPlan({
    openTextTab = false,
    generateAfter = false,
  }: { openTextTab?: boolean; generateAfter?: boolean } = {}) {
    if (isPreparingReviewPlan) {
      return;
    }
    const normalizedText = (attachedDocumentText || rawText).trim();
    if (!normalizedText) {
      setError("Введите текст или загрузите документ.");
      return;
    }

    setError("");
    setGenerationResult(null);
    setIsPreparingReviewPlan(true);
    startTransition(() => {
      const effectiveChartSelectionByTableId = savedChartSelectionByTableId;
      const effectiveChartModeByTableId = savedChartModeByTableId;
      const effectiveHiddenSeriesByTableId = savedHiddenSeriesByTableId;

      const chartOverrides: ChartOverride[] = chartAssessments.map((assessment) => {
        const mode = effectiveChartModeByTableId[assessment.table_id] ?? "table";
        const selectedChart =
          mode === "chart"
            ? selectedChartSpec(
                assessment,
                effectiveChartSelectionByTableId,
                effectiveHiddenSeriesByTableId,
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
        .then((plan: PresentationPlan) => {
          setReviewPlan(plan);
          setCardSlideIndexes([]);
          if (generateAfter) {
            setIsGeneratingPresentation(true);
            return generatePresentation(plan).then((result) => {
              setGenerationResult(result);
            }).finally(() => setIsGeneratingPresentation(false));
          }
          if (openTextTab) {
            setDrawerTab("text");
            setIsStructureDrawerOpen(true);
          }
        })
        .catch((err: Error) => setError(err.message))
        .finally(() => setIsPreparingReviewPlan(false));
    });
  }

  function resetReviewPlan() {
    setReviewPlan(null);
    setCardSlideIndexes([]);
  }

  function clearAttachedDocument() {
    setAttachedDocumentName("");
    setAttachedDocumentText("");
    setDocumentTables([]);
    setDocumentBlocks([]);
    setChartAssessments([]);
    setChartSelectionByTableId({});
    setChartModeByTableId({});
    setHiddenSeriesByTableId({});
    setSavedChartSelectionByTableId({});
    setSavedChartModeByTableId({});
    setSavedHiddenSeriesByTableId({});
    setIsStructureDrawerOpen(false);
    setShowAllTablesInDrawer(false);
    resetReviewPlan();
  }

  function selectedChartSpec(
    assessment: ChartabilityAssessment,
    selectionByTableId = chartSelectionByTableId,
    hiddenByTableId = hiddenSeriesByTableId,
  ): ChartSpec | null {
    const chartId = selectionByTableId[assessment.table_id];
    const baseSpec = assessment.candidate_specs.find((item) => item.chart_id === chartId) ?? preferredChartSpec(assessment) ?? null;
    if (!baseSpec) {
      return null;
    }

    const hiddenSeries = new Set(hiddenByTableId[assessment.table_id] ?? []);
    return {
      ...baseSpec,
      chart_id: `${baseSpec.chart_id}:default`,
      series: baseSpec.series.map((series) => ({
        ...series,
        hidden: hiddenSeries.has(series.name),
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
    JSON.stringify(hiddenSeriesByTableId) !== JSON.stringify(savedHiddenSeriesByTableId);
  const chartableAssessments = chartAssessments.filter((assessment) => assessment.chartable);
  const visibleAssessments = showAllTablesInDrawer ? chartAssessments : chartableAssessments;
  const cardSlideChoices = reviewPlan ? eligibleCardSlides(reviewPlan) : [];

  function handleSaveStructureChoices() {
    setSavedChartSelectionByTableId(chartSelectionByTableId);
    setSavedChartModeByTableId(chartModeByTableId);
    setSavedHiddenSeriesByTableId(hiddenSeriesByTableId);
    setIsStructureDrawerOpen(false);
  }

  function editableCardItems(slide: SlideSpec): string[] {
    const explicitBullets = (slide.bullets ?? []).map((item) => item.trim()).filter(Boolean);
    if (explicitBullets.length >= 2) {
      return explicitBullets.slice(0, 3);
    }

    const blockItems = (slide.content_blocks ?? []).flatMap((block) => {
      if (block.kind === "bullet_list") {
        return block.items;
      }
      return block.text ? [block.text] : [];
    }).map((item) => item.trim()).filter(Boolean);
    if (blockItems.length >= 2) {
      return blockItems.slice(0, 3);
    }

    const text = (slide.text ?? "").replace(/\s+/g, " ").trim();
    if (!text) {
      return [];
    }

    const sentences = text.match(/[^.!?。！？]+[.!?。！？]?/g)?.map((item) => item.trim()).filter(Boolean) ?? [text];
    if (sentences.length >= 2) {
      return compactCardItems(sentences, 3);
    }
    if (text.length >= 90) {
      return compactCardItems(text.split(/[,;:]\s+|\s+-\s+/).map((item) => item.trim()).filter(Boolean), 3);
    }
    return [];
  }

  function compactCardItems(items: string[], maxItems: number): string[] {
    const cleaned = items.map((item) => item.trim()).filter(Boolean);
    if (cleaned.length <= maxItems) {
      return cleaned;
    }

    const buckets = Array.from({ length: maxItems }, () => "");
    cleaned.forEach((item, index) => {
      const bucketIndex = Math.min(maxItems - 1, Math.floor((index * maxItems) / cleaned.length));
      buckets[bucketIndex] = `${buckets[bucketIndex]} ${item}`.trim();
    });
    return buckets.filter(Boolean);
  }

  function splitCardItem(item: string): { title: string; description: string } {
    const normalized = item.replace(/\s+/g, " ").trim();
    const colonMatch = normalized.match(/^(.{4,54}?):\s+(.{12,})$/);
    if (colonMatch) {
      return { title: colonMatch[1].trim(), description: colonMatch[2].trim() };
    }

    const dashMatch = normalized.match(/^(.{4,54}?)\s+[—-]\s+(.{12,})$/);
    if (dashMatch) {
      return { title: dashMatch[1].trim(), description: dashMatch[2].trim() };
    }

    return { title: normalized, description: "" };
  }

  function encodeCardItem(item: string): string {
    const { title, description } = splitCardItem(item);
    return description ? `${title}\n${description}` : title;
  }

  function scoreCardSlide(slide: SlideSpec, items: string[]): { fit: CardSlideFit | null; reason: string } {
    if (items.length < 2 || items.length > 4) {
      return { fit: null, reason: "" };
    }

    const title = (slide.title ?? "").toLowerCase();
    const text = (slide.text ?? "").replace(/\s+/g, " ").trim();
    const lengths = items.map((item) => item.length);
    const longest = Math.max(...lengths);
    const shortest = Math.min(...lengths);
    const average = lengths.reduce((sum, length) => sum + length, 0) / lengths.length;
    const spread = longest / Math.max(shortest, 1);
    const hasListBlocks = (slide.content_blocks ?? []).some((block) => block.kind === "bullet_list");
    const hasExplicitBullets = (slide.bullets ?? []).filter((item) => item.trim()).length >= 2;
    const hasCardTitleCue = /фактор|преимуществ|этап|шаг|риск|направлен|принцип|драйвер|причин|задач|решени|сценари|вариант|метрик|эффект/.test(title);
    const hasDenseNarrative = text.length > 420 && !hasListBlocks && !hasExplicitBullets;
    const hasLongItems = longest > 150 || average > 115;
    const hasUnevenItems = spread > 3.2 && longest > 110;
    const numericSignals = items.filter((item) => /\d/.test(item)).length;

    let score = 0;
    if (items.length === 3) score += 3;
    if (items.length === 2 || items.length === 4) score += 2;
    if (hasExplicitBullets || hasListBlocks) score += 3;
    if (hasCardTitleCue) score += 2;
    if (average >= 18 && average <= 95) score += 2;
    if (spread <= 2.4) score += 1;
    if (hasDenseNarrative) score -= 4;
    if (hasLongItems) score -= 3;
    if (hasUnevenItems) score -= 2;
    if (numericSignals >= 2 && !hasExplicitBullets && !hasListBlocks) score -= 2;

    if (score >= 7) {
      return { fit: "high", reason: "Рекомендовано: короткие равноправные тезисы." };
    }
    if (score >= 4) {
      return { fit: "medium", reason: "Можно разложить на карточки." };
    }
    return { fit: null, reason: "" };
  }

  function eligibleCardSlides(plan: PresentationPlan): CardSlideChoice[] {
    return plan.slides
      .map((slide, index) => ({ index, slide, items: editableCardItems(slide) }))
      .filter(({ slide, items }) => {
        const layoutKey = slide.preferred_layout_key ?? "";
        const isDataSlide =
          slide.kind === "table" ||
          slide.kind === "chart" ||
          Boolean(slide.table) ||
          Boolean(slide.chart) ||
          Boolean(slide.source_table_id) ||
          layoutKey.includes("table") ||
          layoutKey.includes("chart");
        return slide.kind !== "title" && !isDataSlide && layoutKey !== "cards_3" && items.length >= 2;
      })
      .map(({ index, slide, items }) => ({
        index,
        slide,
        items,
        ...scoreCardSlide(slide, items),
      }))
      .filter((choice): choice is CardSlideChoice => {
        return choice.fit !== null;
      });
  }

  function toggleCardSlide(index: number) {
    setCardSlideIndexes((current) => {
      if (current.includes(index)) {
        return current.filter((item) => item !== index);
      }
      return [...current, index].sort((left, right) => left - right);
    });
  }

  function applyCardSlideChoices(plan: PresentationPlan, selectedIndexes: number[]): PresentationPlan {
    const selected = new Set(selectedIndexes);
    return {
      ...plan,
      slides: plan.slides.map((slide, index) => {
        if (!selected.has(index)) {
          return slide;
        }
        const cardItems = editableCardItems(slide).map(encodeCardItem);
        if (cardItems.length < 2) {
          return slide;
        }
        return {
          ...slide,
          kind: "bullets",
          text: null,
          bullets: cardItems,
          content_blocks: [],
          left_bullets: [],
          right_bullets: [],
          preferred_layout_key: "cards_3",
        };
      }),
    };
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

  function chartOptionLabel(spec: ChartSpec): string {
    const variantLabel = spec.variant_label?.trim();
    if (spec.chart_type === "combo") {
      return `Сравнение: ${spec.series.map((series) => series.name).join(", ")}`;
    }
    if (variantLabel) {
      return variantLabel;
    }
    return chartTypeLabel(spec.chart_type);
  }

  function chartTypeOptions(assessment: ChartabilityAssessment): string[] {
    return assessment.candidate_specs.reduce<string[]>((accumulator, spec) => {
      const chartType = chartControlType(spec);
      if (isSelectableChartType(chartType) && !accumulator.includes(chartType)) {
        accumulator.push(chartType);
      }
      return accumulator;
    }, []);
  }

  function selectedChartType(assessment: ChartabilityAssessment, selectedSpec: ChartSpec | null): string {
    if (selectedSpec) {
      const chartType = chartControlType(selectedSpec);
      if (isSelectableChartType(chartType)) {
        return chartType;
      }
    }
    return chartTypeOptions(assessment)[0] ?? preferredChartSpec(assessment)?.chart_type ?? "";
  }

  function candidateSpecsForChartType(assessment: ChartabilityAssessment, chartType: string): ChartSpec[] {
    return assessment.candidate_specs
      .filter((spec) => chartControlType(spec) === chartType)
      .sort((left, right) => chartVariantPriority(left) - chartVariantPriority(right));
  }

  function hasVariantChoices(assessment: ChartabilityAssessment, chartType: string): boolean {
    if (!isSelectableChartType(chartType)) {
      return false;
    }
    return candidateSpecsForChartType(assessment, chartType).length > 1;
  }

  function preferredChartSpec(assessment: ChartabilityAssessment): ChartSpec | null {
    const selectable = assessment.candidate_specs.filter((spec) => isSelectableChartType(chartControlType(spec)));
    const ordered = (selectable.length ? selectable : assessment.candidate_specs).sort(
      (left, right) => chartVariantPriority(left) - chartVariantPriority(right),
    );
    return ordered[0] ?? null;
  }

  function chartCardTitle(assessment: ChartabilityAssessment, selectedSpec: ChartSpec | null): string {
    if (selectedSpec?.title?.trim()) {
      return selectedSpec.title.trim();
    }

    const headers = assessment.structured_table?.cells[0]?.map((cell) => cell.text.trim()).filter(Boolean) ?? [];
    if (headers.length >= 2) {
      return `${headers[0]} · ${headers.slice(1, 3).join(" / ")}`;
    }
    if (headers.length === 1) {
      return `Таблица: ${headers[0]}`;
    }
    return assessment.chartable ? "Таблица для графика" : "Таблица документа";
  }

  function chartCardDescription(assessment: ChartabilityAssessment, selectedSpec: ChartSpec | null): string {
    if (!assessment.chartable) {
      return "Для этой таблицы лучше оставить табличный вид, чтобы не исказить данные.";
    }
    return "Эту таблицу можно использовать для графика.";
  }

  function assessmentHints(assessment: ChartabilityAssessment): string[] {
    const rawHints = [...assessment.reasons, ...assessment.warnings].filter((hint) => !isSystemHint(hint));
    if (rawHints.length === 0) {
      return assessment.chartable ? [] : ["Для этой таблицы график лучше не использовать."];
    }

    return rawHints.map((hint) => assessmentHintLabels[hint] ?? hint);
  }

  return (
    <main className="app-shell" data-testid="app-shell">
      <section className="hero-block" data-node-id="634:1739">
        <h1 className="hero-title" data-node-id="633:1698">
          Создай свою презентацию
        </h1>
        <p className="hero-description" data-node-id="633:3191">
          Превращайте документ в готовую презентацию в корпоративном стиле
          <br />
          Сервис извлекает структуру документа, собирает план слайдов и генерирует без ручной верстки.
        </p>
      </section>

      <div className="composer-stack">
        {generationResult ? (
          <div className="status-panel status-success" data-testid="generation-success">
            <button type="button" className="status-close" onClick={() => setGenerationResult(null)} aria-label="Закрыть">
              ×
            </button>
            <div className="status-title">Презентация готова</div>
            <div className="status-text" data-testid="generated-file-name">{generationResult.file_name}</div>
            <button
              type="button"
              className="primary-button status-download"
              data-testid="download-presentation"
              onClick={() => window.open(buildDownloadUrl(generationResult.download_url), "_blank", "noopener,noreferrer")}
            >
              Скачать презентацию
            </button>
          </div>
        ) : null}

        {error ? (
          <div className="status-panel status-error" data-testid="error-panel">
            <button type="button" className="status-close" onClick={() => setError("")} aria-label="Закрыть">
              ×
            </button>
            <div className="status-title">Ошибка</div>
            <div className="status-text" data-testid="error-text">{error}</div>
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

        <section className={`composer-card${rawText.trim() || attachedDocumentName ? " is-active" : ""}`} data-node-id="633:1701">
          <div className="composer-inner" data-node-id="634:1782">
            <div className="textarea-wrap" data-node-id="633:3164">
              <textarea
                value={rawText}
                onChange={(event) => {
                  setRawText(event.target.value);
                  resetReviewPlan();
                }}
                placeholder={initialText}
                className="main-textarea"
                data-testid="raw-text-input"
                rows={10}
              />
            </div>

            <div className="actions-row" data-node-id="634:1772">
              {attachedDocumentName ? (
                <div className="attached-file" data-testid="attached-document">
                  <span className="attached-file-name" title={attachedDocumentName}>
                    {attachedDocumentName}
                  </span>
                  <button
                    type="button"
                    className="attached-file-remove"
                    data-testid="remove-attached-document"
                    aria-label="Удалить файл"
                    onClick={clearAttachedDocument}
                  >
                    ×
                  </button>
                </div>
              ) : (
                <label className="secondary-button file-button" data-node-id="644:3605" data-testid="upload-document-trigger">
                  <input
                    type="file"
                    accept=".txt,.md,.markdown,.pdf,.docx,text/plain,text/markdown,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    className="sr-only"
                    data-testid="upload-document-input"
                    aria-label="Загрузить документ"
                    onChange={handleTextFileUpload}
                  />
                  <svg className="file-button-icon" viewBox="0 0 16 16" aria-hidden="true" focusable="false">
                    <path
                      d="M8 2.25a.75.75 0 0 1 .75.75v5.19l1.72-1.72a.75.75 0 1 1 1.06 1.06l-3 3a.75.75 0 0 1-1.06 0l-3-3a.75.75 0 0 1 1.06-1.06l1.72 1.72V3A.75.75 0 0 1 8 2.25ZM3.25 12a.75.75 0 0 1 .75.75h8a.75.75 0 0 1 1.5 0V13a1.25 1.25 0 0 1-1.25 1.25h-8.5A1.25 1.25 0 0 1 2.5 13v-.25A.75.75 0 0 1 3.25 12Z"
                      fill="currentColor"
                    />
                  </svg>
                  <span>Файл</span>
                </label>
              )}

              <div className="actions-group">
                {chartAssessments.length > 0 ? (
                  <button
                    type="button"
                    className="secondary-button"
                    data-testid="open-structure-drawer"
                    onClick={() => {
                      setDrawerTab("charts");
                      setIsStructureDrawerOpen(true);
                    }}
                  >
                    Подготовить ({chartAssessments.length})
                  </button>
                ) : null}

                <button
                  type="button"
                  className="primary-button"
                  data-node-id="634:1743"
                  data-testid="generate-presentation"
                  onClick={handleGenerate}
                  disabled={isPending || isPreparingReviewPlan || isGeneratingPresentation}
                >
                  {isPreparingReviewPlan || isGeneratingPresentation ? <span className="button-spinner" aria-hidden="true" /> : null}
                  Сгенерировать
                </button>
              </div>
            </div>
          </div>
        </section>

        {chartAssessments.length > 0 || reviewPlan ? (
          <StructureDrawer
            open={isStructureDrawerOpen}
            onOpenChange={setIsStructureDrawerOpen}
            title="Структура слайдов"
            description={`Таблиц: ${chartAssessments.length}. Текстовых вариантов: ${cardSlideChoices.length}.`}
            footer={
              <div className="drawer-footer-actions">
                <div className="drawer-footer-note">
                  {hasUnsavedStructureChanges ? "Есть несохранённые изменения по таблицам." : "Выбор будет использован при генерации."}
                </div>
                <button
                  type="button"
                  className="primary-button"
                  data-testid="save-structure-choices"
                  onClick={handleSaveStructureChoices}
                >
                  Сохранить
                </button>
              </div>
            }
          >
            <div className="drawer-tabs" role="tablist" aria-label="Настройки структуры">
              <button
                type="button"
                className={`drawer-tab${drawerTab === "charts" ? " is-active" : ""}`}
                data-testid="drawer-tab-charts"
                role="tab"
                aria-selected={drawerTab === "charts"}
                onClick={() => setDrawerTab("charts")}
              >
                Таблица / График
              </button>
              <button
                type="button"
                className={`drawer-tab${drawerTab === "text" ? " is-active" : ""}`}
                data-testid="drawer-tab-text"
                role="tab"
                aria-selected={drawerTab === "text"}
                onClick={() => {
                  setDrawerTab("text");
                  if (!reviewPlan && (attachedDocumentText.trim() || rawText.trim())) {
                    prepareReviewPlan();
                  }
                }}
              >
                Вид текста
              </button>
            </div>

            {drawerTab === "charts" ? (
              <>
                <div className="drawer-toolbar">
                  <div className="drawer-toolbar-copy">
                    {chartableAssessments.length > 0
                      ? "Сначала показаны только таблицы, которые можно превратить в графики."
                      : "В документе не найдено таблиц, которые система может уверенно превратить в графики."}
                  </div>
                  {chartAssessments.length > chartableAssessments.length ? (
                    <label className="drawer-switch">
                      <span>Показать все таблицы</span>
                      <input
                        type="checkbox"
                        className="drawer-switch-input"
                        checked={showAllTablesInDrawer}
                        onChange={(event) => setShowAllTablesInDrawer(event.target.checked)}
                      />
                      <span className="drawer-switch-track" aria-hidden="true">
                        <span className="drawer-switch-thumb" />
                      </span>
                    </label>
                  ) : null}
                </div>
                <section className="preview-gallery">
                  {visibleAssessments.map((assessment) => {
                const selectedSpec = selectedChartSpec(assessment);
                const mode = chartModeByTableId[assessment.table_id] ?? "table";
                const selectedChartId =
                  chartSelectionByTableId[assessment.table_id] ?? preferredChartSpec(assessment)?.chart_id ?? "";
                const selectedType = selectedChartType(assessment, selectedSpec);
                const typeOptions = chartTypeOptions(assessment);
                const variantOptions = candidateSpecsForChartType(assessment, selectedType);

                return (
                  <Card className="preview-card" key={assessment.table_id}>
                    <div data-testid={`assessment-card-${assessment.table_id}`}>
                    <CardHeader>
                      <CardTitle>
                        {chartCardTitle(assessment, selectedSpec)}
                      </CardTitle>
                      <CardDescription>
                        {chartCardDescription(assessment, selectedSpec)}
                      </CardDescription>
                    </CardHeader>
                    <CardContent className="preview-card-content">
                      <div className="preview-toolbar">
                        {assessment.chartable ? (
                          <>
                            <button
                              type="button"
                              className={`secondary-button${mode === "table" ? " is-active" : ""}`}
                              data-testid={`mode-table-${assessment.table_id}`}
                              onClick={() => setChartModeByTableId((current) => ({ ...current, [assessment.table_id]: "table" }))}
                            >
                              Таблица
                            </button>
                            <button
                              type="button"
                              className={`secondary-button${mode === "chart" ? " is-active" : ""}`}
                              data-testid={`mode-chart-${assessment.table_id}`}
                              onClick={() => setChartModeByTableId((current) => ({ ...current, [assessment.table_id]: "chart" }))}
                              disabled={!selectedSpec}
                            >
                              График
                            </button>
                          </>
                        ) : (
                          <div className="preview-toolbar-note">Для этой таблицы вариант с графиком не предлагается.</div>
                        )}
                        {assessment.chartable && mode === "chart" && typeOptions.length > 1 ? (
                          <Select
                            className="chart-type-select"
                            data-testid={`chart-type-${assessment.table_id}`}
                            value={selectedType}
                            onChange={(event) => {
                              const nextType = event.target.value;
                              const nextVariant = candidateSpecsForChartType(assessment, nextType)[0];
                              if (!nextVariant) {
                                return;
                              }
                              setChartSelectionByTableId((current) => ({ ...current, [assessment.table_id]: nextVariant.chart_id }));
                              setHiddenSeriesByTableId((current) => ({ ...current, [assessment.table_id]: [] }));
                            }}
                          >
                            {typeOptions.map((chartType) => (
                              <option key={chartType} value={chartType}>
                                {chartTypeLabel(chartType)}
                              </option>
                            ))}
                          </Select>
                        ) : null}
                        {assessment.chartable && mode === "chart" && hasVariantChoices(assessment, selectedType) ? (
                          <Select
                            className="chart-type-select"
                            data-testid={`chart-variant-${assessment.table_id}`}
                            value={selectedChartId}
                            onChange={(event) =>
                              setChartSelectionByTableId((current) => ({ ...current, [assessment.table_id]: event.target.value }))
                            }
                          >
                            {variantOptions.map((spec) => (
                              <option key={spec.chart_id} value={spec.chart_id}>
                                {chartOptionLabel(spec)}
                              </option>
                            ))}
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
                                  data-testid={`series-toggle-${assessment.table_id}-${series.name}`}
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
                    </div>
                  </Card>
                );
              })}
                </section>
              </>
            ) : (
              <section className="slide-review-panel" data-testid="slide-review-panel">
                <div className="slide-review-head">
                  <div>
                    <div className="slide-review-title">Вид текста</div>
                    <div className="slide-review-text">
                      Отметь слайды, где текст нужно заменить на карточки.
                    </div>
                  </div>
                  <button type="button" className="secondary-button" onClick={() => setCardSlideIndexes([])}>
                    Сбросить выбор
                  </button>
                </div>

                {cardSlideChoices.length > 0 ? (
                  <div className="slide-choice-list">
                    {cardSlideChoices.map(({ index, slide, items, fit, reason }) => {
                      const previewItems = items.map(splitCardItem);
                      return (
                        <label className="slide-choice" data-testid={`card-slide-choice-${index}`} key={`slide-choice-${index}`}>
                          <Input
                            type="checkbox"
                            className="slide-choice-input"
                            checked={cardSlideIndexes.includes(index)}
                            onChange={() => toggleCardSlide(index)}
                          />
                          <span className="slide-choice-body">
                            <span className="slide-choice-title">
                              {index + 1}. {slide.title || "Слайд без заголовка"}
                            </span>
                            <span className={`slide-choice-fit is-${fit}`}>{reason}</span>
                            <span className="slide-choice-preview">
                              {previewItems.map(({ title, description }) => description ? `${title}: ${description}` : title).join(" · ")}
                            </span>
                          </span>
                        </label>
                      );
                    })}
                  </div>
                ) : isPreparingReviewPlan ? (
                  <div className="slide-review-text" data-testid="preparing-card-slide-choices">
                    Подготавливаю текстовые слайды...
                  </div>
                ) : (
                  <div className="slide-review-text" data-testid="no-card-slide-choices">
                    В плане нет текстовых слайдов, которые можно безопасно разложить на карточки.
                  </div>
                )}
              </section>
            )}
          </StructureDrawer>
        ) : null}
      </div>
    </main>
  );
}
