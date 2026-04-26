import { ChangeEvent, type CSSProperties, useEffect, useState, useTransition } from "react";

import { buildDownloadUrl, buildPlan, buildPlanWithTemplate, extractTextFromDocument, fetchTemplate, fetchTemplates, generatePresentation, generatePresentationWithTemplate } from "@/api";
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
  SlideLayoutReview,
  SlideSpec,
  TableBlock,
  TemplateManifest,
  TemplateSummary,
} from "@/types";

type ManifestCardSlot = {
  editable_role?: string | null;
  editable_capabilities: string[];
  left_emu?: number | null;
  top_emu?: number | null;
  width_emu?: number | null;
  height_emu?: number | null;
};

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

const editableRoleLabels: Record<string, string> = {
  title: "заголовки",
  subtitle: "подзаголовки",
  body: "текст",
  bullet_list: "списки",
  bullet_item: "пункты",
  image: "изображения",
  table: "таблицы",
  chart: "графики",
};

const representationHintLabels: Record<string, string> = {
  cards: "карточки",
  table: "таблица",
  chart: "график",
  image: "изображение",
  contacts: "контакты",
  two_column: "две колонки",
};

const slideKindLabels: Record<string, string> = {
  title: "титульный",
  text: "текст",
  bullets: "список",
  table: "таблица",
  chart: "график",
  image: "изображение",
  two_column: "две колонки",
};

function displayLayoutSourceLabel(source: "layout" | "prototype", sourceLabel?: string | null): string {
  const normalized = sourceLabel?.trim().toLowerCase() ?? "";
  if (normalized.startsWith("prototype slide ")) {
    return `Прототипный слайд ${sourceLabel?.trim().slice("prototype slide ".length)}`;
  }
  if (normalized.startsWith("layout ")) {
    return `Макет ${sourceLabel?.trim().slice("layout ".length)}`;
  }
  if (sourceLabel?.trim()) {
    return sourceLabel.trim();
  }
  return source === "prototype" ? "Прототипный слайд" : "Макет";
}

function displayLayoutSourceType(source: "layout" | "prototype"): string {
  return source === "prototype" ? "Прототип" : "Макет";
}

function templateOriginBadgeLabel(hasAttachedTemplate: boolean, hasManifest: boolean): string | null {
  if (hasAttachedTemplate) {
    return "Пользовательский шаблон";
  }
  if (hasManifest) {
    return "Шаблон из каталога";
  }
  return null;
}

function layoutPurposeLabel(option: {
  representation_hints: string[];
  editable_roles: string[];
  supports_current_slide_kind: boolean;
}): string | null {
  const hint = option.representation_hints[0];
  if (hint) {
    return representationHintLabels[hint] ?? hint;
  }
  const role = option.editable_roles[0];
  if (role) {
    return editableRoleLabels[role] ?? role;
  }
  if (option.supports_current_slide_kind) {
    return "подходит по типу слайда";
  }
  return null;
}

function layoutOptionMeta(option: {
  source: "layout" | "prototype";
  source_label?: string | null;
  representation_hints: string[];
  editable_roles: string[];
  supports_current_slide_kind: boolean;
  estimated_text_capacity_chars?: number | null;
  match_summary?: string | null;
  recommendation_label?: string | null;
}): string {
  const parts = [
    option.recommendation_label,
    displayLayoutSourceLabel(option.source, option.source_label),
    layoutPurposeLabel(option),
    option.estimated_text_capacity_chars ? `до ~${option.estimated_text_capacity_chars} символов текста` : null,
  ].filter(Boolean);
  return parts.join(" · ");
}

function layoutRecommendationText(option: {
  recommendation_label?: string | null;
  recommendation_reasons: string[];
  match_summary?: string | null;
}): string {
  if (option.recommendation_reasons.length) {
    return option.recommendation_reasons.slice(0, 2).join(" ");
  }
  if (option.match_summary?.trim()) {
    return option.match_summary.trim();
  }
  return "Сервис отсортировал варианты по типу слайда, смыслу и запасу по вместимости.";
}

function resolveTemplateUiVars(manifest: TemplateManifest | null): CSSProperties {
  if (!manifest) {
    return {};
  }

  const tokens = manifest.design_tokens ?? {};
  const scheme = manifest.theme?.color_scheme ?? {};
  const primary = String(tokens.primary_color ?? scheme.accent3 ?? "#679aea");
  const accent = String(tokens.accent_color ?? scheme.accent4 ?? primary);
  const background = String(tokens.background_color ?? scheme.lt1 ?? "#f9f9f9");
  const surface = String(tokens.surface_color ?? scheme.accent1 ?? "#f5f7fb");
  const text = String(tokens.text_color ?? scheme.dk1 ?? "#171717");
  const muted = String(tokens.muted_text_color ?? scheme.dk2 ?? primary);
  const border = String(tokens.table_border_color ?? scheme.accent2 ?? "#dbe3f1");
  const title = String(tokens.title_color ?? text);
  const subtitle = String(tokens.subtitle_color ?? text);

  return {
    "--template-ui-primary": primary,
    "--template-ui-primary-strong": accent,
    "--template-ui-background": background,
    "--template-ui-surface": surface,
    "--template-ui-text": text,
    "--template-ui-muted": muted,
    "--template-ui-border": border,
    "--template-ui-title": title,
    "--template-ui-subtitle": subtitle,
    "--template-ui-primary-soft": `${primary}14`,
    "--template-ui-accent-soft": `${accent}12`,
  } as CSSProperties;
}

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

function transientTemplateIdFromFilename(fileName: string): string {
  const stem = fileName.replace(/\.[^.]+$/, "").trim().toLowerCase();
  const normalized = stem
    .replace(/[^a-z0-9а-яё_-]+/gi, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 64);
  return `uploaded_${normalized || "template"}`;
}

function summarizeEditableSlots(manifest: TemplateManifest): {
  total: number;
  grouped: number;
  roleLines: string[];
} {
  const roleCounts = new Map<string, number>();
  let total = 0;
  let grouped = 0;

  const registerSlot = (role?: string | null, capabilities?: string[], slotGroup?: string | null) => {
    if (!role && (!capabilities || capabilities.length === 0)) {
      return;
    }
    total += 1;
    if (slotGroup) {
      grouped += 1;
    }
    const normalizedRole = role ?? "body";
    roleCounts.set(normalizedRole, (roleCounts.get(normalizedRole) ?? 0) + 1);
  };

  manifest.layouts.forEach((layout) => {
    layout.placeholders.forEach((placeholder) => {
      registerSlot(placeholder.editable_role, placeholder.editable_capabilities, placeholder.slot_group);
    });
  });
  manifest.prototype_slides.forEach((slide) => {
    slide.tokens.forEach((token) => {
      registerSlot(token.editable_role, token.editable_capabilities, token.slot_group);
    });
  });

  const roleLines = Array.from(roleCounts.entries())
    .sort((left, right) => right[1] - left[1])
    .slice(0, 4)
    .map(([role, count]) => `${editableRoleLabels[role] ?? role}: ${count}`);

  return { total, grouped, roleLines };
}

function summarizeRepresentationHints(manifest: TemplateManifest): string[] {
  const counts = new Map<string, number>();
  const registerHints = (hints: string[]) => {
    hints.forEach((hint) => {
      counts.set(hint, (counts.get(hint) ?? 0) + 1);
    });
  };

  manifest.layouts.forEach((layout) => registerHints(layout.representation_hints));
  manifest.prototype_slides.forEach((slide) => registerHints(slide.representation_hints));

  return Array.from(counts.entries())
    .sort((left, right) => right[1] - left[1] || left[0].localeCompare(right[0]))
    .map(([hint, count]) => `${representationHintLabels[hint] ?? hint}: ${count} вариантов`);
}

function summarizeDetectedLayouts(manifest: TemplateManifest): string[] {
  return manifest.layouts
    .slice(0, 8)
    .map((layout) => {
      const hints = layout.representation_hints.length
        ? layout.representation_hints.map((hint) => representationHintLabels[hint] ?? hint).join(", ")
        : layout.supported_slide_kinds.join(", ");
      return hints ? `${layout.name} (${hints})` : layout.name;
    });
}

function isCardTextSlot(slot: ManifestCardSlot): boolean {
  if (slot.editable_role === "body" || slot.editable_role === "bullet_item" || slot.editable_role === "bullet_list") {
    return true;
  }
  if (slot.editable_role === "title" || slot.editable_role === "subtitle" || slot.editable_role === "image" || slot.editable_role === "table" || slot.editable_role === "chart") {
    return false;
  }
  return slot.editable_capabilities.includes("text") || slot.editable_capabilities.includes("list_item");
}

function scoreCardSlotCollection(slots: ManifestCardSlot[]): number | null {
  const textSlots = slots.filter(
    (slot) =>
      isCardTextSlot(slot) &&
      typeof slot.left_emu === "number" &&
      typeof slot.top_emu === "number" &&
      typeof slot.width_emu === "number" &&
      typeof slot.height_emu === "number" &&
      slot.width_emu > 0 &&
      slot.height_emu > 0,
  );
  if (textSlots.length < 2) {
    return null;
  }

  let bestScore: number | null = null;
  for (const baseSlot of textSlots) {
    const sameRow = textSlots.filter((slot) => {
      const topDelta = Math.abs((slot.top_emu ?? 0) - (baseSlot.top_emu ?? 0));
      const heightReference = Math.max(slot.height_emu ?? 0, baseSlot.height_emu ?? 0);
      return topDelta <= heightReference * 0.45;
    });
    if (sameRow.length < 2 || sameRow.length > 4) {
      continue;
    }

    const widths = sameRow.map((slot) => slot.width_emu ?? 0);
    const heights = sameRow.map((slot) => slot.height_emu ?? 0);
    const lefts = sameRow.map((slot) => slot.left_emu ?? 0).sort((left, right) => left - right);
    const tops = sameRow.map((slot) => slot.top_emu ?? 0);
    const widthSpread = Math.max(...widths) / Math.max(Math.min(...widths), 1);
    const heightSpread = Math.max(...heights) / Math.max(Math.min(...heights), 1);
    const topSpread = Math.max(...tops) - Math.min(...tops);
    const distinctColumns = new Set(lefts.map((value) => Math.round(value / 10000))).size;

    if (distinctColumns < sameRow.length) {
      continue;
    }
    if (widthSpread > 1.8 || heightSpread > 1.8) {
      continue;
    }

    const score =
      sameRow.length * 12 -
      Math.abs(sameRow.length - 3) * 4 -
      Math.round(widthSpread * 3) -
      Math.round(heightSpread * 2) -
      Math.round(topSpread / 100000);
    if (bestScore === null || score > bestScore) {
      bestScore = score;
    }
  }
  return bestScore;
}

function findCardCapableLayoutKey(manifest: TemplateManifest): string | null {
  const candidates: Array<{ key: string; score: number }> = [];

  manifest.layouts.forEach((layout) => {
    if (!layout.supported_slide_kinds.includes("bullets") && !layout.supported_slide_kinds.includes("text")) {
      return;
    }
    if (layout.representation_hints.includes("cards")) {
      candidates.push({ key: layout.key, score: 100 });
      return;
    }
    const score = scoreCardSlotCollection(layout.placeholders);
    if (score === null) {
      return;
    }
    candidates.push({ key: layout.key, score });
  });

  manifest.prototype_slides.forEach((slide) => {
    if (!slide.supported_slide_kinds.includes("bullets") && !slide.supported_slide_kinds.includes("text")) {
      return;
    }
    if (slide.representation_hints.includes("cards")) {
      candidates.push({ key: slide.key, score: 101 });
      return;
    }
    const score = scoreCardSlotCollection(slide.tokens);
    if (score === null) {
      return;
    }
    candidates.push({ key: slide.key, score: score + 1 });
  });

  candidates.sort((left, right) => right.score - left.score || left.key.localeCompare(right.key));
  return candidates[0]?.key ?? null;
}

function manifestSlideTarget(manifest: TemplateManifest, layoutKey: string) {
  if (!layoutKey) {
    return null;
  }
  return manifest.layouts.find((layout) => layout.key === layoutKey)
    ?? manifest.prototype_slides.find((slide) => slide.key === layoutKey)
    ?? null;
}

function targetSupportsDataRepresentation(manifest: TemplateManifest | null, layoutKey: string): boolean {
  if (!manifest || !layoutKey) {
    return false;
  }
  const target = manifestSlideTarget(manifest, layoutKey);
  if (!target) {
    return false;
  }
  if (target.representation_hints.includes("table") || target.representation_hints.includes("chart") || target.representation_hints.includes("image")) {
    return true;
  }
  if (target.supported_slide_kinds.includes("table") || target.supported_slide_kinds.includes("chart") || target.supported_slide_kinds.includes("image")) {
    return true;
  }
  if ("placeholders" in target) {
    return target.placeholders.some((placeholder) => (
      placeholder.editable_role === "table"
      || placeholder.editable_role === "chart"
      || placeholder.editable_role === "image"
      || placeholder.editable_capabilities.includes("table")
      || placeholder.editable_capabilities.includes("chart")
      || placeholder.editable_capabilities.includes("image")
    ));
  }
  return target.tokens.some((token) => (
    token.editable_role === "table"
    || token.editable_role === "chart"
    || token.editable_role === "image"
    || token.editable_capabilities.includes("table")
    || token.editable_capabilities.includes("chart")
    || token.editable_capabilities.includes("image")
  ));
}

function currentLayoutOption(review: SlideLayoutReview | null, slide: SlideSpec) {
  if (!review?.available_layouts.length) {
    return null;
  }
  const key = slide.preferred_layout_key ?? review.current_layout_key ?? review.available_layouts[0]?.key ?? "";
  return review.available_layouts.find((option) => option.key === key) ?? review.available_layouts[0] ?? null;
}

function inventoryTargetRuntimeProfileKey(manifest: TemplateManifest | null, targetKey: string | null): string | null {
  if (!manifest || !targetKey) {
    return null;
  }
  const layout = manifest.layouts.find((item) => item.key === targetKey);
  const prototype = manifest.prototype_slides.find((item) => item.key === targetKey);
  const representationHints = new Set(layout?.representation_hints ?? prototype?.representation_hints ?? []);
  const editableRoles = new Set(
    layout
      ? layout.placeholders.map((item) => item.editable_role).filter((item): item is string => Boolean(item))
      : (prototype?.tokens.map((item) => item.editable_role).filter((item): item is string => Boolean(item)) ?? []),
  );
  const supportedKinds = new Set(layout?.supported_slide_kinds ?? prototype?.supported_slide_kinds ?? []);

  if (representationHints.has("contacts")) {
    return "contacts";
  }
  if (representationHints.has("cards")) {
    return editableRoles.has("metric_value") ? "cards_kpi" : "cards_3";
  }
  if (representationHints.has("two_column")) {
    return editableRoles.has("icon") ? "list_with_icons" : "two_column";
  }
  if (representationHints.has("table") || editableRoles.has("table") || editableRoles.has("chart")) {
    return "table";
  }
  if (representationHints.has("image") || editableRoles.has("image")) {
    return "image_text";
  }
  if (editableRoles.has("bullet_list") || editableRoles.has("bullet_item")) {
    return representationHints.has("icons") ? "list_with_icons" : "list_full_width";
  }
  if (supportedKinds.has("title")) {
    return "cover";
  }
  return "text_full_width";
}

export function App() {
  const [templates, setTemplates] = useState<TemplateSummary[]>([]);
  const [selectedTemplateId, setSelectedTemplateId] = useState("");
  const [rawText, setRawText] = useState("");
  const [attachedDocumentName, setAttachedDocumentName] = useState("");
  const [attachedDocumentText, setAttachedDocumentText] = useState("");
  const [attachedTemplateFile, setAttachedTemplateFile] = useState<File | null>(null);
  const [attachedTemplateName, setAttachedTemplateName] = useState("");
  const [attachedTemplateManifest, setAttachedTemplateManifest] = useState<TemplateManifest | null>(null);
  const [selectedTemplateManifest, setSelectedTemplateManifest] = useState<TemplateManifest | null>(null);
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
  const [slideLayoutReviews, setSlideLayoutReviews] = useState<SlideLayoutReview[]>([]);
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

  useEffect(() => {
    if (!selectedTemplateId || attachedTemplateFile) {
      return;
    }

    let isActive = true;
    fetchTemplate(selectedTemplateId)
      .then((response) => {
        if (!isActive) {
          return;
        }
        setSelectedTemplateManifest(response.manifest);
      })
      .catch(() => {
        if (!isActive) {
          return;
        }
        setSelectedTemplateManifest(null);
      });

    return () => {
      isActive = false;
    };
  }, [attachedTemplateFile, selectedTemplateId]);

  const effectiveTemplateId = attachedTemplateFile
    ? transientTemplateIdFromFilename(attachedTemplateFile.name)
    : selectedTemplateId;
  const effectiveTemplateManifest = attachedTemplateManifest ?? selectedTemplateManifest;
  const selectedTemplateSummary = templates.find((item) => item.template_id === selectedTemplateId) ?? null;
  const templateUiVars = resolveTemplateUiVars(effectiveTemplateManifest);

  function handleTextFileUpload(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }
    if (!file.name.toLowerCase().endsWith(".docx")) {
      setError("Загрузите документ в формате .docx.");
      event.target.value = "";
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
          setSlideLayoutReviews([]);
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

  function handleTemplateUpload(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }
    if (!file.name.toLowerCase().endsWith(".pptx")) {
      setError("Загрузите шаблон в формате .pptx.");
      event.target.value = "";
      return;
    }

    setError("");
    setGenerationResult(null);
    setAttachedTemplateFile(file);
    setAttachedTemplateName(file.name);
    setAttachedTemplateManifest(null);
    setSelectedTemplateManifest(null);
    event.target.value = "";
  }

  function generateCurrentPlan(plan: PresentationPlan): Promise<GeneratePresentationResponse> {
    return attachedTemplateFile
      ? generatePresentationWithTemplate(plan, attachedTemplateFile)
      : generatePresentation(plan);
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
        generateCurrentPlan(applyCardSlideChoices(reviewPlan, cardSlideIndexes, cardTargetLayoutKey, effectiveTemplateManifest))
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
    if (!effectiveTemplateId) {
      setError("Выберите шаблон или загрузите .pptx шаблон.");
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

      const planRequest = {
        template_id: effectiveTemplateId,
        raw_text: normalizedText,
        title: "A3 Presentation",
        tables: documentTables,
        blocks: documentBlocks,
        chart_overrides: chartOverrides,
      };

      const buildPlanPromise = attachedTemplateFile
        ? buildPlanWithTemplate(planRequest, attachedTemplateFile).then((response) => {
            setAttachedTemplateManifest(response.manifest);
            setSlideLayoutReviews(response.slide_layout_reviews ?? []);
            return response.plan;
          })
        : buildPlan(planRequest).then((plan) => {
            setAttachedTemplateManifest(null);
            setSlideLayoutReviews([]);
            return plan;
          });

      buildPlanPromise
        .then((plan: PresentationPlan) => {
          setReviewPlan(plan);
          setCardSlideIndexes([]);
          if (generateAfter) {
            setIsGeneratingPresentation(true);
            return generateCurrentPlan(plan).then((result) => {
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
    setSlideLayoutReviews([]);
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

  function clearAttachedTemplate() {
    setAttachedTemplateFile(null);
    setAttachedTemplateName("");
    setAttachedTemplateManifest(null);
    setGenerationResult(null);
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
  const cardTargetLayoutKey = effectiveTemplateManifest ? findCardCapableLayoutKey(effectiveTemplateManifest) : null;
  const cardSlideChoices = reviewPlan ? eligibleCardSlides(reviewPlan, cardTargetLayoutKey, effectiveTemplateManifest) : [];
  const attachedTemplateSlotSummary = effectiveTemplateManifest ? summarizeEditableSlots(effectiveTemplateManifest) : null;
  const activeTemplateRepresentationSummary = effectiveTemplateManifest ? summarizeRepresentationHints(effectiveTemplateManifest) : [];
  const detectedLayoutSummary = effectiveTemplateManifest ? summarizeDetectedLayouts(effectiveTemplateManifest) : [];
  const activeSlideLayoutReviews = reviewPlan
    ? reviewPlan.slides.map((slide, index) => ({
        index,
        slide,
        review: slideLayoutReviews.find((item) => item.slide_index === index) ?? null,
      }))
    : [];

  function handleSaveStructureChoices() {
    const shouldRebuildReviewPlan = hasUnsavedStructureChanges;
    setSavedChartSelectionByTableId(chartSelectionByTableId);
    setSavedChartModeByTableId(chartModeByTableId);
    setSavedHiddenSeriesByTableId(hiddenSeriesByTableId);
    if (shouldRebuildReviewPlan) {
      resetReviewPlan();
      setGenerationResult(null);
    }
    setIsStructureDrawerOpen(false);
  }

  function handleSlideLayoutChange(slideIndex: number, layoutKey: string) {
    const selectedReview = slideLayoutReviews.find((review) => review.slide_index === slideIndex) ?? null;
    const selectedOption = selectedReview?.available_layouts.find((option) => option.key === layoutKey) ?? null;
    const runtimeProfileKey = selectedOption?.runtime_profile_key
      ?? inventoryTargetRuntimeProfileKey(effectiveTemplateManifest, layoutKey);
    setReviewPlan((current) => {
      if (!current) {
        return current;
      }
      return {
        ...current,
        slides: current.slides.map((slide, index) => (
          index === slideIndex
            ? {
              ...slide,
              preferred_layout_key: layoutKey || null,
              runtime_profile_key: runtimeProfileKey ?? slide.runtime_profile_key ?? null,
            }
            : slide
        )),
      };
    });
    setSlideLayoutReviews((current) => current.map((review) => (
      review.slide_index === slideIndex
        ? {
          ...review,
          current_layout_key: layoutKey || null,
          current_runtime_profile_key: runtimeProfileKey ?? review.current_runtime_profile_key ?? null,
        }
        : review
    )));
  }

  function editableCardItems(slide: SlideSpec): string[] {
    const explicitBullets = (slide.bullets ?? []).map((item) => item.trim()).filter(Boolean);
    if (explicitBullets.length >= 2) {
      return explicitBullets.slice(0, 4);
    }

    const blockItems = (slide.content_blocks ?? []).flatMap((block) => {
      if (block.kind === "bullet_list") {
        return block.items;
      }
      return block.text ? [block.text] : [];
    }).map((item) => item.trim()).filter(Boolean);
    if (blockItems.length >= 2) {
      return blockItems.slice(0, 4);
    }

    const text = (slide.text ?? "").replace(/\s+/g, " ").trim();
    if (!text) {
      return [];
    }

    const sentences = text.match(/[^.!?。！？]+[.!?。！？]?/g)?.map((item) => item.trim()).filter(Boolean) ?? [text];
    if (sentences.length >= 2) {
      return compactCardItems(sentences, 4);
    }
    if (text.length >= 90) {
      return compactCardItems(text.split(/[,;:]\s+|\s+-\s+/).map((item) => item.trim()).filter(Boolean), 4);
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
    const metricMatch = normalized.match(
      /^([<>~≈]?\s*\d+(?:[.,]\d+)?(?:\s*(?:%|‰|млн|млрд|тыс|трлн|сек(?:унд[аы]?)?|с|мин|ч|дн(?:ей|я)?|₽|руб(?:\.|лей|ля|ль)?))*)\s+(.+)$/iu,
    );
    if (metricMatch) {
      return { title: metricMatch[1].replace(/\s+/g, " ").trim(), description: metricMatch[2].trim() };
    }

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
    const multiline = item.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
    if (multiline.length >= 2) {
      return multiline.join("\n");
    }
    const { title, description } = splitCardItem(item);
    return description ? `${title}\n${description}` : title;
  }

  function numericCardSignals(items: string[]): number {
    return items.map(splitCardItem).filter(({ title, description }) => /\d/.test(title) && description.length >= 3).length;
  }

  function hasNumericCardLayout(items: string[]): boolean {
    return numericCardSignals(items) >= 2;
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
    const numericSignals = numericCardSignals(items);
    const isNumericLayout = numericSignals >= 2;

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
    if (isNumericLayout) score += 3;

    if (score >= 7) {
      return { fit: "high", reason: isNumericLayout ? "Рекомендовано: KPI-карточки с числами." : "Рекомендовано: короткие равноправные тезисы." };
    }
    if (score >= 4) {
      return { fit: "medium", reason: isNumericLayout ? "Можно разложить как KPI-карточки." : "Можно разложить на карточки." };
    }
    return { fit: null, reason: "" };
  }

  function eligibleCardSlides(
    plan: PresentationPlan,
    targetLayoutKey: string | null,
    manifest: TemplateManifest | null,
  ): CardSlideChoice[] {
    if (!targetLayoutKey) {
      return [];
    }
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
          targetSupportsDataRepresentation(manifest, layoutKey);
        return slide.kind !== "title" && !isDataSlide && layoutKey !== targetLayoutKey && items.length >= 2;
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

  function applyCardSlideChoices(
    plan: PresentationPlan,
    selectedIndexes: number[],
    targetLayoutKey: string | null,
    manifest: TemplateManifest | null,
  ): PresentationPlan {
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
          preferred_layout_key: targetLayoutKey ?? slide.preferred_layout_key,
          runtime_profile_key: inventoryTargetRuntimeProfileKey(manifest, targetLayoutKey) ?? slide.runtime_profile_key ?? "cards_3",
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
    <main className="app-shell" data-testid="app-shell" style={templateUiVars}>
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
            <div className="template-config-panel" data-testid="template-config-panel">
                <div className="template-config-head">
                  <div>
                    <div className="template-config-title">Активный шаблон</div>
                    <div className="template-config-text">
                      Можно выбрать системный шаблон или загрузить свой `.pptx`.
                    </div>
                  </div>
                {templateOriginBadgeLabel(Boolean(attachedTemplateFile), Boolean(effectiveTemplateManifest)) ? (
                  <div className="template-config-badge">
                    {templateOriginBadgeLabel(Boolean(attachedTemplateFile), Boolean(effectiveTemplateManifest))}
                  </div>
                ) : null}
              </div>

              <div className="template-config-controls">
                <Select
                  value={selectedTemplateId}
                  className="template-select"
                  data-testid="template-select"
                  onChange={(event) => {
                    setSelectedTemplateId(event.target.value);
                    setGenerationResult(null);
                    resetReviewPlan();
                  }}
                  disabled={Boolean(attachedTemplateFile) || templates.length === 0}
                >
                  {templates.map((template) => (
                    <option key={template.template_id} value={template.template_id}>
                      {template.display_name}
                    </option>
                  ))}
                </Select>
              </div>

              {effectiveTemplateManifest ? (
                <div className="template-analysis-grid" data-testid="template-analysis-summary">
                  <div className="template-analysis-card">
                    <div className="template-analysis-label">Активный шаблон</div>
                    <div className="template-analysis-value">
                      {attachedTemplateName || selectedTemplateSummary?.display_name || effectiveTemplateManifest.display_name}
                    </div>
                    <div className="template-analysis-meta">
                      {attachedTemplateFile ? "Загружен пользователем для текущей генерации." : "Выбран из каталога шаблонов."}
                    </div>
                  </div>
                  <div className="template-analysis-card">
                    <div className="template-analysis-label">Что можно менять</div>
                    <div className="template-analysis-value">{attachedTemplateSlotSummary?.total ?? 0}</div>
                    <div className="template-analysis-meta">
                      {attachedTemplateSlotSummary?.grouped ? `сгруппированных областей: ${attachedTemplateSlotSummary.grouped}` : "Отдельные области без явных групп."}
                    </div>
                  </div>
                  <div className="template-analysis-card">
                    <div className="template-analysis-label">Карточный режим</div>
                    <div className="template-analysis-value">{cardTargetLayoutKey ? "Доступен" : "Пока не найден"}</div>
                    <div className="template-analysis-meta">
                      {cardSlideChoices.length > 0 ? `Можно применить к ${cardSlideChoices.length} слайдам.` : "Оценка появится после построения плана."}
                    </div>
                  </div>
                </div>
              ) : null}

              {attachedTemplateSlotSummary?.roleLines.length || activeTemplateRepresentationSummary.length ? (
                <div className="template-analysis-lines">
                  {attachedTemplateSlotSummary?.roleLines.length ? (
                    <div className="template-analysis-line" data-testid="template-slot-roles">
                      <span className="template-analysis-line-label">Редактируемые области:</span>
                      <span>{attachedTemplateSlotSummary.roleLines.join(" · ")}</span>
                    </div>
                  ) : null}
                  {activeTemplateRepresentationSummary.length ? (
                    <div className="template-analysis-line" data-testid="template-representation-hints">
                      <span className="template-analysis-line-label">Подходит для:</span>
                      <span>{activeTemplateRepresentationSummary.slice(0, 5).join(" · ")}</span>
                    </div>
                  ) : null}
                  {detectedLayoutSummary.length ? (
                    <div className="template-analysis-line" data-testid="template-detected-layouts">
                      <span className="template-analysis-line-label">Найденные макеты:</span>
                      <span>{detectedLayoutSummary.join(" · ")}</span>
                    </div>
                  ) : null}
                </div>
              ) : null}
            </div>

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
              <div className="upload-actions">
                {attachedDocumentName ? (
                  <div className="attached-file" data-testid="attached-document">
                    <span className="attached-file-name" title={attachedDocumentName}>
                      {attachedDocumentName}
                    </span>
                    <button
                      type="button"
                      className="attached-file-remove"
                      data-testid="remove-attached-document"
                      aria-label="Удалить текстовый документ"
                      onClick={clearAttachedDocument}
                    >
                      ×
                    </button>
                  </div>
                ) : (
                  <label className="secondary-button file-button" data-node-id="644:3605" data-testid="upload-document-trigger">
                    <input
                      type="file"
                      accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                      className="sr-only"
                      data-testid="upload-document-input"
                      aria-label="Загрузить текст"
                      onChange={handleTextFileUpload}
                    />
                    <svg className="file-button-icon" viewBox="0 0 16 16" aria-hidden="true" focusable="false">
                      <path
                        d="M8 2.25a.75.75 0 0 1 .75.75v5.19l1.72-1.72a.75.75 0 1 1 1.06 1.06l-3 3a.75.75 0 0 1-1.06 0l-3-3a.75.75 0 0 1 1.06-1.06l1.72 1.72V3A.75.75 0 0 1 8 2.25ZM3.25 12a.75.75 0 0 1 .75.75h8a.75.75 0 0 1 1.5 0V13a1.25 1.25 0 0 1-1.25 1.25h-8.5A1.25 1.25 0 0 1 2.5 13v-.25A.75.75 0 0 1 3.25 12Z"
                        fill="currentColor"
                      />
                    </svg>
                    <span>Загрузка текста</span>
                  </label>
                )}

                {attachedTemplateName ? (
                  <div className="attached-file" data-testid="attached-template">
                    <div className="attached-file-text">
                      <span className="attached-file-name" title={attachedTemplateName}>
                        {attachedTemplateName}
                      </span>
                      {attachedTemplateManifest?.display_name ? (
                        <span className="attached-file-meta" title={attachedTemplateManifest.display_name}>
                          {attachedTemplateManifest.display_name}
                        </span>
                      ) : null}
                      {attachedTemplateSlotSummary ? (
                        <span
                          className="attached-file-meta attached-file-meta-slots"
                          title={`Редактируемые слоты: ${attachedTemplateSlotSummary.total}`}
                        >
                          {attachedTemplateSlotSummary.total} редактируемых слотов
                          {attachedTemplateSlotSummary.grouped > 0 ? ` · групп: ${attachedTemplateSlotSummary.grouped}` : ""}
                        </span>
                      ) : null}
                      {attachedTemplateSlotSummary?.roleLines.length ? (
                        <span className="attached-file-meta attached-file-meta-slots" title={attachedTemplateSlotSummary.roleLines.join(", ")}>
                          {attachedTemplateSlotSummary.roleLines.join(" · ")}
                        </span>
                      ) : null}
                    </div>
                    <button
                      type="button"
                      className="attached-file-remove"
                      data-testid="remove-attached-template"
                      aria-label="Удалить шаблон"
                      onClick={clearAttachedTemplate}
                    >
                      ×
                    </button>
                  </div>
                ) : (
                  <label className="secondary-button file-button" data-testid="upload-template-trigger">
                    <input
                      type="file"
                      accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"
                      className="sr-only"
                      data-testid="upload-template-input"
                      aria-label="Загрузить шаблон"
                      onChange={handleTemplateUpload}
                    />
                    <svg className="file-button-icon" viewBox="0 0 16 16" aria-hidden="true" focusable="false">
                      <path
                        d="M2.75 3.5A1.75 1.75 0 0 1 4.5 1.75h4.19c.46 0 .9.18 1.24.5l2.82 2.82c.32.33.5.78.5 1.24v6.19a1.75 1.75 0 0 1-1.75 1.75h-7A1.75 1.75 0 0 1 2.75 12.5v-9Zm6.5-.15V5.5c0 .14.11.25.25.25h2.15L9.25 3.35ZM5 8a.75.75 0 0 0 0 1.5h6A.75.75 0 0 0 11 8H5Zm0 2.5A.75.75 0 0 0 5 12h4a.75.75 0 0 0 0-1.5H5Z"
                        fill="currentColor"
                      />
                    </svg>
                    <span>Загрузка шаблона</span>
                  </label>
                )}
              </div>

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
                      Выберите вариант оформления для каждого текстового слайда. Самые подходящие макеты показаны сверху, а карточный режим можно включить отдельно.
                    </div>
                  </div>
                  <button type="button" className="secondary-button" onClick={() => setCardSlideIndexes([])}>
                    Сбросить выбор
                  </button>
                </div>

                {activeSlideLayoutReviews.length > 0 ? (
                  <div className="slide-choice-list">
                    {activeSlideLayoutReviews.map(({ index, slide, review }) => (
                      <div className="slide-choice" data-testid={`layout-slide-choice-${index}`} key={`layout-slide-choice-${index}`}>
                        <span className="slide-choice-body">
                          <span className="slide-choice-title">
                            {index + 1}. {slide.title || "Слайд без заголовка"}
                          </span>
                          <span className="slide-choice-meta-row">
                            <span className="slide-choice-fit">
                              {slideKindLabels[slide.kind] ?? slide.kind}
                              {slide.runtime_profile_key ? ` · profile ${slide.runtime_profile_key}` : ""}
                              {slide.preferred_layout_key ? ` · target ${slide.preferred_layout_key}` : ""}
                            </span>
                            {currentLayoutOption(review, slide) ? (
                              <span
                                className={`slide-choice-source is-${currentLayoutOption(review, slide)?.source}`}
                                data-testid={`layout-source-badge-${index}`}
                                title={displayLayoutSourceLabel(
                                  currentLayoutOption(review, slide)?.source ?? "layout",
                                  currentLayoutOption(review, slide)?.source_label,
                                )}
                              >
                                {displayLayoutSourceType(currentLayoutOption(review, slide)?.source ?? "layout")}
                              </span>
                            ) : null}
                          </span>
                          {currentLayoutOption(review, slide) ? (
                            <span className="slide-choice-source-label" data-testid={`layout-source-label-${index}`}>
                              {displayLayoutSourceLabel(
                                currentLayoutOption(review, slide)?.source ?? "layout",
                                currentLayoutOption(review, slide)?.source_label,
                              )}
                            </span>
                          ) : null}
                          {review?.available_layouts.length ? (
                            <>
                              <Select
                                className="chart-type-select"
                                data-testid={`slide-layout-select-${index}`}
                                value={slide.preferred_layout_key ?? review.current_layout_key ?? review.available_layouts[0]?.key ?? ""}
                                onChange={(event) => handleSlideLayoutChange(index, event.target.value)}
                              >
                                {review.available_layouts.map((option) => {
                                  const meta = layoutOptionMeta(option);
                                  return (
                                    <option key={`${index}-${option.key}`} value={option.key}>
                                      {option.name} · {meta || displayLayoutSourceLabel(option.source, option.source_label)}
                                    </option>
                                  );
                                })}
                              </Select>
                              <span className="slide-choice-preview">{layoutRecommendationText(currentLayoutOption(review, slide) ?? review.available_layouts[0])}</span>
                            </>
                          ) : (
                            <span className="slide-choice-preview">Для этого слайда пока не нашлось подходящих вариантов макета.</span>
                          )}
                        </span>
                      </div>
                    ))}
                  </div>
                ) : null}

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
