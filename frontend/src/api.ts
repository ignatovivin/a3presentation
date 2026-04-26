import type {
  AnalyzeTemplateResponse,
  AutoUploadTemplateResponse,
  ChartOverride,
  DocumentBlock,
  ExtractTextResponse,
  GeneratePresentationResponse,
  PlanWithTemplateResponse,
  PresentationPlan,
  TableBlock,
  TemplateDetailsResponse,
  TemplateSummary,
} from "./types";

const rawApiBaseUrl = import.meta.env.VITE_API_BASE_URL?.trim();
const API_BASE_URL = rawApiBaseUrl ? rawApiBaseUrl.replace(/\/+$/, "") : "/api";

function buildApiUrl(path: string): string {
  return `${API_BASE_URL}${path}`;
}

function normalizeApiErrorMessage(message: string, status: number): string {
  const normalized = message.trim();
  if (!normalized) {
    return status >= 500
      ? "Сервис временно недоступен. Повторите попытку через минуту."
      : "Запрос не удалось выполнить. Проверьте введённые данные и повторите попытку.";
  }

  if (normalized.includes("template_file must be a .pptx")) {
    return "Шаблон должен быть файлом PowerPoint в формате .pptx.";
  }
  if (normalized.includes("file must be a .docx")) {
    return "Документ должен быть файлом Word в формате .docx.";
  }
  if (normalized.includes("template_file is empty")) {
    return "Загруженный шаблон пуст. Выберите корректный файл .pptx.";
  }
  if (normalized.includes("Failed to analyze uploaded template")) {
    return "Не удалось прочитать структуру шаблона PowerPoint. Проверьте, что файл не повреждён и сохранён в формате .pptx.";
  }
  if (normalized.includes("No extractable text found")) {
    return "В документе не найден текст для извлечения. Проверьте, что файл содержит текст, а не только изображения.";
  }
  if (normalized.includes("Failed to extract text")) {
    return "Не удалось прочитать документ. Проверьте, что файл не повреждён и сохранён в формате .docx.";
  }
  if (normalized.includes("Invalid payload_json") || normalized.includes("Invalid plan_json")) {
    return "Внутренние данные запроса повреждены. Обновите страницу и повторите попытку.";
  }
  if (normalized.includes("Generated deck failed layout quality gate")) {
    if (normalized.includes("missing_table_shape")) {
      return "Шаблон не содержит подходящую область для таблицы на одном из слайдов. Выберите другой макет или другой шаблон.";
    }
    if (normalized.includes("missing_chart_shape")) {
      return "Шаблон не содержит подходящую область для графика на одном из слайдов. Выберите другой макет или другой шаблон.";
    }
    if (normalized.includes("missing_image_shape")) {
      return "Шаблон не содержит подходящую область для изображения на одном из слайдов. Выберите другой макет или другой шаблон.";
    }
    if (normalized.includes("content_order_mismatch")) {
      return "Не удалось сохранить порядок текста на одном из слайдов. Попробуйте другой макет или сократите исходный текст.";
    }
    if (normalized.includes("card_overlap") || normalized.includes("two_column_overlap") || normalized.includes("image_text_overlap")) {
      return "Не удалось собрать слайд без пересечения блоков. Попробуйте другой макет или сократите содержимое.";
    }
    if (normalized.includes("chart_type_mismatch") || normalized.includes("chart_series_count_mismatch")) {
      return "Не удалось корректно построить график для выбранного варианта. Попробуйте другой тип графика или оставьте таблицу.";
    }
    return "Не удалось собрать презентацию без ошибок вёрстки. Попробуйте сократить текст, выбрать другой макет или использовать другой шаблон.";
  }
  if (normalized.includes("Failed to generate a valid PowerPoint file")) {
    return "Не удалось собрать корректный PowerPoint-файл для выбранного шаблона. Попробуйте другой макет или шаблон.";
  }
  if (normalized.includes("Failed to generate PowerPoint file")) {
    return "Не удалось сгенерировать PowerPoint-файл. Повторите попытку или измените входные данные.";
  }
  if (normalized.includes("Template '") && normalized.includes("' not found")) {
    return "Выбранный шаблон больше недоступен. Выберите другой шаблон и повторите попытку.";
  }
  if (normalized.includes("Template PPTX not found")) {
    return "Файл выбранного шаблона недоступен. Выберите другой шаблон.";
  }

  return normalized;
}

async function readErrorMessage(response: Response): Promise<string> {
  const contentType = response.headers.get("content-type") ?? "";
  if (contentType.includes("application/json")) {
    try {
      const payload = await response.json() as { detail?: unknown; message?: unknown };
      if (typeof payload.detail === "string") {
        return payload.detail;
      }
      if (typeof payload.message === "string") {
        return payload.message;
      }
    } catch {
      return "";
    }
  }
  try {
    return await response.text();
  } catch {
    return "";
  }
}

async function readJson<T>(response: Response): Promise<T> {
  if (!response.ok) {
    const message = await readErrorMessage(response);
    throw new Error(normalizeApiErrorMessage(message, response.status));
  }
  return response.json() as Promise<T>;
}

export async function fetchTemplates(): Promise<TemplateSummary[]> {
  const response = await fetch(buildApiUrl("/templates"));
  return readJson<TemplateSummary[]>(response);
}

export async function fetchTemplate(templateId: string): Promise<TemplateDetailsResponse> {
  const response = await fetch(buildApiUrl(`/templates/${templateId}`));
  return readJson<TemplateDetailsResponse>(response);
}

export async function buildPlan(payload: {
  template_id: string;
  raw_text: string;
  title?: string;
  tables?: TableBlock[];
  blocks?: DocumentBlock[];
  chart_overrides?: ChartOverride[];
}): Promise<PresentationPlan> {
  const response = await fetch(buildApiUrl("/plans/from-text"), {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });
  return readJson<PresentationPlan>(response);
}

export async function buildPlanWithTemplate(
  payload: {
    template_id: string;
    raw_text: string;
    title?: string;
    tables?: TableBlock[];
    blocks?: DocumentBlock[];
    chart_overrides?: ChartOverride[];
  },
  templateFile: File,
): Promise<PlanWithTemplateResponse> {
  const formData = new FormData();
  formData.set("payload_json", JSON.stringify(payload));
  formData.set("template_file", templateFile);

  const response = await fetch(buildApiUrl("/plans/from-text-with-template"), {
    method: "POST",
    body: formData,
  });
  return readJson<PlanWithTemplateResponse>(response);
}

export async function generatePresentation(plan: PresentationPlan): Promise<GeneratePresentationResponse> {
  const response = await fetch(buildApiUrl("/presentations/generate"), {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(plan),
  });
  return readJson<GeneratePresentationResponse>(response);
}

export async function generatePresentationWithTemplate(
  plan: PresentationPlan,
  templateFile: File,
): Promise<GeneratePresentationResponse> {
  const formData = new FormData();
  formData.set("plan_json", JSON.stringify(plan));
  formData.set("template_file", templateFile);

  const response = await fetch(buildApiUrl("/presentations/generate-with-template"), {
    method: "POST",
    body: formData,
  });
  return readJson<GeneratePresentationResponse>(response);
}

export async function analyzeTemplate(templateId: string, displayName?: string): Promise<AnalyzeTemplateResponse> {
  const query = displayName ? `?display_name=${encodeURIComponent(displayName)}` : "";
  const response = await fetch(buildApiUrl(`/templates/${templateId}/analyze${query}`), {
    method: "POST",
  });
  return readJson<AnalyzeTemplateResponse>(response);
}

export async function uploadTemplateFromUi(payload: {
  template_id: string;
  display_name: string;
  description?: string;
  template_file: File;
}): Promise<AutoUploadTemplateResponse> {
  const formData = new FormData();
  formData.set("template_id", payload.template_id);
  formData.set("display_name", payload.display_name);
  if (payload.description) {
    formData.set("description", payload.description);
  }
  formData.set("template_file", payload.template_file);

  const response = await fetch(buildApiUrl("/templates/auto"), {
    method: "POST",
    body: formData,
  });
  return readJson<AutoUploadTemplateResponse>(response);
}

export async function extractTextFromDocument(file: File): Promise<ExtractTextResponse> {
  const formData = new FormData();
  formData.set("file", file);

  const response = await fetch(buildApiUrl("/documents/extract-text"), {
    method: "POST",
    body: formData,
  });
  return readJson<ExtractTextResponse>(response);
}

export function buildDownloadUrl(downloadUrl: string): string {
  return buildApiUrl(downloadUrl);
}
