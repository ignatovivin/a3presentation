import type {
  AnalyzeTemplateResponse,
  AutoUploadTemplateResponse,
  ChartOverride,
  DocumentBlock,
  ExtractTextResponse,
  GeneratePresentationResponse,
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

async function readJson<T>(response: Response): Promise<T> {
  if (!response.ok) {
    const message = await response.text();
    throw new Error(message || `Request failed with status ${response.status}`);
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
