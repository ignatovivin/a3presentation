export type TemplateSummary = {
  template_id: string;
  display_name: string;
  description?: string | null;
};

export type PlaceholderSpec = {
  name: string;
  kind: string;
  idx?: number | null;
  shape_name?: string | null;
  max_chars?: number | null;
};

export type LayoutSpec = {
  key: string;
  name: string;
  slide_layout_index: number;
  preview_path?: string | null;
  supported_slide_kinds: string[];
  placeholders: PlaceholderSpec[];
};

export type TemplateManifest = {
  template_id: string;
  display_name: string;
  source_pptx: string;
  description?: string | null;
  default_layout_key?: string | null;
  layouts: LayoutSpec[];
};

export type TemplateDetailsResponse = {
  manifest: TemplateManifest;
  has_template_file: boolean;
};

export type AnalyzeTemplateResponse = {
  template_id: string;
  manifest_path: string;
};

export type GeneratePresentationResponse = {
  output_path: string;
  file_name: string;
  download_url: string;
};

export type AutoUploadTemplateResponse = {
  template_id: string;
  manifest_path: string;
  template_path: string;
  analyzed: boolean;
};

export type ExtractTextResponse = {
  file_name: string;
  text: string;
  tables: TableBlock[];
  blocks: DocumentBlock[];
  chart_assessments: ChartabilityAssessment[];
};

export type DocumentBlock = {
  kind: string;
  text?: string | null;
  level?: number | null;
  style_name?: string | null;
  style_id?: string | null;
  items: string[];
  table?: TableBlock | null;
  hyperlinks?: string[];
  run_count?: number | null;
  image_name?: string | null;
  image_content_type?: string | null;
  image_base64?: string | null;
};

export type TableBlock = {
  headers: string[];
  rows: string[][];
};

export type ChartConfidence = "high" | "medium" | "low" | "none";

export type ChartTableClassification =
  | "single_series_category"
  | "multi_series_category"
  | "time_series"
  | "composition"
  | "ranking"
  | "matrix_numeric"
  | "text_dominant"
  | "mixed_ambiguous"
  | "not_chartable";

export type ChartType = "bar" | "column" | "line" | "stacked_bar" | "stacked_column" | "combo" | "pie";

export type StructuredCell = {
  text: string;
  normalized_text: string;
  value_type: string;
  parsed_value?: number | null;
  unit?: string | null;
  annotation?: string | null;
  is_header_like: boolean;
};

export type StructuredTable = {
  table_id: string;
  header_rows: number[];
  label_columns: number[];
  numeric_columns: number[];
  time_columns: number[];
  data_start_row: number;
  cells: StructuredCell[][];
  summary_rows: number[];
  warnings: string[];
};

export type ChartSeries = {
  name: string;
  values: number[];
  unit?: string | null;
  axis: string;
  hidden: boolean;
};

export type ChartSpec = {
  chart_id: string;
  source_table_id: string;
  chart_type: ChartType;
  title?: string | null;
  categories: string[];
  series: ChartSeries[];
  x_axis_title?: string | null;
  y_axis_title?: string | null;
  legend_visible: boolean;
  data_labels_visible: boolean;
  value_format: string;
  stacking: string;
  confidence: ChartConfidence;
  warnings: string[];
};

export type ChartabilityAssessment = {
  table_id: string;
  chartable: boolean;
  classification: ChartTableClassification;
  confidence: ChartConfidence;
  reasons: string[];
  warnings: string[];
  candidate_specs: ChartSpec[];
  structured_table?: StructuredTable | null;
};

export type SlideSpec = {
  kind: string;
  title?: string | null;
  subtitle?: string | null;
  text?: string | null;
  bullets: string[];
  left_bullets: string[];
  right_bullets: string[];
  table?: TableBlock | null;
  notes?: string | null;
  preferred_layout_key?: string | null;
  image_base64?: string | null;
  image_content_type?: string | null;
};

export type PresentationPlan = {
  template_id: string;
  title: string;
  author?: string | null;
  subject?: string | null;
  slides: SlideSpec[];
};
