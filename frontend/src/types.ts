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
  binding?: string | null;
  max_chars?: number | null;
  left_emu?: number | null;
  top_emu?: number | null;
  width_emu?: number | null;
  height_emu?: number | null;
  margin_left_emu?: number | null;
  margin_right_emu?: number | null;
  margin_top_emu?: number | null;
  margin_bottom_emu?: number | null;
};

export type LayoutSpec = {
  key: string;
  name: string;
  slide_master_index: number;
  slide_layout_index: number;
  preview_path?: string | null;
  supported_slide_kinds: string[];
  placeholders: PlaceholderSpec[];
};

export type PrototypeTokenSpec = {
  token: string;
  binding: string;
  shape_name?: string | null;
  left_emu?: number | null;
  top_emu?: number | null;
  width_emu?: number | null;
  height_emu?: number | null;
  margin_left_emu?: number | null;
  margin_right_emu?: number | null;
  margin_top_emu?: number | null;
  margin_bottom_emu?: number | null;
};

export type PrototypeSlideSpec = {
  key: string;
  name: string;
  source_slide_index: number;
  supported_slide_kinds: string[];
  tokens: PrototypeTokenSpec[];
};

export type TemplateManifest = {
  template_id: string;
  display_name: string;
  source_pptx: string;
  description?: string | null;
  generation_mode: "layout" | "prototype";
  default_layout_key?: string | null;
  layouts: LayoutSpec[];
  prototype_slides: PrototypeSlideSpec[];
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
  header_fill_colors: Array<string | null>;
  rows: string[][];
  row_fill_colors: Array<Array<string | null>>;
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
  variant_label?: string | null;
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
  transpose_allowed: boolean;
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

export type ChartOverride = {
  table_id: string;
  mode: "table" | "chart";
  selected_chart?: ChartSpec | null;
};

export type SlideContentBlockKind = "paragraph" | "bullet_list" | "callout" | "qa_item";

export type SlideContentBlock = {
  kind: SlideContentBlockKind;
  text?: string | null;
  items: string[];
};

export type SlideSpec = {
  kind: string;
  title?: string | null;
  subtitle?: string | null;
  text?: string | null;
  bullets: string[];
  content_blocks: SlideContentBlock[];
  left_bullets: string[];
  right_bullets: string[];
  table?: TableBlock | null;
  chart?: ChartSpec | null;
  source_table_id?: string | null;
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
