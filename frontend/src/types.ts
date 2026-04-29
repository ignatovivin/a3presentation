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
  editable_role?: string | null;
  editable_capabilities: string[];
  slot_group?: string | null;
  slot_group_order?: number | null;
  max_chars?: number | null;
  left_emu?: number | null;
  top_emu?: number | null;
  width_emu?: number | null;
  height_emu?: number | null;
  margin_left_emu?: number | null;
  margin_right_emu?: number | null;
  margin_top_emu?: number | null;
  margin_bottom_emu?: number | null;
  text_style?: TemplateTextStyleSpec | null;
  paragraph_styles?: TemplateParagraphStyleCatalog | null;
  shape_style?: TemplateShapeStyleSpec | null;
};

export type LayoutSpec = {
  key: string;
  name: string;
  slide_master_index: number;
  slide_layout_index: number;
  preview_path?: string | null;
  preview_image_base64?: string | null;
  preview_image_content_type?: string | null;
  supported_slide_kinds: string[];
  representation_hints: string[];
  placeholders: PlaceholderSpec[];
  background_color?: string | null;
  background_style?: TemplateShapeStyleSpec | null;
  background_xml?: string | null;
  background_image_base64?: string | null;
  background_image_content_type?: string | null;
};

export type PrototypeTokenSpec = {
  token: string;
  binding: string;
  shape_name?: string | null;
  editable_role?: string | null;
  editable_capabilities: string[];
  slot_group?: string | null;
  slot_group_order?: number | null;
  left_emu?: number | null;
  top_emu?: number | null;
  width_emu?: number | null;
  height_emu?: number | null;
  margin_left_emu?: number | null;
  margin_right_emu?: number | null;
  margin_top_emu?: number | null;
  margin_bottom_emu?: number | null;
  text_style?: TemplateTextStyleSpec | null;
  paragraph_styles?: TemplateParagraphStyleCatalog | null;
  shape_style?: TemplateShapeStyleSpec | null;
};

export type PrototypeSlideSpec = {
  key: string;
  name: string;
  source_slide_index: number;
  preview_image_base64?: string | null;
  preview_image_content_type?: string | null;
  supported_slide_kinds: string[];
  representation_hints: string[];
  tokens: PrototypeTokenSpec[];
};

export type TemplateThemeSpec = {
  color_scheme: Record<string, string>;
  font_scheme: Record<string, string>;
  master_text_styles: Record<string, TemplateTextStyleSpec>;
  master_paragraph_styles: Record<string, TemplateParagraphStyleCatalog>;
  master_shape_styles: Record<string, TemplateShapeStyleSpec>;
};

export type TemplateTextStyleSpec = {
  source?: string | null;
  role?: string | null;
  font_family?: string | null;
  font_size_pt?: number | null;
  font_weight?: number | null;
  bold?: boolean | null;
  italic?: boolean | null;
  underline?: boolean | null;
  color?: string | null;
  alignment?: string | null;
  vertical_anchor?: string | null;
  word_wrap?: boolean | null;
  auto_size?: string | null;
  line_spacing?: number | null;
  space_before_pt?: number | null;
  space_after_pt?: number | null;
  margin_left_emu?: number | null;
  margin_right_emu?: number | null;
  indent_emu?: number | null;
  default_tab_size_emu?: number | null;
  rtl?: boolean | null;
  bullet_type?: string | null;
  bullet_font?: string | null;
  bullet_char?: string | null;
  hanging_emu?: number | null;
  level?: number | null;
  kerning_pt?: number | null;
};

export type TemplateParagraphStyleCatalog = {
  level_styles: Record<string, TemplateTextStyleSpec>;
};

export type TemplateShapeStyleSpec = {
  role?: string | null;
  fill_type?: string | null;
  fill_color?: string | null;
  fill_transparency?: number | null;
  line_color?: string | null;
  line_width_pt?: number | null;
  line_transparency?: number | null;
  line_compound?: string | null;
  line_cap?: string | null;
  line_join?: string | null;
  geometry_preset?: string | null;
  rotation?: number | null;
  inset_left_emu?: number | null;
  inset_right_emu?: number | null;
  inset_top_emu?: number | null;
  inset_bottom_emu?: number | null;
  vertical_anchor?: string | null;
  horizontal_anchor?: string | null;
  shadow_type?: string | null;
  shadow_color?: string | null;
  glow_radius_pt?: number | null;
  soft_edge_radius_pt?: number | null;
  reflection_type?: string | null;
  effect_list: string[];
  theme_fill_ref?: string | null;
  theme_line_ref?: string | null;
  table_cell_margin_left_emu?: number | null;
  table_cell_margin_right_emu?: number | null;
  table_cell_margin_top_emu?: number | null;
  table_cell_margin_bottom_emu?: number | null;
  chart_plot_left_factor?: number | null;
  chart_plot_top_factor?: number | null;
  chart_plot_width_factor?: number | null;
  chart_plot_height_factor?: number | null;
  chart_legend_offset_x_emu?: number | null;
  chart_legend_offset_y_emu?: number | null;
  chart_category_axis_label_offset?: number | null;
  chart_value_axis_label_offset?: number | null;
};

export type TemplateComponentStyleSpec = {
  text_styles: Record<string, TemplateTextStyleSpec>;
  shape_style?: TemplateShapeStyleSpec | null;
  spacing_tokens: Record<string, string | number | boolean | null>;
  behavior_tokens: Record<string, string | number | boolean | null>;
};

export type ComponentGeometry = {
  left_emu?: number | null;
  top_emu?: number | null;
  width_emu?: number | null;
  height_emu?: number | null;
  margin_left_emu?: number | null;
  margin_right_emu?: number | null;
  margin_top_emu?: number | null;
  margin_bottom_emu?: number | null;
};

export type ComponentStyle = {
  text_style?: TemplateTextStyleSpec | null;
  paragraph_styles?: TemplateParagraphStyleCatalog | null;
  shape_style?: TemplateShapeStyleSpec | null;
};

export type ExtractedComponent = {
  component_id: string;
  source_kind: "layout" | "slide";
  source_index: number;
  source_name?: string | null;
  shape_name?: string | null;
  component_type: string;
  role: string;
  binding?: string | null;
  confidence: "high" | "medium" | "low";
  editability: "editable" | "semi_editable" | "decorative" | "unsafe";
  capabilities: string[];
  geometry: ComponentGeometry;
  style: ComponentStyle;
  text_excerpt?: string | null;
  child_component_ids: string[];
};

export type ExtractedSlideInventory = {
  source_kind: "layout" | "slide";
  source_index: number;
  name?: string | null;
  component_ids: string[];
  supported_slide_kinds: string[];
  representation_hints: string[];
};

export type ExtractedPresentationInventory = {
  components: ExtractedComponent[];
  slides: ExtractedSlideInventory[];
  warnings: string[];
  degradation_mode?: string | null;
  has_usable_layout_inventory: boolean;
  has_prototype_inventory: boolean;
};

export type TemplateManifest = {
  template_id: string;
  display_name: string;
  source_pptx: string;
  description?: string | null;
  generation_mode: "layout" | "prototype";
  default_layout_key?: string | null;
  design_tokens: Record<string, string | number | boolean | null>;
  component_styles: Record<string, TemplateComponentStyleSpec>;
  theme: TemplateThemeSpec;
  layouts: LayoutSpec[];
  prototype_slides: PrototypeSlideSpec[];
  inventory: ExtractedPresentationInventory;
};

export type InventoryTargetSummary = {
  key: string;
  name: string;
  source: "layout" | "prototype" | "direct_shape_binding";
  source_label?: string | null;
  supported_slide_kinds: string[];
  representation_hints: string[];
  editable_slot_count: number;
  editable_roles: string[];
};

export type TemplateInventorySummary = {
  generation_mode: "layout" | "prototype";
  usability_status: "usable" | "usable_with_degradation" | "not_safely_editable";
  has_usable_layout_inventory: boolean;
  has_prototype_inventory: boolean;
  degradation_mode?: string | null;
  warnings: string[];
  layout_target_count: number;
  prototype_target_count: number;
  direct_target_count: number;
  targets: InventoryTargetSummary[];
};

export type EditableTargetSummary = {
  key: string;
  name: string;
  source: "layout" | "prototype" | "direct_shape_binding";
  source_label?: string | null;
  runtime_profile_key?: string | null;
  supported_slide_kinds: string[];
  representation_hints: string[];
  editable_slot_count: number;
  editable_roles: string[];
};

export type DetectedComponentSummary = {
  component_id: string;
  source_kind: "layout" | "slide";
  source_index: number;
  source_name?: string | null;
  shape_name?: string | null;
  component_type: string;
  role: string;
  binding?: string | null;
  confidence: "high" | "medium" | "low";
  editability: "editable" | "semi_editable" | "decorative" | "unsafe";
  capabilities: string[];
  geometry: ComponentGeometry;
  text_excerpt?: string | null;
  child_component_ids: string[];
};

export type TemplateDetailsResponse = {
  manifest: TemplateManifest;
  has_template_file: boolean;
  inventory_summary: TemplateInventorySummary;
  editable_targets: EditableTargetSummary[];
  detected_components: DetectedComponentSummary[];
};

export type AnalyzeTemplateResponse = {
  template_id: string;
  manifest_path: string;
  inventory_summary: TemplateInventorySummary;
  editable_targets: EditableTargetSummary[];
  detected_components: DetectedComponentSummary[];
};

export type PlanWithTemplateResponse = {
  plan: PresentationPlan;
  manifest: TemplateManifest;
  inventory_summary: TemplateInventorySummary;
  editable_targets: EditableTargetSummary[];
  detected_components: DetectedComponentSummary[];
  slide_layout_reviews: SlideLayoutReview[];
};

export type SlideLayoutOption = {
  key: string;
  name: string;
  source: "layout" | "prototype" | "direct_shape_binding";
  source_label?: string | null;
  runtime_profile_key?: string | null;
  supported_slide_kinds: string[];
  representation_hints: string[];
  editable_slot_count: number;
  editable_roles: string[];
  supports_current_slide_kind: boolean;
  estimated_text_capacity_chars?: number | null;
  match_summary?: string | null;
  recommendation_label?: string | null;
  recommendation_reasons: string[];
};

export type SlideLayoutReview = {
  slide_index: number;
  current_target_key?: string | null;
  current_target_type?: "layout" | "prototype" | "direct_shape_binding" | "auto_layout" | null;
  current_target_source?: string | null;
  current_target_explanation?: string | null;
  current_target_confidence?: string | null;
  current_target_degradation_reasons: string[];
  current_runtime_profile_key?: string | null;
  available_layouts: SlideLayoutOption[];
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
  inventory_summary: TemplateInventorySummary;
  editable_targets: EditableTargetSummary[];
  detected_components: DetectedComponentSummary[];
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

export type SlideRenderTarget = {
  type: "layout" | "prototype" | "direct_shape_binding" | "auto_layout";
  key?: string | null;
  label?: string | null;
  source?: string | null;
  binding_keys: string[];
  degradation_reasons: string[];
  confidence?: string | null;
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
  runtime_profile_key?: string | null;
  render_target?: SlideRenderTarget | null;
  background_only?: boolean;
  background_xml?: string | null;
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
