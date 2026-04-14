from __future__ import annotations

from enum import Enum

from pydantic import BaseModel, Field


class TemplateTextStyleSpec(BaseModel):
    source: str | None = None
    role: str | None = None
    font_family: str | None = None
    font_size_pt: float | None = None
    font_weight: int | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    color: str | None = None
    alignment: str | None = None
    vertical_anchor: str | None = None
    word_wrap: bool | None = None
    auto_size: str | None = None
    line_spacing: float | None = None
    space_before_pt: float | None = None
    space_after_pt: float | None = None
    margin_left_emu: int | None = None
    margin_right_emu: int | None = None
    indent_emu: int | None = None
    default_tab_size_emu: int | None = None
    rtl: bool | None = None
    bullet_type: str | None = None
    bullet_font: str | None = None
    bullet_char: str | None = None
    hanging_emu: int | None = None
    level: int | None = None
    kerning_pt: float | None = None

    @property
    def bullet_font_family(self) -> str | None:
        return self.bullet_font

    @bullet_font_family.setter
    def bullet_font_family(self, value: str | None) -> None:
        self.bullet_font = value

    @property
    def bullet_character(self) -> str | None:
        return self.bullet_char

    @bullet_character.setter
    def bullet_character(self, value: str | None) -> None:
        self.bullet_char = value

    @property
    def bullet_indent_emu(self) -> int | None:
        return self.indent_emu

    @bullet_indent_emu.setter
    def bullet_indent_emu(self, value: int | None) -> None:
        self.indent_emu = value

    @property
    def hanging_indent_emu(self) -> int | None:
        return self.hanging_emu

    @hanging_indent_emu.setter
    def hanging_indent_emu(self, value: int | None) -> None:
        self.hanging_emu = value


class TemplateParagraphStyleCatalog(BaseModel):
    level_styles: dict[str, TemplateTextStyleSpec] = Field(default_factory=dict)


class TemplateShapeStyleSpec(BaseModel):
    role: str | None = None
    fill_type: str | None = None
    fill_color: str | None = None
    fill_transparency: float | None = None
    line_color: str | None = None
    line_width_pt: float | None = None
    line_transparency: float | None = None
    line_compound: str | None = None
    line_cap: str | None = None
    line_join: str | None = None
    geometry_preset: str | None = None
    rotation: float | None = None
    inset_left_emu: int | None = None
    inset_right_emu: int | None = None
    inset_top_emu: int | None = None
    inset_bottom_emu: int | None = None
    vertical_anchor: str | None = None
    horizontal_anchor: str | None = None
    shadow_type: str | None = None
    shadow_color: str | None = None
    glow_radius_pt: float | None = None
    soft_edge_radius_pt: float | None = None
    reflection_type: str | None = None
    effect_list: list[str] = Field(default_factory=list)
    theme_fill_ref: str | None = None
    theme_line_ref: str | None = None
    table_cell_margin_left_emu: int | None = None
    table_cell_margin_right_emu: int | None = None
    table_cell_margin_top_emu: int | None = None
    table_cell_margin_bottom_emu: int | None = None
    chart_plot_left_factor: float | None = None
    chart_plot_top_factor: float | None = None
    chart_plot_width_factor: float | None = None
    chart_plot_height_factor: float | None = None
    chart_legend_offset_x_emu: int | None = None
    chart_legend_offset_y_emu: int | None = None
    chart_category_axis_label_offset: int | None = None
    chart_value_axis_label_offset: int | None = None


class TemplateThemeSpec(BaseModel):
    color_scheme: dict[str, str] = Field(default_factory=dict)
    font_scheme: dict[str, str] = Field(default_factory=dict)
    master_text_styles: dict[str, TemplateTextStyleSpec] = Field(default_factory=dict)
    master_paragraph_styles: dict[str, TemplateParagraphStyleCatalog] = Field(default_factory=dict)
    master_shape_styles: dict[str, TemplateShapeStyleSpec] = Field(default_factory=dict)


class GenerationMode(str, Enum):
    LAYOUT = "layout"
    PROTOTYPE = "prototype"


class PlaceholderKind(str, Enum):
    TITLE = "title"
    SUBTITLE = "subtitle"
    BODY = "body"
    IMAGE = "image"
    TABLE = "table"
    CHART = "chart"
    FOOTER = "footer"
    UNKNOWN = "unknown"


class PlaceholderSpec(BaseModel):
    name: str
    kind: PlaceholderKind = PlaceholderKind.UNKNOWN
    idx: int | None = None
    shape_name: str | None = None
    binding: str | None = None
    max_chars: int | None = None
    left_emu: int | None = None
    top_emu: int | None = None
    width_emu: int | None = None
    height_emu: int | None = None
    margin_left_emu: int | None = None
    margin_right_emu: int | None = None
    margin_top_emu: int | None = None
    margin_bottom_emu: int | None = None
    text_style: TemplateTextStyleSpec | None = None
    paragraph_styles: TemplateParagraphStyleCatalog | None = None
    shape_style: TemplateShapeStyleSpec | None = None


class LayoutSpec(BaseModel):
    key: str
    name: str
    slide_master_index: int = 0
    slide_layout_index: int
    preview_path: str | None = None
    supported_slide_kinds: list[str] = Field(default_factory=list)
    placeholders: list[PlaceholderSpec] = Field(default_factory=list)
    background_color: str | None = None
    background_style: TemplateShapeStyleSpec | None = None
    background_xml: str | None = None
    background_image_base64: str | None = None
    background_image_content_type: str | None = None


class PrototypeTokenSpec(BaseModel):
    token: str
    binding: str
    shape_name: str | None = None
    left_emu: int | None = None
    top_emu: int | None = None
    width_emu: int | None = None
    height_emu: int | None = None
    margin_left_emu: int | None = None
    margin_right_emu: int | None = None
    margin_top_emu: int | None = None
    margin_bottom_emu: int | None = None
    text_style: TemplateTextStyleSpec | None = None
    paragraph_styles: TemplateParagraphStyleCatalog | None = None
    shape_style: TemplateShapeStyleSpec | None = None


class PrototypeSlideSpec(BaseModel):
    key: str
    name: str
    source_slide_index: int
    supported_slide_kinds: list[str] = Field(default_factory=list)
    tokens: list[PrototypeTokenSpec] = Field(default_factory=list)


class TemplateManifest(BaseModel):
    template_id: str
    display_name: str
    source_pptx: str = "template.pptx"
    description: str | None = None
    generation_mode: GenerationMode = GenerationMode.LAYOUT
    default_layout_key: str | None = None
    design_tokens: dict[str, str | float | int | bool | None] = Field(default_factory=dict)
    theme: TemplateThemeSpec = Field(default_factory=TemplateThemeSpec)
    layouts: list[LayoutSpec] = Field(default_factory=list)
    prototype_slides: list[PrototypeSlideSpec] = Field(default_factory=list)
