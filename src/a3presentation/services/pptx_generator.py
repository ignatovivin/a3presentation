from __future__ import annotations

import base64
import copy
import math
import re
import zipfile
from datetime import UTC, datetime
from io import BytesIO
from pathlib import Path
from copy import deepcopy

from pptx.chart.axis import ValueAxis
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE, XL_MARKER_STYLE
from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
from pptx.util import Pt

from a3presentation.domain.chart import ChartSpec, ChartType
from a3presentation.domain.presentation import (
    PresentationPlan,
    SlideContentBlock,
    SlideContentBlockKind,
    SlideKind,
    SlideSpec,
    TableBlock,
)
from a3presentation.domain.template import (
    GenerationMode,
    LayoutSpec,
    PlaceholderKind,
    PrototypeSlideSpec,
    TemplateManifest,
)
from a3presentation.services.chart_style import CHART_STYLE_CONFIG
from a3presentation.services.chart_render_contract import (
    PRIMARY_AXIS,
    SECONDARY_AXIS,
    chart_axis_number_format,
    chart_axis_number_format_for_axis,
    render_chart_spec,
    uses_secondary_value_axis,
)
from a3presentation.services.layout_capacity import (
    LayoutCapacityProfile,
    derive_capacity_profile_for_geometry,
    geometry_policy_for_layout,
    profile_for_layout,
    spacing_policy_for_layout,
)


class PptxGenerator:
    TOKEN_PATTERN = re.compile(r"{{\s*([a-zA-Z0-9_]+)\s*}}")
    RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    SKIPPED_RELATIONSHIP_TYPES = {
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
    }
    EMU_PER_PT = 12700
    TITLE_CONTENT_GAP_EMU = 180000
    TITLE_BODY_GAP_NO_SUBTITLE_EMU = 300000
    CONTENT_FOOTER_GAP_EMU = 180000
    COVER_TITLE_TOP_EMU = 651176
    COVER_TITLE_LEFT_EMU = 444249
    COVER_TITLE_WIDTH_EMU = 10693901
    COVER_TITLE_MIN_HEIGHT_EMU = 1422646
    COVER_META_LEFT_EMU = 444249
    COVER_META_WIDTH_EMU = 8200000
    COVER_META_DEFAULT_TOP_EMU = 2438400
    COVER_META_MIN_HEIGHT_EMU = 700000
    COVER_META_GAP_EMU = 220000
    COVER_BOTTOM_LIMIT_EMU = 6200000
    FULL_CONTENT_LEFT_EMU = 442913
    FULL_CONTENT_WIDTH_EMU = 11198224
    FOOTER_TOP_EMU = 6384626
    FOOTER_HEIGHT_EMU = 260000
    DEFAULT_TEXT_MARGIN_X_EMU = 91440
    DEFAULT_TEXT_MARGIN_Y_EMU = 45720
    GEOMETRY_PROFILE_TOLERANCE_EMU = 120000
    BUILTIN_LAYOUT_KEYS = {
        "text_full_width",
        "dense_text_full_width",
        "list_full_width",
        "table",
        "image_text",
        "cards_3",
        "list_with_icons",
        "contacts",
        "cover",
    }
    _CHART_STYLE_CONFIG = None

    @classmethod
    def _chart_style_config(cls) -> dict:
        if cls._CHART_STYLE_CONFIG is None:
            cls._CHART_STYLE_CONFIG = copy.deepcopy(CHART_STYLE_CONFIG)
        return cls._CHART_STYLE_CONFIG

    @classmethod
    def _style_rgb(cls, key: str) -> RGBColor:
        value = cls._chart_style_config()[key].lstrip("#")
        return RGBColor.from_string(value)

    @classmethod
    def _series_color(cls, index: int) -> RGBColor:
        palette = cls._chart_style_config()["palette"]
        return RGBColor.from_string(palette[index % len(palette)].lstrip("#"))

    def generate(self, template_path: Path, manifest: TemplateManifest, plan: PresentationPlan, output_dir: Path) -> Path:
        if manifest.generation_mode == GenerationMode.PROTOTYPE and manifest.prototype_slides:
            presentation = self._generate_from_prototypes(template_path, manifest, plan)
        else:
            presentation = self._generate_from_layouts(template_path, manifest, plan)

        self._apply_core_properties(presentation, plan)
        output_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now(UTC).strftime("%Y%m%dT%H%M%S%fZ")
        output_stem = self._build_output_stem(plan.title or plan.template_id, timestamp)
        output_path = output_dir / f"{output_stem}.pptx"
        presentation.save(str(output_path))
        self._validate_output_file(output_path, expected_slide_count=len(plan.slides))
        return output_path

    def _build_output_stem(self, title: str, timestamp: str) -> str:
        normalized = re.sub(r"\s+", " ", (title or "").strip())
        safe = re.sub(r"[^A-Za-z0-9А-Яа-яЁё _-]+", "", normalized).strip(" ._-")
        safe = safe.replace(" ", "_")
        if not safe:
            safe = "presentation"
        return f"{safe[:48]}_{timestamp}"

    def _generate_from_prototypes(self, template_path: Path, manifest: TemplateManifest, plan: PresentationPlan) -> Presentation:
        output_presentation = Presentation(str(template_path))
        source_slides = [output_presentation.slides[index] for index in range(len(output_presentation.slides))]
        self._remove_all_slides(output_presentation)

        for slide_spec in plan.slides:
            prototype = self._resolve_prototype_slide(manifest, slide_spec)
            source_slide = source_slides[prototype.source_slide_index]
            target_slide = self._clone_slide(output_presentation, source_slide)
            self._replace_tokens_in_slide(target_slide, prototype, slide_spec, plan.title)
            self._apply_layout_expansion_and_flow(target_slide, slide_spec.preferred_layout_key or prototype.key, slide_spec)

        return output_presentation

    def _generate_from_layouts(self, template_path: Path, manifest: TemplateManifest, plan: PresentationPlan) -> Presentation:
        presentation = Presentation(str(template_path))
        self._remove_all_slides(presentation)
        self._apply_core_properties(presentation, plan)

        for slide_spec in plan.slides:
            layout = self._resolve_layout(manifest, slide_spec)
            slide_layout = presentation.slide_masters[layout.slide_master_index].slide_layouts[layout.slide_layout_index]
            slide = presentation.slides.add_slide(slide_layout)
            self._fill_slide_from_layout(slide, slide_spec, layout, plan.title)

        return presentation

    def _apply_core_properties(self, presentation: Presentation, plan: PresentationPlan) -> None:
        props = presentation.core_properties
        props.title = plan.title
        props.author = plan.author or "a3presentation"
        if plan.subject:
            props.subject = plan.subject

    def _remove_all_slides(self, presentation: Presentation) -> None:
        for index in range(len(presentation.slides) - 1, -1, -1):
            slide_id = presentation.slides._sldIdLst[index]
            relationship_id = slide_id.rId
            presentation.part.drop_rel(relationship_id)
            del presentation.slides._sldIdLst[index]

    def _clone_slide(self, presentation: Presentation, source_slide):
        blank_layout = presentation.slide_layouts[6] if len(presentation.slide_layouts) > 6 else presentation.slide_layouts[-1]
        target_slide = presentation.slides.add_slide(blank_layout)
        self._clear_slide_shapes(target_slide)

        relationship_map: dict[str, str] = {}
        for relationship in source_slide.part.rels.values():
            if relationship.reltype in self.SKIPPED_RELATIONSHIP_TYPES:
                continue
            new_relationship_id = target_slide.part.relate_to(
                relationship.target_ref if relationship.is_external else relationship.target_part,
                relationship.reltype,
                relationship.is_external,
            )
            relationship_map[relationship.rId] = new_relationship_id

        source_background = source_slide._element.cSld.bg
        if source_background is not None:
            target_background = copy.deepcopy(source_background)
            self._remap_relationship_ids(target_background, relationship_map)
            if target_slide._element.cSld.bg is not None:
                target_slide._element.cSld.remove(target_slide._element.cSld.bg)
            target_slide._element.cSld.insert(0, target_background)

        for shape_element in list(source_slide.shapes._spTree.iterchildren()):
            if shape_element.tag.endswith("nvGrpSpPr") or shape_element.tag.endswith("grpSpPr"):
                continue
            cloned_element = copy.deepcopy(shape_element)
            self._remap_relationship_ids(cloned_element, relationship_map)
            target_slide.shapes._spTree.insert_element_before(cloned_element, "p:extLst")

        return target_slide

    def _clear_slide_shapes(self, slide) -> None:
        for shape_element in list(slide.shapes._spTree.iterchildren()):
            if shape_element.tag.endswith("nvGrpSpPr") or shape_element.tag.endswith("grpSpPr"):
                continue
            slide.shapes._spTree.remove(shape_element)

    def _remap_relationship_ids(self, element, relationship_map: dict[str, str]) -> None:
        for current_element in element.iter():
            for attr_name, attr_value in list(current_element.attrib.items()):
                if attr_value in relationship_map and attr_name.startswith(f"{{{self.RELATIONSHIP_NAMESPACE}}}"):
                    current_element.set(attr_name, relationship_map[attr_value])

    def _resolve_prototype_slide(self, manifest: TemplateManifest, slide_spec: SlideSpec) -> PrototypeSlideSpec:
        if slide_spec.kind == SlideKind.CHART:
            for prototype_slide in manifest.prototype_slides:
                if self._prototype_supports_chart(prototype_slide):
                    return prototype_slide

        if slide_spec.preferred_layout_key:
            for prototype_slide in manifest.prototype_slides:
                if prototype_slide.key == slide_spec.preferred_layout_key:
                    return prototype_slide

        for prototype_slide in manifest.prototype_slides:
            if slide_spec.kind.value in prototype_slide.supported_slide_kinds:
                return prototype_slide

        return manifest.prototype_slides[0]

    def _prototype_supports_chart(self, prototype: PrototypeSlideSpec) -> bool:
        if prototype.key == "chart" or SlideKind.CHART.value in prototype.supported_slide_kinds:
            return True
        return any(token.binding in {"chart", "chart_image"} for token in prototype.tokens)

    def _replace_tokens_in_slide(self, slide, prototype: PrototypeSlideSpec, slide_spec: SlideSpec, presentation_title: str) -> None:
        token_values = self._build_token_value_map(slide_spec, presentation_title)
        used_shapes: set[str] = set()
        layout_key = slide_spec.preferred_layout_key or "text_full_width"
        layout_profile = profile_for_layout(layout_key)

        # Preferred path for real templates: bind by explicit shape name from manifest.
        for token_spec in prototype.tokens:
            if not token_spec.shape_name:
                continue
            target_shape = next((shape for shape in slide.shapes if shape.name == token_spec.shape_name), None)
            if target_shape is None:
                continue
            self._apply_shape_spec_metadata(target_shape, token_spec)
            shape_profile = self._capacity_profile_for_shape(layout_key, target_shape, layout_profile)
            self._fill_shape_by_binding(target_shape, token_spec.binding, slide_spec, presentation_title, shape_profile)
            self._apply_shape_spec_metadata(target_shape, token_spec)
            used_shapes.add(token_spec.shape_name)

        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            if shape.name in used_shapes:
                continue

            original_text = shape.text or ""
            matches = self.TOKEN_PATTERN.findall(original_text)
            if not matches:
                continue

            normalized = original_text.strip()
            single_token_match = self.TOKEN_PATTERN.fullmatch(normalized)
            if single_token_match:
                token_name = single_token_match.group(1)
                token_value = token_values.get(token_name, "")
                shape_profile = self._capacity_profile_for_shape(layout_key, shape, layout_profile)
                if isinstance(token_value, list):
                    self._set_bullets(shape, token_value, shape_profile)
                else:
                    self._set_text(shape, str(token_value), shape_profile)
                continue

            replaced_text = original_text
            for token_name in matches:
                token_value = token_values.get(token_name, "")
                if isinstance(token_value, list):
                    token_value = "\n".join(token_value)
                replaced_text = re.sub(r"{{\s*" + re.escape(token_name) + r"\s*}}", str(token_value), replaced_text)
            shape_profile = self._capacity_profile_for_shape(layout_key, shape, layout_profile)
            self._set_text(shape, replaced_text, shape_profile)

    def _build_token_value_map(self, slide_spec: SlideSpec, presentation_title: str) -> dict[str, str | list[str]]:
        token_map: dict[str, str | list[str]] = {
            "title": slide_spec.title or "",
            "subtitle": slide_spec.subtitle or "",
            "text": slide_spec.text or "",
            "body": slide_spec.text or "",
            "summary": slide_spec.text or "",
            "notes": slide_spec.notes or "",
            "footer": slide_spec.notes or presentation_title,
            "presentation_title": presentation_title,
            "presentation_name": presentation_title,
            "cover_title": presentation_title if not slide_spec.title else slide_spec.title,
            "cover_meta": slide_spec.notes or presentation_title,
            "left_note": slide_spec.left_bullets[0] if slide_spec.left_bullets else "",
            "right_note": slide_spec.right_bullets[0] if slide_spec.right_bullets else "",
            "main_text": slide_spec.text or "",
            "secondary_text": slide_spec.notes or "",
            "left_text": "\n".join(slide_spec.left_bullets) if slide_spec.left_bullets else slide_spec.text or "",
            "right_list": slide_spec.right_bullets or slide_spec.bullets,
            "contact_title": slide_spec.title or "",
            "contact_name_or_title": slide_spec.title or "",
            "contact_role": slide_spec.subtitle or "",
            "contact_phone": slide_spec.left_bullets[0] if slide_spec.left_bullets else "",
            "contact_email": slide_spec.right_bullets[0] if slide_spec.right_bullets else "",
            "address": slide_spec.text or "",
            "phone": slide_spec.left_bullets[0] if slide_spec.left_bullets else "",
            "email": slide_spec.right_bullets[0] if slide_spec.right_bullets else "",
            "website": slide_spec.notes or "",
            "bullets": slide_spec.bullets,
            "left_bullets": slide_spec.left_bullets,
            "right_bullets": slide_spec.right_bullets,
        }

        for index in range(1, 11):
            token_map[f"bullet_{index}"] = slide_spec.bullets[index - 1] if len(slide_spec.bullets) >= index else ""
            token_map[f"left_bullet_{index}"] = slide_spec.left_bullets[index - 1] if len(slide_spec.left_bullets) >= index else ""
            token_map[f"right_bullet_{index}"] = slide_spec.right_bullets[index - 1] if len(slide_spec.right_bullets) >= index else ""
            token_map[f"card_{index}"] = slide_spec.bullets[index - 1] if len(slide_spec.bullets) >= index else ""
            token_map[f"icon_{index}"] = ""

        if slide_spec.table is not None:
            token_map["table"] = [" | ".join(row) for row in slide_spec.table.rows]
            token_map["table_headers"] = slide_spec.table.headers
            token_map["table_rows"] = [" | ".join(row) for row in slide_spec.table.rows]
        if slide_spec.image_base64:
            token_map["image"] = slide_spec.image_base64

        return token_map

    def _resolve_layout(self, manifest: TemplateManifest, slide_spec: SlideSpec) -> LayoutSpec:
        if slide_spec.preferred_layout_key:
            for layout in manifest.layouts:
                if layout.key == slide_spec.preferred_layout_key:
                    return layout

        for layout in manifest.layouts:
            if slide_spec.kind.value in layout.supported_slide_kinds:
                return layout

        if manifest.default_layout_key:
            for layout in manifest.layouts:
                if layout.key == manifest.default_layout_key:
                    return layout

        if not manifest.layouts:
            raise ValueError(f"Template '{manifest.template_id}' does not contain layouts")
        return manifest.layouts[0]

    def _fill_slide_from_layout(self, slide, slide_spec: SlideSpec, layout: LayoutSpec, presentation_title: str) -> None:
        if layout.key == "cover":
            self._populate_cover_slide(slide, slide_spec)
            return
        layout_key = slide_spec.preferred_layout_key or layout.key
        layout_profile = profile_for_layout(layout_key)
        placeholders = {placeholder.placeholder_format.idx: placeholder for placeholder in slide.placeholders}
        used_placeholder_indices: set[int] = set()

        for placeholder_spec in layout.placeholders:
            if placeholder_spec.idx is None or placeholder_spec.idx not in placeholders:
                continue
            shape = placeholders[placeholder_spec.idx]
            used_placeholder_indices.add(placeholder_spec.idx)
            self._apply_shape_spec_metadata(shape, placeholder_spec)
            shape_profile = self._capacity_profile_for_shape(layout_key, shape, layout_profile)
            if placeholder_spec.binding:
                self._fill_shape_by_binding(shape, placeholder_spec.binding, slide_spec, presentation_title, shape_profile)
                if placeholder_spec.binding != "table":
                    self._apply_shape_spec_metadata(shape, placeholder_spec)
                continue
            if placeholder_spec.kind == PlaceholderKind.TITLE:
                self._set_text(shape, slide_spec.title or "", shape_profile)
            elif placeholder_spec.kind == PlaceholderKind.SUBTITLE:
                self._set_text(shape, slide_spec.subtitle or "", shape_profile)
            elif placeholder_spec.kind == PlaceholderKind.BODY:
                self._fill_body(shape, slide_spec, shape_profile)
            elif placeholder_spec.kind == PlaceholderKind.FOOTER:
                self._set_text(shape, slide_spec.notes or "", shape_profile)
            elif placeholder_spec.kind == PlaceholderKind.TABLE:
                self._fill_table_or_chart(shape, slide_spec)
            elif placeholder_spec.kind == PlaceholderKind.CHART:
                self._fill_chart(shape, slide_spec)
            self._apply_shape_spec_metadata(shape, placeholder_spec)

        for placeholder in slide.placeholders:
            placeholder_idx = placeholder.placeholder_format.idx
            if placeholder_idx in used_placeholder_indices:
                continue
            self._clear_placeholder(placeholder)

        self._apply_layout_expansion_and_flow(slide, slide_spec.preferred_layout_key or layout.key, slide_spec)

    def _capacity_profile_for_shape(
        self,
        layout_key: str,
        shape,
        base_profile: LayoutCapacityProfile,
    ) -> LayoutCapacityProfile:
        width = getattr(shape, "width", None)
        height = getattr(shape, "height", None)
        if width is None or height is None:
            return base_profile
        if layout_key in self.BUILTIN_LAYOUT_KEYS:
            return base_profile
        reference_width = None
        reference_height = None
        placeholder_format = getattr(shape, "placeholder_format", None)
        placeholder_idx = getattr(placeholder_format, "idx", None)
        base_geometry = geometry_policy_for_layout(layout_key)
        reference_placeholder = base_geometry.placeholders.get(placeholder_idx or -1)
        if reference_placeholder is None and placeholder_idx in {None, 14}:
            reference_placeholder = base_geometry.placeholders.get(14)
        if reference_placeholder is not None:
            reference_width = reference_placeholder.width_emu
            reference_height = reference_placeholder.height_emu
            if (
                abs(width - reference_width) <= self.GEOMETRY_PROFILE_TOLERANCE_EMU
                and abs(height - reference_height) <= self.GEOMETRY_PROFILE_TOLERANCE_EMU
            ):
                return base_profile
        return derive_capacity_profile_for_geometry(
            layout_key,
            width_emu=width,
            height_emu=height,
            reference_width_emu=reference_width,
            reference_height_emu=reference_height,
        )

    def _apply_layout_expansion_and_flow(self, slide, layout_key: str, slide_spec: SlideSpec | None = None) -> None:
        geometry_layout_key = "text_full_width" if layout_key == "dense_text_full_width" else layout_key
        if geometry_layout_key in {"text_full_width", "list_full_width"}:
            self._expand_text_full_width_layout(slide)
        elif geometry_layout_key == "image_text":
            self._expand_image_text_layout(slide)
        elif geometry_layout_key == "table":
            self._expand_table_layout(slide)
        elif geometry_layout_key == "cards_3":
            self._expand_cards_layout(slide)
        elif geometry_layout_key == "list_with_icons":
            self._expand_list_with_icons_layout(slide)
        elif geometry_layout_key == "contacts":
            self._expand_contacts_layout(slide)
        else:
            return

        self._adjust_title_and_flow(slide, layout_key, slide_spec)

    def _apply_shape_spec_metadata(self, shape, spec) -> None:
        geometry_values = (spec.left_emu, spec.top_emu, spec.width_emu, spec.height_emu)
        if all(isinstance(value, int) and value > 0 for value in geometry_values):
            shape.left = spec.left_emu
            shape.top = spec.top_emu
            shape.width = spec.width_emu
            shape.height = spec.height_emu

        if not getattr(shape, "has_text_frame", False):
            return

        text_frame = shape.text_frame
        if isinstance(spec.margin_left_emu, int) and spec.margin_left_emu >= 0:
            text_frame.margin_left = spec.margin_left_emu
        if isinstance(spec.margin_right_emu, int) and spec.margin_right_emu >= 0:
            text_frame.margin_right = spec.margin_right_emu
        if isinstance(spec.margin_top_emu, int) and spec.margin_top_emu >= 0:
            text_frame.margin_top = spec.margin_top_emu
        if isinstance(spec.margin_bottom_emu, int) and spec.margin_bottom_emu >= 0:
            text_frame.margin_bottom = spec.margin_bottom_emu

    def _populate_cover_slide(self, slide, slide_spec: SlideSpec) -> None:
        title_shape = self._find_cover_title_shape(slide)
        meta_shape = self._find_cover_meta_shape(slide)
        keep_shape_ids = {
            shape.shape_id
            for shape in (title_shape, meta_shape)
            if shape is not None
        }

        if title_shape is not None:
            title_shape.left = self.COVER_TITLE_LEFT_EMU
            title_shape.top = self.COVER_TITLE_TOP_EMU
            title_shape.width = self.COVER_TITLE_WIDTH_EMU
            title_shape.height = self.COVER_TITLE_MIN_HEIGHT_EMU
            self._set_cover_text(
                title_shape,
                slide_spec.title or "",
                font_size=Pt(46),
                bold=True,
                color=RGBColor(0xF5, 0xF9, 0xFE),
                align=PP_ALIGN.LEFT,
            )

        if meta_shape is not None:
            meta_text = (slide_spec.notes or "").strip()
            if meta_text:
                meta_shape.left = self.COVER_META_LEFT_EMU
                meta_shape.top = self.COVER_META_DEFAULT_TOP_EMU
                meta_shape.width = self.COVER_META_WIDTH_EMU
                meta_shape.height = 1400000
                self._set_cover_text(
                    meta_shape,
                    meta_text,
                    font_size=Pt(22),
                    bold=False,
                    color=RGBColor(0xF5, 0xF9, 0xFE),
                    align=PP_ALIGN.LEFT,
                )
            else:
                self._clear_placeholder(meta_shape)

        for shape in slide.shapes:
            if shape.shape_id in keep_shape_ids:
                continue
            if getattr(shape, "is_placeholder", False):
                self._clear_placeholder(shape)

        if title_shape is None:
            title_shape = slide.shapes.add_textbox(
                self.COVER_TITLE_LEFT_EMU,
                self.COVER_TITLE_TOP_EMU,
                self.COVER_TITLE_WIDTH_EMU,
                self.COVER_TITLE_MIN_HEIGHT_EMU,
            )
            self._set_cover_text(
                title_shape,
                slide_spec.title or "",
                font_size=Pt(46),
                bold=True,
                color=RGBColor(0xF5, 0xF9, 0xFE),
                align=PP_ALIGN.LEFT,
            )

        if meta_shape is None and (slide_spec.notes or "").strip():
            meta_shape = slide.shapes.add_textbox(442913, 6120605, 3371850, 277813)
            meta_shape.left = self.COVER_META_LEFT_EMU
            meta_shape.top = self.COVER_META_DEFAULT_TOP_EMU
            meta_shape.width = self.COVER_META_WIDTH_EMU
            meta_shape.height = 1400000
            self._set_cover_text(
                meta_shape,
                (slide_spec.notes or "").strip(),
                font_size=Pt(22),
                bold=False,
                color=RGBColor(0xF5, 0xF9, 0xFE),
                align=PP_ALIGN.LEFT,
            )

        self._adjust_cover_layout(title_shape, meta_shape)

    def _find_cover_title_shape(self, slide):
        candidates = [
            shape
            for shape in slide.shapes
            if getattr(shape, "has_text_frame", False)
            and shape.top < 2500000
            and shape.width > 5000000
        ]
        if candidates:
            return min(candidates, key=lambda shape: (shape.top, shape.left))
        return None

    def _find_cover_meta_shape(self, slide):
        candidates = [
            shape
            for shape in slide.shapes
            if getattr(shape, "has_text_frame", False)
            and shape.top > 5000000
            and shape.width < 5000000
        ]
        if candidates:
            return min(candidates, key=lambda shape: (shape.top, shape.left))
        return None

    def _set_cover_text(self, shape, text: str, *, font_size, bold: bool, color: RGBColor, align) -> None:
        if not getattr(shape, "has_text_frame", False):
            return
        text_frame = shape.text_frame
        text_frame.clear()
        text_frame.word_wrap = True
        text_frame.auto_size = MSO_AUTO_SIZE.NONE
        self._apply_text_frame_margins(text_frame)
        lines = text.splitlines() or [text]
        for index, line in enumerate(lines):
            paragraph = text_frame.paragraphs[0] if index == 0 else text_frame.add_paragraph()
            paragraph.alignment = align
            self._apply_paragraph_spacing(paragraph, "cover", "text_full_width")
            run = paragraph.add_run()
            run.text = line
            run.font.size = font_size
            run.font.bold = bold
            run.font.color.rgb = color

    def _adjust_cover_layout(self, title_shape, meta_shape) -> None:
        if title_shape is None or not getattr(title_shape, "has_text_frame", False):
            return

        title_text = (getattr(title_shape, "text", "") or "").strip()
        if not title_text:
            return

        title_font_size_pt = self._fit_cover_title_font_size_points(title_text, title_shape.width)
        self._apply_font_size(title_shape, title_font_size_pt)
        self._configure_title_text_frame(title_shape)
        required_title_height = self._estimate_title_height_emu(title_shape, title_text, title_font_size_pt)
        title_shape.height = max(self.COVER_TITLE_MIN_HEIGHT_EMU, required_title_height)

        if meta_shape is None or not getattr(meta_shape, "has_text_frame", False):
            return

        meta_text = (getattr(meta_shape, "text", "") or "").strip()
        if not meta_text:
            return

        meta_shape.left = self.COVER_META_LEFT_EMU
        meta_shape.width = self.COVER_META_WIDTH_EMU
        desired_meta_top = title_shape.top + title_shape.height + self.COVER_META_GAP_EMU
        meta_shape.top = max(self.COVER_META_DEFAULT_TOP_EMU, desired_meta_top)
        meta_required_height = self._estimate_text_height_emu(meta_text, meta_shape.width, 22.0)
        meta_shape.height = max(self.COVER_META_MIN_HEIGHT_EMU, meta_required_height)

        available_meta_height = self.COVER_BOTTOM_LIMIT_EMU - meta_shape.top
        if available_meta_height < self.COVER_META_MIN_HEIGHT_EMU:
            # If the title becomes too tall, tighten the title first before collapsing the meta block.
            title_font_size_pt = self._fit_cover_title_font_size_points(title_text, title_shape.width, max_height_emu=1900000)
            self._apply_font_size(title_shape, title_font_size_pt)
            required_title_height = self._estimate_title_height_emu(title_shape, title_text, title_font_size_pt)
            title_shape.height = max(self.COVER_TITLE_MIN_HEIGHT_EMU, required_title_height)
            meta_shape.top = max(self.COVER_META_DEFAULT_TOP_EMU, title_shape.top + title_shape.height + self.COVER_META_GAP_EMU)
            available_meta_height = self.COVER_BOTTOM_LIMIT_EMU - meta_shape.top

        meta_shape.height = max(self.COVER_META_MIN_HEIGHT_EMU, min(meta_shape.height, available_meta_height))

    def _fit_cover_title_font_size_points(self, text: str, width_emu: int, max_height_emu: int = 2200000) -> float:
        for candidate in (46.0, 42.0, 38.0, 34.0, 32.0, 30.0, 28.0):
            estimated_height = self._estimate_text_height_emu(text, width_emu, candidate)
            if estimated_height <= max_height_emu:
                return candidate
        return 28.0

    def _fill_body(self, shape, slide_spec: SlideSpec, layout_profile: LayoutCapacityProfile) -> None:
        if slide_spec.content_blocks:
            self._set_content_blocks(shape, slide_spec.content_blocks, layout_profile)
            return
        if slide_spec.kind == SlideKind.BULLETS:
            if not slide_spec.bullets:
                self._clear_placeholder(shape)
                return
            self._set_bullets(shape, slide_spec.bullets, layout_profile)
            return
        if slide_spec.kind == SlideKind.TWO_COLUMN:
            merged = [*slide_spec.left_bullets, "", *slide_spec.right_bullets]
            if not any(item.strip() for item in merged):
                self._clear_placeholder(shape)
                return
            self._set_bullets(shape, merged, layout_profile)
            return
        if slide_spec.kind == SlideKind.TEXT:
            if not (slide_spec.text or "").strip():
                self._clear_placeholder(shape)
                return
            self._set_text(shape, slide_spec.text or "", layout_profile)
            return
        if slide_spec.kind == SlideKind.TITLE:
            if not (slide_spec.text or "").strip():
                self._clear_placeholder(shape)
                return
            self._set_text(shape, slide_spec.text or "", layout_profile)
            return
        if slide_spec.table is not None:
            rows = [" | ".join(row) for row in slide_spec.table.rows]
            if not rows:
                self._clear_placeholder(shape)
                return
            self._set_bullets(shape, rows, layout_profile)
            return
        if not (slide_spec.text or "").strip():
            self._clear_placeholder(shape)
            return
        self._set_text(shape, slide_spec.text or "", layout_profile)

    def _set_content_blocks(
        self,
        shape,
        content_blocks: list[SlideContentBlock],
        layout_profile: LayoutCapacityProfile,
    ) -> None:
        text_frame = shape.text_frame
        text_frame.clear()

        entries: list[tuple[str, bool, SlideContentBlockKind]] = []
        for block in content_blocks:
            if block.kind in {SlideContentBlockKind.PARAGRAPH, SlideContentBlockKind.CALLOUT, SlideContentBlockKind.QA_ITEM}:
                text = (block.text or "").strip()
                if text:
                    entries.append((text, False, block.kind))
                continue
            if block.kind == SlideContentBlockKind.BULLET_LIST:
                entries.extend((item.strip(), True, block.kind) for item in block.items if item and item.strip())

        if not entries:
            self._clear_placeholder(shape)
            return

        first = True
        for text, is_bullet, block_kind in entries:
            paragraph = text_frame.paragraphs[0] if first else text_frame.add_paragraph()
            paragraph.text = text
            if is_bullet:
                paragraph.level = 0
                self._apply_bullet_format(paragraph, layout_profile.layout_key)
                self._apply_paragraph_spacing(paragraph, "body", layout_profile.layout_key)
            else:
                self._apply_paragraph_spacing(paragraph, "body", layout_profile.layout_key)
            self._style_content_block_paragraph(paragraph, block_kind)
            first = False

        self._configure_body_text_frame(text_frame)
        self._apply_body_font_size(text_frame, [text for text, _, _ in entries], shape, layout_profile)

    def _style_content_block_paragraph(self, paragraph, block_kind: SlideContentBlockKind) -> None:
        if not paragraph.runs and paragraph.text:
            run = paragraph.add_run()
            run.text = paragraph.text
            paragraph.text = ""

        if block_kind == SlideContentBlockKind.QA_ITEM:
            paragraph.level = 0
            for run in paragraph.runs:
                run.font.bold = True
                run.font.italic = False
            return

        if block_kind == SlideContentBlockKind.CALLOUT:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.italic = True
            return

    def _fill_shape_by_binding(self, shape, binding: str, slide_spec: SlideSpec, presentation_title: str, layout_profile: LayoutCapacityProfile) -> None:
        placeholder_idx = None
        if getattr(shape, "is_placeholder", False):
            placeholder_format = getattr(shape, "placeholder_format", None)
            placeholder_idx = getattr(placeholder_format, "idx", None)

        first_block_text = ""
        if slide_spec.content_blocks:
            for block in slide_spec.content_blocks:
                candidate = (block.text or "").strip()
                if candidate:
                    first_block_text = candidate
                    break

        if slide_spec.content_blocks and placeholder_idx == 14 and binding in {"body", "main_text", "bullets"}:
            self._set_content_blocks(shape, slide_spec.content_blocks, layout_profile)
            return
        if slide_spec.content_blocks and placeholder_idx == 15 and binding in {"secondary_text", "notes"}:
            self._clear_placeholder(shape)
            return
        if slide_spec.content_blocks and placeholder_idx == 13 and binding == "subtitle":
            subtitle_text = (slide_spec.subtitle or "").strip()
            if not subtitle_text:
                self._clear_placeholder(shape)
                return
            if first_block_text and first_block_text.startswith(subtitle_text):
                self._clear_placeholder(shape)
                return

        binding_value = self._build_token_value_map(slide_spec, presentation_title).get(binding, "")
        if binding == "table":
            self._fill_table_or_chart(shape, slide_spec)
            return
        if binding in {"chart", "chart_image"}:
            self._fill_chart(shape, slide_spec)
            return
        if binding == "image":
            self._fill_image(shape, slide_spec)
            return
        if binding == "icon_grid":
            self._clear_placeholder(shape)
            return
        if self._is_empty_binding_value(binding_value) and binding not in {"presentation_name", "cover_title", "title"}:
            self._clear_placeholder(shape)
            return
        if not getattr(shape, "has_text_frame", False):
            return
        if isinstance(binding_value, list):
            self._set_bullets(shape, [str(item) for item in binding_value], layout_profile)
            return
        if slide_spec.kind == SlideKind.CHART and binding == "title":
            self._set_text(shape, str(binding_value), layout_profile)
            self._configure_title_text_frame(shape)
            self._apply_font_size(shape, self._title_font_size_points(slide_spec.preferred_layout_key or "table"))
            return
        if slide_spec.kind == SlideKind.CHART and binding == "subtitle":
            self._set_text(shape, str(binding_value), layout_profile)
            self._configure_subtitle_text_frame(shape)
            self._apply_font_size(shape, self._table_subtitle_font_size_points(str(binding_value)))
            return
        self._set_text(shape, str(binding_value), layout_profile)

    def _expand_text_full_width_layout(self, slide) -> None:
        geometry = geometry_policy_for_layout("text_full_width")
        placeholders = {
            shape.placeholder_format.idx: shape
            for shape in slide.placeholders
            if getattr(shape, "is_placeholder", False)
        }

        title = placeholders.get(0)
        subtitle = placeholders.get(13)
        main_text = placeholders.get(14)
        secondary_text = placeholders.get(15)

        if title is not None:
            policy = geometry.placeholders[0]
            title.left = policy.left_emu
            title.top = policy.top_emu
            title.width = policy.width_emu
            title.height = max(title.height or 0, policy.height_emu)

        if subtitle is not None:
            policy = geometry.placeholders[13]
            subtitle.left = policy.left_emu
            subtitle.top = policy.top_emu
            subtitle.width = policy.width_emu
            subtitle.height = policy.height_emu

        if main_text is not None:
            policy = geometry.placeholders[14]
            main_text.left = policy.left_emu
            main_text.top = policy.top_emu
            main_text.width = policy.width_emu
            main_text.height = 1700000 if secondary_text is not None and getattr(secondary_text, "text", "").strip() else policy.height_emu

        if secondary_text is not None and getattr(secondary_text, "text", "").strip():
            policy = geometry.placeholders[15]
            secondary_text.left = policy.left_emu
            secondary_text.top = policy.top_emu
            secondary_text.width = policy.width_emu
            secondary_text.height = policy.height_emu

        footer = placeholders.get(17)
        if footer is not None:
            policy = geometry.placeholders[17]
            footer.left = policy.left_emu
            footer.top = policy.top_emu
            footer.width = policy.width_emu
            footer.height = policy.height_emu

    def _expand_table_layout(self, slide) -> None:
        geometry = geometry_policy_for_layout("table")
        placeholders = {
            shape.placeholder_format.idx: shape
            for shape in slide.placeholders
            if getattr(shape, "is_placeholder", False)
        }

        title = placeholders.get(0)
        subtitle = placeholders.get(13)
        left_note = placeholders.get(11)
        right_note = placeholders.get(12)
        footer = placeholders.get(15)

        if title is not None:
            policy = geometry.placeholders[0]
            title.left = policy.left_emu
            title.top = policy.top_emu
            title.width = policy.width_emu
            title.height = max(title.height or 0, policy.height_emu)

        if footer is not None:
            policy = geometry.placeholders[15]
            footer.left = policy.left_emu
            footer.top = policy.top_emu
            footer.width = policy.width_emu
            footer.height = policy.height_emu

        if subtitle is None or not getattr(subtitle, "text", "").strip():
            return

        has_side_notes = any(
            shape is not None and getattr(shape, "text", "").strip()
            for shape in (left_note, right_note)
        )
        if has_side_notes:
            return

        policy = geometry.placeholders[13]
        subtitle.left = policy.left_emu
        subtitle.top = policy.top_emu
        subtitle.width = policy.width_emu
        subtitle.height = policy.height_emu

    def _expand_image_text_layout(self, slide) -> None:
        geometry = geometry_policy_for_layout("image_text")
        placeholders = {
            shape.placeholder_format.idx: shape
            for shape in slide.placeholders
            if getattr(shape, "is_placeholder", False)
        }
        for idx, policy in geometry.placeholders.items():
            shape = placeholders.get(idx)
            if shape is None:
                continue
            shape.left = policy.left_emu
            shape.top = policy.top_emu
            shape.width = policy.width_emu
            shape.height = policy.height_emu

    def _expand_cards_layout(self, slide) -> None:
        geometry = geometry_policy_for_layout("cards_3")
        placeholders = {
            shape.placeholder_format.idx: shape
            for shape in slide.placeholders
            if getattr(shape, "is_placeholder", False)
        }
        for idx, policy in geometry.placeholders.items():
            shape = placeholders.get(idx)
            if shape is None:
                continue
            shape.left = policy.left_emu
            shape.top = policy.top_emu
            shape.width = policy.width_emu
            shape.height = policy.height_emu

    def _expand_list_with_icons_layout(self, slide) -> None:
        geometry = geometry_policy_for_layout("list_with_icons")
        placeholders = {
            shape.placeholder_format.idx: shape
            for shape in slide.placeholders
            if getattr(shape, "is_placeholder", False)
        }
        for idx, policy in geometry.placeholders.items():
            shape = placeholders.get(idx)
            if shape is None:
                continue
            shape.left = policy.left_emu
            shape.top = policy.top_emu
            shape.width = policy.width_emu
            shape.height = policy.height_emu

    def _expand_contacts_layout(self, slide) -> None:
        geometry = geometry_policy_for_layout("contacts")
        placeholders = {
            shape.placeholder_format.idx: shape
            for shape in slide.placeholders
            if getattr(shape, "is_placeholder", False)
        }
        for idx, policy in geometry.placeholders.items():
            shape = placeholders.get(idx)
            if shape is None:
                continue
            shape.left = policy.left_emu
            shape.top = policy.top_emu
            shape.width = policy.width_emu
            shape.height = policy.height_emu

    def _expand_cards_layout(self, slide) -> None:
        geometry = geometry_policy_for_layout("cards_3")
        placeholders = {
            shape.placeholder_format.idx: shape
            for shape in slide.placeholders
            if getattr(shape, "is_placeholder", False)
        }
        for idx, policy in geometry.placeholders.items():
            shape = placeholders.get(idx)
            if shape is None:
                continue
            shape.left = policy.left_emu
            shape.top = policy.top_emu
            shape.width = policy.width_emu
            shape.height = policy.height_emu

    def _expand_list_with_icons_layout(self, slide) -> None:
        geometry = geometry_policy_for_layout("list_with_icons")
        placeholders = {
            shape.placeholder_format.idx: shape
            for shape in slide.placeholders
            if getattr(shape, "is_placeholder", False)
        }
        for idx, policy in geometry.placeholders.items():
            shape = placeholders.get(idx)
            if shape is None:
                continue
            shape.left = policy.left_emu
            shape.top = policy.top_emu
            shape.width = policy.width_emu
            shape.height = policy.height_emu

    def _expand_contacts_layout(self, slide) -> None:
        geometry = geometry_policy_for_layout("contacts")
        placeholders = {
            shape.placeholder_format.idx: shape
            for shape in slide.placeholders
            if getattr(shape, "is_placeholder", False)
        }
        for idx, policy in geometry.placeholders.items():
            shape = placeholders.get(idx)
            if shape is None:
                continue
            shape.left = policy.left_emu
            shape.top = policy.top_emu
            shape.width = policy.width_emu
            shape.height = policy.height_emu

    def _adjust_title_and_flow(self, slide, layout_key: str, slide_spec: SlideSpec | None = None) -> None:
        if layout_key in {"text_full_width", "dense_text_full_width", "list_full_width"}:
            self._stack_text_content(slide, layout_key)
            return
        if layout_key == "image_text":
            self._stack_image_text_content(slide, layout_key)
            return
        if layout_key == "table":
            self._stack_table_content(slide, layout_key, slide_spec)
            return
        if layout_key == "cards_3":
            self._stack_cards_content(slide, layout_key)
            return
        if layout_key == "list_with_icons":
            self._stack_two_column_content(slide, layout_key)
            return
        if layout_key == "contacts":
            self._stack_contacts_content(slide, layout_key)
            return
        placeholders = {
            shape.placeholder_format.idx: shape
            for shape in slide.placeholders
            if getattr(shape, "is_placeholder", False)
        }
        title_shape = placeholders.get(0)
        if title_shape is None or not getattr(title_shape, "has_text_frame", False):
            return

        title_text = (getattr(title_shape, "text", "") or "").strip()
        if not title_text:
            return

        font_size_pt = self._fit_title_font_size_points(title_text, title_shape.width, layout_key)
        self._apply_font_size(title_shape, font_size_pt)
        required_height = self._estimate_text_height_emu(title_text, title_shape.width, font_size_pt)
        base_height = title_shape.height or 0
        final_title_height = max(base_height, required_height)
        title_shape.height = final_title_height

        protected_indices = {0, 17 if layout_key in {"text_full_width", "list_full_width"} else -1}
        flow_shapes = [
            shape
            for placeholder_idx, shape in placeholders.items()
            if placeholder_idx not in protected_indices and shape.top > title_shape.top
        ]
        if not flow_shapes:
            return

        current_flow_top = min(shape.top for shape in flow_shapes)
        desired_flow_top = title_shape.top + final_title_height + self.TITLE_CONTENT_GAP_EMU
        delta = max(0, desired_flow_top - current_flow_top)

        for shape in flow_shapes:
            shape.top += delta

    def _stack_text_content(self, slide, layout_key: str) -> None:
        geometry = geometry_policy_for_layout(layout_key)
        placeholders = {
            shape.placeholder_format.idx: shape
            for shape in slide.placeholders
            if getattr(shape, "is_placeholder", False)
        }
        title = placeholders.get(0)
        body = placeholders.get(14)
        footer = placeholders.get(17)
        subtitle = placeholders.get(13)
        secondary = placeholders.get(15)

        if title is None or body is None or footer is None:
            return

        title_text = (getattr(title, "text", "") or "").strip()
        if title_text:
            font_size_pt = self._fit_title_font_size_points(title_text, title.width, layout_key)
            self._apply_font_size(title, font_size_pt)
            self._configure_title_text_frame(title)
            required_height = self._estimate_title_height_emu(title, title_text, font_size_pt)
            title.height = max(self._minimum_title_height_emu(layout_key), required_height)

        if subtitle is not None and self._subtitle_duplicates_body(subtitle, body):
            self._clear_placeholder(subtitle)
            subtitle = None

        has_subtitle = subtitle is not None and getattr(subtitle, "text", "").strip()
        subtitle_body_gap = geometry.title_content_gap_emu
        compact_subtitle = False
        subtitle_text = ""
        body_text = (getattr(body, "text", "") or "").strip()
        if has_subtitle and subtitle is not None:
            subtitle_text = subtitle.text.strip()
            compact_subtitle = self._should_compact_subtitle_stack(title_text, subtitle_text, body_text, layout_key=layout_key)
        title_gap = geometry.title_content_gap_emu if has_subtitle else geometry.title_body_gap_no_subtitle_emu
        if compact_subtitle:
            title_gap = min(title_gap, 120000)
            subtitle_body_gap = 100000
        cursor = title.top + title.height + title_gap

        if has_subtitle:
            subtitle_font_pt = 18.0
            self._configure_subtitle_text_frame(subtitle)
            self._apply_font_size(subtitle, subtitle_font_pt)
            subtitle.height = max(280000 if compact_subtitle else 360000, self._estimate_text_height_emu(subtitle_text, subtitle.width, subtitle_font_pt))
            subtitle.top = cursor
            cursor = subtitle.top + subtitle.height + subtitle_body_gap

        secondary_has_text = secondary is not None and getattr(secondary, "text", "").strip()
        available_bottom = footer.top - geometry.content_footer_gap_emu

        if secondary_has_text and secondary is not None:
            secondary_text = secondary.text.strip()
            secondary.height = max(secondary.height or 0, self._estimate_text_height_emu(secondary_text, secondary.width, 16.0))
            secondary.height = min(secondary.height, max(700000, available_bottom - cursor - 900000))
            body.top = cursor
            body.height = max(900000, secondary.top - geometry.title_content_gap_emu - body.top)
            secondary.top = body.top + body.height + geometry.title_content_gap_emu
            secondary.height = max(700000, min(secondary.height, available_bottom - secondary.top))
            return

        body.top = cursor
        body.height = max(900000, available_bottom - body.top)

    def _should_compact_subtitle_stack(self, title_text: str, subtitle_text: str, body_text: str, *, layout_key: str) -> bool:
        if layout_key not in {"text_full_width", "dense_text_full_width"}:
            return False
        if not subtitle_text or not body_text:
            return False
        if len(subtitle_text) > 72 or len(subtitle_text.split()) > 10:
            return False
        if len(body_text) > 900:
            return False
        if body_text.count("\n") >= 5:
            return False
        return True

    def _subtitle_duplicates_body(self, subtitle, body) -> bool:
        subtitle_text = (getattr(subtitle, "text", "") or "").strip()
        body_text = (getattr(body, "text", "") or "").strip()
        if not subtitle_text or not body_text:
            return False
        if body_text.startswith(subtitle_text):
            return True
        if len(subtitle_text) >= 24 and body_text.startswith(subtitle_text[:-1]):
            return True
        return False

    def _stack_table_content(self, slide, layout_key: str, slide_spec: SlideSpec | None = None) -> None:
        geometry = geometry_policy_for_layout(layout_key)
        placeholders = {
            shape.placeholder_format.idx: shape
            for shape in slide.placeholders
            if getattr(shape, "is_placeholder", False)
        }
        title = placeholders.get(0)
        subtitle = placeholders.get(13)
        table = placeholders.get(14)
        chart = next((shape for shape in slide.shapes if getattr(shape, "has_chart", False)), None)
        content = table or chart
        footer = placeholders.get(15)

        if title is None or content is None or footer is None:
            return

        title_text = (getattr(title, "text", "") or "").strip()
        if title_text:
            font_size_pt = self._fit_title_font_size_points(title_text, title.width, layout_key)
            self._apply_font_size(title, font_size_pt)
            self._configure_title_text_frame(title)
            required_height = self._estimate_title_height_emu(title, title_text, font_size_pt)
            title.height = max(self._minimum_title_height_emu(layout_key), required_height)

        cursor = title.top + title.height + geometry.title_content_gap_emu
        if subtitle is not None and getattr(subtitle, "text", "").strip():
            subtitle_text = subtitle.text.strip()
            subtitle_font_size = self._table_subtitle_font_size_points(subtitle_text)
            self._apply_font_size(subtitle, subtitle_font_size)
            self._configure_subtitle_text_frame(subtitle)
            subtitle.height = max(360000, self._estimate_text_height_emu(subtitle_text, subtitle.width, subtitle_font_size))
            subtitle.top = cursor
            cursor = subtitle.top + subtitle.height + geometry.title_content_gap_emu

        content.top = cursor
        content.height = max(900000, footer.top - geometry.content_footer_gap_emu - content.top)

    def _stack_image_text_content(self, slide, layout_key: str) -> None:
        geometry = geometry_policy_for_layout(layout_key)
        placeholders = {
            shape.placeholder_format.idx: shape
            for shape in slide.placeholders
            if getattr(shape, "is_placeholder", False)
        }
        title = placeholders.get(0)
        body = placeholders.get(14)
        footer = placeholders.get(17)
        subtitle = placeholders.get(13)
        secondary = placeholders.get(15)
        image = placeholders.get(16)

        if title is None or body is None or footer is None:
            return

        title_text = (getattr(title, "text", "") or "").strip()
        if title_text:
            font_size_pt = self._fit_title_font_size_points(title_text, title.width, layout_key)
            self._apply_font_size(title, font_size_pt)
            self._configure_title_text_frame(title)
            required_height = self._estimate_title_height_emu(title, title_text, font_size_pt)
            title.height = max(self._minimum_title_height_emu(layout_key), required_height)

        has_subtitle = subtitle is not None and getattr(subtitle, "text", "").strip()
        title_gap = geometry.title_content_gap_emu if has_subtitle else geometry.title_body_gap_no_subtitle_emu
        cursor = title.top + title.height + title_gap
        if has_subtitle:
            subtitle_text = subtitle.text.strip()
            self._configure_subtitle_text_frame(subtitle)
            self._apply_font_size(subtitle, 18.0)
            subtitle.height = max(360000, self._estimate_text_height_emu(subtitle_text, subtitle.width, 18.0))
            subtitle.top = cursor
            cursor = subtitle.top + subtitle.height + geometry.title_content_gap_emu

        available_bottom = footer.top - geometry.content_footer_gap_emu
        if image is not None:
            image.top = cursor
            image.height = max(1200000, available_bottom - image.top)

        secondary_has_text = secondary is not None and getattr(secondary, "text", "").strip()
        if secondary_has_text and secondary is not None:
            secondary_text = secondary.text.strip()
            secondary.height = max(secondary.height or 0, self._estimate_text_height_emu(secondary_text, secondary.width, 16.0))
            secondary.height = min(secondary.height, max(700000, available_bottom - cursor - 900000))
            body.top = cursor
            body.height = max(900000, secondary.top - geometry.title_content_gap_emu - body.top)
            secondary.top = body.top + body.height + geometry.title_content_gap_emu
            secondary.height = max(700000, min(secondary.height, available_bottom - secondary.top))
            return

        body.top = cursor
        body.height = max(900000, available_bottom - body.top)

    def _stack_cards_content(self, slide, layout_key: str) -> None:
        geometry = geometry_policy_for_layout(layout_key)
        placeholders = {
            shape.placeholder_format.idx: shape
            for shape in slide.placeholders
            if getattr(shape, "is_placeholder", False)
        }
        title = placeholders.get(0)
        footer = placeholders.get(15)
        cards = [placeholders.get(idx) for idx in (11, 12, 13)]
        cards = [card for card in cards if card is not None]
        if title is None or footer is None or not cards:
            return

        title_text = (getattr(title, "text", "") or "").strip()
        if title_text:
            font_size_pt = self._fit_title_font_size_points(title_text, title.width, layout_key)
            self._apply_font_size(title, font_size_pt)
            self._configure_title_text_frame(title)
            required_height = self._estimate_title_height_emu(title, title_text, font_size_pt)
            title.height = max(self._minimum_title_height_emu(layout_key), required_height)

        cursor = title.top + title.height + geometry.title_body_gap_no_subtitle_emu
        base_top = min(card.top for card in cards)
        delta = max(0, cursor - base_top)
        max_height = max(900000, footer.top - geometry.content_footer_gap_emu - (base_top + delta))
        for card in cards:
            card.top += delta
            card.height = min(card.height, max_height)

    def _stack_two_column_content(self, slide, layout_key: str) -> None:
        geometry = geometry_policy_for_layout(layout_key)
        placeholders = {
            shape.placeholder_format.idx: shape
            for shape in slide.placeholders
            if getattr(shape, "is_placeholder", False)
        }
        title = placeholders.get(0)
        subtitle = placeholders.get(13)
        left = placeholders.get(12)
        right = placeholders.get(14)
        footer = placeholders.get(21)
        if title is None or left is None or right is None or footer is None:
            return

        title_text = (getattr(title, "text", "") or "").strip()
        if title_text:
            font_size_pt = self._fit_title_font_size_points(title_text, title.width, layout_key)
            self._apply_font_size(title, font_size_pt)
            self._configure_title_text_frame(title)
            required_height = self._estimate_title_height_emu(title, title_text, font_size_pt)
            title.height = max(self._minimum_title_height_emu(layout_key), required_height)

        has_subtitle = subtitle is not None and getattr(subtitle, "text", "").strip()
        title_gap = geometry.title_content_gap_emu if has_subtitle else geometry.title_body_gap_no_subtitle_emu
        cursor = title.top + title.height + title_gap
        if has_subtitle:
            subtitle_text = subtitle.text.strip()
            self._configure_subtitle_text_frame(subtitle)
            self._apply_font_size(subtitle, 18.0)
            subtitle.height = max(360000, self._estimate_text_height_emu(subtitle_text, subtitle.width, 18.0))
            subtitle.top = cursor
            cursor = subtitle.top + subtitle.height + geometry.title_content_gap_emu

        content_indices = [12, 14, 15, 16, 17, 18, 19, 20]
        content_shapes = [placeholders[idx] for idx in content_indices if idx in placeholders]
        if not content_shapes:
            return
        base_top = min(shape.top for shape in content_shapes)
        delta = max(0, cursor - base_top)
        max_height = max(900000, footer.top - geometry.content_footer_gap_emu - (base_top + delta))
        for shape in content_shapes:
            shape.top += delta
            if shape.placeholder_format.idx in {12, 14}:
                shape.height = min(shape.height, max_height)

    def _stack_contacts_content(self, slide, layout_key: str) -> None:
        geometry = geometry_policy_for_layout(layout_key)
        placeholders = {
            shape.placeholder_format.idx: shape
            for shape in slide.placeholders
            if getattr(shape, "is_placeholder", False)
        }
        for idx in (10, 11, 12, 13):
            shape = placeholders.get(idx)
            policy = geometry.placeholders.get(idx)
            if shape is None or policy is None:
                continue
            text = (getattr(shape, "text", "") or "").strip()
            if not text:
                continue
            font_size_pt = 18.0 if idx == 10 else 14.0
            if len(text) >= (60 if idx == 10 else 40):
                font_size_pt -= 2.0
            self._apply_font_size(shape, font_size_pt)
            self._configure_subtitle_text_frame(shape)
            shape.left = policy.left_emu
            shape.top = policy.top_emu
            shape.width = policy.width_emu
            shape.height = policy.height_emu

    def _title_font_size_points(self, layout_key: str) -> float:
        return 28.0

    def _table_subtitle_font_size_points(self, text: str) -> float:
        return 18.0

    def _configure_subtitle_text_frame(self, shape) -> None:
        if not getattr(shape, "has_text_frame", False):
            return
        text_frame = shape.text_frame
        text_frame.word_wrap = True
        text_frame.auto_size = MSO_AUTO_SIZE.NONE
        self._apply_text_frame_margins(text_frame)
        for paragraph in text_frame.paragraphs:
            self._apply_paragraph_spacing(paragraph, "subtitle", "text_full_width")

    def _minimum_title_height_emu(self, layout_key: str) -> int:
        if layout_key == "table":
            return 500000
        return 520000

    def _fit_title_font_size_points(self, text: str, width_emu: int, layout_key: str) -> float:
        base_size = self._title_font_size_points(layout_key)
        for candidate in (base_size, base_size - 2, base_size - 4, base_size - 6, base_size - 8):
            font_size = max(candidate, 22.0)
            estimated_height = self._estimate_text_height_emu(text, width_emu, font_size)
            if estimated_height <= 1650000:
                return font_size
        return 22.0

    def _apply_font_size(self, shape, font_size_pt: float) -> None:
        if not getattr(shape, "has_text_frame", False):
            return
        text_frame = shape.text_frame
        text_frame.word_wrap = True
        for paragraph in text_frame.paragraphs:
            if not paragraph.runs and paragraph.text:
                run = paragraph.add_run()
                run.text = paragraph.text
                paragraph.text = ""
            for run in paragraph.runs:
                run.font.size = Pt(font_size_pt)

    def _configure_title_text_frame(self, shape) -> None:
        if not getattr(shape, "has_text_frame", False):
            return
        text_frame = shape.text_frame
        text_frame.word_wrap = True
        text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        self._apply_text_frame_margins(text_frame)
        for paragraph in text_frame.paragraphs:
            self._apply_paragraph_spacing(paragraph, "title", "text_full_width")

    def _estimate_title_height_emu(self, shape, text: str, font_size_pt: float) -> int:
        text_frame = shape.text_frame
        effective_width = max(shape.width - text_frame.margin_left - text_frame.margin_right, shape.width // 2)
        width_pt = effective_width / self.EMU_PER_PT
        average_char_width_pt = max(font_size_pt * 0.52, 1.0)
        wrapped_lines = self._estimate_wrapped_line_count(text, width_pt, average_char_width_pt, min_chars_per_line=6)

        if len(text) >= 90:
            wrapped_lines = max(wrapped_lines, 2)
        if len(text) >= 135:
            wrapped_lines = max(wrapped_lines, 3)

        line_height_pt = font_size_pt * 1.24
        vertical_padding_pt = font_size_pt * 1.15
        return int((wrapped_lines * line_height_pt + vertical_padding_pt) * self.EMU_PER_PT)

    def _estimate_text_height_emu(self, text: str, width_emu: int, font_size_pt: float) -> int:
        if not text.strip() or width_emu <= 0:
            return 0

        width_pt = width_emu / self.EMU_PER_PT
        average_char_width_pt = max(font_size_pt * 0.52, 1.0)
        wrapped_lines = self._estimate_wrapped_line_count(text, width_pt, average_char_width_pt, min_chars_per_line=8)

        line_height_pt = font_size_pt * 1.18
        vertical_padding_pt = font_size_pt * 0.7
        return int((wrapped_lines * line_height_pt + vertical_padding_pt) * self.EMU_PER_PT)

    def _estimate_wrapped_line_count(
        self,
        text: str,
        width_pt: float,
        average_char_width_pt: float,
        *,
        min_chars_per_line: int,
    ) -> int:
        chars_per_line = max(int(width_pt / average_char_width_pt), min_chars_per_line)
        wrapped_lines = 0

        for paragraph in text.splitlines() or [text]:
            normalized = paragraph.strip()
            if not normalized:
                wrapped_lines += 1
                continue

            words = normalized.split()
            if len(words) <= 1:
                wrapped_lines += max(1, math.ceil(len(normalized) / chars_per_line))
                continue

            current_line_len = 0
            paragraph_lines = 1
            for word in words:
                word_len = len(word)
                projected = word_len if current_line_len == 0 else current_line_len + 1 + word_len
                if projected <= chars_per_line:
                    current_line_len = projected
                    continue
                if current_line_len == 0:
                    paragraph_lines += max(math.ceil(word_len / chars_per_line) - 1, 0)
                    current_line_len = word_len % chars_per_line or chars_per_line
                    continue
                paragraph_lines += 1
                current_line_len = word_len

            wrapped_lines += paragraph_lines

        return wrapped_lines

    def _fill_table(self, shape, slide_spec: SlideSpec) -> None:
        if slide_spec.table is None:
            if getattr(shape, "has_text_frame", False):
                self._set_text(shape, "", profile_for_layout("text_full_width"))
            return

        headers = slide_spec.table.headers
        rows = slide_spec.table.rows
        row_count = len(rows) + (1 if headers else 0)
        col_count = len(headers) if headers else max((len(row) for row in rows), default=0)
        if row_count == 0 or col_count == 0:
            if getattr(shape, "has_text_frame", False):
                self._set_text(shape, "", profile_for_layout("text_full_width"))
            return

        if hasattr(shape, "insert_table"):
            try:
                target_width = shape.width
                target_height = shape.height
                graphic_frame = shape.insert_table(row_count, col_count)
                graphic_frame.width = target_width
                graphic_frame.height = target_height
                table = graphic_frame.table
                current_row = 0
                if headers:
                    for col_index, value in enumerate(headers):
                        table.cell(0, col_index).text = value
                    current_row = 1
                for row_index, row in enumerate(rows, start=current_row):
                    for col_index, value in enumerate(row):
                        if col_index < col_count:
                            table.cell(row_index, col_index).text = value
                final_height = self._format_table(table, slide_spec.table, graphic_frame.width, graphic_frame.height)
                graphic_frame.height = final_height
                return
            except (AttributeError, TypeError, ValueError):
                pass

        if getattr(shape, "has_table", False):
            table = shape.table
            max_rows = len(table.rows)
            max_cols = len(table.columns)
            current_row = 0
            if headers and max_rows > 0:
                for col_index, value in enumerate(headers[:max_cols]):
                    table.cell(0, col_index).text = value
                current_row = 1
            for row_index, row in enumerate(rows, start=current_row):
                if row_index >= max_rows:
                    break
                for col_index, value in enumerate(row[:max_cols]):
                    table.cell(row_index, col_index).text = value
            final_height = self._format_table(table, slide_spec.table, shape.width, shape.height)
            shape.height = final_height
            return

        as_lines = []
        if headers:
            as_lines.append(" | ".join(headers))
        as_lines.extend(" | ".join(row) for row in rows)
        if getattr(shape, "has_text_frame", False):
            self._set_bullets(shape, as_lines, profile_for_layout("list_full_width"))

    def _fill_table_or_chart(self, shape, slide_spec: SlideSpec) -> None:
        if slide_spec.chart is not None:
            self._fill_chart(shape, slide_spec)
            return
        self._fill_table(shape, slide_spec)

    def _fill_chart(self, shape, slide_spec: SlideSpec) -> None:
        if slide_spec.chart is None:
            self._clear_placeholder(shape)
            return

        chart_spec = slide_spec.chart
        if not chart_spec.categories or not chart_spec.series:
            self._clear_placeholder(shape)
            return
        resolved_chart_spec = render_chart_spec(chart_spec)
        if resolved_chart_spec is None or not resolved_chart_spec.series:
            self._clear_placeholder(shape)
            return

        try:
            chart_type = self._resolve_chart_type(resolved_chart_spec)
            chart_data = CategoryChartData()
            chart_data.categories = resolved_chart_spec.categories
            for series in resolved_chart_spec.series:
                chart_data.add_series(series.name or "Ряд", series.values)

            slide_shapes = shape.part.slide.shapes
            left = shape.left
            top = shape.top
            width = shape.width
            height = shape.height
            if getattr(shape, "is_placeholder", False):
                self._clear_placeholder(shape)

            graphic_frame = slide_shapes.add_chart(chart_type, left, top, width, height, chart_data)
            chart = graphic_frame.chart
            if resolved_chart_spec.chart_type == ChartType.COMBO:
                self._convert_chart_to_combo(chart, resolved_chart_spec)
            self._style_chart(chart, resolved_chart_spec)
        except Exception:
            if getattr(shape, "has_text_frame", False):
                self._set_text(shape, "Не удалось построить график", profile_for_layout("text_full_width"))
            else:
                self._clear_placeholder(shape)
    def _convert_chart_to_combo(self, chart, chart_spec: ChartSpec) -> None:
        chart_space = chart._chartSpace
        plot_area = chart_space.chart.plotArea
        bar_charts = plot_area.xpath("./c:barChart")
        if not bar_charts:
            return

        bar_chart = bar_charts[0]
        series = bar_chart.xpath("./c:ser")
        if len(series) < 2:
            return

        line_series = series[-1]
        bar_chart.remove(line_series)

        line_chart = OxmlElement("c:lineChart")
        grouping = OxmlElement("c:grouping")
        grouping.set("val", "standard")
        line_chart.append(grouping)

        self._ensure_line_series_shape(line_series)
        line_chart.append(line_series)

        bar_axis_ids = [ax_id.get("val") for ax_id in bar_chart.xpath("./c:axId")]
        if uses_secondary_value_axis(chart_spec):
            secondary_axis_id = self._ensure_secondary_value_axis(chart)
            for axis_id in [bar_axis_ids[0], secondary_axis_id]:
                cloned_ax_id = OxmlElement("c:axId")
                cloned_ax_id.set("val", axis_id)
                line_chart.append(cloned_ax_id)
        else:
            for ax_id in bar_chart.xpath("./c:axId"):
                cloned_ax_id = OxmlElement("c:axId")
                cloned_ax_id.set("val", ax_id.get("val"))
                line_chart.append(cloned_ax_id)

        insert_at = list(plot_area).index(bar_chart) + 1
        plot_area.insert(insert_at, line_chart)

    def _ensure_secondary_value_axis(self, chart) -> str:
        chart_space = chart._chartSpace
        plot_area = chart_space.chart.plotArea
        value_axes = plot_area.xpath("./c:valAx")
        if len(value_axes) > 1:
            return value_axes[1].xpath("./c:axId")[0].get("val")

        category_axis = (plot_area.xpath("./c:catAx") or plot_area.xpath("./c:dateAx"))[0]
        primary_value_axis = value_axes[0]
        category_axis_id = category_axis.xpath("./c:axId")[0].get("val")
        secondary_axis_id = self._next_chart_axis_id(chart)
        secondary_axis = deepcopy(primary_value_axis)
        secondary_axis.xpath("./c:axId")[0].set("val", secondary_axis_id)
        secondary_axis.xpath("./c:axPos")[0].set("val", "r")
        secondary_axis.xpath("./c:crossAx")[0].set("val", category_axis_id)
        for gridlines in secondary_axis.xpath("./c:majorGridlines"):
            secondary_axis.remove(gridlines)
        plot_area.insert(list(plot_area).index(primary_value_axis) + 1, secondary_axis)
        return secondary_axis_id

    def _next_chart_axis_id(self, chart) -> str:
        axis_ids = [
            int(axis_id.get("val"))
            for axis_id in chart._chartSpace.xpath(".//c:axId")
            if axis_id.get("val") is not None
        ]
        next_id = max((abs(value) for value in axis_ids), default=1) + 1
        return str(next_id)

    def _ensure_line_series_shape(self, series_element) -> None:
        if not series_element.xpath("./c:marker"):
            marker = OxmlElement("c:marker")
            symbol = OxmlElement("c:symbol")
            symbol.set("val", "none")
            marker.append(symbol)
            insert_at = 0
            for idx, child in enumerate(series_element):
                if child.tag.endswith("}spPr"):
                    insert_at = idx + 1
                    break
            series_element.insert(insert_at, marker)

        if not series_element.xpath("./c:smooth"):
            smooth = OxmlElement("c:smooth")
            smooth.set("val", "0")
            series_element.append(smooth)

    def _resolve_chart_type(self, chart_spec: ChartSpec) -> XL_CHART_TYPE:
        chart_type_map = {
            ChartType.BAR: XL_CHART_TYPE.BAR_CLUSTERED,
            ChartType.COLUMN: XL_CHART_TYPE.COLUMN_CLUSTERED,
            ChartType.LINE: XL_CHART_TYPE.LINE,
            ChartType.STACKED_BAR: XL_CHART_TYPE.BAR_STACKED,
            ChartType.STACKED_COLUMN: XL_CHART_TYPE.COLUMN_STACKED,
            ChartType.COMBO: XL_CHART_TYPE.COLUMN_CLUSTERED,
            ChartType.PIE: XL_CHART_TYPE.PIE,
        }
        return chart_type_map.get(chart_spec.chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)

    def _style_chart(self, chart, chart_spec: ChartSpec) -> None:
        self._style_chart_plot(chart, chart_spec)
        chart.has_legend = chart_spec.legend_visible and len(chart.series) > 1
        if chart.has_legend:
            chart.legend.include_in_layout = False
            self._style_chart_legend(chart)

        if chart_spec.title:
            chart.has_title = True
            chart.chart_title.text_frame.text = chart_spec.title
            self._style_chart_title(chart)
        else:
            chart.has_title = False

        for index, series in enumerate(chart.series):
            self._style_chart_series(series, chart_spec, index)
            if chart_spec.data_labels_visible:
                series.has_data_labels = True
                self._style_data_labels(series, chart_spec)

        if hasattr(chart, "category_axis"):
            try:
                chart.category_axis.has_title = bool(chart_spec.x_axis_title)
                if chart_spec.x_axis_title:
                    chart.category_axis.axis_title.text_frame.text = chart_spec.x_axis_title
                    self._style_axis_title(chart.category_axis)
                self._style_category_axis(chart.category_axis, chart_spec)
            except Exception:
                pass

        for axis_index, axis in enumerate(self._value_axes(chart)):
            axis_role = PRIMARY_AXIS if axis_index == 0 else SECONDARY_AXIS
            try:
                if axis_role == PRIMARY_AXIS:
                    axis.has_title = bool(chart_spec.y_axis_title)
                    if chart_spec.y_axis_title:
                        axis.axis_title.text_frame.text = chart_spec.y_axis_title
                        self._style_axis_title(axis)
                else:
                    axis.has_title = False
                axis.has_major_gridlines = axis_role == PRIMARY_AXIS
                self._style_value_axis(axis, chart_spec, axis_role=axis_role)
            except Exception:
                pass

    def _style_chart_plot(self, chart, chart_spec: ChartSpec) -> None:
        if not chart.plots:
            return
        plot = chart.plots[0]
        dense_threshold = self._chart_style_config()["denseCategoryThreshold"]
        very_dense_threshold = self._chart_style_config()["veryDenseCategoryThreshold"]
        try:
            if chart_spec.chart_type in {ChartType.COLUMN, ChartType.BAR, ChartType.STACKED_BAR, ChartType.STACKED_COLUMN}:
                if len(chart_spec.categories) >= very_dense_threshold:
                    plot.gap_width = 124
                elif len(chart_spec.categories) >= dense_threshold:
                    plot.gap_width = 96
                else:
                    plot.gap_width = 72
                plot.overlap = 100 if chart_spec.chart_type in {ChartType.STACKED_BAR, ChartType.STACKED_COLUMN} else 0
        except Exception:
            pass

    def _style_chart_title(self, chart) -> None:
        try:
            text_frame = chart.chart_title.text_frame
            self._style_text_frame_runs(text_frame, font_size=Pt(18), bold=True, color=self._style_rgb("textColor"))
        except Exception:
            pass

    def _style_chart_legend(self, chart) -> None:
        try:
            chart.legend.font.size = Pt(11)
            chart.legend.font.color.rgb = self._style_rgb("mutedTextColor")
        except Exception:
            pass

    def _style_chart_series(self, series, chart_spec: ChartSpec, index: int) -> None:
        color = self._series_color(index)
        if chart_spec.chart_type == ChartType.PIE:
            self._style_pie_points(series)
            return

        is_combo_line_series = chart_spec.chart_type == ChartType.COMBO and index == len(chart_spec.series) - 1

        try:
            if not is_combo_line_series:
                series.format.fill.solid()
                series.format.fill.fore_color.rgb = color
        except Exception:
            pass

        try:
            series.format.line.color.rgb = color
            if chart_spec.chart_type == ChartType.LINE or is_combo_line_series:
                series.format.line.width = Pt(2.5)
                self._style_line_marker(series, color)
        except Exception:
            pass

    def _style_pie_points(self, series) -> None:
        for index, point in enumerate(getattr(series, "points", [])):
            color = self._series_color(index)
            try:
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = color
            except Exception:
                pass
            try:
                point.format.line.color.rgb = self._style_rgb("surfaceColor")
                point.format.line.width = Pt(1)
            except Exception:
                pass

    def _style_line_marker(self, series, color: RGBColor) -> None:
        try:
            series.marker.style = XL_MARKER_STYLE.CIRCLE
            series.marker.size = 7
            series.marker.format.fill.solid()
            series.marker.format.fill.fore_color.rgb = color
            series.marker.format.line.color.rgb = self._style_rgb("surfaceColor")
            series.marker.format.line.width = Pt(1)
        except Exception:
            pass

    def _style_data_labels(self, series, chart_spec: ChartSpec) -> None:
        try:
            data_labels = series.data_labels
            data_labels.font.size = Pt(10)
            data_labels.font.color.rgb = self._style_rgb("mutedTextColor")
            data_labels.number_format = self._chart_number_format(chart_spec)
        except Exception:
            pass

    def _style_category_axis(self, axis, chart_spec: ChartSpec) -> None:
        self._style_axis_line(axis)
        self._style_tick_labels(axis, chart_spec)

    def _style_value_axis(self, axis, chart_spec: ChartSpec, *, axis_role: str = PRIMARY_AXIS) -> None:
        self._style_axis_line(axis)
        self._style_tick_labels(axis, chart_spec)
        try:
            if axis_role == PRIMARY_AXIS and axis.has_major_gridlines:
                axis.major_gridlines.format.line.color.rgb = self._style_rgb("gridColor")
                axis.major_gridlines.format.line.width = Pt(0.8)
        except Exception:
            pass
        try:
            axis.tick_labels.number_format = chart_axis_number_format_for_axis(chart_spec, axis_role) or self._chart_axis_number_format(chart_spec)
        except Exception:
            pass

    def _value_axes(self, chart) -> list[ValueAxis]:
        return [ValueAxis(element) for element in chart._chartSpace.valAx_lst]

    def _style_axis_line(self, axis) -> None:
        try:
            axis.format.line.color.rgb = self._style_rgb("axisColor")
            axis.format.line.width = Pt(0.8)
        except Exception:
            pass

    def _style_tick_labels(self, axis, chart_spec: ChartSpec) -> None:
        tick_size = self._chart_tick_font_size(chart_spec)
        try:
            axis.tick_labels.font.size = Pt(tick_size)
            axis.tick_labels.font.color.rgb = self._style_rgb("mutedTextColor")
        except Exception:
            pass

    def _chart_number_format(self, chart_spec: ChartSpec) -> str:
        if chart_spec.value_format == "currency":
            return '#,##0" ₽"'
        if chart_spec.value_format == "percent":
            return '0"%"'
        return "General"

    def _chart_axis_number_format(self, chart_spec: ChartSpec) -> str:
        return chart_axis_number_format(chart_spec) or "#,##0"

    def _chart_tick_font_size(self, chart_spec: ChartSpec) -> int:
        config = self._chart_style_config()
        if len(chart_spec.categories) >= config["veryDenseCategoryThreshold"]:
            return config["tickFontSizes"]["veryDense"]
        if len(chart_spec.categories) >= config["denseCategoryThreshold"]:
            return config["tickFontSizes"]["dense"]
        return config["tickFontSizes"]["default"]

    def _style_axis_title(self, axis) -> None:
        try:
            self._style_text_frame_runs(axis.axis_title.text_frame, font_size=Pt(11), bold=False, color=self._style_rgb("textColor"))
        except Exception:
            pass

    def _style_text_frame_runs(self, text_frame, *, font_size, bold: bool, color: RGBColor) -> None:
        for paragraph in text_frame.paragraphs:
            if not paragraph.runs and paragraph.text:
                run = paragraph.add_run()
                run.text = paragraph.text
                paragraph.text = ""
            for run in paragraph.runs:
                run.font.size = font_size
                run.font.bold = bold
                run.font.color.rgb = color

    def _fill_image(self, shape, slide_spec: SlideSpec) -> None:
        if not slide_spec.image_base64:
            self._clear_placeholder(shape)
            return

        try:
            image_stream = BytesIO(base64.b64decode(slide_spec.image_base64))
        except Exception:
            self._clear_placeholder(shape)
            return

        if getattr(shape, "is_placeholder", False) and hasattr(shape, "insert_picture"):
            try:
                shape.insert_picture(image_stream)
                return
            except Exception:
                image_stream.seek(0)

        left = shape.left
        top = shape.top
        width = shape.width
        height = shape.height
        if getattr(shape, "is_placeholder", False):
            self._clear_placeholder(shape)

        try:
            slide_shapes = shape.part.slide.shapes
            slide_shapes.add_picture(image_stream, left, top, width=width, height=height)
        except Exception:
            if getattr(shape, "has_text_frame", False):
                self._set_text(shape, "Изображение из документа", profile_for_layout("text_full_width"))

    def _set_text(self, shape, text: str, layout_profile: LayoutCapacityProfile) -> None:
        text_frame = shape.text_frame
        text_frame.clear()
        text_frame.text = text
        if getattr(shape, "is_placeholder", False):
            placeholder_format = getattr(shape, "placeholder_format", None)
            if placeholder_format is not None and placeholder_format.idx in {15, 17}:
                self._apply_footer_font_size(text_frame, text)
                return
        self._configure_body_text_frame(text_frame)
        self._apply_body_font_size(text_frame, [text], shape, layout_profile)

    def _set_bullets(self, shape, items: list[str], layout_profile: LayoutCapacityProfile) -> None:
        text_frame = shape.text_frame
        text_frame.clear()
        if not items:
            return

        first = True
        for item in items:
            paragraph = text_frame.paragraphs[0] if first else text_frame.add_paragraph()
            paragraph.text = item
            if item:
                paragraph.level = 0
                self._apply_bullet_format(paragraph, layout_profile.layout_key)
                self._apply_paragraph_spacing(paragraph, "body", layout_profile.layout_key)
            first = False
        self._configure_body_text_frame(text_frame)
        self._apply_body_font_size(text_frame, items, shape, layout_profile)

    def _configure_body_text_frame(self, text_frame) -> None:
        text_frame.word_wrap = True
        text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        self._apply_text_frame_margins(text_frame)

    def _apply_paragraph_spacing(self, paragraph, role: str, layout_key: str) -> None:
        spacing = spacing_policy_for_layout(layout_key)
        role_policy = getattr(spacing, role)
        paragraph.line_spacing = role_policy.line_spacing
        paragraph.space_after = Pt(role_policy.space_after_pt)

    def _apply_text_frame_margins(self, text_frame) -> None:
        text_frame.margin_left = self.DEFAULT_TEXT_MARGIN_X_EMU
        text_frame.margin_right = self.DEFAULT_TEXT_MARGIN_X_EMU
        text_frame.margin_top = self.DEFAULT_TEXT_MARGIN_Y_EMU
        text_frame.margin_bottom = self.DEFAULT_TEXT_MARGIN_Y_EMU

    def _apply_body_font_size(self, text_frame, items: list[str], shape, layout_profile: LayoutCapacityProfile) -> None:
        non_empty_items = [item.strip() for item in items if item and item.strip()]
        if not non_empty_items:
            return

        total_chars = sum(len(item) for item in non_empty_items)
        max_item_len = max(len(item) for item in non_empty_items)
        item_count = len(non_empty_items)
        shape_height = getattr(shape, "height", 0) if shape is not None else 0
        shape_width = getattr(shape, "width", 0) if shape is not None else 0

        points = layout_profile.max_font_pt
        if total_chars >= layout_profile.max_chars * 2 or max_item_len >= 320:
            points = layout_profile.min_font_pt
        elif total_chars >= int(layout_profile.max_chars * 1.55) or max_item_len >= 240 or item_count >= layout_profile.max_items:
            points = max(layout_profile.max_font_pt - 3, layout_profile.min_font_pt)
        elif total_chars >= int(layout_profile.max_chars * 1.2) or max_item_len >= 180 or item_count >= max(layout_profile.max_items - 2, 1):
            points = max(layout_profile.max_font_pt - 2, layout_profile.min_font_pt)
        elif total_chars >= int(layout_profile.max_chars * 0.8) or max_item_len >= 120 or item_count >= max(layout_profile.max_items - 4, 1):
            points = max(layout_profile.max_font_pt - 1, layout_profile.min_font_pt)

        # Tight containers need one extra step down to avoid overflow on dense appendix-like slides.
        if shape_height and shape_height < 4000000 and (total_chars >= 900 or item_count >= 7):
            points = max(points - 1, layout_profile.min_font_pt)
        if shape_width and shape_width < 8000000 and total_chars >= 600:
            points = max(points - 1, layout_profile.min_font_pt)

        font_size = Pt(points)

        for paragraph in text_frame.paragraphs:
            if not paragraph.runs and paragraph.text:
                run = paragraph.add_run()
                run.text = paragraph.text
                paragraph.text = ""
            self._apply_paragraph_spacing(paragraph, "body", layout_profile.layout_key)
            paragraph.font.size = font_size
            for run in paragraph.runs:
                run.font.size = font_size

    def _apply_footer_font_size(self, text_frame, text: str) -> None:
        normalized = (text or "").strip()
        if not normalized:
            return

        font_size = None
        if len(normalized) >= 160:
            font_size = Pt(10)
        elif len(normalized) >= 120:
            font_size = Pt(11)
        elif len(normalized) >= 80:
            font_size = Pt(12)

        if font_size is None:
            return

        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = font_size

    def _apply_bullet_format(self, paragraph, layout_key: str) -> None:
        spacing = spacing_policy_for_layout(layout_key).bullet
        paragraph_properties = paragraph._p.get_or_add_pPr()
        for child in list(paragraph_properties):
            if child.tag.endswith("}buNone") or child.tag.endswith("}buChar") or child.tag.endswith("}buAutoNum"):
                paragraph_properties.remove(child)

        bullet = OxmlElement("a:buChar")
        bullet.set("char", "•")
        paragraph_properties.insert(0, bullet)
        paragraph_properties.set("marL", str(spacing.margin_left_emu))
        paragraph_properties.set("indent", str(spacing.indent_emu))

    def _format_table(self, table, table_block: TableBlock | None, target_width: int, target_height: int) -> int:
        if table_block is None:
            headers: list[str] = []
            rows: list[list[str]] = []
            header_fill_colors: list[str | None] = []
            row_fill_colors: list[list[str | None]] = []
        else:
            headers = table_block.headers
            rows = table_block.rows
            header_fill_colors = table_block.header_fill_colors
            row_fill_colors = table_block.row_fill_colors
        all_rows = [headers, *rows] if headers else rows
        column_stats = self._column_stats(headers, rows)
        max_lengths = [item["max_len"] for item in column_stats]
        row_count = len(all_rows)
        col_count = len(max_lengths)
        max_cell_length = max(max_lengths, default=0)
        avg_cell_length = (
            sum(len((headers[col] if headers and col < len(headers) else "")) for col in range(col_count))
            + sum(len(value or "") for row in rows for value in row[:col_count])
        ) / max(1, row_count * max(1, col_count))
        self._apply_table_geometry(table, column_stats, target_width, target_height, row_count, avg_cell_length=avg_cell_length)
        font_size = self._estimate_table_font_size(
            row_count=row_count,
            col_count=col_count,
            max_cell_length=max_cell_length,
            avg_cell_length=avg_cell_length,
        )
        margins = self._estimate_table_margins(
            row_count=row_count,
            col_count=col_count,
            max_cell_length=max_cell_length,
            avg_cell_length=avg_cell_length,
        )

        for row_index, row in enumerate(table.rows):
            is_header = headers and row_index == 0
            for col_index, cell in enumerate(row.cells):
                fill_color = (
                    header_fill_colors[col_index]
                    if is_header and col_index < len(header_fill_colors)
                    else (
                        row_fill_colors[row_index - 1][col_index]
                        if not is_header
                        and row_index - 1 < len(row_fill_colors)
                        and col_index < len(row_fill_colors[row_index - 1])
                        else None
                    )
                )
                self._style_table_cell(
                    cell,
                    is_header=is_header,
                    font_size=font_size,
                    margins=margins,
                    fill_color=fill_color,
                )
        return max(sum(row.height for row in table.rows), 600000)

    def _column_stats(self, headers: list[str], rows: list[list[str]]) -> list[dict[str, float]]:
        col_count = len(headers) if headers else max((len(row) for row in rows), default=0)
        stats: list[dict[str, float]] = []
        for col_index in range(col_count):
            values: list[str] = []
            if col_index < len(headers):
                values.append(headers[col_index])
            for row in rows:
                if col_index < len(row):
                    values.append(row[col_index])
            lengths = [len(value or "") for value in values]
            stats.append(
                {
                    "max_len": max(lengths, default=8),
                    "avg_len": (sum(lengths) / len(lengths)) if lengths else 8.0,
                    "header_len": len(headers[col_index]) if col_index < len(headers) else 0,
                }
            )
        return stats

    def _apply_table_geometry(
        self,
        table,
        column_stats: list[dict[str, float]],
        target_width: int,
        target_height: int,
        row_count: int,
        *,
        avg_cell_length: float,
    ) -> None:
        if column_stats and target_width > 0:
            weights = self._column_width_weights(column_stats)
            weight_sum = sum(weights) or len(weights)
            assigned = 0
            for index, column in enumerate(table.columns):
                if index == len(weights) - 1:
                    width = max(target_width - assigned, int(target_width * 0.08))
                else:
                    min_share = 0.14 if len(weights) >= 3 else 0.1
                    width = max(int(target_width * weights[index] / weight_sum), int(target_width * min_share))
                    assigned += width
                column.width = width

        if row_count > 0 and target_height > 0:
            row_height = self._table_row_height(target_height, row_count, avg_cell_length)
            for row in table.rows:
                row.height = row_height

    def _table_row_height(self, target_height: int, row_count: int, avg_cell_length: float) -> int:
        computed = max(int(target_height / row_count), 200000)
        if row_count <= 3:
            cap = 360000 if avg_cell_length < 45 else 420000
            return min(computed, cap)
        if row_count <= 5:
            cap = 340000 if avg_cell_length < 45 else 400000
            return min(computed, cap)
        if row_count <= 8:
            cap = 320000 if avg_cell_length < 45 else 360000
            return min(computed, cap)
        return computed

    def _column_width_weights(self, column_stats: list[dict[str, float]]) -> list[float]:
        col_count = len(column_stats)
        weights: list[float] = []
        for stat in column_stats:
            # Bias toward typical cell size, not one outlier.
            weight = (stat["header_len"] * 0.9) + (stat["avg_len"] * 1.35) + (min(stat["max_len"], stat["avg_len"] * 1.8) * 0.55)
            weights.append(max(weight, 8.0))

        if col_count == 2:
            first_weight = weights[0]
            second_weight = weights[1]
            total = first_weight + second_weight
            if total > 0:
                first_share = first_weight / total
                first_share = min(max(first_share, 0.22), 0.42)
                return [first_share, 1.0 - first_share]

        if col_count == 3:
            total = sum(weights) or 1.0
            shares = [weight / total for weight in weights]
            normalized = []
            for share in shares:
                normalized.append(min(max(share, 0.14), 0.58))
            scale = sum(normalized) or 1.0
            return [share / scale for share in normalized]

        return weights

    def _estimate_table_font_size(
        self,
        *,
        row_count: int,
        col_count: int,
        max_cell_length: int,
        avg_cell_length: float,
    ) -> Pt:
        points = 11
        if row_count >= 8:
            points = 10
        if row_count >= 12:
            points = 9
        if row_count >= 16:
            points = 8

        if col_count >= 3:
            points -= 1
        if max_cell_length >= 90 or avg_cell_length >= 45:
            points -= 1
        if row_count >= 10 and (max_cell_length >= 140 or avg_cell_length >= 60):
            points -= 1

        if row_count <= 4:
            points = max(points, 9)
        elif row_count <= 7:
            points = max(points, 8)
        else:
            points = max(points, 8)

        points = min(points, 11)
        return Pt(points)

    def _estimate_table_margins(
        self,
        *,
        row_count: int,
        col_count: int,
        max_cell_length: int,
        avg_cell_length: float,
    ) -> tuple[int, int, int, int]:
        if row_count >= 8 or col_count >= 3 or max_cell_length >= 90 or avg_cell_length >= 40:
            return (40000, 40000, 20000, 20000)
        if row_count >= 5 or max_cell_length >= 60:
            return (60000, 60000, 30000, 30000)
        return (80000, 80000, 40000, 40000)

    def _style_table_cell(
        self,
        cell,
        *,
        is_header: bool,
        font_size: Pt,
        margins: tuple[int, int, int, int],
        fill_color: str | None = None,
    ) -> None:
        fill = cell.fill
        fill.solid()
        resolved_fill = self._table_fill_rgb(fill_color, is_header=is_header)
        fill.fore_color.rgb = resolved_fill

        text_frame = cell.text_frame
        text_frame.word_wrap = True
        cell.margin_left, cell.margin_right, cell.margin_top, cell.margin_bottom = margins

        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = font_size
                run.font.bold = bool(is_header)
                run.font.color.rgb = self._table_text_rgb_for_fill(resolved_fill, is_header=is_header)

    def _table_fill_rgb(self, fill_color: str | None, *, is_header: bool) -> RGBColor:
        if fill_color:
            normalized = fill_color.strip().lstrip("#").upper()
            if re.fullmatch(r"[0-9A-F]{6}", normalized):
                return RGBColor.from_string(normalized)
        return RGBColor(0x08, 0x1C, 0x4F) if is_header else RGBColor(0xEB, 0xF3, 0xFE)

    def _table_text_rgb_for_fill(self, fill_rgb: RGBColor, *, is_header: bool) -> RGBColor:
        red, green, blue = fill_rgb[0], fill_rgb[1], fill_rgb[2]
        luminance = (0.2126 * red + 0.7152 * green + 0.0722 * blue) / 255
        if luminance < 0.45:
            return RGBColor(0xFF, 0xFF, 0xFF)
        if is_header:
            return RGBColor(0x08, 0x1C, 0x4F)
        return RGBColor(0x08, 0x1C, 0x4F)

    def _clear_placeholder(self, shape) -> None:
        if getattr(shape, "has_text_frame", False):
            self._remove_shape(shape)
            return
        if getattr(shape, "has_table", False):
            return
        self._remove_shape(shape)

    def _remove_shape(self, shape) -> None:
        parent = shape._element.getparent()
        if parent is not None:
            parent.remove(shape._element)

    def _is_empty_binding_value(self, value) -> bool:
        if isinstance(value, list):
            return not any(str(item).strip() for item in value)
        return not str(value or "").strip()

    def _validate_output_file(self, output_path: Path, expected_slide_count: int) -> None:
        try:
            with zipfile.ZipFile(output_path) as archive:
                bad_entry = archive.testzip()
                if bad_entry is not None:
                    raise ValueError(f"Generated PPTX contains a corrupted archive entry: {bad_entry}")

                archive_names = archive.namelist()
                duplicates = sorted({name for name in archive_names if archive_names.count(name) > 1})
                if duplicates:
                    raise ValueError(f"Generated PPTX contains duplicate package entries: {', '.join(duplicates[:5])}")

            presentation = Presentation(str(output_path))
            if len(presentation.slides) != expected_slide_count:
                raise ValueError(
                    f"Generated PPTX slide count mismatch: expected {expected_slide_count}, got {len(presentation.slides)}"
                )
        except Exception:
            output_path.unlink(missing_ok=True)
            raise
