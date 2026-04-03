from __future__ import annotations

import base64
import copy
import math
import re
import zipfile
from datetime import UTC, datetime
from io import BytesIO
from pathlib import Path

from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
from pptx.util import Pt

from a3presentation.domain.presentation import PresentationPlan, SlideKind, SlideSpec
from a3presentation.domain.template import (
    GenerationMode,
    LayoutSpec,
    PlaceholderKind,
    PrototypeSlideSpec,
    TemplateManifest,
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
        if slide_spec.preferred_layout_key:
            for prototype_slide in manifest.prototype_slides:
                if prototype_slide.key == slide_spec.preferred_layout_key:
                    return prototype_slide

        for prototype_slide in manifest.prototype_slides:
            if slide_spec.kind.value in prototype_slide.supported_slide_kinds:
                return prototype_slide

        return manifest.prototype_slides[0]

    def _replace_tokens_in_slide(self, slide, prototype: PrototypeSlideSpec, slide_spec: SlideSpec, presentation_title: str) -> None:
        token_values = self._build_token_value_map(slide_spec, presentation_title)
        used_shapes: set[str] = set()

        # Preferred path for real templates: bind by explicit shape name from manifest.
        for token_spec in prototype.tokens:
            if not token_spec.shape_name:
                continue
            target_shape = next((shape for shape in slide.shapes if shape.name == token_spec.shape_name), None)
            if target_shape is None:
                continue
            self._fill_shape_by_binding(target_shape, token_spec.binding, slide_spec, presentation_title)
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
                if isinstance(token_value, list):
                    self._set_bullets(shape, token_value)
                else:
                    self._set_text(shape, str(token_value))
                continue

            replaced_text = original_text
            for token_name in matches:
                token_value = token_values.get(token_name, "")
                if isinstance(token_value, list):
                    token_value = "\n".join(token_value)
                replaced_text = re.sub(r"{{\s*" + re.escape(token_name) + r"\s*}}", str(token_value), replaced_text)
            self._set_text(shape, replaced_text)

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
        placeholders = {placeholder.placeholder_format.idx: placeholder for placeholder in slide.placeholders}
        used_placeholder_indices: set[int] = set()

        for placeholder_spec in layout.placeholders:
            if placeholder_spec.idx is None or placeholder_spec.idx not in placeholders:
                continue
            shape = placeholders[placeholder_spec.idx]
            used_placeholder_indices.add(placeholder_spec.idx)
            if placeholder_spec.binding:
                self._fill_shape_by_binding(shape, placeholder_spec.binding, slide_spec, presentation_title)
                continue
            if placeholder_spec.kind == PlaceholderKind.TITLE:
                self._set_text(shape, slide_spec.title or "")
            elif placeholder_spec.kind == PlaceholderKind.SUBTITLE:
                self._set_text(shape, slide_spec.subtitle or "")
            elif placeholder_spec.kind == PlaceholderKind.BODY:
                self._fill_body(shape, slide_spec)
            elif placeholder_spec.kind == PlaceholderKind.FOOTER:
                self._set_text(shape, slide_spec.notes or "")
            elif placeholder_spec.kind == PlaceholderKind.TABLE:
                self._fill_table(shape, slide_spec)

        for placeholder in slide.placeholders:
            placeholder_idx = placeholder.placeholder_format.idx
            if placeholder_idx in used_placeholder_indices:
                continue
            self._clear_placeholder(placeholder)

        if layout.key in {"text_full_width", "list_full_width"}:
            self._expand_text_full_width_layout(slide)
        elif layout.key == "table":
            self._expand_table_layout(slide)

        self._adjust_title_and_flow(slide, layout.key)

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
        lines = text.splitlines() or [text]
        for index, line in enumerate(lines):
            paragraph = text_frame.paragraphs[0] if index == 0 else text_frame.add_paragraph()
            paragraph.alignment = align
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

    def _fill_body(self, shape, slide_spec: SlideSpec) -> None:
        if slide_spec.kind == SlideKind.BULLETS:
            if not slide_spec.bullets:
                self._clear_placeholder(shape)
                return
            self._set_bullets(shape, slide_spec.bullets)
            return
        if slide_spec.kind == SlideKind.TWO_COLUMN:
            merged = [*slide_spec.left_bullets, "", *slide_spec.right_bullets]
            if not any(item.strip() for item in merged):
                self._clear_placeholder(shape)
                return
            self._set_bullets(shape, merged)
            return
        if slide_spec.kind == SlideKind.TEXT:
            if not (slide_spec.text or "").strip():
                self._clear_placeholder(shape)
                return
            self._set_text(shape, slide_spec.text or "")
            return
        if slide_spec.kind == SlideKind.TITLE:
            if not (slide_spec.text or "").strip():
                self._clear_placeholder(shape)
                return
            self._set_text(shape, slide_spec.text or "")
            return
        if slide_spec.table is not None:
            rows = [" | ".join(row) for row in slide_spec.table.rows]
            if not rows:
                self._clear_placeholder(shape)
                return
            self._set_bullets(shape, rows)
            return
        if not (slide_spec.text or "").strip():
            self._clear_placeholder(shape)
            return
        self._set_text(shape, slide_spec.text or "")

    def _fill_shape_by_binding(self, shape, binding: str, slide_spec: SlideSpec, presentation_title: str) -> None:
        binding_value = self._build_token_value_map(slide_spec, presentation_title).get(binding, "")
        if binding == "table":
            self._fill_table(shape, slide_spec)
            return
        if binding == "image":
            self._fill_image(shape, slide_spec)
            return
        if binding in {"chart", "chart_image", "icon_grid"}:
            self._clear_placeholder(shape)
            return
        if self._is_empty_binding_value(binding_value) and binding not in {"presentation_name", "cover_title", "title"}:
            self._clear_placeholder(shape)
            return
        if not getattr(shape, "has_text_frame", False):
            return
        if isinstance(binding_value, list):
            self._set_bullets(shape, [str(item) for item in binding_value])
            return
        self._set_text(shape, str(binding_value))

    def _expand_text_full_width_layout(self, slide) -> None:
        placeholders = {
            shape.placeholder_format.idx: shape
            for shape in slide.placeholders
            if getattr(shape, "is_placeholder", False)
        }

        title = placeholders.get(0)
        subtitle = placeholders.get(13)
        main_text = placeholders.get(14)
        secondary_text = placeholders.get(15)

        title_top = 671247
        title_height = 1120247

        if title is not None:
            title.left = self.FULL_CONTENT_LEFT_EMU
            title.top = title_top
            title.width = self.FULL_CONTENT_WIDTH_EMU
            title.height = max(title.height or 0, title_height)

        if subtitle is not None:
            subtitle.left = self.FULL_CONTENT_LEFT_EMU
            subtitle.top = 1228230
            subtitle.width = self.FULL_CONTENT_WIDTH_EMU
            subtitle.height = 552402

        if main_text is not None:
            main_text.left = self.FULL_CONTENT_LEFT_EMU
            main_text.top = 1791494
            main_text.width = self.FULL_CONTENT_WIDTH_EMU
            main_text.height = 1700000 if secondary_text is not None and getattr(secondary_text, "text", "").strip() else 3550000

        if secondary_text is not None and getattr(secondary_text, "text", "").strip():
            secondary_text.left = self.FULL_CONTENT_LEFT_EMU
            secondary_text.top = 3800000
            secondary_text.width = self.FULL_CONTENT_WIDTH_EMU
            secondary_text.height = 1850000

        footer = placeholders.get(17)
        if footer is not None:
            footer.left = self.FULL_CONTENT_LEFT_EMU
            footer.top = self.FOOTER_TOP_EMU
            footer.width = self.FULL_CONTENT_WIDTH_EMU
            footer.height = self.FOOTER_HEIGHT_EMU

    def _expand_table_layout(self, slide) -> None:
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
            title.left = self.FULL_CONTENT_LEFT_EMU
            title.top = 671247
            title.width = self.FULL_CONTENT_WIDTH_EMU
            title.height = max(title.height or 0, 584960)

        if subtitle is None or not getattr(subtitle, "text", "").strip():
            return

        has_side_notes = any(
            shape is not None and getattr(shape, "text", "").strip()
            for shape in (left_note, right_note)
        )
        if has_side_notes:
            return

        subtitle.left = self.FULL_CONTENT_LEFT_EMU
        subtitle.top = 1228230
        subtitle.width = self.FULL_CONTENT_WIDTH_EMU
        subtitle.height = 700000

        if footer is not None:
            footer.left = self.FULL_CONTENT_LEFT_EMU
            footer.top = self.FOOTER_TOP_EMU
            footer.width = self.FULL_CONTENT_WIDTH_EMU
            footer.height = self.FOOTER_HEIGHT_EMU

    def _adjust_title_and_flow(self, slide, layout_key: str) -> None:
        if layout_key in {"text_full_width", "list_full_width"}:
            self._stack_text_content(slide, layout_key)
            return
        if layout_key == "table":
            self._stack_table_content(slide, layout_key)
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

        has_subtitle = subtitle is not None and getattr(subtitle, "text", "").strip()
        title_gap = self.TITLE_CONTENT_GAP_EMU if has_subtitle else self.TITLE_BODY_GAP_NO_SUBTITLE_EMU
        cursor = title.top + title.height + title_gap

        if has_subtitle:
            subtitle_text = subtitle.text.strip()
            subtitle.height = max(360000, self._estimate_text_height_emu(subtitle_text, subtitle.width, 18.0))
            subtitle.top = cursor
            cursor = subtitle.top + subtitle.height + self.TITLE_CONTENT_GAP_EMU

        secondary_has_text = secondary is not None and getattr(secondary, "text", "").strip()
        available_bottom = footer.top - self.CONTENT_FOOTER_GAP_EMU

        if secondary_has_text and secondary is not None:
            secondary_text = secondary.text.strip()
            secondary.height = max(secondary.height or 0, self._estimate_text_height_emu(secondary_text, secondary.width, 16.0))
            secondary.height = min(secondary.height, max(700000, available_bottom - cursor - 900000))
            body.top = cursor
            body.height = max(900000, secondary.top - self.TITLE_CONTENT_GAP_EMU - body.top)
            secondary.top = body.top + body.height + self.TITLE_CONTENT_GAP_EMU
            secondary.height = max(700000, min(secondary.height, available_bottom - secondary.top))
            return

        body.top = cursor
        body.height = max(900000, available_bottom - body.top)

    def _stack_table_content(self, slide, layout_key: str) -> None:
        placeholders = {
            shape.placeholder_format.idx: shape
            for shape in slide.placeholders
            if getattr(shape, "is_placeholder", False)
        }
        title = placeholders.get(0)
        subtitle = placeholders.get(13)
        table = placeholders.get(14)
        footer = placeholders.get(15)

        if title is None or table is None or footer is None:
            return

        title_text = (getattr(title, "text", "") or "").strip()
        if title_text:
            font_size_pt = self._fit_title_font_size_points(title_text, title.width, layout_key)
            self._apply_font_size(title, font_size_pt)
            self._configure_title_text_frame(title)
            required_height = self._estimate_title_height_emu(title, title_text, font_size_pt)
            title.height = max(self._minimum_title_height_emu(layout_key), required_height)

        cursor = title.top + title.height + self.TITLE_CONTENT_GAP_EMU
        if subtitle is not None and getattr(subtitle, "text", "").strip():
            subtitle_text = subtitle.text.strip()
            subtitle.height = max(360000, self._estimate_text_height_emu(subtitle_text, subtitle.width, 16.0))
            subtitle.top = cursor
            cursor = subtitle.top + subtitle.height + self.TITLE_CONTENT_GAP_EMU

        table.top = cursor
        table.height = max(900000, footer.top - self.CONTENT_FOOTER_GAP_EMU - table.top)

    def _title_font_size_points(self, layout_key: str) -> float:
        if layout_key == "table":
            return 28.0
        if layout_key in {"text_full_width", "list_full_width", "image_text"}:
            return 30.0
        if layout_key == "cards_3":
            return 28.0
        if layout_key == "list_with_icons":
            return 28.0
        return 30.0

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

    def _estimate_title_height_emu(self, shape, text: str, font_size_pt: float) -> int:
        text_frame = shape.text_frame
        effective_width = max(shape.width - text_frame.margin_left - text_frame.margin_right, shape.width // 2)
        width_pt = effective_width / self.EMU_PER_PT
        average_char_width_pt = max(font_size_pt * 0.6, 1.0)
        chars_per_line = max(int(width_pt / average_char_width_pt), 6)

        wrapped_lines = 0
        for paragraph in text.splitlines() or [text]:
            normalized = paragraph.strip()
            if not normalized:
                wrapped_lines += 1
                continue
            wrapped_lines += max(1, math.ceil(len(normalized) / chars_per_line))

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
        chars_per_line = max(int(width_pt / average_char_width_pt), 8)

        wrapped_lines = 0
        for paragraph in text.splitlines() or [text]:
            normalized = paragraph.strip()
            if not normalized:
                wrapped_lines += 1
                continue
            wrapped_lines += max(1, math.ceil(len(normalized) / chars_per_line))

        line_height_pt = font_size_pt * 1.18
        vertical_padding_pt = font_size_pt * 0.7
        return int((wrapped_lines * line_height_pt + vertical_padding_pt) * self.EMU_PER_PT)

    def _fill_table(self, shape, slide_spec: SlideSpec) -> None:
        if slide_spec.table is None:
            if getattr(shape, "has_text_frame", False):
                self._set_text(shape, "")
            return

        headers = slide_spec.table.headers
        rows = slide_spec.table.rows
        row_count = len(rows) + (1 if headers else 0)
        col_count = len(headers) if headers else max((len(row) for row in rows), default=0)
        if row_count == 0 or col_count == 0:
            if getattr(shape, "has_text_frame", False):
                self._set_text(shape, "")
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
                final_height = self._format_table(table, headers, rows, graphic_frame.width, graphic_frame.height)
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
            final_height = self._format_table(table, headers, rows, shape.width, shape.height)
            shape.height = final_height
            return

        as_lines = []
        if headers:
            as_lines.append(" | ".join(headers))
        as_lines.extend(" | ".join(row) for row in rows)
        if getattr(shape, "has_text_frame", False):
            self._set_bullets(shape, as_lines)

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
                self._set_text(shape, "Изображение из документа")

    def _set_text(self, shape, text: str) -> None:
        text_frame = shape.text_frame
        text_frame.clear()
        text_frame.text = text
        if getattr(shape, "is_placeholder", False):
            placeholder_format = getattr(shape, "placeholder_format", None)
            if placeholder_format is not None and placeholder_format.idx in {15, 17}:
                self._apply_footer_font_size(text_frame, text)
                return
        self._apply_body_font_size(text_frame, [text])

    def _set_bullets(self, shape, items: list[str]) -> None:
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
                self._apply_bullet_format(paragraph)
            first = False
        self._apply_body_font_size(text_frame, items)

    def _apply_body_font_size(self, text_frame, items: list[str]) -> None:
        non_empty_items = [item.strip() for item in items if item and item.strip()]
        if not non_empty_items:
            return

        total_chars = sum(len(item) for item in non_empty_items)
        max_item_len = max(len(item) for item in non_empty_items)
        font_size = None

        if total_chars >= 1100 or max_item_len >= 320:
            font_size = Pt(12)
        elif total_chars >= 850 or max_item_len >= 240:
            font_size = Pt(13)
        elif total_chars >= 650 or max_item_len >= 180:
            font_size = Pt(14)

        if font_size is None:
            return

        for paragraph in text_frame.paragraphs:
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

    def _apply_bullet_format(self, paragraph) -> None:
        paragraph_properties = paragraph._p.get_or_add_pPr()
        for child in list(paragraph_properties):
            if child.tag.endswith("}buNone") or child.tag.endswith("}buChar") or child.tag.endswith("}buAutoNum"):
                paragraph_properties.remove(child)

        bullet = OxmlElement("a:buChar")
        bullet.set("char", "•")
        paragraph_properties.insert(0, bullet)
        paragraph_properties.set("marL", "342900")
        paragraph_properties.set("indent", "-171450")

    def _format_table(self, table, headers: list[str], rows: list[list[str]], target_width: int, target_height: int) -> int:
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
                self._style_table_cell(cell, is_header=is_header, font_size=font_size, margins=margins)
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

    def _style_table_cell(self, cell, *, is_header: bool, font_size: Pt, margins: tuple[int, int, int, int]) -> None:
        fill = cell.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(0x08, 0x1C, 0x4F) if is_header else RGBColor(0xEB, 0xF3, 0xFE)

        text_frame = cell.text_frame
        text_frame.word_wrap = True
        cell.margin_left, cell.margin_right, cell.margin_top, cell.margin_bottom = margins

        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = font_size
                run.font.bold = bool(is_header)
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF) if is_header else RGBColor(0x08, 0x1C, 0x4F)

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
