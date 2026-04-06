from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from pptx import Presentation
from a3presentation.domain.presentation import PresentationPlan, SlideKind, SlideSpec
from a3presentation.domain.template import PlaceholderKind, TemplateManifest
from a3presentation.services.layout_capacity import (
    LayoutCapacityProfile,
    LayoutGeometryPolicy,
    PlaceholderGeometryPolicy,
    geometry_policy_for_layout,
    profile_for_layout,
)
from a3presentation.services.pptx_generator import PptxGenerator

CONTINUATION_FONT_DELTA_TOLERANCE_PT = 2.0
CONTINUATION_UNDERFILL_GRACE = 0.09
CONTINUATION_AUDIT_EPSILON = 0.01
GEOMETRY_TOLERANCE_EMU = 90000
BODY_HEIGHT_UNDERFILL_RATIO = 0.45


@dataclass(frozen=True)
class SlideAudit:
    slide_index: int
    title: str
    kind: str
    layout_key: str
    body_char_count: int
    body_font_sizes: tuple[float, ...]
    profile: LayoutCapacityProfile
    content_width: int | None = None
    footer_width: int | None = None
    has_table: bool = False
    has_chart: bool = False
    has_image: bool = False
    expected_items: tuple[str, ...] = ()
    rendered_items: tuple[str, ...] = ()
    title_top: int | None = None
    title_height: int | None = None
    title_width: int | None = None
    subtitle_top: int | None = None
    subtitle_height: int | None = None
    subtitle_width: int | None = None
    body_top: int | None = None
    body_height: int | None = None
    body_left: int | None = None
    footer_top: int | None = None
    footer_left: int | None = None
    auxiliary_widths: dict[int, int] | None = None
    auxiliary_lefts: dict[int, int] | None = None
    auxiliary_tops: dict[int, int] | None = None
    image_width: int | None = None
    image_left: int | None = None
    image_top: int | None = None
    body_margin_left: int | None = None
    body_margin_right: int | None = None
    body_margin_top: int | None = None
    body_margin_bottom: int | None = None
    expected_body_margin_left: int | None = None
    expected_body_margin_right: int | None = None
    expected_body_margin_top: int | None = None
    expected_body_margin_bottom: int | None = None
    title_placeholder_idx: int | None = None
    subtitle_placeholder_idx: int | None = None
    body_placeholder_idx: int | None = None
    footer_placeholder_idx: int | None = None
    geometry: LayoutGeometryPolicy | None = None

    @property
    def fill_ratio(self) -> float:
        if self.profile.max_chars <= 0:
            return 0.0
        return self.body_char_count / self.profile.max_chars

    @property
    def min_font_size(self) -> float | None:
        if not self.body_font_sizes:
            return None
        return min(self.body_font_sizes)

    @property
    def max_font_size(self) -> float | None:
        if not self.body_font_sizes:
            return None
        return max(self.body_font_sizes)

    @property
    def representative_font_size(self) -> float | None:
        if not self.body_font_sizes:
            return None
        return min(self.body_font_sizes)

    @property
    def within_font_bounds(self) -> bool:
        if not self.body_font_sizes:
            return True
        return (
            min(self.body_font_sizes) >= self.profile.min_font_pt
            and max(self.body_font_sizes) <= self.profile.max_font_pt
        )

    @property
    def content_width_ratio(self) -> float:
        if not self.content_width:
            return 0.0
        return self.content_width / PptxGenerator.FULL_CONTENT_WIDTH_EMU

    @property
    def footer_width_ratio(self) -> float:
        if not self.footer_width:
            return 0.0
        return self.footer_width / PptxGenerator.FULL_CONTENT_WIDTH_EMU

    @property
    def title_bottom(self) -> int | None:
        if self.title_top is None or self.title_height is None:
            return None
        return self.title_top + self.title_height

    @property
    def subtitle_bottom(self) -> int | None:
        if self.subtitle_top is None or self.subtitle_height is None:
            return None
        return self.subtitle_top + self.subtitle_height


@dataclass(frozen=True)
class CapacityViolation:
    slide_index: int
    title: str
    rule: str
    details: str


def audit_generated_presentation(
    output_path: Path,
    plan: PresentationPlan,
    manifest: TemplateManifest | None = None,
) -> list[SlideAudit]:
    presentation = Presentation(str(output_path))
    audits: list[SlideAudit] = []

    for slide_index, slide_spec in enumerate(plan.slides, start=1):
        if slide_index > len(presentation.slides):
            break
        if slide_spec.kind not in {
            SlideKind.TEXT,
            SlideKind.BULLETS,
            SlideKind.TWO_COLUMN,
            SlideKind.TABLE,
            SlideKind.CHART,
            SlideKind.IMAGE,
        }:
            continue

        slide = presentation.slides[slide_index - 1]
        placeholders = {
            shape.placeholder_format.idx: shape
            for shape in slide.placeholders
            if getattr(shape, "is_placeholder", False)
        }
        title_idx = _placeholder_idx_for_role(slide_spec, manifest, PlaceholderKind.TITLE, (0,))
        subtitle_idx = _placeholder_idx_for_role(slide_spec, manifest, PlaceholderKind.SUBTITLE, (13,))
        body_idx = _placeholder_idx_for_role(slide_spec, manifest, PlaceholderKind.BODY, (14,))
        footer_idx = _placeholder_idx_for_role(slide_spec, manifest, PlaceholderKind.FOOTER, (15, 17, 21))
        body = placeholders.get(body_idx) if body_idx is not None else None
        footer = placeholders.get(footer_idx) if footer_idx is not None else None
        title = placeholders.get(title_idx) if title_idx is not None else None
        subtitle = placeholders.get(subtitle_idx) if subtitle_idx is not None else None
        resolved_geometry = _geometry_policy_for_slide(slide_spec, manifest)
        expected_body_spec = _placeholder_spec_for_role(slide_spec, manifest, PlaceholderKind.BODY)
        auxiliary_widths = {
            idx: getattr(shape, "width", None)
            for idx, shape in placeholders.items()
            if idx not in {value for value in {title_idx, subtitle_idx, body_idx, footer_idx, 17} if value is not None}
        }
        auxiliary_lefts = {
            idx: getattr(shape, "left", None)
            for idx, shape in placeholders.items()
            if idx not in {value for value in {title_idx, subtitle_idx, body_idx, footer_idx, 17} if value is not None}
        }
        auxiliary_tops = {
            idx: getattr(shape, "top", None)
            for idx, shape in placeholders.items()
            if idx not in {value for value in {title_idx, subtitle_idx, body_idx, footer_idx, 17} if value is not None}
        }
        content_width = getattr(body, "width", None) if body is not None else None
        footer_width = getattr(footer, "width", None) if footer is not None else None

        body_char_count = 0
        body_font_sizes: tuple[float, ...] = ()
        rendered_items: tuple[str, ...] = ()
        if body is not None and getattr(body, "has_text_frame", False):
            body_paragraphs = [paragraph.text.strip() for paragraph in body.text_frame.paragraphs if paragraph.text.strip()]
            body_char_count = sum(len(paragraph) for paragraph in body_paragraphs)
            rendered_items = tuple(body_paragraphs)
            body_font_sizes = tuple(
                sorted(
                    {
                        run.font.size.pt
                        for paragraph in body.text_frame.paragraphs
                        for run in paragraph.runs
                        if run.font.size is not None
                    }
                )
            )

        has_table = any(getattr(shape, "has_table", False) for shape in slide.shapes)
        has_chart = any(getattr(shape, "has_chart", False) for shape in slide.shapes)
        has_image = any(self_has_image(shape) for shape in slide.shapes)
        image_shape = next((shape for shape in slide.shapes if self_has_image(shape)), None)
        if content_width is None:
            if has_table:
                table_shape = next((shape for shape in slide.shapes if getattr(shape, "has_table", False)), None)
                content_width = getattr(table_shape, "width", None) if table_shape is not None else None
            elif has_chart:
                chart_shape = next((shape for shape in slide.shapes if getattr(shape, "has_chart", False)), None)
                content_width = getattr(chart_shape, "width", None) if chart_shape is not None else None
            elif has_image:
                content_width = getattr(image_shape, "width", None) if image_shape is not None else None
        layout_key = slide_spec.preferred_layout_key or _infer_layout_key(slide_spec.kind.value)
        expected_items = _expected_items_for_slide(slide_spec)
        audits.append(
            SlideAudit(
                slide_index=slide_index,
                title=slide_spec.title or "",
                kind=slide_spec.kind.value,
                layout_key=layout_key,
                body_char_count=body_char_count,
                body_font_sizes=body_font_sizes,
                profile=profile_for_layout(layout_key),
                content_width=content_width,
                footer_width=footer_width,
                has_table=has_table,
                has_chart=has_chart,
                has_image=has_image,
                expected_items=expected_items,
                rendered_items=rendered_items,
                title_top=getattr(title, "top", None) if title is not None else None,
                title_height=getattr(title, "height", None) if title is not None else None,
                title_width=getattr(title, "width", None) if title is not None else None,
                subtitle_top=getattr(subtitle, "top", None) if subtitle is not None else None,
                subtitle_height=getattr(subtitle, "height", None) if subtitle is not None else None,
                subtitle_width=getattr(subtitle, "width", None) if subtitle is not None else None,
                body_top=getattr(body, "top", None) if body is not None else None,
                body_height=getattr(body, "height", None) if body is not None else None,
                body_left=getattr(body, "left", None) if body is not None else None,
                footer_top=getattr(footer, "top", None) if footer is not None else None,
                footer_left=getattr(footer, "left", None) if footer is not None else None,
                auxiliary_widths=auxiliary_widths,
                auxiliary_lefts=auxiliary_lefts,
                auxiliary_tops=auxiliary_tops,
                image_width=getattr(image_shape, "width", None) if image_shape is not None else None,
                image_left=getattr(image_shape, "left", None) if image_shape is not None else None,
                image_top=getattr(image_shape, "top", None) if image_shape is not None else None,
                body_margin_left=getattr(getattr(body, "text_frame", None), "margin_left", None) if body is not None and getattr(body, "has_text_frame", False) else None,
                body_margin_right=getattr(getattr(body, "text_frame", None), "margin_right", None) if body is not None and getattr(body, "has_text_frame", False) else None,
                body_margin_top=getattr(getattr(body, "text_frame", None), "margin_top", None) if body is not None and getattr(body, "has_text_frame", False) else None,
                body_margin_bottom=getattr(getattr(body, "text_frame", None), "margin_bottom", None) if body is not None and getattr(body, "has_text_frame", False) else None,
                expected_body_margin_left=getattr(expected_body_spec, "margin_left_emu", None),
                expected_body_margin_right=getattr(expected_body_spec, "margin_right_emu", None),
                expected_body_margin_top=getattr(expected_body_spec, "margin_top_emu", None),
                expected_body_margin_bottom=getattr(expected_body_spec, "margin_bottom_emu", None),
                title_placeholder_idx=title_idx,
                subtitle_placeholder_idx=subtitle_idx,
                body_placeholder_idx=body_idx,
                footer_placeholder_idx=footer_idx,
                geometry=resolved_geometry,
            )
        )

    return audits


def continuation_groups(audits: list[SlideAudit]) -> dict[str, list[SlideAudit]]:
    groups: dict[str, list[SlideAudit]] = {}
    for audit in audits:
        if audit.kind not in {SlideKind.TEXT.value, SlideKind.BULLETS.value, SlideKind.TWO_COLUMN.value}:
            continue
        base_title = re.sub(r"\s+\(\d+\)$", "", audit.title.strip())
        groups.setdefault(base_title, []).append(audit)
    return {title: items for title, items in groups.items() if len(items) > 1}


def find_capacity_violations(audits: list[SlideAudit]) -> list[CapacityViolation]:
    violations: list[CapacityViolation] = []

    for audit in audits:
        geometry = audit.geometry or geometry_policy_for_layout(audit.layout_key)
        if not audit.within_font_bounds:
            violations.append(
                CapacityViolation(
                    slide_index=audit.slide_index,
                    title=audit.title,
                    rule="font_bounds",
                    details=(
                        f"fonts={audit.body_font_sizes} profile={audit.profile.min_font_pt}-{audit.profile.max_font_pt}"
                    ),
                )
            )

        if audit.fill_ratio > (audit.profile.max_fill_ratio + 0.02):
            violations.append(
                CapacityViolation(
                    slide_index=audit.slide_index,
                    title=audit.title,
                    rule="overflow_risk",
                    details=f"fill_ratio={audit.fill_ratio:.2f} max={audit.profile.max_fill_ratio:.2f}",
                )
            )

        if audit.kind == SlideKind.TABLE.value:
            if not audit.has_table:
                violations.append(
                    CapacityViolation(
                        slide_index=audit.slide_index,
                        title=audit.title,
                        rule="missing_table_shape",
                        details="table slide does not contain rendered table shape",
                    )
                )
            if audit.content_width_ratio and audit.content_width_ratio < 0.9:
                violations.append(
                    CapacityViolation(
                        slide_index=audit.slide_index,
                        title=audit.title,
                        rule="narrow_table_content",
                        details=f"content_width_ratio={audit.content_width_ratio:.2f}",
                    )
                )

        if audit.kind == SlideKind.CHART.value:
            if not audit.has_chart:
                violations.append(
                    CapacityViolation(
                        slide_index=audit.slide_index,
                        title=audit.title,
                        rule="missing_chart_shape",
                        details="chart slide does not contain rendered chart shape",
                    )
                )
            if audit.content_width_ratio and audit.content_width_ratio < 0.9:
                violations.append(
                    CapacityViolation(
                        slide_index=audit.slide_index,
                        title=audit.title,
                        rule="narrow_chart_content",
                        details=f"content_width_ratio={audit.content_width_ratio:.2f}",
                    )
                )

        if audit.kind == SlideKind.IMAGE.value:
            if not audit.has_image:
                violations.append(
                    CapacityViolation(
                        slide_index=audit.slide_index,
                        title=audit.title,
                        rule="missing_image_shape",
                        details="image slide does not contain rendered image shape",
                    )
                )
            if audit.content_width_ratio and audit.content_width_ratio < 0.35:
                violations.append(
                    CapacityViolation(
                        slide_index=audit.slide_index,
                        title=audit.title,
                        rule="narrow_image_content",
                        details=f"content_width_ratio={audit.content_width_ratio:.2f}",
                    )
                )
            if audit.layout_key == "image_text" and audit.image_left is not None and audit.body_left is not None and audit.content_width:
                separation = audit.image_left - (audit.body_left + audit.content_width)
                if separation < 300000:
                    violations.append(
                        CapacityViolation(
                            slide_index=audit.slide_index,
                            title=audit.title,
                            rule="image_text_overlap",
                            details=f"separation={separation} min=300000",
                        )
                    )

        if audit.kind in {SlideKind.TABLE.value, SlideKind.CHART.value}:
            if audit.footer_width_ratio and audit.footer_width_ratio < 0.9:
                violations.append(
                    CapacityViolation(
                        slide_index=audit.slide_index,
                        title=audit.title,
                        rule="narrow_table_footer",
                        details=f"footer_width_ratio={audit.footer_width_ratio:.2f}",
                    )
                )

        if audit.layout_key in {"text_full_width", "list_full_width"}:
            if audit.footer_width_ratio and audit.footer_width_ratio < 0.9:
                violations.append(
                    CapacityViolation(
                        slide_index=audit.slide_index,
                        title=audit.title,
                        rule="narrow_text_footer",
                        details=f"footer_width_ratio={audit.footer_width_ratio:.2f}",
                    )
                )

            expected_body_geometry = geometry.placeholders.get(audit.body_placeholder_idx or 14)
            if expected_body_geometry is not None and audit.body_left is not None and abs(audit.body_left - expected_body_geometry.left_emu) > GEOMETRY_TOLERANCE_EMU:
                violations.append(
                    CapacityViolation(
                        slide_index=audit.slide_index,
                        title=audit.title,
                        rule="body_left_misalignment",
                        details=(
                            f"body_left={audit.body_left} expected={expected_body_geometry.left_emu} "
                            f"tolerance={GEOMETRY_TOLERANCE_EMU}"
                        ),
                    )
                )
            expected_margin_left = audit.expected_body_margin_left or PptxGenerator.DEFAULT_TEXT_MARGIN_X_EMU
            expected_margin_right = audit.expected_body_margin_right or PptxGenerator.DEFAULT_TEXT_MARGIN_X_EMU
            expected_margin_top = audit.expected_body_margin_top or PptxGenerator.DEFAULT_TEXT_MARGIN_Y_EMU
            expected_margin_bottom = audit.expected_body_margin_bottom or PptxGenerator.DEFAULT_TEXT_MARGIN_Y_EMU
            if audit.body_margin_left is not None and abs(audit.body_margin_left - expected_margin_left) > GEOMETRY_TOLERANCE_EMU:
                violations.append(
                    CapacityViolation(
                        slide_index=audit.slide_index,
                        title=audit.title,
                        rule="body_margin_mismatch",
                        details=f"margin_left={audit.body_margin_left} expected={expected_margin_left}",
                    )
                )
            if audit.body_margin_right is not None and abs(audit.body_margin_right - expected_margin_right) > GEOMETRY_TOLERANCE_EMU:
                violations.append(
                    CapacityViolation(
                        slide_index=audit.slide_index,
                        title=audit.title,
                        rule="body_margin_mismatch",
                        details=f"margin_right={audit.body_margin_right} expected={expected_margin_right}",
                    )
                )
            if audit.body_margin_top is not None and abs(audit.body_margin_top - expected_margin_top) > GEOMETRY_TOLERANCE_EMU:
                violations.append(
                    CapacityViolation(
                        slide_index=audit.slide_index,
                        title=audit.title,
                        rule="body_margin_mismatch",
                        details=f"margin_top={audit.body_margin_top} expected={expected_margin_top}",
                    )
                )
            if audit.body_margin_bottom is not None and abs(audit.body_margin_bottom - expected_margin_bottom) > GEOMETRY_TOLERANCE_EMU:
                violations.append(
                    CapacityViolation(
                        slide_index=audit.slide_index,
                        title=audit.title,
                        rule="body_margin_mismatch",
                        details=f"margin_bottom={audit.body_margin_bottom} expected={expected_margin_bottom}",
                    )
                )

        if audit.layout_key == "image_text":
            if audit.image_width is not None and audit.image_width < geometry.placeholders[16].width_emu - GEOMETRY_TOLERANCE_EMU:
                violations.append(
                    CapacityViolation(
                        slide_index=audit.slide_index,
                        title=audit.title,
                        rule="narrow_image_panel",
                        details=f"image_width={audit.image_width} min={geometry.placeholders[16].width_emu - GEOMETRY_TOLERANCE_EMU}",
                    )
                )
            if audit.body_left is not None and abs(audit.body_left - geometry.placeholders[14].left_emu) > GEOMETRY_TOLERANCE_EMU:
                violations.append(
                    CapacityViolation(
                        slide_index=audit.slide_index,
                        title=audit.title,
                        rule="image_text_body_misalignment",
                        details=f"body_left={audit.body_left} expected={geometry.placeholders[14].left_emu}",
                    )
                )

        if audit.layout_key == "cards_3":
            card_lefts = audit.auxiliary_lefts or {}
            card_widths = audit.auxiliary_widths or {}
            card_positions = [(idx, card_lefts.get(idx), card_widths.get(idx)) for idx in (11, 12, 13)]
            if all(left is not None and width is not None for _, left, width in card_positions):
                previous_right = None
                for idx, left, width in card_positions:
                    if width < geometry.placeholders[idx].width_emu - GEOMETRY_TOLERANCE_EMU:
                        violations.append(
                            CapacityViolation(
                                slide_index=audit.slide_index,
                                title=audit.title,
                                rule="narrow_card_placeholder",
                                details=f"idx={idx} width={width}",
                            )
                        )
                    if previous_right is not None and left < previous_right + 120000:
                        violations.append(
                            CapacityViolation(
                                slide_index=audit.slide_index,
                                title=audit.title,
                                rule="card_overlap",
                                details=f"idx={idx} left={left} previous_right={previous_right}",
                            )
                        )
                    previous_right = left + width

        if audit.layout_key == "list_with_icons":
            aux_lefts = audit.auxiliary_lefts or {}
            aux_widths = audit.auxiliary_widths or {}
            left_left = aux_lefts.get(12)
            left_width = aux_widths.get(12)
            right_left = aux_lefts.get(14)
            right_width = aux_widths.get(14)
            if None not in {left_left, left_width, right_left, right_width}:
                if left_width < geometry.placeholders[12].width_emu - GEOMETRY_TOLERANCE_EMU:
                    violations.append(
                        CapacityViolation(
                            slide_index=audit.slide_index,
                            title=audit.title,
                            rule="narrow_left_column",
                            details=f"width={left_width}",
                        )
                    )
                if right_width < geometry.placeholders[14].width_emu - GEOMETRY_TOLERANCE_EMU:
                    violations.append(
                        CapacityViolation(
                            slide_index=audit.slide_index,
                            title=audit.title,
                            rule="narrow_right_column",
                            details=f"width={right_width}",
                        )
                    )
                if right_left < left_left + left_width + 300000:
                    violations.append(
                        CapacityViolation(
                            slide_index=audit.slide_index,
                            title=audit.title,
                            rule="two_column_overlap",
                            details=f"left_right_gap={right_left - (left_left + left_width)}",
                        )
                    )

        if audit.layout_key == "contacts":
            aux_lefts = audit.auxiliary_lefts or {}
            aux_widths = audit.auxiliary_widths or {}
            for idx in (10, 11, 12, 13):
                expected = geometry.placeholders[idx]
                left = aux_lefts.get(idx)
                width = aux_widths.get(idx)
                if left is not None and abs(left - expected.left_emu) > GEOMETRY_TOLERANCE_EMU:
                    violations.append(
                        CapacityViolation(
                            slide_index=audit.slide_index,
                            title=audit.title,
                            rule="contact_block_misalignment",
                            details=f"idx={idx} left={left} expected={expected.left_emu}",
                        )
                    )
                if width is not None and width < expected.width_emu - GEOMETRY_TOLERANCE_EMU:
                    violations.append(
                        CapacityViolation(
                            slide_index=audit.slide_index,
                            title=audit.title,
                            rule="narrow_contact_block",
                            details=f"idx={idx} width={width}",
                        )
                    )

        if audit.layout_key in {"text_full_width", "list_full_width", "image_text"}:
            expected_body_geometry = geometry.placeholders.get(audit.body_placeholder_idx or 14)
            expected_body_height = expected_body_geometry.height_emu if expected_body_geometry is not None else None
            if (
                expected_body_height is not None
                and
                audit.body_height is not None
                and audit.body_height < int(expected_body_height * BODY_HEIGHT_UNDERFILL_RATIO)
                and audit.fill_ratio < max(audit.profile.target_fill_ratio - 0.2, 0.0)
            ):
                violations.append(
                    CapacityViolation(
                        slide_index=audit.slide_index,
                        title=audit.title,
                        rule="underfilled_body_height",
                        details=f"body_height={audit.body_height} expected_min={int(expected_body_height * BODY_HEIGHT_UNDERFILL_RATIO)}",
                    )
                )

        if audit.body_top is not None and audit.body_height is not None and audit.footer_top is not None:
            minimum_bottom_gap = geometry.content_footer_gap_emu - GEOMETRY_TOLERANCE_EMU
            body_bottom = audit.body_top + audit.body_height
            footer_gap = audit.footer_top - body_bottom
            if footer_gap < minimum_bottom_gap:
                violations.append(
                    CapacityViolation(
                        slide_index=audit.slide_index,
                        title=audit.title,
                        rule="content_footer_overlap",
                        details=f"gap={footer_gap} min={minimum_bottom_gap}",
                    )
                )

        if audit.title_bottom is not None and audit.body_top is not None:
            if audit.subtitle_top is not None and audit.subtitle_height is not None:
                title_subtitle_gap = audit.subtitle_top - audit.title_bottom
                minimum_gap = geometry.title_content_gap_emu - GEOMETRY_TOLERANCE_EMU
                if title_subtitle_gap < minimum_gap:
                    violations.append(
                        CapacityViolation(
                            slide_index=audit.slide_index,
                            title=audit.title,
                            rule="title_subtitle_overlap",
                            details=f"gap={title_subtitle_gap} min={minimum_gap}",
                        )
                    )
                subtitle_body_gap = audit.body_top - audit.subtitle_bottom
                if subtitle_body_gap < minimum_gap:
                    violations.append(
                        CapacityViolation(
                            slide_index=audit.slide_index,
                            title=audit.title,
                            rule="subtitle_body_overlap",
                            details=f"gap={subtitle_body_gap} min={minimum_gap}",
                        )
                    )
            else:
                title_body_gap = audit.body_top - audit.title_bottom
                if audit.kind in {SlideKind.TABLE.value, SlideKind.CHART.value}:
                    minimum_gap = geometry.title_content_gap_emu - GEOMETRY_TOLERANCE_EMU
                else:
                    minimum_gap = geometry.title_body_gap_no_subtitle_emu - GEOMETRY_TOLERANCE_EMU
                if title_body_gap < minimum_gap:
                    violations.append(
                        CapacityViolation(
                            slide_index=audit.slide_index,
                            title=audit.title,
                            rule="title_body_overlap",
                            details=f"gap={title_body_gap} min={minimum_gap}",
                        )
                    )

        should_check_body_order = audit.kind == SlideKind.BULLETS.value or (
            audit.kind == SlideKind.TEXT.value and audit.expected_items and any(audit.expected_items)
        )
        if should_check_body_order and audit.expected_items:
            expected = [_normalize_audit_text(item) for item in audit.expected_items if _normalize_audit_text(item)]
            rendered = [_normalize_audit_text(item) for item in audit.rendered_items if _normalize_audit_text(item)]
            if rendered and expected != rendered:
                violations.append(
                    CapacityViolation(
                        slide_index=audit.slide_index,
                        title=audit.title,
                        rule="content_order_mismatch",
                        details=f"expected={expected} rendered={rendered}",
                    )
                )

    for title, items in continuation_groups(audits).items():
        fills = [item.fill_ratio for item in items]
        fill_delta = max(fills) - min(fills)
        tolerance = items[0].profile.continuation_balance_tolerance
        has_group_overflow = any(item.fill_ratio > (item.profile.max_fill_ratio + 0.02) for item in items)
        has_material_underfill = any(
            item.fill_ratio + CONTINUATION_AUDIT_EPSILON
            < max(item.profile.target_fill_ratio - item.profile.continuation_balance_tolerance - CONTINUATION_UNDERFILL_GRACE, 0.0)
            for item in items[1:]
        )
        if fill_delta > tolerance and (has_group_overflow or has_material_underfill):
            violations.append(
                CapacityViolation(
                    slide_index=items[-1].slide_index,
                    title=title,
                    rule="continuation_balance",
                    details=f"delta={fill_delta:.2f} tolerance={tolerance:.2f}",
                )
            )

        for item in items[1:]:
            min_fill = max(
                item.profile.target_fill_ratio
                - item.profile.continuation_balance_tolerance
                - CONTINUATION_UNDERFILL_GRACE,
                0.0,
            )
            if item.fill_ratio + CONTINUATION_AUDIT_EPSILON < min_fill:
                violations.append(
                    CapacityViolation(
                        slide_index=item.slide_index,
                        title=item.title,
                        rule="underfilled_continuation",
                        details=f"fill_ratio={item.fill_ratio:.2f} min={min_fill:.2f}",
                    )
                )
            elif item.fill_ratio > (item.profile.max_fill_ratio + 0.02):
                violations.append(
                    CapacityViolation(
                        slide_index=item.slide_index,
                        title=item.title,
                        rule="overflow_continuation",
                        details=f"fill_ratio={item.fill_ratio:.2f} max={item.profile.max_fill_ratio:.2f}",
                    )
                )

        for previous, current in zip(items, items[1:]):
            previous_font = previous.representative_font_size
            current_font = current.representative_font_size
            if previous_font is None or current_font is None:
                continue
            delta = abs(previous_font - current_font)
            if delta > CONTINUATION_FONT_DELTA_TOLERANCE_PT:
                violations.append(
                    CapacityViolation(
                        slide_index=current.slide_index,
                        title=current.title,
                        rule="continuation_font_delta",
                        details=(
                            f"delta={delta:.2f} tolerance={CONTINUATION_FONT_DELTA_TOLERANCE_PT:.2f} "
                            f"previous={previous_font:.2f} current={current_font:.2f}"
                        ),
                    )
                )

        tracked_items = [item for item in items if item.expected_items]
        expected_group = [
            _normalize_audit_text(entry)
            for item in tracked_items
            for entry in item.expected_items
            if _normalize_audit_text(entry)
        ]
        rendered_group = [
            _normalize_audit_text(entry)
            for item in tracked_items
            for entry in item.rendered_items
            if _normalize_audit_text(entry)
        ]
        if expected_group and rendered_group and expected_group != rendered_group:
            violations.append(
                CapacityViolation(
                    slide_index=items[-1].slide_index,
                    title=title,
                    rule="continuation_order_mismatch",
                    details=f"expected={expected_group} rendered={rendered_group}",
                )
            )

    return violations


def _placeholder_idx_for_role(
    slide_spec: SlideSpec,
    manifest: TemplateManifest | None,
    kind: PlaceholderKind,
    fallback_indices: tuple[int, ...],
) -> int | None:
    placeholder = _placeholder_spec_for_role(slide_spec, manifest, kind)
    if placeholder is not None and placeholder.idx is not None:
        return placeholder.idx
    return fallback_indices[0] if fallback_indices else None


def _placeholder_spec_for_role(
    slide_spec: SlideSpec,
    manifest: TemplateManifest | None,
    kind: PlaceholderKind,
):
    if manifest is None:
        return None
    layout = next(
        (item for item in manifest.layouts if item.key == (slide_spec.preferred_layout_key or "")),
        None,
    )
    if layout is None:
        return None
    typed = [placeholder for placeholder in layout.placeholders if placeholder.kind == kind and placeholder.idx is not None]
    if typed:
        return typed[0]
    return None


def _geometry_policy_for_slide(slide_spec: SlideSpec, manifest: TemplateManifest | None) -> LayoutGeometryPolicy:
    base_layout_key = slide_spec.preferred_layout_key or _infer_layout_key(slide_spec.kind.value)
    base_policy = geometry_policy_for_layout(base_layout_key)
    if manifest is None:
        return base_policy
    layout = next((item for item in manifest.layouts if item.key == base_layout_key), None)
    if layout is None:
        return base_policy
    placeholders: dict[int, PlaceholderGeometryPolicy] = {}
    for placeholder in layout.placeholders:
        if placeholder.idx is None:
            continue
        if None in {placeholder.left_emu, placeholder.top_emu, placeholder.width_emu, placeholder.height_emu}:
            continue
        placeholders[placeholder.idx] = PlaceholderGeometryPolicy(
            placeholder_idx=placeholder.idx,
            left_emu=placeholder.left_emu,
            top_emu=placeholder.top_emu,
            width_emu=placeholder.width_emu,
            height_emu=placeholder.height_emu,
        )
    if not placeholders:
        return base_policy
    return LayoutGeometryPolicy(
        layout_key=base_layout_key,
        placeholders=placeholders,
        title_content_gap_emu=base_policy.title_content_gap_emu,
        title_body_gap_no_subtitle_emu=base_policy.title_body_gap_no_subtitle_emu,
        content_footer_gap_emu=base_policy.content_footer_gap_emu,
    )


def _infer_layout_key(kind: str) -> str:
    if kind == SlideKind.BULLETS.value:
        return "list_full_width"
    if kind == SlideKind.IMAGE.value:
        return "image_text"
    return "text_full_width"


def self_has_image(shape) -> bool:
    return hasattr(shape, "image") or "Picture Placeholder" in (getattr(shape, "name", "") or "")


def _expected_items_for_slide(slide_spec: SlideSpec) -> tuple[str, ...]:
    if slide_spec.content_blocks:
        has_list_like_block = any(block.items for block in slide_spec.content_blocks)
        if slide_spec.kind == SlideKind.TEXT and not has_list_like_block:
            return ()
        expected: list[str] = []
        for block in slide_spec.content_blocks:
            if block.text and block.text.strip():
                expected.append(block.text.strip())
            expected.extend(item.strip() for item in block.items if item.strip())
        return tuple(expected)
    if slide_spec.kind == SlideKind.BULLETS:
        return tuple(item.strip() for item in slide_spec.bullets if item.strip())
    return ()


def _normalize_audit_text(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip())
