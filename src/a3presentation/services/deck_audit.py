from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from pptx import Presentation
from a3presentation.domain.presentation import PresentationPlan, SlideKind, SlideSpec
from a3presentation.services.layout_capacity import LayoutCapacityProfile, profile_for_layout
from a3presentation.services.pptx_generator import PptxGenerator


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


@dataclass(frozen=True)
class CapacityViolation:
    slide_index: int
    title: str
    rule: str
    details: str


def audit_generated_presentation(output_path: Path, plan: PresentationPlan) -> list[SlideAudit]:
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
        body = placeholders.get(14)
        footer = placeholders.get(15) or placeholders.get(17)
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
        if content_width is None:
            if has_table:
                table_shape = next((shape for shape in slide.shapes if getattr(shape, "has_table", False)), None)
                content_width = getattr(table_shape, "width", None) if table_shape is not None else None
            elif has_chart:
                chart_shape = next((shape for shape in slide.shapes if getattr(shape, "has_chart", False)), None)
                content_width = getattr(chart_shape, "width", None) if chart_shape is not None else None
            elif has_image:
                image_shape = next(
                    (shape for shape in slide.shapes if self_has_image(shape)),
                    None,
                )
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

        if audit.kind == SlideKind.BULLETS.value and audit.expected_items:
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
        if fill_delta > tolerance:
            violations.append(
                CapacityViolation(
                    slide_index=items[-1].slide_index,
                    title=title,
                    rule="continuation_balance",
                    details=f"delta={fill_delta:.2f} tolerance={tolerance:.2f}",
                )
            )

        for item in items[1:]:
            min_fill = max(item.profile.target_fill_ratio - item.profile.continuation_balance_tolerance, 0.0)
            if item.fill_ratio < min_fill:
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

        expected_group = [
            _normalize_audit_text(entry)
            for item in items
            for entry in item.expected_items
            if _normalize_audit_text(entry)
        ]
        rendered_group = [
            _normalize_audit_text(entry)
            for item in items
            for entry in item.rendered_items
            if _normalize_audit_text(entry)
        ]
        if rendered_group and expected_group != rendered_group:
            violations.append(
                CapacityViolation(
                    slide_index=items[-1].slide_index,
                    title=title,
                    rule="continuation_order_mismatch",
                    details=f"expected={expected_group} rendered={rendered_group}",
                )
            )

    return violations


def _infer_layout_key(kind: str) -> str:
    if kind == SlideKind.BULLETS.value:
        return "list_full_width"
    if kind == SlideKind.IMAGE.value:
        return "image_text"
    return "text_full_width"


def self_has_image(shape) -> bool:
    return hasattr(shape, "image") or "Picture Placeholder" in (getattr(shape, "name", "") or "")


def _expected_items_for_slide(slide_spec: SlideSpec) -> tuple[str, ...]:
    if slide_spec.kind == SlideKind.BULLETS:
        return tuple(item.strip() for item in slide_spec.bullets if item.strip())
    if slide_spec.kind == SlideKind.TEXT:
        parts = [part.strip() for part in (slide_spec.text or "", slide_spec.notes or "") if part and part.strip()]
        return tuple(parts)
    return ()


def _normalize_audit_text(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip())
