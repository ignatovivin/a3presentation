from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from a3presentation.domain.diagnostics import (
    GenerationDiagnosticRule,
    GenerationDiagnosticSeverity,
    GenerationDiagnosticSource,
    GenerationErrorCode,
)
from a3presentation.domain.presentation import PresentationPlan, SlideKind
from a3presentation.domain.template import TemplateManifest
from a3presentation.services.deck_audit import audit_generated_presentation, find_capacity_violations
from a3presentation.services.pptx_generator import PptxGenerator
from a3presentation.services.slide_text_policy import slide_text_demand_chars
from a3presentation.services.style_audit import audit_presentation_styles
from a3presentation.services.template_registry import TemplateRegistry


@dataclass(frozen=True)
class GenerationDiagnosticItem:
    slide_index: int
    title: str
    severity: GenerationDiagnosticSeverity
    rule: GenerationDiagnosticRule
    details: str
    source: GenerationDiagnosticSource


@dataclass(frozen=True)
class GeneratedPresentationResult:
    output_path: Path
    warnings: list[str]
    diagnostics: list[GenerationDiagnosticItem] | None = None
    attempt_count: int = 1


class PresentationGenerationError(ValueError):
    def __init__(self, code: GenerationErrorCode, message: str, *, attempt_count: int = 0) -> None:
        super().__init__(message)
        self.code = code
        self.message = message
        self.attempt_count = attempt_count


class PresentationGenerationService:
    SEVERITY_RANK = {"warning": 1, "retryable": 2, "blocking": 3}
    RETRY_RULES = {"overflow_risk", "font_bounds", "rendered_text_overflow"}
    BLOCKING_RULES = {
        "table_overlay_text_overflow",
        "missing_table_shape",
        "missing_chart_shape",
        "missing_image_shape",
        "unexpected_table_shape",
        "unexpected_chart_shape",
        "two_column_overlap",
        "image_text_overlap",
        "chart_type_mismatch",
        "chart_series_count_mismatch",
        "rendered_text_overflow",
    }

    def __init__(
        self,
        *,
        generator: PptxGenerator,
        template_registry: TemplateRegistry,
        output_dir: Path,
        max_audit_attempts: int = 2,
    ) -> None:
        self.generator = generator
        self.template_registry = template_registry
        self.output_dir = output_dir
        self.max_audit_attempts = max_audit_attempts

    def generate_checked_presentation(
        self,
        plan: PresentationPlan,
        manifest: TemplateManifest,
        template_path: Path,
    ) -> Path:
        return self.generate_checked_result(plan, manifest, template_path).output_path

    def generate_checked_result(
        self,
        plan: PresentationPlan,
        manifest: TemplateManifest,
        template_path: Path,
    ) -> GeneratedPresentationResult:
        active_plan = plan
        last_output_path: Path | None = None
        last_violations = []
        attempt_count = 0
        for attempt in range(self.max_audit_attempts):
            attempt_count = attempt + 1
            output_path = self.generator.generate(
                template_path=template_path,
                manifest=manifest,
                plan=active_plan,
                output_dir=self.output_dir,
            )
            last_output_path = output_path
            audits = audit_generated_presentation(output_path, active_plan, manifest)
            violations = find_capacity_violations(audits)
            blocking_violations = self.blocking_generation_violations(violations)
            last_violations = blocking_violations
            if not blocking_violations:
                retry_plan = self.plan_with_capacity_retry(active_plan, manifest, audits, violations)
                if retry_plan is not None and attempt + 1 < self.max_audit_attempts:
                    active_plan = retry_plan
                    continue
                diagnostics = self.generation_diagnostics(output_path, active_plan, manifest, violations)
                return GeneratedPresentationResult(
                    output_path=output_path,
                    warnings=self.warning_messages(diagnostics),
                    diagnostics=diagnostics,
                    attempt_count=attempt_count,
                )
            if attempt + 1 < self.max_audit_attempts:
                retry_plan = self.plan_with_capacity_retry(active_plan, manifest, audits, violations)
                if retry_plan is not None:
                    active_plan = retry_plan
                    continue
            break

        if last_violations:
            details = "; ".join(f"slide {item.slide_index}: {item.rule}" for item in last_violations[:6])
            code: GenerationErrorCode = "retry_exhausted" if attempt_count > 1 else "blocking_quality_gate"
            raise PresentationGenerationError(
                code,
                f"Generated deck failed layout quality gate: {details}",
                attempt_count=attempt_count,
            )
        if last_output_path is not None:
            diagnostics = self.generation_diagnostics(last_output_path, active_plan, manifest, last_violations)
            return GeneratedPresentationResult(
                output_path=last_output_path,
                warnings=self.warning_messages(diagnostics),
                diagnostics=diagnostics,
                attempt_count=attempt_count,
            )
        raise PresentationGenerationError("generation_no_output", "generation produced no output", attempt_count=attempt_count)

    def plan_with_capacity_retry(
        self,
        plan: PresentationPlan,
        manifest: TemplateManifest,
        audits,
        violations,
    ) -> PresentationPlan | None:
        retry_slide_indexes = {item.slide_index for item in violations if item.rule in self.RETRY_RULES}
        if not retry_slide_indexes:
            return None

        audit_by_slide_index = {audit.slide_index: audit for audit in audits}
        retried_slides = []
        changed = False
        for plan_index, slide in enumerate(plan.slides, start=1):
            audit = audit_by_slide_index.get(plan_index)
            if plan_index not in retry_slide_indexes or audit is None or slide.kind not in {
                SlideKind.TEXT,
                SlideKind.BULLETS,
                SlideKind.TWO_COLUMN,
            }:
                retried_slides.append(slide)
                continue

            retry_capacity = max(180, int(audit.profile.max_chars * 0.9))
            split_slides = self.template_registry.split_slide_for_capacity_retry(
                manifest,
                slide,
                retry_capacity,
                overflow_details=getattr(audit, "body_text_overflow_details", ()),
            )
            if len(split_slides) > 1:
                retried_slides.extend(split_slides)
                changed = True
            else:
                retried_slides.append(slide)

        if not changed:
            return None
        return plan.model_copy(update={"slides": retried_slides}, deep=True)

    def blocking_generation_violations(self, violations):
        return [item for item in violations if item.rule in self.BLOCKING_RULES]

    def preflight_diagnostics(
        self,
        plan: PresentationPlan,
        manifest: TemplateManifest,
    ) -> list[GenerationDiagnosticItem]:
        diagnostics: list[GenerationDiagnosticItem] = []
        reviews = self.template_registry.build_slide_layout_reviews(manifest, plan)
        for slide_index, slide in enumerate(plan.slides, start=1):
            review = reviews[slide_index - 1] if slide_index - 1 < len(reviews) else None
            title = slide.title or ""
            if review is not None:
                for reason in review.current_target_degradation_reasons:
                    diagnostics.append(
                        GenerationDiagnosticItem(
                            slide_index=slide_index,
                            title=title,
                            severity="retryable" if reason.startswith("capacity_retry:") else "warning",
                            rule=self._preflight_reason_rule(reason),
                            details=reason,
                            source="capacity",
                        )
                    )

            capacity = self._preflight_capacity_for_slide(review)
            demand = slide_text_demand_chars(slide)
            if capacity is not None and demand > int(capacity * 1.05):
                diagnostics.append(
                    GenerationDiagnosticItem(
                        slide_index=slide_index,
                        title=title,
                        severity="retryable",
                        rule="overflow_risk",
                        details=f"text_chars={demand} estimated_capacity={capacity}",
                        source="capacity",
                    )
                )
            missing_roles = self._missing_required_editable_roles(slide, review)
            if missing_roles:
                diagnostics.append(
                    GenerationDiagnosticItem(
                        slide_index=slide_index,
                        title=title,
                        severity="warning",
                        rule="missing_required_editable_slot",
                        details=f"missing_roles={','.join(missing_roles)}",
                        source="capacity",
                    )
                )
        return self.deduplicate_diagnostics(diagnostics)

    def generation_diagnostics(
        self,
        output_path: Path,
        plan: PresentationPlan,
        manifest: TemplateManifest,
        violations,
    ) -> list[GenerationDiagnosticItem]:
        return self.deduplicate_diagnostics([
            *self.capacity_diagnostics(violations),
            *self.style_diagnostics(output_path, plan, manifest),
        ])

    def capacity_diagnostics(self, violations) -> list[GenerationDiagnosticItem]:
        diagnostics: list[GenerationDiagnosticItem] = []
        for item in violations:
            if item.rule in self.BLOCKING_RULES:
                severity = "blocking"
            elif item.rule in self.RETRY_RULES:
                severity = "retryable"
            else:
                severity = "warning"
            diagnostics.append(
                GenerationDiagnosticItem(
                    slide_index=item.slide_index,
                    title=item.title,
                    severity=severity,
                    rule=item.rule,
                    details=item.details,
                    source="capacity",
                )
            )
        return diagnostics

    def style_diagnostics(
        self,
        output_path: Path,
        plan: PresentationPlan,
        manifest: TemplateManifest,
    ) -> list[GenerationDiagnosticItem]:
        return [
            GenerationDiagnosticItem(
                slide_index=item.slide_index,
                title=item.title,
                severity="warning",
                rule=item.rule,
                details=item.details,
                source="style",
            )
            for item in audit_presentation_styles(output_path, plan, manifest)
        ]

    def warning_messages(self, diagnostics: list[GenerationDiagnosticItem]) -> list[str]:
        return [
            f"slide {item.slide_index}: {item.rule}: {item.details}"
            for item in diagnostics
            if item.severity == "warning"
        ]

    def style_warning_messages(
        self,
        output_path: Path,
        plan: PresentationPlan,
        manifest: TemplateManifest,
    ) -> list[str]:
        return self.warning_messages(self.style_diagnostics(output_path, plan, manifest))

    def deduplicate_diagnostics(
        self,
        diagnostics: list[GenerationDiagnosticItem],
    ) -> list[GenerationDiagnosticItem]:
        deduped: dict[tuple[int, str, str], GenerationDiagnosticItem] = {}
        order: list[tuple[int, str, str]] = []
        for item in diagnostics:
            key = (item.slide_index, item.source, item.rule)
            existing = deduped.get(key)
            if existing is None:
                deduped[key] = item
                order.append(key)
                continue

            severity = self._stronger_severity(existing.severity, item.severity)
            details = self._merge_diagnostic_details(existing.details, item.details)
            title = existing.title or item.title
            deduped[key] = GenerationDiagnosticItem(
                slide_index=existing.slide_index,
                title=title,
                severity=severity,
                rule=existing.rule,
                details=details,
                source=existing.source,
            )
        return [deduped[key] for key in order]

    def _stronger_severity(
        self,
        left: GenerationDiagnosticSeverity,
        right: GenerationDiagnosticSeverity,
    ) -> GenerationDiagnosticSeverity:
        return left if self.SEVERITY_RANK.get(left, 0) >= self.SEVERITY_RANK.get(right, 0) else right

    def _merge_diagnostic_details(self, left: str, right: str) -> str:
        parts = []
        for value in (left, right):
            text = value.strip()
            if text and text not in parts:
                parts.append(text)
        return "; ".join(parts)

    def _preflight_capacity_for_slide(self, review) -> int | None:
        if review is None:
            return None
        current_key = review.current_target_key
        current_option = next((item for item in review.available_layouts if item.key == current_key), None)
        if current_option is not None:
            return current_option.estimated_text_capacity_chars
        if review.available_layouts:
            return review.available_layouts[0].estimated_text_capacity_chars
        return None

    def _missing_required_editable_roles(self, slide, review) -> list[str]:
        required_roles = self._required_editable_roles_for_slide(slide)
        if not required_roles or review is None:
            return []
        current_option = self._current_layout_option(review)
        if current_option is None:
            return required_roles
        available_roles = set(current_option.editable_roles)
        return [role for role in required_roles if role not in available_roles]

    def _current_layout_option(self, review):
        current_key = review.current_target_key
        if current_key:
            current_option = next((item for item in review.available_layouts if item.key == current_key), None)
            if current_option is not None:
                return current_option
        return review.available_layouts[0] if review.available_layouts else None

    def _required_editable_roles_for_slide(self, slide) -> list[str]:
        if slide.kind == SlideKind.TITLE:
            return ["title"]
        if slide.kind in {SlideKind.TEXT, SlideKind.BULLETS, SlideKind.TWO_COLUMN}:
            return ["body"]
        if slide.kind == SlideKind.TABLE:
            return ["table"]
        if slide.kind == SlideKind.CHART:
            return ["chart"]
        if slide.kind == SlideKind.IMAGE:
            return ["image"]
        return []

    def _preflight_reason_rule(self, reason: str) -> str:
        if reason.startswith("capacity_retry:"):
            parts = reason.split(":")
            if len(parts) >= 2 and parts[1]:
                return parts[1]
            return "capacity_retry"
        if reason.startswith("split_for_capacity"):
            return "split_for_capacity"
        if reason.startswith("split_overflow_slot"):
            return "split_overflow_slot"
        return "layout_degradation"
