from __future__ import annotations

import argparse
import json
import traceback
from collections import Counter
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from a3presentation.domain.api import ChartOverride, DocumentBlock
from a3presentation.domain.presentation import (
    PresentationPlan,
    RenderTargetType,
    SlideContentBlock,
    SlideContentBlockKind,
    SlideKind,
    SlideRenderTarget,
    SlideSpec,
    TableBlock,
)
from a3presentation.services.deck_audit import audit_generated_presentation, find_capacity_violations
from a3presentation.services.document_text_extractor import DocumentTextExtractor
from a3presentation.services.planner import TextToPlanService
from a3presentation.services.pptx_generator import PptxGenerator
from a3presentation.services.table_chart_analyzer import TableChartAnalyzer
from a3presentation.services.template_analyzer import TemplateAnalyzer
from a3presentation.services.template_registry import TemplateRegistry


BLOCKING_RULES = {
    "table_overlay_text_overflow",
    "missing_table_shape",
    "missing_chart_shape",
    "missing_image_shape",
    "unexpected_table_shape",
    "unexpected_chart_shape",
    "content_order_mismatch",
    "two_column_overlap",
    "image_text_overlap",
    "chart_type_mismatch",
    "chart_series_count_mismatch",
    "combo_chart_structure_mismatch",
    "missing_secondary_value_axis",
}


@dataclass(frozen=True)
class InputDocument:
    path: Path
    raw_text: str
    tables: list[Any]
    blocks: list[Any]
    chart_overrides: list[ChartOverride] | None = None
    case_kind: str = "file"


def main() -> int:
    parser = argparse.ArgumentParser(description="Run local self-check corpus across uploaded templates and inputs.")
    parser.add_argument("--templates-dir", type=Path, default=Path("examples/selfcheck/templates"))
    parser.add_argument("--inputs-dir", type=Path, default=Path("examples/selfcheck/inputs"))
    parser.add_argument("--out-dir", type=Path, default=Path("run-logs/selfcheck"))
    parser.add_argument("--max-options-per-slide", type=int, default=3)
    parser.add_argument("--max-variants-per-case", type=int, default=40)
    parser.add_argument("--include-synthetic-inputs", action="store_true")
    parser.add_argument("--include-stress-variants", action="store_true")
    parser.add_argument("--stop-after-first-pass", action="store_true")
    args = parser.parse_args()

    templates = sorted(args.templates_dir.glob("*.pptx"))
    inputs = sorted([*args.inputs_dir.glob("*.docx"), *args.inputs_dir.glob("*.txt"), *args.inputs_dir.glob("*.md")])
    args.out_dir.mkdir(parents=True, exist_ok=True)
    output_pptx_dir = args.out_dir / "pptx"
    output_pptx_dir.mkdir(parents=True, exist_ok=True)

    extractor = DocumentTextExtractor()
    analyzer = TemplateAnalyzer()
    registry = TemplateRegistry(Path("storage/templates"))
    planner = TextToPlanService()
    generator = PptxGenerator()

    loaded_inputs = [_load_input(path, extractor) for path in inputs]
    if args.include_synthetic_inputs:
        loaded_inputs.extend(_synthetic_inputs())
    results: list[dict[str, Any]] = []

    for template_path in templates:
        manifest_result = _analyze_template(template_path, analyzer, registry)
        if manifest_result["status"] != "ok":
            for input_doc in loaded_inputs:
                results.append({
                    "template": str(template_path),
                    "input": str(input_doc.path),
                    "stage": "template_analysis",
                    **manifest_result,
                })
            continue
        manifest = manifest_result["manifest"]

        for input_doc in loaded_inputs:
            case_results = _run_case(
                template_path=template_path,
                input_doc=input_doc,
                manifest=manifest,
                registry=registry,
                planner=planner,
                generator=generator,
                output_pptx_dir=output_pptx_dir,
                max_options_per_slide=max(args.max_options_per_slide, 1),
                max_variants=max(args.max_variants_per_case, 1),
                include_stress_variants=args.include_stress_variants,
                stop_after_first_pass=args.stop_after_first_pass,
            )
            results.extend(case_results)

    report = {
        "generated_at": datetime.now(UTC).isoformat(),
        "templates_dir": str(args.templates_dir),
        "inputs_dir": str(args.inputs_dir),
        "templates": [str(path) for path in templates],
        "inputs": [str(input_doc.path) for input_doc in loaded_inputs],
        "summary": _summarize(results),
        "results": results,
    }
    json_path = args.out_dir / "selfcheck-report.json"
    md_path = args.out_dir / "selfcheck-report.md"
    json_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    md_path.write_text(_render_markdown(report), encoding="utf-8")
    print(f"wrote {json_path}")
    print(f"wrote {md_path}")
    print(json.dumps(report["summary"], ensure_ascii=False, indent=2))
    return 0 if report["summary"]["failed_cases"] == 0 else 1


def _load_input(path: Path, extractor: DocumentTextExtractor) -> InputDocument:
    if path.suffix.lower() == ".docx":
        raw_text, tables, blocks = extractor.extract(path.name, path.read_bytes())
        return InputDocument(path=path, raw_text=raw_text, tables=tables, blocks=blocks)
    return InputDocument(path=path, raw_text=path.read_text(encoding="utf-8"), tables=[], blocks=[])


def _synthetic_inputs() -> list[InputDocument]:
    chart_table = TableBlock(
        headers=["Квартал", "Выручка", "Заявки", "Конверсия"],
        rows=[
            ["Q1 2026", "120", "48", "18"],
            ["Q2 2026", "156", "63", "21"],
            ["Q3 2026", "181", "72", "24"],
            ["Q4 2026", "214", "89", "27"],
        ],
    )
    text_table = TableBlock(
        headers=["Риск", "Причина", "Действие"],
        rows=[
            ["Сроки", "Зависимость от подрядчика", "Еженедельный контроль"],
            ["Качество", "Разные форматы исходников", "Нормализация перед генерацией"],
            ["Данные", "Неполные поля", "Проверка обязательных блоков"],
        ],
    )
    long_paragraph = (
        "Команда должна сохранить визуальный стиль шаблона: размеры текста, внутренние и внешние отступы, "
        "расположение элементов, структуру заголовков и пропорции контентных областей. При этом количество "
        "слайдов не ограничивается исходным шаблоном, потому что длинный материал обязан раскладываться на "
        "дополнительные слайды без сжатия до нечитаемого состояния."
    )
    short_text = "\n".join([
        "# Короткий отчет",
        "Цель: проверить генерацию на малом объеме текста.",
        "- Один ключевой тезис",
        "- Второй короткий тезис",
        "- Риск и следующий шаг",
    ])
    long_text = "\n\n".join(
        [f"## Раздел {index}\n{long_paragraph}\n\n- Контроль качества {index}\n- Проверка верстки {index}\n- Резервный сценарий {index}" for index in range(1, 11)]
    )
    table_only_text = "# Только таблицы\nПроверка сценария без нарратива, где основной контент должен стать табличными слайдами."
    chart_only_text = "# Только графики\nПроверка сценария, где числовая таблица должна быть показана как график."
    mixed_text = "\n\n".join([
        "# Смешанный отчет",
        "Нужно проверить текст, списки, таблицы и графики в одном прогоне.",
        "## Итоги",
        long_paragraph,
        "- Выручка растет",
        "- Конверсия улучшается",
        "- Риски требуют отдельного контроля",
    ])

    chart_override = _chart_override(chart_table, "table_1")
    return [
        InputDocument(path=Path("synthetic/short_text.md"), raw_text=short_text, tables=[], blocks=_text_blocks(short_text), case_kind="synthetic_short_text"),
        InputDocument(path=Path("synthetic/long_text.md"), raw_text=long_text, tables=[], blocks=_text_blocks(long_text), case_kind="synthetic_long_text"),
        InputDocument(path=Path("synthetic/table_only.md"), raw_text=table_only_text, tables=[text_table, chart_table], blocks=_table_blocks(table_only_text, [text_table, chart_table]), case_kind="synthetic_table_only"),
        InputDocument(path=Path("synthetic/chart_only.md"), raw_text=chart_only_text, tables=[chart_table], blocks=_table_blocks(chart_only_text, [chart_table]), chart_overrides=[chart_override] if chart_override else [], case_kind="synthetic_chart_only"),
        InputDocument(path=Path("synthetic/mixed_text_table_chart.md"), raw_text=mixed_text, tables=[text_table, chart_table], blocks=[*_text_blocks(mixed_text), *_table_blocks("", [text_table, chart_table])], chart_overrides=[chart_override] if chart_override else [], case_kind="synthetic_mixed"),
    ]


def _text_blocks(raw_text: str) -> list[DocumentBlock]:
    blocks: list[DocumentBlock] = []
    for line in raw_text.splitlines():
        value = line.strip()
        if not value:
            continue
        if value.startswith("#"):
            blocks.append(DocumentBlock(kind="heading", text=value.lstrip("# "), level=value.count("#")))
        elif value.startswith("- "):
            blocks.append(DocumentBlock(kind="list", items=[value[2:].strip()]))
        else:
            blocks.append(DocumentBlock(kind="paragraph", text=value))
    return blocks


def _table_blocks(raw_text: str, tables: list[TableBlock]) -> list[DocumentBlock]:
    blocks = _text_blocks(raw_text)
    blocks.extend(DocumentBlock(kind="table", table=table) for table in tables)
    return blocks


def _chart_override(table: TableBlock, table_id: str) -> ChartOverride | None:
    assessment = TableChartAnalyzer().analyze(table, table_id=table_id)
    if not assessment.candidate_specs:
        return None
    chart = assessment.candidate_specs[0].model_copy(deep=True)
    chart.title = chart.title or "Динамика показателей"
    return ChartOverride(table_id=table_id, mode="chart", selected_chart=chart)


def _analyze_template(template_path: Path, analyzer: TemplateAnalyzer, registry: TemplateRegistry) -> dict[str, Any]:
    try:
        manifest = analyzer.analyze(
            template_id=f"selfcheck_{template_path.stem}",
            template_path=template_path,
            display_name=template_path.stem,
        )
        return {"status": "ok", "manifest": registry.normalize_manifest(manifest)}
    except Exception as exc:
        return {
            "status": "error",
            "error_type": type(exc).__name__,
            "error": str(exc),
            "traceback": traceback.format_exc(limit=6),
        }


def _run_case(
    *,
    template_path: Path,
    input_doc: InputDocument,
    manifest: Any,
    registry: TemplateRegistry,
    planner: TextToPlanService,
    generator: PptxGenerator,
    output_pptx_dir: Path,
    max_options_per_slide: int,
    max_variants: int,
    include_stress_variants: bool,
    stop_after_first_pass: bool,
) -> list[dict[str, Any]]:
    try:
        base_plan = planner.build_plan(
            template_id=manifest.template_id,
            raw_text=input_doc.raw_text,
            title=input_doc.path.stem,
            tables=input_doc.tables,
            blocks=input_doc.blocks,
            chart_overrides=input_doc.chart_overrides,
        )
        base_plan = registry.apply_layout_inventory_to_plan(manifest, base_plan)
        reviews = registry.build_slide_layout_reviews(manifest, base_plan)
    except Exception as exc:
        return [{
            "template": str(template_path),
            "input": str(input_doc.path),
            "case_kind": input_doc.case_kind,
            "variant": "plan",
            "stage": "planning",
            "status": "error",
            "error_type": type(exc).__name__,
            "error": str(exc),
            "traceback": traceback.format_exc(limit=6),
        }]

    variants = _build_variants(
        base_plan,
        reviews,
        max_options_per_slide=max_options_per_slide,
        max_variants=max_variants,
        include_stress_variants=include_stress_variants,
    )
    results = []
    for variant_name, plan in variants:
        result = _generate_and_audit(
            template_path=template_path,
            input_path=input_doc.path,
            case_kind=input_doc.case_kind,
            manifest=manifest,
            plan=plan,
            generator=generator,
            output_pptx_dir=output_pptx_dir,
            variant_name=variant_name,
        )
        results.append(result)
        if stop_after_first_pass and result["status"] == "pass":
            break
    return results


def _build_variants(
    base_plan: PresentationPlan,
    reviews: list[Any],
    *,
    max_options_per_slide: int,
    max_variants: int,
    include_stress_variants: bool,
) -> list[tuple[str, PresentationPlan]]:
    variants: list[tuple[str, PresentationPlan]] = [("baseline", base_plan)]
    if include_stress_variants:
        for name, plan in _stress_variants(base_plan):
            variants.append((name, plan))
            if len(variants) >= max_variants:
                return variants
    for review in reviews:
        current_key = review.current_target_key
        for option_index, option in enumerate(review.available_layouts[:max_options_per_slide], start=1):
            if option.key == current_key:
                continue
            plan = base_plan.model_copy(deep=True)
            slide = plan.slides[review.slide_index]
            slide.preferred_layout_key = option.key
            slide.runtime_profile_key = option.runtime_profile_key
            slide.render_target = SlideRenderTarget(
                type=RenderTargetType(option.source),
                key=option.key,
                label=option.name,
                source=option.source,
                confidence=option.recommendation_label,
            )
            variants.append((f"slide_{review.slide_index}_option_{option_index}_{option.key}", plan))
            if len(variants) >= max_variants:
                return variants
    return variants


def _generate_and_audit(
    *,
    template_path: Path,
    input_path: Path,
    case_kind: str,
    manifest: Any,
    plan: PresentationPlan,
    generator: PptxGenerator,
    output_pptx_dir: Path,
    variant_name: str,
) -> dict[str, Any]:
    case_id = _safe_name(f"{template_path.stem}__{input_path.stem}__{variant_name}")
    try:
        output_path = generator.generate(
            template_path=template_path,
            manifest=manifest,
            plan=plan,
            output_dir=output_pptx_dir,
        )
        audits = audit_generated_presentation(output_path, plan, manifest)
        violations = find_capacity_violations(audits)
        blocking = [item for item in violations if item.rule in BLOCKING_RULES]
        return {
            "template": str(template_path),
            "input": str(input_path),
            "case_kind": case_kind,
            "variant": variant_name,
            "stage": "audit" if blocking else "complete",
            "status": "fail" if blocking else "pass",
            "output": str(output_path),
            "slides": len(plan.slides),
            "rules": [item.rule for item in blocking],
            "all_rules": [item.rule for item in violations],
            "targets": _targets_for_plan(plan),
            "audits": [_audit_summary(item) for item in audits],
        }
    except Exception as exc:
        return {
            "template": str(template_path),
            "input": str(input_path),
            "case_kind": case_kind,
            "variant": variant_name,
            "stage": "generation",
            "status": "error",
            "error_type": type(exc).__name__,
            "error": str(exc),
            "traceback": traceback.format_exc(limit=8),
            "case_id": case_id,
            "targets": _targets_for_plan(plan),
        }


def _stress_variants(base_plan: PresentationPlan) -> list[tuple[str, PresentationPlan]]:
    variants: list[tuple[str, PresentationPlan]] = []
    for name, builder in [
        ("stress_more_slides", _variant_more_slides),
        ("stress_fewer_slides", _variant_fewer_slides),
        ("stress_text_as_bullets", _variant_text_as_bullets),
        ("stress_bullets_as_text", _variant_bullets_as_text),
        ("stress_tables_only", _variant_tables_only),
        ("stress_charts_only", _variant_charts_only),
    ]:
        plan = builder(base_plan)
        if plan is not None:
            variants.append((name, plan))
    return variants


def _variant_more_slides(base_plan: PresentationPlan) -> PresentationPlan | None:
    plan = base_plan.model_copy(deep=True)
    for index, slide in enumerate(plan.slides):
        payload = _slide_text_payload(slide)
        if len(payload) < 360:
            continue
        midpoint = max(1, len(payload.split()) // 2)
        words = payload.split()
        first = " ".join(words[:midpoint])
        second = " ".join(words[midpoint:])
        slide.kind = SlideKind.TEXT
        slide.text = first
        slide.bullets = []
        slide.content_blocks = [SlideContentBlock(kind=SlideContentBlockKind.PARAGRAPH, text=first)]
        clone = slide.model_copy(deep=True)
        clone.title = f"{slide.title or 'Продолжение'} - продолжение"
        clone.text = second
        clone.content_blocks = [SlideContentBlock(kind=SlideContentBlockKind.PARAGRAPH, text=second)]
        plan.slides.insert(index + 1, clone)
        return plan
    return None


def _variant_fewer_slides(base_plan: PresentationPlan) -> PresentationPlan | None:
    plan = base_plan.model_copy(deep=True)
    for index in range(1, len(plan.slides) - 1):
        first = plan.slides[index]
        second = plan.slides[index + 1]
        if first.kind in {SlideKind.TABLE, SlideKind.CHART, SlideKind.IMAGE, SlideKind.TITLE}:
            continue
        if second.kind in {SlideKind.TABLE, SlideKind.CHART, SlideKind.IMAGE, SlideKind.TITLE}:
            continue
        merged = "\n\n".join(part for part in [_slide_text_payload(first), _slide_text_payload(second)] if part)
        if not merged:
            continue
        first.kind = SlideKind.TEXT
        first.text = merged
        first.bullets = []
        first.content_blocks = [SlideContentBlock(kind=SlideContentBlockKind.PARAGRAPH, text=merged)]
        del plan.slides[index + 1]
        return plan
    return None


def _variant_text_as_bullets(base_plan: PresentationPlan) -> PresentationPlan | None:
    plan = base_plan.model_copy(deep=True)
    for slide in plan.slides:
        payload = _slide_text_payload(slide)
        if slide.kind != SlideKind.TEXT or len(payload) < 120:
            continue
        bullets = [item.strip() for item in payload.replace("\n", " ").split(".") if len(item.strip()) > 20][:6]
        if len(bullets) < 2:
            continue
        slide.kind = SlideKind.BULLETS
        slide.text = None
        slide.bullets = bullets
        slide.content_blocks = [SlideContentBlock(kind=SlideContentBlockKind.BULLET_LIST, items=bullets)]
        return plan
    return None


def _variant_bullets_as_text(base_plan: PresentationPlan) -> PresentationPlan | None:
    plan = base_plan.model_copy(deep=True)
    for slide in plan.slides:
        if slide.kind != SlideKind.BULLETS or len(slide.bullets) < 2:
            continue
        text = ". ".join(slide.bullets) + "."
        slide.kind = SlideKind.TEXT
        slide.text = text
        slide.bullets = []
        slide.content_blocks = [SlideContentBlock(kind=SlideContentBlockKind.PARAGRAPH, text=text)]
        return plan
    return None


def _variant_tables_only(base_plan: PresentationPlan) -> PresentationPlan | None:
    plan = base_plan.model_copy(deep=True)
    media = [slide for slide in plan.slides if slide.kind == SlideKind.TABLE and slide.table is not None]
    if not media:
        return None
    plan.slides = [plan.slides[0], *[slide.model_copy(deep=True) for slide in media]]
    return plan


def _variant_charts_only(base_plan: PresentationPlan) -> PresentationPlan | None:
    plan = base_plan.model_copy(deep=True)
    media = [slide for slide in plan.slides if slide.kind == SlideKind.CHART and slide.chart is not None]
    if not media:
        return None
    plan.slides = [plan.slides[0], *[slide.model_copy(deep=True) for slide in media]]
    return plan


def _slide_text_payload(slide: SlideSpec) -> str:
    parts = [slide.text or "", *slide.bullets, *slide.left_bullets, *slide.right_bullets]
    for block in slide.content_blocks:
        if block.text:
            parts.append(block.text)
        parts.extend(block.items)
    return "\n".join(part.strip() for part in parts if part and part.strip())


def _targets_for_plan(plan: PresentationPlan) -> list[dict[str, Any]]:
    return [
        {
            "slide": index,
            "kind": slide.kind.value,
            "title": slide.title,
            "preferred_layout_key": slide.preferred_layout_key,
            "runtime_profile_key": slide.runtime_profile_key,
            "target_type": slide.render_target.type.value if slide.render_target is not None else None,
            "target_key": slide.render_target.key if slide.render_target is not None else None,
            "target_source": slide.render_target.source if slide.render_target is not None else None,
        }
        for index, slide in enumerate(plan.slides, start=1)
    ]


def _audit_summary(audit: Any) -> dict[str, Any]:
    return {
        "slide": audit.slide_index,
        "kind": audit.kind,
        "title": audit.title,
        "has_table": audit.has_table,
        "has_chart": audit.has_chart,
        "has_image": audit.has_image,
        "target_type": audit.target_type,
        "layout_key": audit.layout_key,
        "runtime_profile_key": audit.runtime_profile_key,
        "fill_ratio": round(audit.fill_ratio, 3),
    }


def _summarize(results: list[dict[str, Any]]) -> dict[str, Any]:
    statuses = Counter(item["status"] for item in results)
    rules = Counter(rule for item in results for rule in item.get("rules", []))
    error_types = Counter(item.get("error_type", "") for item in results if item.get("error_type"))
    case_keys = {(item["template"], item["input"]) for item in results}
    failed_case_keys = {
        (item["template"], item["input"])
        for item in results
        if item["status"] in {"fail", "error"}
    }
    passed_case_keys = {
        (item["template"], item["input"])
        for item in results
        if item["status"] == "pass"
    }
    unresolved_case_keys = failed_case_keys - passed_case_keys
    return {
        "total_variants": len(results),
        "total_cases": len(case_keys),
        "passed_cases": len(passed_case_keys),
        "failed_cases": len(unresolved_case_keys),
        "statuses": dict(statuses),
        "blocking_rules": dict(rules),
        "error_types": dict(error_types),
    }


def _render_markdown(report: dict[str, Any]) -> str:
    lines = [
        "# Selfcheck Report",
        "",
        f"Generated: {report['generated_at']}",
        "",
        "## Summary",
        "",
    ]
    for key, value in report["summary"].items():
        lines.append(f"- {key}: `{value}`")
    lines.extend(["", "## Failing Variants", ""])
    failing = [item for item in report["results"] if item["status"] != "pass"]
    if not failing:
        lines.append("No failing variants.")
    for item in failing[:200]:
        lines.append(
            f"- `{item['status']}` `{item.get('stage')}` template=`{Path(item['template']).name}` "
            f"input=`{Path(item['input']).name}` variant=`{item.get('variant')}` "
            f"rules=`{item.get('rules') or item.get('error_type')}`"
        )
    return "\n".join(lines) + "\n"


def _safe_name(value: str) -> str:
    return "".join(char if char.isalnum() or char in {"-", "_"} else "_" for char in value)[:180]


if __name__ == "__main__":
    raise SystemExit(main())
