from __future__ import annotations

import os
import tempfile
from pathlib import Path
from typing import get_args

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse

from a3presentation.domain.api import (
    AnalyzeTemplateResponse,
    AutoUploadTemplateResponse,
    ExtractTextResponse,
    GenerationDiagnostic,
    GenerationDiagnosticRuleMetadata,
    GenerationDiagnosticsSummary,
    GenerationDiagnosticsMetadataResponse,
    GeneratePresentationResponse,
    PresentationDiagnosticsResponse,
    PlanWithTemplateResponse,
    TemplateDetailsResponse,
    TemplateSummary,
    TextPlanRequest,
    UploadTemplateResponse,
)
from a3presentation.domain.diagnostics import GenerationDiagnosticRule, GenerationDiagnosticSeverity, GenerationDiagnosticSource
from a3presentation.domain.presentation import PresentationPlan
from a3presentation.domain.template import TemplateManifest
from a3presentation.services.diagnostic_catalog import diagnostic_rule_action, diagnostic_rule_label
from a3presentation.services.document_text_extractor import DocumentTextExtractor
from a3presentation.services.planner import TextToPlanService
from a3presentation.services.presentation_generation import PresentationGenerationError, PresentationGenerationService
from a3presentation.services.pptx_generator import PptxGenerator
from a3presentation.services.table_chart_analyzer import TableChartAnalyzer
from a3presentation.services.template_analyzer import TemplateAnalyzer
from a3presentation.services.template_registry import TemplateRegistry
from a3presentation.settings import get_settings

router = APIRouter()

settings = get_settings()
template_registry = TemplateRegistry(settings.templates_dir)
planner = TextToPlanService()
analyzer = TemplateAnalyzer()
document_text_extractor = DocumentTextExtractor()
table_chart_analyzer = TableChartAnalyzer()
generator = PptxGenerator()
generation_service = PresentationGenerationService(
    generator=generator,
    template_registry=template_registry,
    output_dir=settings.outputs_dir,
)


@router.get("/health")
def healthcheck() -> dict[str, str]:
    return {
        "status": "ok",
        "commit": os.getenv("APP_COMMIT_SHA", "unknown"),
        "branch": os.getenv("APP_COMMIT_BRANCH", "unknown"),
    }


@router.get("/diagnostics/metadata")
def diagnostics_metadata() -> GenerationDiagnosticsMetadataResponse:
    rules = sorted(get_args(GenerationDiagnosticRule))
    return GenerationDiagnosticsMetadataResponse(
        severities=list(get_args(GenerationDiagnosticSeverity)),
        sources=list(get_args(GenerationDiagnosticSource)),
        rules=[
            GenerationDiagnosticRuleMetadata(
                rule=rule,
                label=diagnostic_rule_label(rule),
                action=diagnostic_rule_action(rule),
            )
            for rule in rules
        ],
    )


@router.get("/templates")
def list_templates() -> list[TemplateSummary]:
    templates = template_registry.list_templates()
    return [
        TemplateSummary(
            template_id=item.template_id,
            display_name=item.display_name,
            description=item.description,
        )
        for item in templates
    ]


@router.get("/templates/{template_id}")
def get_template(template_id: str) -> TemplateDetailsResponse:
    try:
        manifest = template_registry.get_template(template_id)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc

    template_dir = settings.templates_dir / template_id
    template_path = template_dir / manifest.source_pptx
    return TemplateDetailsResponse(
        manifest=manifest,
        has_template_file=template_path.exists(),
        inventory_summary=template_registry.build_inventory_summary(manifest),
        editable_targets=template_registry.build_editable_targets(manifest),
        detected_components=template_registry.build_detected_components(manifest),
    )


@router.post("/templates", status_code=201)
async def upload_template(
    manifest_json: str = Form(...),
    template_file: UploadFile = File(...),
) -> UploadTemplateResponse:
    try:
        manifest = TemplateManifest.model_validate_json(manifest_json)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Invalid manifest_json: {exc}") from exc

    if not template_file.filename or not template_file.filename.lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="template_file must be a .pptx")

    template_bytes = await template_file.read()
    if not template_bytes:
        raise HTTPException(status_code=400, detail="template_file is empty")
    try:
        template_path = template_registry.save_template_file(
            template_id=manifest.template_id,
            filename=manifest.source_pptx,
            content=template_bytes,
        )
        manifest_path = template_registry.save_manifest(manifest)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return UploadTemplateResponse(
        template_id=manifest.template_id,
        manifest_path=str(manifest_path),
        template_path=str(template_path),
    )


@router.post("/templates/auto", status_code=201)
async def upload_template_auto(
    template_id: str = Form(...),
    display_name: str = Form(...),
    description: str | None = Form(default=None),
    template_file: UploadFile = File(...),
) -> AutoUploadTemplateResponse:
    if not template_file.filename or not template_file.filename.lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="template_file must be a .pptx")

    template_bytes = await template_file.read()
    if not template_bytes:
        raise HTTPException(status_code=400, detail="template_file is empty")
    try:
        template_path = template_registry.save_template_file(
            template_id=template_id,
            filename="template.pptx",
            content=template_bytes,
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    try:
        manifest = analyzer.analyze(
            template_id=template_id,
            template_path=template_path,
            display_name=display_name,
        )
    except Exception as exc:
        raise HTTPException(status_code=400, detail="Failed to analyze uploaded template") from exc
    manifest.description = description
    try:
        manifest_path = template_registry.save_manifest(manifest)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return AutoUploadTemplateResponse(
        template_id=template_id,
        manifest_path=str(manifest_path),
        template_path=str(template_path),
        analyzed=True,
        inventory_summary=template_registry.build_inventory_summary(manifest),
        editable_targets=template_registry.build_editable_targets(manifest),
        detected_components=template_registry.build_detected_components(manifest),
    )


@router.post("/templates/{template_id}/analyze")
def analyze_template(template_id: str, display_name: str | None = None) -> AnalyzeTemplateResponse:
    try:
        template_path = template_registry.get_template_pptx_path(template_id)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except FileNotFoundError:
        template_dir = settings.templates_dir / template_id
        if not template_dir.exists():
            raise HTTPException(status_code=404, detail=f"Template '{template_id}' not found")
        pptx_candidates = sorted(template_dir.glob("*.pptx"))
        if not pptx_candidates:
            raise HTTPException(status_code=404, detail=f"Template PPTX not found for '{template_id}'")
        template_path = pptx_candidates[0]

    manifest = analyzer.analyze(
        template_id=template_id,
        template_path=template_path,
        display_name=display_name,
    )
    try:
        manifest_path = template_registry.save_manifest(manifest)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return AnalyzeTemplateResponse(
        template_id=template_id,
        manifest_path=str(manifest_path),
        inventory_summary=template_registry.build_inventory_summary(manifest),
        editable_targets=template_registry.build_editable_targets(manifest),
        detected_components=template_registry.build_detected_components(manifest),
    )


@router.post("/plans/from-text")
def plan_from_text(payload: TextPlanRequest) -> PresentationPlan:
    plan = planner.build_plan(
        template_id=payload.template_id,
        raw_text=payload.raw_text,
        title=payload.title,
        tables=payload.tables,
        blocks=payload.blocks,
        chart_overrides=payload.chart_overrides,
    )
    try:
        manifest = template_registry.get_template(payload.template_id)
    except (ValueError, FileNotFoundError):
        return plan
    return template_registry.apply_layout_inventory_to_plan(manifest, plan)


@router.post("/plans/from-text-with-template")
async def plan_from_text_with_template(
    payload_json: str = Form(...),
    template_file: UploadFile = File(...),
) -> PlanWithTemplateResponse:
    if not template_file.filename or not template_file.filename.lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="template_file must be a .pptx")

    try:
        payload = TextPlanRequest.model_validate_json(payload_json)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Invalid payload_json: {exc}") from exc

    template_bytes = await template_file.read()
    if not template_bytes:
        raise HTTPException(status_code=400, detail="template_file is empty")

    with tempfile.TemporaryDirectory() as temp_dir:
        template_path = Path(temp_dir) / "template.pptx"
        template_path.write_bytes(template_bytes)
        try:
            manifest = analyzer.analyze(
                template_id=f"uploaded_{Path(template_file.filename).stem or 'template'}",
                template_path=template_path,
                display_name=Path(template_file.filename).stem or "Uploaded template",
            )
        except Exception as exc:
            raise HTTPException(status_code=400, detail="Failed to analyze uploaded template") from exc
        manifest = template_registry.normalize_manifest(manifest)
        plan = planner.build_plan(
            template_id=manifest.template_id,
            raw_text=payload.raw_text,
            title=payload.title,
            tables=payload.tables,
            blocks=payload.blocks,
            chart_overrides=payload.chart_overrides,
        )
        plan = template_registry.apply_layout_inventory_to_plan(manifest, plan)
    slide_layout_reviews = template_registry.build_slide_layout_reviews(manifest, plan)
    return PlanWithTemplateResponse(
        plan=plan,
        manifest=manifest,
        inventory_summary=template_registry.build_inventory_summary(manifest),
        editable_targets=template_registry.build_editable_targets(manifest),
        detected_components=template_registry.build_detected_components(manifest),
        slide_layout_reviews=slide_layout_reviews,
    )


@router.post("/documents/extract-text")
async def extract_document_text(file: UploadFile = File(...)) -> ExtractTextResponse:
    if not file.filename:
        raise HTTPException(status_code=400, detail="File name is required")
    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="file must be a .docx")

    content = await file.read()
    try:
        text, tables, blocks = document_text_extractor.extract(file.filename, content)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Failed to extract text from '{file.filename}'") from exc

    if not text.strip():
        raise HTTPException(status_code=400, detail=f"No extractable text found in '{file.filename}'")

    chart_assessments = [
        table_chart_analyzer.analyze(table, table_id=f"table_{index}")
        for index, table in enumerate(tables, start=1)
    ]

    return ExtractTextResponse(
        file_name=file.filename,
        text=text,
        tables=tables,
        blocks=blocks,
        chart_assessments=chart_assessments,
    )


@router.post("/presentations/generate")
def generate_presentation(plan: PresentationPlan) -> GeneratePresentationResponse:
    try:
        manifest = template_registry.get_template(plan.template_id)
        template_path = template_registry.get_template_pptx_path(plan.template_id)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc

    target_plan = template_registry.apply_layout_inventory_to_plan(manifest, plan)
    result = _generate_checked_result(target_plan, manifest, template_path)
    output_path = result.output_path
    diagnostics = _generation_diagnostics(result)
    return GeneratePresentationResponse(
        output_path=str(output_path),
        file_name=output_path.name,
        download_url=f"/presentations/files/{output_path.name}",
        warnings=result.warnings,
        diagnostics=diagnostics,
        diagnostics_summary=_diagnostics_summary(diagnostics),
        attempt_count=result.attempt_count,
    )


@router.post("/presentations/diagnose")
def diagnose_presentation(plan: PresentationPlan) -> PresentationDiagnosticsResponse:
    try:
        manifest = template_registry.get_template(plan.template_id)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc

    target_plan = template_registry.apply_layout_inventory_to_plan(manifest, plan)
    diagnostics = _diagnostic_items(generation_service.preflight_diagnostics(target_plan, manifest))
    return PresentationDiagnosticsResponse(diagnostics=diagnostics, diagnostics_summary=_diagnostics_summary(diagnostics))


@router.post("/presentations/generate-with-template")
async def generate_presentation_with_template(
    plan_json: str = Form(...),
    template_file: UploadFile = File(...),
) -> GeneratePresentationResponse:
    if not template_file.filename or not template_file.filename.lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="template_file must be a .pptx")

    try:
        plan = PresentationPlan.model_validate_json(plan_json)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Invalid plan_json: {exc}") from exc

    template_bytes = await template_file.read()
    if not template_bytes:
        raise HTTPException(status_code=400, detail="template_file is empty")

    with tempfile.TemporaryDirectory() as temp_dir:
        template_path = Path(temp_dir) / "template.pptx"
        template_path.write_bytes(template_bytes)
        try:
            manifest = analyzer.analyze(
                template_id=f"uploaded_{Path(template_file.filename).stem or 'template'}",
                template_path=template_path,
                display_name=Path(template_file.filename).stem or "Uploaded template",
            )
        except Exception as exc:
            raise HTTPException(status_code=400, detail="Failed to analyze uploaded template") from exc
        manifest = template_registry.normalize_manifest(manifest)
        transient_plan = plan.model_copy(update={"template_id": manifest.template_id}, deep=True)
        transient_plan = template_registry.apply_layout_inventory_to_plan(manifest, transient_plan)
        result = _generate_checked_result(transient_plan, manifest, template_path)
        output_path = result.output_path

    diagnostics = _generation_diagnostics(result)
    return GeneratePresentationResponse(
        output_path=str(output_path),
        file_name=output_path.name,
        download_url=f"/presentations/files/{output_path.name}",
        warnings=result.warnings,
        diagnostics=diagnostics,
        diagnostics_summary=_diagnostics_summary(diagnostics),
        attempt_count=result.attempt_count,
    )


@router.post("/presentations/diagnose-with-template")
async def diagnose_presentation_with_template(
    plan_json: str = Form(...),
    template_file: UploadFile = File(...),
) -> PresentationDiagnosticsResponse:
    if not template_file.filename or not template_file.filename.lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="template_file must be a .pptx")

    try:
        plan = PresentationPlan.model_validate_json(plan_json)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Invalid plan_json: {exc}") from exc

    template_bytes = await template_file.read()
    if not template_bytes:
        raise HTTPException(status_code=400, detail="template_file is empty")

    with tempfile.TemporaryDirectory() as temp_dir:
        template_path = Path(temp_dir) / "template.pptx"
        template_path.write_bytes(template_bytes)
        try:
            manifest = analyzer.analyze(
                template_id=f"uploaded_{Path(template_file.filename).stem or 'template'}",
                template_path=template_path,
                display_name=Path(template_file.filename).stem or "Uploaded template",
            )
        except Exception as exc:
            raise HTTPException(status_code=400, detail="Failed to analyze uploaded template") from exc
        manifest = template_registry.normalize_manifest(manifest)
        transient_plan = plan.model_copy(update={"template_id": manifest.template_id}, deep=True)
        transient_plan = template_registry.apply_layout_inventory_to_plan(manifest, transient_plan)

    diagnostics = _diagnostic_items(generation_service.preflight_diagnostics(transient_plan, manifest))
    return PresentationDiagnosticsResponse(diagnostics=diagnostics, diagnostics_summary=_diagnostics_summary(diagnostics))


def _generate_checked_presentation(plan: PresentationPlan, manifest: TemplateManifest, template_path: Path) -> Path:
    return _generate_checked_result(plan, manifest, template_path).output_path


def _generate_checked_result(plan: PresentationPlan, manifest: TemplateManifest, template_path: Path):
    try:
        return generation_service.generate_checked_result(plan, manifest, template_path)
    except PresentationGenerationError as exc:
        raise HTTPException(
            status_code=500,
            detail={
                "code": exc.code,
                "message": exc.message,
                "attempt_count": exc.attempt_count,
            },
        ) from exc
    except ValueError as exc:
        raise HTTPException(
            status_code=500,
            detail={
                "code": "generation_failed",
                "message": f"Failed to generate a valid PowerPoint file: {exc}",
                "attempt_count": 0,
            },
        ) from exc
    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail={
                "code": "generation_failed",
                "message": "Failed to generate PowerPoint file",
                "attempt_count": 0,
            },
        ) from exc


def _generation_diagnostics(result) -> list[GenerationDiagnostic]:
    return _diagnostic_items(result.diagnostics or [])


def _diagnostic_items(items) -> list[GenerationDiagnostic]:
    return [
        GenerationDiagnostic(
            slide_index=item.slide_index,
            title=item.title,
            severity=item.severity,
            rule=item.rule,
            label=diagnostic_rule_label(item.rule),
            action=diagnostic_rule_action(item.rule),
            details=item.details,
            source=item.source,
        )
        for item in items
    ]


def _diagnostics_summary(items: list[GenerationDiagnostic]) -> GenerationDiagnosticsSummary:
    return GenerationDiagnosticsSummary(
        total=len(items),
        blocking=sum(1 for item in items if item.severity == "blocking"),
        retryable=sum(1 for item in items if item.severity == "retryable"),
        warning=sum(1 for item in items if item.severity == "warning"),
        capacity=sum(1 for item in items if item.source == "capacity"),
        style=sum(1 for item in items if item.source == "style"),
    )


@router.get("/presentations/files/{file_name}")
def download_presentation(file_name: str) -> FileResponse:
    safe_name = Path(file_name).name
    file_path = settings.outputs_dir / safe_name
    if not file_path.exists() or not file_path.is_file():
        raise HTTPException(status_code=404, detail=f"Generated file '{safe_name}' not found")
    return FileResponse(
        path=file_path,
        filename=safe_name,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
