from __future__ import annotations

import os
import tempfile
from pathlib import Path

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse

from a3presentation.domain.api import (
    AnalyzeTemplateResponse,
    AutoUploadTemplateResponse,
    ExtractTextResponse,
    GeneratePresentationResponse,
    PlanWithTemplateResponse,
    TemplateDetailsResponse,
    TemplateSummary,
    TextPlanRequest,
    UploadTemplateResponse,
)
from a3presentation.domain.presentation import PresentationPlan
from a3presentation.domain.template import TemplateManifest
from a3presentation.services.document_text_extractor import DocumentTextExtractor
from a3presentation.services.deck_audit import audit_generated_presentation, find_capacity_violations
from a3presentation.services.planner import TextToPlanService
from a3presentation.services.table_chart_analyzer import TableChartAnalyzer
from a3presentation.services.template_analyzer import TemplateAnalyzer
from a3presentation.services.pptx_generator import PptxGenerator
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


@router.get("/health")
def healthcheck() -> dict[str, str]:
    return {
        "status": "ok",
        "commit": os.getenv("APP_COMMIT_SHA", "unknown"),
        "branch": os.getenv("APP_COMMIT_BRANCH", "unknown"),
    }


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

    output_path = _generate_checked_presentation(plan, manifest, template_path)
    return GeneratePresentationResponse(
        output_path=str(output_path),
        file_name=output_path.name,
        download_url=f"/presentations/files/{output_path.name}",
    )


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
        output_path = _generate_checked_presentation(transient_plan, manifest, template_path)

    return GeneratePresentationResponse(
        output_path=str(output_path),
        file_name=output_path.name,
        download_url=f"/presentations/files/{output_path.name}",
    )


def _generate_checked_presentation(plan: PresentationPlan, manifest: TemplateManifest, template_path: Path) -> Path:
    try:
        output_path = generator.generate(
            template_path=template_path,
            manifest=manifest,
            plan=plan,
            output_dir=settings.outputs_dir,
        )
        audits = audit_generated_presentation(output_path, plan, manifest)
        violations = _blocking_generation_violations(find_capacity_violations(audits))
        if violations:
            details = "; ".join(f"slide {item.slide_index}: {item.rule}" for item in violations[:6])
            raise ValueError(f"Generated deck failed layout quality gate: {details}")
        return output_path
    except ValueError as exc:
        raise HTTPException(status_code=500, detail=f"Failed to generate a valid PowerPoint file: {exc}") from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail="Failed to generate PowerPoint file") from exc


def _blocking_generation_violations(violations):
    blocking_rules = {
        "table_overlay_text_overflow",
        "missing_table_shape",
        "missing_chart_shape",
        "missing_image_shape",
        "content_order_mismatch",
        "card_overlap",
        "two_column_overlap",
        "image_text_overlap",
        "chart_type_mismatch",
        "chart_series_count_mismatch",
    }
    return [item for item in violations if item.rule in blocking_rules]


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
