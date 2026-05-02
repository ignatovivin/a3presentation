from __future__ import annotations

import asyncio
import importlib
import os
import shutil
import tempfile
import unittest
from io import BytesIO
from pathlib import Path
from typing import get_args
from unittest.mock import patch

from docx import Document
from fastapi import HTTPException
from pptx import Presentation
from pydantic import ValidationError
from starlette.datastructures import UploadFile

from a3presentation import main as main_module
from a3presentation import settings as settings_module
from a3presentation.api import routes as routes_module
from a3presentation.services import presentation_generation as generation_module
from a3presentation.services.presentation_generation import (
    GeneratedPresentationResult,
    GenerationDiagnosticItem,
    PresentationGenerationError,
)
from a3presentation.domain.template import LayoutSpec, PlaceholderKind, PlaceholderSpec, TemplateManifest, TemplateTextStyleSpec
from a3presentation.domain.api import GenerationDiagnostic, TextPlanRequest
from a3presentation.domain.diagnostics import GenerationDiagnosticRule
from a3presentation.domain.presentation import (
    PresentationPlan,
    RenderTargetType,
    SlideContentBlock,
    SlideContentBlockKind,
    SlideKind,
    SlideRenderTarget,
    SlideSpec,
)
from a3presentation.services.deck_audit import CapacityViolation, SlideAudit, TextOverflowDetail, find_capacity_violations
from a3presentation.services.diagnostic_catalog import DIAGNOSTIC_RULE_ACTIONS, DIAGNOSTIC_RULE_LABELS
from a3presentation.services.layout_capacity import TEXT_FULL_WIDTH_PROFILE
from a3presentation.services.slide_text_policy import slide_content_chunks, slide_text_demand_chars
from a3presentation.services.text_fit import FitBox, TextBoxMetrics, assign_text_chunks_to_boxes


class ApiContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._env_backup = {key: os.environ.get(key) for key in ("TEMPLATES_DIR", "OUTPUTS_DIR", "STORAGE_DIR")}
        cls._temp_dir = tempfile.TemporaryDirectory()
        cls._root = Path(cls._temp_dir.name)
        cls._templates_dir = cls._root / "templates"
        cls._outputs_dir = cls._root / "outputs"
        cls._templates_dir.mkdir(parents=True, exist_ok=True)
        cls._outputs_dir.mkdir(parents=True, exist_ok=True)

        source_templates = Path(__file__).resolve().parent / "fixtures" / "templates"
        for template_id in ("deterministic_layout_fixture",):
            shutil.copytree(source_templates / template_id, cls._templates_dir / template_id)
        manifest_path = cls._templates_dir / "missing_source" / "manifest.json"
        manifest_path.parent.mkdir(parents=True, exist_ok=True)
        manifest_path.write_text(
            TemplateManifest(
                template_id="missing_source",
                display_name="Missing Source",
                description="Manifest without pptx source file",
                source_pptx="missing.pptx",
                default_layout_key="cover",
                layouts=[],
            ).model_dump_json(),
            encoding="utf-8",
        )

        os.environ["TEMPLATES_DIR"] = str(cls._templates_dir)
        os.environ["OUTPUTS_DIR"] = str(cls._outputs_dir)
        os.environ["STORAGE_DIR"] = str(cls._root)

        importlib.reload(settings_module)
        importlib.reload(routes_module)
        importlib.reload(main_module)

    @classmethod
    def tearDownClass(cls) -> None:
        cls._temp_dir.cleanup()
        for key, value in cls._env_backup.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value
        importlib.reload(settings_module)
        importlib.reload(routes_module)
        importlib.reload(main_module)

    def test_health_endpoint_returns_ok(self) -> None:
        response = routes_module.healthcheck()
        self.assertEqual(response["status"], "ok")
        self.assertIn("commit", response)
        self.assertIn("branch", response)

    def test_create_app_starts_with_empty_template_storage(self) -> None:
        empty_root = self._root / "empty"
        templates_dir = empty_root / "templates"
        outputs_dir = empty_root / "outputs"
        templates_dir.mkdir(parents=True, exist_ok=True)
        outputs_dir.mkdir(parents=True, exist_ok=True)
        os.environ["TEMPLATES_DIR"] = str(templates_dir)
        os.environ["OUTPUTS_DIR"] = str(outputs_dir)
        os.environ["STORAGE_DIR"] = str(empty_root)

        importlib.reload(settings_module)
        importlib.reload(main_module)

        app = main_module.create_app()

        self.assertEqual(app.title, "A3 Presentation API")
        self.assertEqual(list(templates_dir.iterdir()), [])

        os.environ["TEMPLATES_DIR"] = str(self._templates_dir)
        os.environ["OUTPUTS_DIR"] = str(self._outputs_dir)
        os.environ["STORAGE_DIR"] = str(self._root)
        importlib.reload(settings_module)
        importlib.reload(main_module)

    def test_templates_endpoint_lists_available_templates(self) -> None:
        templates = routes_module.list_templates()
        template_ids = {item.template_id for item in templates}
        self.assertIn("deterministic_layout_fixture", template_ids)

    def test_template_details_expose_missing_template_file(self) -> None:
        response = routes_module.get_template("missing_source")
        self.assertFalse(response.has_template_file)
        self.assertEqual(response.inventory_summary.usability_status, "not_safely_editable")
        self.assertEqual(response.inventory_summary.generation_mode, "layout")
        self.assertEqual(response.inventory_summary.layout_target_count, 0)
        self.assertEqual(response.inventory_summary.prototype_target_count, 0)
        self.assertEqual(response.editable_targets, [])
        self.assertEqual(response.detected_components, [])

    def test_template_details_reject_path_traversal_template_id(self) -> None:
        with self.assertRaises(HTTPException) as error:
            routes_module.get_template("..\\..\\outside")

        self.assertEqual(error.exception.status_code, 400)
        self.assertIn("escapes the storage root", error.exception.detail)

    def test_extract_text_endpoint_returns_blocks_tables_and_chart_assessments(self) -> None:
        document = Document()
        document.add_paragraph("Отчет по сегментам")
        document.add_heading("Рынок", level=1)
        table = document.add_table(rows=3, cols=2)
        table.cell(0, 0).text = "Сегмент"
        table.cell(0, 1).text = "Выручка"
        table.cell(1, 0).text = "SMB"
        table.cell(1, 1).text = "120"
        table.cell(2, 0).text = "Enterprise"
        table.cell(2, 1).text = "250"
        buffer = BytesIO()
        document.save(buffer)
        upload = UploadFile(filename="report.docx", file=BytesIO(buffer.getvalue()))

        payload = asyncio.run(routes_module.extract_document_text(upload))
        self.assertEqual(payload.file_name, "report.docx")
        self.assertTrue(payload.text)
        self.assertEqual(len(payload.tables), 1)
        self.assertEqual(len(payload.chart_assessments), 1)

    def test_extract_text_endpoint_rejects_non_docx_uploads(self) -> None:
        upload = UploadFile(filename="report.txt", file=BytesIO(b"plain text"))

        with self.assertRaises(HTTPException) as error:
            asyncio.run(routes_module.extract_document_text(upload))

        self.assertEqual(error.exception.status_code, 400)
        self.assertIn(".docx", error.exception.detail)

    def test_extract_text_endpoint_offers_safe_combo_variants_for_mixed_unit_table(self) -> None:
        document = Document()
        document.add_heading("Метрики", level=1)
        table = document.add_table(rows=4, cols=3)
        table.cell(0, 0).text = "Квартал"
        table.cell(0, 1).text = "Выручка"
        table.cell(0, 2).text = "Маржа"
        for row_index, values in enumerate(
            [
                ("Q1", "120 млн руб", "18%"),
                ("Q2", "150 млн руб", "22%"),
                ("Q3", "190 млн руб", "27%"),
            ],
            start=1,
        ):
            for col_index, value in enumerate(values):
                table.cell(row_index, col_index).text = value
        buffer = BytesIO()
        document.save(buffer)
        upload = UploadFile(filename="mixed-units.docx", file=BytesIO(buffer.getvalue()))

        payload = asyncio.run(routes_module.extract_document_text(upload))

        self.assertEqual(len(payload.chart_assessments), 1)
        chart_specs = payload.chart_assessments[0].candidate_specs
        self.assertEqual(
            [spec.chart_type.value for spec in chart_specs],
            ["combo", "combo", "column", "column", "line", "line"],
        )
        self.assertEqual(
            [spec.variant_label for spec in chart_specs],
            [
                "Комбинированный: столбцы Выручка; линия Маржа",
                "Комбинированный: столбцы Маржа; линия Выручка",
                "Единичный: Выручка",
                "Единичный: Маржа",
                "Единичный: Выручка",
                "Единичный: Маржа",
            ],
        )

    def test_extract_text_endpoint_rejects_too_ambiguous_mixed_unit_chart(self) -> None:
        document = Document()
        document.add_heading("Смешанные метрики", level=1)
        table = document.add_table(rows=4, cols=4)
        for col_index, value in enumerate(["Метрика", "Деньги", "Доля", "Количество"]):
            table.cell(0, col_index).text = value
        for row_index, values in enumerate(
            [
                ("A", "120 млн руб", "18%", "25"),
                ("B", "150 млн руб", "22%", "31"),
                ("C", "190 млн руб", "27%", "44"),
            ],
            start=1,
        ):
            for col_index, value in enumerate(values):
                table.cell(row_index, col_index).text = value
        buffer = BytesIO()
        document.save(buffer)
        upload = UploadFile(filename="too-mixed.docx", file=BytesIO(buffer.getvalue()))

        payload = asyncio.run(routes_module.extract_document_text(upload))

        self.assertEqual(len(payload.chart_assessments), 1)
        self.assertFalse(payload.chart_assessments[0].chartable)
        self.assertEqual(payload.chart_assessments[0].candidate_specs, [])

    def test_plan_from_text_returns_presentation_plan(self) -> None:
        payload = routes_module.plan_from_text(
            TextPlanRequest(
                template_id="deterministic_layout_fixture",
                title="Demo",
                raw_text="Основные выводы\n- Рост выручки\n- Снижение churn",
            )
        )
        self.assertEqual(payload.template_id, "deterministic_layout_fixture")
        self.assertGreaterEqual(len(payload.slides), 1)

    def test_plan_from_text_accepts_transient_uploaded_template_id(self) -> None:
        payload = routes_module.plan_from_text(
            TextPlanRequest(
                template_id="uploaded_customer_template",
                title="Demo",
                raw_text="Основные выводы\n- Рост выручки\n- Снижение churn",
            )
        )
        self.assertEqual(payload.template_id, "uploaded_customer_template")
        self.assertGreaterEqual(len(payload.slides), 1)
        legacy_target_keys = {"cover", "text_full_width", "dense_text_full_width", "list_full_width", "table", "image_text"}
        self.assertTrue(all(slide.preferred_layout_key not in legacy_target_keys for slide in payload.slides))

    def test_plan_from_text_with_uploaded_template_returns_plan_and_manifest(self) -> None:
        template_path = self._templates_dir / "deterministic_layout_fixture" / "template.pptx"
        upload = UploadFile(filename="customer-template.pptx", file=BytesIO(template_path.read_bytes()))
        payload = TextPlanRequest(
            template_id="ignored_template_id",
            title="Demo",
            raw_text="Основные выводы\n- Рост выручки\n- Снижение churn",
        )

        response = asyncio.run(
            routes_module.plan_from_text_with_template(
                payload_json=payload.model_dump_json(),
                template_file=upload,
            )
        )

        self.assertTrue(response.manifest.template_id.startswith("uploaded_customer-template"))
        self.assertEqual(response.plan.template_id, response.manifest.template_id)
        self.assertGreaterEqual(len(response.plan.slides), 1)
        self.assertTrue(response.inventory_summary.targets)
        self.assertIn(response.inventory_summary.usability_status, {"usable", "usable_with_degradation"})
        self.assertTrue(response.editable_targets)
        self.assertTrue(response.detected_components)
        self.assertEqual(response.inventory_summary.generation_mode, response.manifest.generation_mode.value)
        self.assertEqual(len(response.slide_layout_reviews), len(response.plan.slides))
        self.assertTrue(all(review.available_layouts for review in response.slide_layout_reviews))
        review_payload = response.model_dump(mode="json")
        self.assertTrue(
            all("current_layout_key" not in review for review in review_payload["slide_layout_reviews"])
        )
        inventory_target_keys = {target.key for target in response.inventory_summary.targets}
        self.assertTrue(
            all(
                review.current_target_key is None or review.current_target_key in inventory_target_keys
                for review in response.slide_layout_reviews
            )
        )
        text_review = next((review for review in response.slide_layout_reviews if review.slide_index > 0), None)
        self.assertIsNotNone(text_review)
        self.assertEqual(text_review.current_target_key, response.plan.slides[text_review.slide_index].render_target.key)
        self.assertEqual(text_review.current_target_type, response.plan.slides[text_review.slide_index].render_target.type.value)
        self.assertEqual(text_review.current_target_source, response.plan.slides[text_review.slide_index].render_target.source)
        self.assertTrue(text_review.current_target_explanation)
        self.assertEqual(text_review.current_target_confidence, response.plan.slides[text_review.slide_index].render_target.confidence)
        self.assertEqual(
            text_review.current_target_degradation_reasons,
            response.plan.slides[text_review.slide_index].render_target.degradation_reasons,
        )
        self.assertEqual(text_review.current_runtime_profile_key, response.plan.slides[text_review.slide_index].runtime_profile_key)
        best_option = text_review.available_layouts[0]
        self.assertTrue(best_option.runtime_profile_key)
        self.assertTrue(best_option.source_label)
        self.assertTrue(best_option.match_summary)
        self.assertTrue(best_option.recommendation_label)
        self.assertTrue(best_option.recommendation_reasons)
        if best_option.supported_slide_kinds and "text" in best_option.supported_slide_kinds:
            self.assertIsNotNone(best_option.estimated_text_capacity_chars)

    def test_plan_from_text_with_arbitrary_uploaded_template_uses_synthesized_prototype_inventory(self) -> None:
        pptx = Presentation()
        slide = pptx.slides.add_slide(pptx.slide_layouts[6])
        title_shape = slide.shapes.add_textbox(600000, 400000, 8000000, 900000)
        title_shape.text_frame.text = "Произвольный заголовок"
        body_shape = slide.shapes.add_textbox(600000, 1900000, 7600000, 2400000)
        body_shape.text_frame.text = "Первый тезис"
        body_shape.text_frame.add_paragraph().text = "Второй тезис"
        body_shape.text_frame.add_paragraph().text = "Третий тезис"
        buffer = BytesIO()
        pptx.save(buffer)

        upload = UploadFile(filename="arbitrary-template.pptx", file=BytesIO(buffer.getvalue()))
        payload = TextPlanRequest(
            template_id="ignored_template_id",
            title="Demo",
            raw_text="Основные выводы\n- Рост выручки\n- Снижение churn",
        )

        response = asyncio.run(
            routes_module.plan_from_text_with_template(
                payload_json=payload.model_dump_json(),
                template_file=upload,
            )
        )

        self.assertEqual(response.manifest.generation_mode.value, "prototype")
        self.assertTrue(response.manifest.prototype_slides)
        self.assertTrue(response.manifest.inventory.components)
        self.assertTrue(response.manifest.inventory.slides)
        self.assertTrue(response.manifest.inventory.has_prototype_inventory)
        self.assertIn(response.manifest.inventory.degradation_mode, {None, "prototype_only"})
        self.assertTrue(response.inventory_summary.has_prototype_inventory)
        self.assertEqual(response.inventory_summary.usability_status, "usable_with_degradation")
        self.assertTrue(response.inventory_summary.prototype_target_count >= 1)
        self.assertTrue(any(target.source == "prototype" for target in response.inventory_summary.targets))
        self.assertTrue(response.editable_targets)
        self.assertTrue(response.detected_components)
        self.assertTrue(all(review.available_layouts for review in response.slide_layout_reviews))
        self.assertTrue(
            any(option.source == "layout" for review in response.slide_layout_reviews for option in review.available_layouts)
        )
        self.assertTrue(
            any(
                slide.render_target is not None
                and slide.render_target.key
                and any(target.key == slide.render_target.key and target.source == "prototype" for target in response.inventory_summary.targets)
                for slide in response.plan.slides[1:]
            )
        )
        self.assertTrue(
            all(
                review.current_target_key == response.plan.slides[review.slide_index].render_target.key
                for review in response.slide_layout_reviews
                if response.plan.slides[review.slide_index].render_target is not None
            )
        )
        self.assertTrue(
            all(
                review.current_target_degradation_reasons
                == response.plan.slides[review.slide_index].render_target.degradation_reasons
                for review in response.slide_layout_reviews
                if response.plan.slides[review.slide_index].render_target is not None
            )
        )
        layout_option = next(
            option
            for review in response.slide_layout_reviews
            for option in review.available_layouts
            if option.source == "layout"
        )
        self.assertTrue(layout_option.source_label)
        self.assertTrue(layout_option.match_summary)
        self.assertTrue(layout_option.recommendation_label)
        self.assertTrue(layout_option.recommendation_reasons)

    def test_slide_layout_reviews_expose_stable_ranking_metadata(self) -> None:
        template_path = self._templates_dir / "deterministic_layout_fixture" / "template.pptx"
        upload = UploadFile(filename="customer-template.pptx", file=BytesIO(template_path.read_bytes()))
        payload = TextPlanRequest(
            template_id="ignored_template_id",
            title="Demo",
            raw_text="Основные выводы\n- Рост выручки\n- Снижение churn",
        )

        response = asyncio.run(
            routes_module.plan_from_text_with_template(
                payload_json=payload.model_dump_json(),
                template_file=upload,
            )
        )

        self.assertTrue(response.slide_layout_reviews)
        for review in response.slide_layout_reviews:
            self.assertGreaterEqual(review.slide_index, 0)
            self.assertTrue(review.available_layouts)
            self.assertEqual(review.current_runtime_profile_key, response.plan.slides[review.slide_index].runtime_profile_key)
            for option in review.available_layouts:
                self.assertIn(option.source, {"layout", "prototype"})
                self.assertTrue(option.runtime_profile_key)
                self.assertTrue(option.source_label)
                self.assertTrue(option.match_summary)
                self.assertIn(option.recommendation_label, {"Рекомендуем", "Подходит", "Запасной вариант"})
                self.assertTrue(option.recommendation_reasons)

    def test_generate_and_download_presentation_for_valid_template(self) -> None:
        payload = routes_module.generate_presentation(
            PresentationPlan(
                template_id="deterministic_layout_fixture",
                title="Smoke Test",
                slides=[
                    SlideSpec(kind=SlideKind.TITLE, title="Smoke Test", subtitle="API contract"),
                    SlideSpec(kind=SlideKind.TEXT, title="Итог", text="Проверка generate/download через API."),
                ],
            )
        )
        self.assertTrue((self._outputs_dir / payload.file_name).exists())
        self.assertEqual(payload.warnings, [])
        self.assertEqual(payload.diagnostics, [])
        self.assertEqual(payload.attempt_count, 1)

        download = routes_module.download_presentation(payload.file_name)
        self.assertEqual(
            download.media_type,
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        self.assertEqual(Path(download.path).name, payload.file_name)

    def test_diagnostics_metadata_exposes_backend_contract(self) -> None:
        metadata = routes_module.diagnostics_metadata()
        contract_rules = set(get_args(GenerationDiagnosticRule))

        self.assertEqual(set(metadata.severities), {"blocking", "retryable", "warning"})
        self.assertEqual(set(metadata.sources), {"capacity", "style"})
        self.assertEqual({item.rule for item in metadata.rules}, contract_rules)
        self.assertTrue(all(item.label for item in metadata.rules))
        self.assertTrue(all(item.action for item in metadata.rules))

    def test_generate_with_uploaded_template_does_not_require_registry_template(self) -> None:
        template_path = self._templates_dir / "deterministic_layout_fixture" / "template.pptx"
        upload = UploadFile(filename="custom-template.pptx", file=BytesIO(template_path.read_bytes()))
        plan = PresentationPlan(
            template_id="deterministic_layout_fixture",
            title="Custom Template Smoke",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="Custom Template Smoke", subtitle="Transient upload"),
                SlideSpec(kind=SlideKind.TEXT, title="Итог", text="Генерация через временно загруженный шаблон."),
            ],
        )

        payload = asyncio.run(
            routes_module.generate_presentation_with_template(
                plan_json=plan.model_dump_json(),
                template_file=upload,
            )
        )

        self.assertTrue((self._outputs_dir / payload.file_name).exists())
        self.assertNotIn("custom-template", {item.template_id for item in routes_module.list_templates()})

    def test_generate_response_exposes_style_audit_warnings(self) -> None:
        output_path = self._outputs_dir / "style-warning.pptx"
        output_path.write_bytes(b"pptx")

        with patch.object(
            routes_module.generation_service,
            "generate_checked_result",
            return_value=GeneratedPresentationResult(
                output_path=output_path,
                warnings=["slide 1: text_color_mismatch: placeholder=Body actual=['000000'] expected=CC3300"],
                diagnostics=[
                    GenerationDiagnosticItem(
                        slide_index=1,
                        title="Warning",
                        severity="warning",
                        rule="text_color_mismatch",
                        details="placeholder=Body actual=['000000'] expected=CC3300",
                        source="style",
                    )
                ],
                attempt_count=2,
            ),
        ):
            payload = routes_module.generate_presentation(
                PresentationPlan(
                    template_id="deterministic_layout_fixture",
                    title="Style Warning",
                    slides=[SlideSpec(kind=SlideKind.TEXT, title="Warning", text="Style warning propagation.")],
                )
            )

        self.assertEqual(payload.file_name, output_path.name)
        self.assertEqual(payload.warnings, ["slide 1: text_color_mismatch: placeholder=Body actual=['000000'] expected=CC3300"])
        self.assertEqual(len(payload.diagnostics), 1)
        self.assertEqual(payload.diagnostics_summary.total, 1)
        self.assertEqual(payload.diagnostics_summary.warning, 1)
        self.assertEqual(payload.diagnostics_summary.style, 1)
        self.assertEqual(payload.attempt_count, 2)
        self.assertEqual(payload.diagnostics[0].severity, "warning")
        self.assertEqual(payload.diagnostics[0].source, "style")
        self.assertEqual(payload.diagnostics[0].rule, "text_color_mismatch")
        self.assertEqual(payload.diagnostics[0].label, "Цвет текста не совпал")
        self.assertEqual(payload.diagnostics[0].action, "Проверьте цвет текста в placeholder style шаблона.")

    def test_capacity_diagnostics_classify_retryable_and_blocking_rules(self) -> None:
        diagnostics = routes_module.generation_service.capacity_diagnostics(
            [
                CapacityViolation(
                    slide_index=1,
                    title="Retry",
                    rule="overflow_risk",
                    details="fill_ratio=1.40 max=1.00",
                ),
                CapacityViolation(
                    slide_index=2,
                    title="Block",
                    rule="missing_chart_shape",
                    details="chart slide does not contain rendered chart shape",
                ),
            ]
        )

        self.assertEqual([item.severity for item in diagnostics], ["retryable", "blocking"])
        self.assertEqual([item.source for item in diagnostics], ["capacity", "capacity"])

    def test_generation_diagnostics_deduplicate_slide_source_rule(self) -> None:
        diagnostics = routes_module.generation_service.deduplicate_diagnostics(
            [
                GenerationDiagnosticItem(
                    slide_index=2,
                    title="Dup",
                    severity="warning",
                    rule="overflow_risk",
                    details="text_chars=920 estimated_capacity=780",
                    source="capacity",
                ),
                GenerationDiagnosticItem(
                    slide_index=2,
                    title="Dup",
                    severity="retryable",
                    rule="overflow_risk",
                    details="fill_ratio=1.30 max=1.00",
                    source="capacity",
                ),
            ]
        )

        self.assertEqual(len(diagnostics), 1)
        self.assertEqual(diagnostics[0].severity, "retryable")
        self.assertIn("text_chars=920", diagnostics[0].details)
        self.assertIn("fill_ratio=1.30", diagnostics[0].details)

    def test_generation_diagnostic_rejects_unknown_severity_and_source(self) -> None:
        with self.assertRaises(ValidationError):
            GenerationDiagnostic(
                slide_index=1,
                severity="info",
                rule="overflow_risk",
                details="details",
                source="capacity",
            )
        with self.assertRaises(ValidationError):
            GenerationDiagnostic(
                slide_index=1,
                severity="warning",
                rule="overflow_risk",
                details="details",
                source="planner",
            )

    def test_generation_diagnostic_rejects_unknown_rule_and_catalog_covers_contract(self) -> None:
        with self.assertRaises(ValidationError):
            GenerationDiagnostic(
                slide_index=1,
                severity="warning",
                rule="unknown_rule",
                details="details",
                source="capacity",
            )

        contract_rules = set(get_args(GenerationDiagnosticRule))
        self.assertEqual(set(DIAGNOSTIC_RULE_LABELS), contract_rules)
        self.assertTrue(contract_rules.issuperset(DIAGNOSTIC_RULE_ACTIONS))

    def test_diagnose_presentation_exposes_preflight_degradation_reasons(self) -> None:
        response = routes_module.diagnose_presentation(
            PresentationPlan(
                template_id="deterministic_layout_fixture",
                title="Preflight Diagnostics",
                slides=[
                    SlideSpec(
                        kind=SlideKind.TEXT,
                        title="Preflight",
                        text="Проверка ранней диагностики.",
                        render_target=SlideRenderTarget(
                            type=RenderTargetType.AUTO_LAYOUT,
                            key="text_full_width",
                            degradation_reasons=["capacity_retry:rendered_text_overflow:shape=Body:idx=14:height=860>650"],
                            confidence="medium",
                        ),
                    )
                ],
            )
        )

        retryable_rules = {
            item.rule
            for item in response.diagnostics
            if item.severity == "retryable" and item.source == "capacity"
        }
        self.assertEqual(response.diagnostics_summary.total, len(response.diagnostics))
        self.assertGreaterEqual(response.diagnostics_summary.retryable, 1)
        self.assertGreaterEqual(response.diagnostics_summary.capacity, 1)
        self.assertIn("rendered_text_overflow", retryable_rules)
        self.assertIn("Текст вышел за границы блока", {item.label for item in response.diagnostics})
        self.assertIn(
            "Система повторит разбиение по фактическому переполненному блоку.",
            {item.action for item in response.diagnostics},
        )

    def test_preflight_diagnostics_report_missing_required_editable_slot(self) -> None:
        manifest = TemplateManifest(
            template_id="missing_slot_demo",
            display_name="Missing Slot Demo",
            source_pptx="template.pptx",
            default_layout_key="title_only",
            layouts=[
                LayoutSpec(
                    key="title_only",
                    name="Title Only",
                    slide_layout_index=0,
                    supported_slide_kinds=["text"],
                    placeholders=[
                        PlaceholderSpec(
                            name="Title",
                            kind=PlaceholderKind.TITLE,
                            idx=0,
                            editable_role="title",
                            editable_capabilities=["text"],
                        )
                    ],
                )
            ],
        )
        plan = PresentationPlan(
            template_id="missing_slot_demo",
            title="Missing Slot Demo",
            slides=[
                SlideSpec(
                    kind=SlideKind.TEXT,
                    title="Needs body",
                    text="Этот текст требует body slot.",
                    preferred_layout_key="title_only",
                )
            ],
        )

        diagnostics = routes_module.generation_service.preflight_diagnostics(plan, manifest)

        missing_slot = [item for item in diagnostics if item.rule == "missing_required_editable_slot"]
        self.assertEqual(len(missing_slot), 1)
        self.assertEqual(missing_slot[0].severity, "warning")
        self.assertIn("missing_roles=body", missing_slot[0].details)

    def test_generate_retries_after_post_generation_capacity_audit(self) -> None:
        first_output = self._outputs_dir / "first.pptx"
        retry_output = self._outputs_dir / "retry.pptx"
        source_parts = [
            "Первый перегруженный фрагмент текста для проверки post-generation audit retry. " * 2,
            "Второй перегруженный фрагмент текста должен уйти на отдельный слайд после аудита. " * 2,
            "Третий перегруженный фрагмент текста сохраняет смысл исходного документа. " * 2,
        ]
        plan = PresentationPlan(
            template_id="deterministic_layout_fixture",
            title="Retry Contract",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="Retry Contract"),
                SlideSpec(
                    kind=SlideKind.TEXT,
                    title="Перегруженный раздел",
                    text="\n".join(source_parts),
                    content_blocks=[
                        SlideContentBlock(kind=SlideContentBlockKind.PARAGRAPH, text=part)
                        for part in source_parts
                    ],
                ),
            ],
        )
        generated_plans: list[PresentationPlan] = []

        def fake_generate(**kwargs):
            generated_plans.append(kwargs["plan"])
            output_path = first_output if len(generated_plans) == 1 else retry_output
            output_path.write_bytes(b"pptx")
            return output_path

        first_audit = [
            SlideAudit(
                slide_index=2,
                title="Перегруженный раздел",
                kind=SlideKind.TEXT.value,
                layout_key="text_full_width",
                body_char_count=TEXT_FULL_WIDTH_PROFILE.max_chars + 200,
                body_font_sizes=(TEXT_FULL_WIDTH_PROFILE.max_font_pt,),
                profile=TEXT_FULL_WIDTH_PROFILE,
            )
        ]
        first_violations = [
            CapacityViolation(
                slide_index=2,
                title="Перегруженный раздел",
                rule="overflow_risk",
                details="fill_ratio=1.40 max=1.00",
            )
        ]

        with (
            patch.object(routes_module.generation_service.generator, "generate", side_effect=fake_generate),
            patch.object(generation_module, "audit_generated_presentation", side_effect=[first_audit, []]),
            patch.object(generation_module, "find_capacity_violations", side_effect=[first_violations, []]),
            patch.object(generation_module, "audit_presentation_styles", return_value=[]),
        ):
            output_path = routes_module._generate_checked_presentation(
                plan,
                routes_module.template_registry.get_template("deterministic_layout_fixture"),
                routes_module.template_registry.get_template_pptx_path("deterministic_layout_fixture"),
            )

        self.assertEqual(output_path.name, retry_output.name)
        self.assertEqual(len(generated_plans), 2)
        self.assertEqual(len(generated_plans[0].slides), 2)
        self.assertGreater(len(generated_plans[1].slides), len(generated_plans[0].slides))

    def test_generate_error_detail_is_structured_for_retry_exhausted(self) -> None:
        with patch.object(
            routes_module.generation_service,
            "generate_checked_result",
            side_effect=PresentationGenerationError(
                "retry_exhausted",
                "Generated deck failed layout quality gate: slide 2: rendered_text_overflow",
                attempt_count=2,
            ),
        ):
            with self.assertRaises(HTTPException) as error:
                routes_module._generate_checked_result(
                    PresentationPlan(
                        template_id="deterministic_layout_fixture",
                        title="Retry Exhausted",
                        slides=[SlideSpec(kind=SlideKind.TEXT, title="Overflow", text="Overflow")],
                    ),
                    routes_module.template_registry.get_template("deterministic_layout_fixture"),
                    routes_module.template_registry.get_template_pptx_path("deterministic_layout_fixture"),
                )

        self.assertEqual(error.exception.status_code, 500)
        self.assertEqual(error.exception.detail["code"], "retry_exhausted")
        self.assertEqual(error.exception.detail["attempt_count"], 2)
        self.assertIn("layout quality gate", error.exception.detail["message"])

    def test_capacity_retry_uses_overflow_shape_diagnostics_to_split_target_slot(self) -> None:
        manifest = TemplateManifest(
            template_id="slot_retry_demo",
            display_name="Slot Retry Demo",
            source_pptx="template.pptx",
            default_layout_key="diagnostic_text",
            layouts=[
                LayoutSpec(
                    key="diagnostic_text",
                    name="Diagnostic Text",
                    slide_layout_index=0,
                    supported_slide_kinds=["text"],
                    placeholders=[
                        PlaceholderSpec(
                            name="Body",
                            kind=PlaceholderKind.BODY,
                            idx=14,
                            editable_role="body",
                            editable_capabilities=["text"],
                            left_emu=600000,
                            top_emu=1300000,
                            width_emu=3600000,
                            height_emu=1000000,
                            text_style=TemplateTextStyleSpec(font_size_pt=18.0, line_spacing=1.18),
                        )
                    ],
                )
            ],
        )
        source_parts = [
            "Первый фрагмент для точечного retry с подробностями и несколькими ограничениями.",
            "Второй фрагмент для точечного retry с рисками и следующим действием.",
        ]
        plan = PresentationPlan(
            template_id="slot_retry_demo",
            title="Slot Retry Demo",
            slides=[
                SlideSpec(
                    kind=SlideKind.TEXT,
                    title="Targeted retry",
                    text="\n".join(source_parts),
                    content_blocks=[
                        SlideContentBlock(kind=SlideContentBlockKind.PARAGRAPH, text=part)
                        for part in source_parts
                    ],
                    preferred_layout_key="diagnostic_text",
                )
            ],
        )
        audit = SlideAudit(
            slide_index=1,
            title="Targeted retry",
            kind=SlideKind.TEXT.value,
            layout_key="diagnostic_text",
            body_char_count=sum(len(part) for part in source_parts),
            body_font_sizes=(18.0,),
            profile=TEXT_FULL_WIDTH_PROFILE,
            body_text_overflow_details=(
                TextOverflowDetail(
                    shape_name="Body",
                    shape_id=7,
                    placeholder_idx=14,
                    left=600000,
                    top=1300000,
                    width=3600000,
                    height=1000000,
                    estimated_height=860,
                    available_height=650,
                    font_size_pt=18.0,
                    text_excerpt=source_parts[0],
                ),
            ),
        )
        violations = [
            CapacityViolation(
                slide_index=1,
                title="Targeted retry",
                rule="rendered_text_overflow",
                details="shape=Body idx=14",
            )
        ]

        retry_plan = routes_module.generation_service.plan_with_capacity_retry(plan, manifest, [audit], violations)

        self.assertIsNotNone(retry_plan)
        assert retry_plan is not None
        self.assertEqual(len(retry_plan.slides), 2)
        self.assertIn(source_parts[0], retry_plan.slides[0].text or "")
        self.assertIn(source_parts[1], retry_plan.slides[1].text or "")
        self.assertIsNotNone(retry_plan.slides[0].render_target)
        assert retry_plan.slides[0].render_target is not None
        self.assertIn(
            "capacity_retry:rendered_text_overflow:shape=Body:idx=14:height=860>650",
            retry_plan.slides[0].render_target.degradation_reasons,
        )
        self.assertEqual(retry_plan.slides[0].render_target.confidence, "medium")

    def test_capacity_retry_splits_only_overflow_assigned_slot_batch(self) -> None:
        manifest = TemplateManifest(
            template_id="slot_batch_retry_demo",
            display_name="Slot Batch Retry Demo",
            source_pptx="template.pptx",
            default_layout_key="two_slot_text",
            layouts=[
                LayoutSpec(
                    key="two_slot_text",
                    name="Two Slot Text",
                    slide_layout_index=0,
                    supported_slide_kinds=["text"],
                    placeholders=[
                        PlaceholderSpec(
                            name="Lead Body",
                            kind=PlaceholderKind.BODY,
                            idx=14,
                            editable_role="body",
                            editable_capabilities=["text"],
                            left_emu=600000,
                            top_emu=1300000,
                            width_emu=3200000,
                            height_emu=900000,
                            text_style=TemplateTextStyleSpec(font_size_pt=18.0, line_spacing=1.18),
                        ),
                        PlaceholderSpec(
                            name="Overflow Body",
                            kind=PlaceholderKind.BODY,
                            idx=18,
                            editable_role="body",
                            editable_capabilities=["text"],
                            left_emu=4200000,
                            top_emu=1300000,
                            width_emu=3200000,
                            height_emu=900000,
                            text_style=TemplateTextStyleSpec(font_size_pt=18.0, line_spacing=1.18),
                        ),
                    ],
                )
            ],
        )
        source_parts = [
            "Префикс первого slot с операционным контекстом, длинным описанием, ограничениями и дополнительными деталями процесса.",
            "Overflow batch первый фрагмент.",
            "Overflow batch второй фрагмент.",
        ]
        plan = PresentationPlan(
            template_id="slot_batch_retry_demo",
            title="Slot Batch Retry Demo",
            slides=[
                SlideSpec(
                    kind=SlideKind.TEXT,
                    title="Batch retry",
                    text="\n".join(source_parts),
                    content_blocks=[
                        SlideContentBlock(kind=SlideContentBlockKind.PARAGRAPH, text=part)
                        for part in source_parts
                    ],
                    preferred_layout_key="two_slot_text",
                )
            ],
        )
        audit = SlideAudit(
            slide_index=1,
            title="Batch retry",
            kind=SlideKind.TEXT.value,
            layout_key="two_slot_text",
            body_char_count=sum(len(part) for part in source_parts),
            body_font_sizes=(18.0,),
            profile=TEXT_FULL_WIDTH_PROFILE,
            body_text_overflow_details=(
                TextOverflowDetail(
                    shape_name="Overflow Body",
                    shape_id=18,
                    placeholder_idx=18,
                    left=4200000,
                    top=1300000,
                    width=3200000,
                    height=900000,
                    estimated_height=860,
                    available_height=450,
                    font_size_pt=18.0,
                    text_excerpt="Overflow batch первый фрагмент. Overflow batch второй фрагмент.",
                ),
            ),
        )
        violations = [
            CapacityViolation(
                slide_index=1,
                title="Batch retry",
                rule="rendered_text_overflow",
                details="shape=Overflow Body idx=18",
            )
        ]

        retry_plan = routes_module.generation_service.plan_with_capacity_retry(plan, manifest, [audit], violations)

        self.assertIsNotNone(retry_plan)
        assert retry_plan is not None
        self.assertEqual(len(retry_plan.slides), 2)
        self.assertIn(source_parts[0], retry_plan.slides[0].text or "")
        self.assertIn(source_parts[1], retry_plan.slides[0].text or "")
        self.assertNotIn(source_parts[0], retry_plan.slides[1].text or "")
        self.assertIn(source_parts[2], retry_plan.slides[1].text or "")

    def test_capacity_retry_splits_multiple_overflow_slot_batches(self) -> None:
        manifest = TemplateManifest(
            template_id="multi_slot_retry_demo",
            display_name="Multi Slot Retry Demo",
            source_pptx="template.pptx",
            default_layout_key="two_slot_text",
            layouts=[
                LayoutSpec(
                    key="two_slot_text",
                    name="Two Slot Text",
                    slide_layout_index=0,
                    supported_slide_kinds=["text"],
                    placeholders=[
                        PlaceholderSpec(
                            name="Lead Body",
                            kind=PlaceholderKind.BODY,
                            idx=14,
                            editable_role="body",
                            editable_capabilities=["text"],
                            left_emu=600000,
                            top_emu=1300000,
                            width_emu=3300000,
                            height_emu=950000,
                            text_style=TemplateTextStyleSpec(font_size_pt=18.0, line_spacing=1.18),
                        ),
                        PlaceholderSpec(
                            name="Overflow Body",
                            kind=PlaceholderKind.BODY,
                            idx=18,
                            editable_role="body",
                            editable_capabilities=["text"],
                            left_emu=4300000,
                            top_emu=1300000,
                            width_emu=3300000,
                            height_emu=950000,
                            text_style=TemplateTextStyleSpec(font_size_pt=18.0, line_spacing=1.18),
                        ),
                    ],
                )
            ],
        )
        source_parts = [
            "Первый slot первый фрагмент с операционным контекстом.",
            "Первый slot второй фрагмент с дополнительными ограничениями.",
            "Второй slot первый фрагмент с рисками внедрения.",
            "Второй slot второй фрагмент с критериями приемки.",
        ]
        plan = PresentationPlan(
            template_id="multi_slot_retry_demo",
            title="Multi Slot Retry Demo",
            slides=[
                SlideSpec(
                    kind=SlideKind.TEXT,
                    title="Multi overflow retry",
                    text="\n".join(source_parts),
                    content_blocks=[
                        SlideContentBlock(kind=SlideContentBlockKind.PARAGRAPH, text=part)
                        for part in source_parts
                    ],
                    preferred_layout_key="two_slot_text",
                )
            ],
        )
        audit = SlideAudit(
            slide_index=1,
            title="Multi overflow retry",
            kind=SlideKind.TEXT.value,
            layout_key="two_slot_text",
            body_char_count=sum(len(part) for part in source_parts),
            body_font_sizes=(18.0,),
            profile=TEXT_FULL_WIDTH_PROFILE,
            body_text_overflow_details=(
                TextOverflowDetail(
                    shape_name="Lead Body",
                    shape_id=14,
                    placeholder_idx=14,
                    left=600000,
                    top=1300000,
                    width=3300000,
                    height=950000,
                    estimated_height=900,
                    available_height=520,
                    font_size_pt=18.0,
                    text_excerpt=" ".join(source_parts[:2]),
                ),
                TextOverflowDetail(
                    shape_name="Overflow Body",
                    shape_id=18,
                    placeholder_idx=18,
                    left=4300000,
                    top=1300000,
                    width=3300000,
                    height=950000,
                    estimated_height=900,
                    available_height=520,
                    font_size_pt=18.0,
                    text_excerpt=" ".join(source_parts[2:]),
                ),
            ),
        )
        violations = [
            CapacityViolation(
                slide_index=1,
                title="Multi overflow retry",
                rule="rendered_text_overflow",
                details="shape=Lead Body idx=14; shape=Overflow Body idx=18",
            )
        ]

        retry_plan = routes_module.generation_service.plan_with_capacity_retry(plan, manifest, [audit], violations)

        self.assertIsNotNone(retry_plan)
        assert retry_plan is not None
        self.assertGreaterEqual(len(retry_plan.slides), 3)
        rendered_texts = [slide.text or "" for slide in retry_plan.slides]
        self.assertFalse(any(source_parts[0] in text and source_parts[1] in text for text in rendered_texts))
        self.assertFalse(any(source_parts[2] in text and source_parts[3] in text for text in rendered_texts))

    def test_text_pack_result_reports_failed_box_and_chunk(self) -> None:
        box = FitBox(
            metrics=TextBoxMetrics(width_emu=1800000, height_emu=650000, font_size_pt=18.0),
            min_font_pt=12,
            max_font_pt=18,
        )

        result = assign_text_chunks_to_boxes(
            [
                "Первый компактный фрагмент.",
                "Второй фрагмент намеренно длиннее и должен не поместиться в единственный target box.",
            ],
            [box],
        )

        self.assertFalse(result.fits)
        self.assertEqual(result.failed_box_index, 0)
        self.assertEqual(result.failed_chunk_index, 1)
        self.assertIn(result.reason, {"chunk_too_large", "no_remaining_box"})

    def test_slide_text_policy_prefers_content_blocks_for_demand_and_chunks(self) -> None:
        slide = SlideSpec(
            kind=SlideKind.TEXT,
            title="Policy",
            text="Fallback text should not be counted when content blocks exist.",
            bullets=["Fallback bullet"],
            content_blocks=[
                SlideContentBlock(kind=SlideContentBlockKind.PARAGRAPH, text="First paragraph."),
                SlideContentBlock(kind=SlideContentBlockKind.BULLET_LIST, items=["First bullet", "Second bullet"]),
            ],
        )

        chunks = slide_content_chunks(slide)

        self.assertEqual(
            chunks,
            [
                ("paragraph", "First paragraph."),
                ("bullet", "First bullet"),
                ("bullet", "Second bullet"),
            ],
        )
        self.assertEqual(slide_text_demand_chars(slide), len("First paragraph.") + len("First bullet") + len("Second bullet"))

    def test_capacity_audit_reports_native_table_cell_text_overflow(self) -> None:
        violations = find_capacity_violations(
            [
                SlideAudit(
                    slide_index=1,
                    title="Table",
                    kind=SlideKind.TABLE.value,
                    layout_key="table",
                    body_char_count=0,
                    body_font_sizes=(),
                    profile=TEXT_FULL_WIDTH_PROFILE,
                    has_table=True,
                    table_cell_overflow_count=1,
                    table_cell_overflow_details=("row=1 col=0: required=900 available=300",),
                )
            ]
        )

        self.assertIn("table_cell_text_overflow", {item.rule for item in violations})

    def test_target_split_uses_assignment_failure_boundary(self) -> None:
        box = FitBox(
            metrics=TextBoxMetrics(width_emu=2400000, height_emu=700000, font_size_pt=18.0),
            min_font_pt=12,
            max_font_pt=18,
        )
        chunks = [("paragraph", "Alpha короткий текст.") for _ in range(5)]

        batches = routes_module.template_registry._split_chunks_by_target_assignment(chunks, [box])

        self.assertIsNotNone(batches)
        assert batches is not None
        self.assertEqual([len(batch) for batch in batches], [2, 2, 1])

    def test_generate_returns_404_for_template_without_pptx(self) -> None:
        with self.assertRaises(HTTPException) as error:
            routes_module.generate_presentation(
                PresentationPlan(
                    template_id="missing_source",
                    title="Broken Template",
                    slides=[SlideSpec(kind=SlideKind.TITLE, title="Broken Template")],
                )
            )

        self.assertEqual(error.exception.status_code, 404)
        self.assertIn("Template PPTX not found", error.exception.detail)

    def test_upload_template_rejects_manifest_path_traversal(self) -> None:
        manifest = TemplateManifest(
            template_id="..\\..\\outside",
            display_name="Broken",
            source_pptx="template.pptx",
        )
        upload = UploadFile(
            filename="template.pptx",
            file=BytesIO(b"not-a-real-pptx"),
        )

        with self.assertRaises(HTTPException) as error:
            asyncio.run(
                routes_module.upload_template(
                    manifest_json=manifest.model_dump_json(),
                    template_file=upload,
                )
            )

        self.assertEqual(error.exception.status_code, 400)
        self.assertIn("escapes the storage root", error.exception.detail)

    def test_upload_template_rejects_empty_pptx(self) -> None:
        manifest = TemplateManifest(
            template_id="empty_template",
            display_name="Empty",
            source_pptx="template.pptx",
        )
        upload = UploadFile(filename="template.pptx", file=BytesIO(b""))

        with self.assertRaises(HTTPException) as error:
            asyncio.run(
                routes_module.upload_template(
                    manifest_json=manifest.model_dump_json(),
                    template_file=upload,
                )
            )

        self.assertEqual(error.exception.status_code, 400)
        self.assertIn("template_file is empty", error.exception.detail)

    def test_upload_template_auto_rejects_empty_pptx(self) -> None:
        upload = UploadFile(filename="empty-template.pptx", file=BytesIO(b""))

        with self.assertRaises(HTTPException) as error:
            asyncio.run(
                routes_module.upload_template_auto(
                    template_id="empty_auto",
                    display_name="Empty Auto",
                    description=None,
                    template_file=upload,
                )
            )

        self.assertEqual(error.exception.status_code, 400)
        self.assertIn("template_file is empty", error.exception.detail)

    def test_upload_template_auto_returns_inventory_contract(self) -> None:
        template_path = self._templates_dir / "deterministic_layout_fixture" / "template.pptx"
        upload = UploadFile(filename="auto-template.pptx", file=BytesIO(template_path.read_bytes()))

        response = asyncio.run(
            routes_module.upload_template_auto(
                template_id="uploaded_api_auto",
                display_name="Uploaded API Auto",
                description="Contract check",
                template_file=upload,
            )
        )

        self.assertTrue(response.analyzed)
        self.assertIn(response.inventory_summary.usability_status, {"usable", "usable_with_degradation"})
        self.assertTrue(response.editable_targets)
        self.assertTrue(response.detected_components)

    def test_analyze_template_returns_inventory_contract(self) -> None:
        response = routes_module.analyze_template("deterministic_layout_fixture")

        self.assertEqual(response.template_id, "deterministic_layout_fixture")
        self.assertIn(
            response.inventory_summary.usability_status,
            {"usable", "usable_with_degradation", "not_safely_editable"},
        )
        self.assertTrue(response.editable_targets)
        self.assertIsNotNone(response.detected_components)

    def test_plan_from_text_with_template_rejects_invalid_payload_json(self) -> None:
        template_path = self._templates_dir / "deterministic_layout_fixture" / "template.pptx"
        upload = UploadFile(filename="customer-template.pptx", file=BytesIO(template_path.read_bytes()))

        with self.assertRaises(HTTPException) as error:
            asyncio.run(
                routes_module.plan_from_text_with_template(
                    payload_json="{not valid json}",
                    template_file=upload,
                )
            )

        self.assertEqual(error.exception.status_code, 400)
        self.assertIn("Invalid payload_json", error.exception.detail)

    def test_plan_from_text_with_template_returns_degradation_error_for_broken_template(self) -> None:
        upload = UploadFile(filename="broken-template.pptx", file=BytesIO(b"not-a-real-pptx"))
        payload = TextPlanRequest(
            template_id="ignored_template_id",
            title="Demo",
            raw_text="Основные выводы\n- Рост выручки\n- Снижение churn",
        )

        with patch.object(routes_module.analyzer, "analyze", side_effect=ValueError("broken template")):
            with self.assertRaises(HTTPException) as error:
                asyncio.run(
                    routes_module.plan_from_text_with_template(
                        payload_json=payload.model_dump_json(),
                        template_file=upload,
                    )
                )

        self.assertEqual(error.exception.status_code, 400)
        self.assertIn("Failed to analyze uploaded template", error.exception.detail)


if __name__ == "__main__":
    unittest.main()
