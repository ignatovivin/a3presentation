from __future__ import annotations

import asyncio
import tempfile
import unittest
from io import BytesIO
from pathlib import Path

from docx import Document
from pptx import Presentation
from starlette.datastructures import UploadFile

from a3presentation.api import routes as routes_module
from a3presentation.domain.api import TextPlanRequest
from a3presentation.domain.chart import ChartConfidence, ChartSeries, ChartSpec, ChartType
from a3presentation.domain.presentation import PresentationPlan, SlideKind, SlideSpec, TableBlock
from a3presentation.services.deck_audit import (
    audit_generated_presentation,
    continuation_groups,
    find_capacity_violations,
)
from a3presentation.services.document_text_extractor import DocumentTextExtractor
from a3presentation.services.layout_capacity import (
    LIST_FULL_WIDTH_PROFILE,
    TEXT_FULL_WIDTH_PROFILE,
    profile_for_layout,
)
from a3presentation.services.planner import TextToPlanService
from a3presentation.services.pptx_generator import PptxGenerator
from a3presentation.services.template_analyzer import TemplateAnalyzer
from a3presentation.services.template_registry import TemplateRegistry
from a3presentation.settings import get_settings


class ProjectContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.settings = get_settings()
        cls.registry = TemplateRegistry(cls.settings.templates_dir)
        cls.analyzer = TemplateAnalyzer()
        cls.extractor = DocumentTextExtractor()
        cls.planner = TextToPlanService()
        cls.generator = PptxGenerator()

    def test_template_registry_entries_are_internally_consistent(self) -> None:
        manifests = self.registry.list_templates()

        self.assertGreaterEqual(len(manifests), 1)
        self.assertEqual(len({manifest.template_id for manifest in manifests}), len(manifests))

        for manifest in manifests:
            with self.subTest(template_id=manifest.template_id):
                template_dir = self.settings.templates_dir / manifest.template_id
                self.assertTrue(template_dir.exists())
                self.assertTrue((template_dir / "manifest.json").exists())
                self.assertEqual(manifest.template_id, template_dir.name)

                details = routes_module.get_template(manifest.template_id)
                self.assertEqual(details.manifest.template_id, manifest.template_id)

                template_path = template_dir / manifest.source_pptx
                self.assertEqual(details.has_template_file, template_path.exists())
                if template_path.exists():
                    self.assertEqual(self.registry.get_template_pptx_path(manifest.template_id), template_path)

    def test_template_manifests_have_valid_internal_structure(self) -> None:
        manifests = self.registry.list_templates()

        for manifest in manifests:
            with self.subTest(template_id=manifest.template_id):
                self.assertTrue(manifest.template_id.strip())
                self.assertTrue(manifest.display_name.strip())
                self.assertTrue(manifest.source_pptx.lower().endswith(".pptx"))

                layout_keys = [layout.key for layout in manifest.layouts]
                self.assertEqual(len(set(layout_keys)), len(layout_keys))

                if manifest.layouts:
                    self.assertTrue(manifest.default_layout_key)
                    self.assertIn(manifest.default_layout_key, layout_keys)

                for layout in manifest.layouts:
                    self.assertTrue(layout.key.strip())
                    self.assertTrue(layout.name.strip())
                    self.assertGreaterEqual(layout.slide_layout_index, 0)
                    self.assertEqual(
                        len(set(layout.supported_slide_kinds)),
                        len(layout.supported_slide_kinds),
                    )

                for prototype in manifest.prototype_slides:
                    self.assertTrue(prototype.key.strip())
                    self.assertGreaterEqual(prototype.source_slide_index, 0)
                    self.assertEqual(
                        len(set(prototype.supported_slide_kinds)),
                        len(prototype.supported_slide_kinds),
                    )

    def test_every_template_with_pptx_supports_smoke_generation(self) -> None:
        manifests = self.registry.list_templates()

        for manifest in manifests:
            template_path = self.settings.templates_dir / manifest.template_id / manifest.source_pptx
            if not template_path.exists():
                continue

            with self.subTest(template_id=manifest.template_id):
                plan = PresentationPlan(
                    template_id=manifest.template_id,
                    title=f"{manifest.display_name} Smoke",
                    slides=self._smoke_slides_for_manifest(manifest),
                )

                with tempfile.TemporaryDirectory() as temp_dir:
                    output_path = self.generator.generate(
                        template_path=template_path,
                        manifest=manifest,
                        plan=plan,
                        output_dir=Path(temp_dir),
                    )
                    presentation = Presentation(str(output_path))
                    self.assertEqual(len(presentation.slides), len(plan.slides))

    def test_template_analyzer_output_stays_compatible_with_generator(self) -> None:
        manifests = self.registry.list_templates()

        for manifest in manifests:
            template_path = self.settings.templates_dir / manifest.template_id / manifest.source_pptx
            if not template_path.exists():
                continue

            with self.subTest(template_id=manifest.template_id):
                analyzed_manifest = self.analyzer.analyze(
                    template_id=manifest.template_id,
                    template_path=template_path,
                    display_name=manifest.display_name,
                )
                self.assertEqual(analyzed_manifest.template_id, manifest.template_id)
                self.assertEqual(analyzed_manifest.source_pptx, template_path.name)

                plan = PresentationPlan(
                    template_id=manifest.template_id,
                    title=f"{manifest.display_name} Analyzed Smoke",
                    slides=self._smoke_slides_for_manifest(analyzed_manifest),
                )

                with tempfile.TemporaryDirectory() as temp_dir:
                    output_path = self.generator.generate(
                        template_path=template_path,
                        manifest=analyzed_manifest,
                        plan=plan,
                        output_dir=Path(temp_dir),
                    )
                    presentation = Presentation(str(output_path))
                    self.assertEqual(len(presentation.slides), len(plan.slides))

    def test_full_pipeline_contract_for_mixed_docx_document(self) -> None:
        document = Document()
        document.add_paragraph("A3")
        document.add_paragraph("Смешанный стратегический документ")
        document.add_heading("1. Основные выводы", level=1)
        document.add_paragraph(
            "Рост выручки обеспечивается за счет новых сегментов, улучшения конверсии и масштабируемой платформы."
        )
        document.add_paragraph(style="List Bullet").add_run("Приоритет 1: рост B2B сегмента")
        document.add_paragraph(style="List Bullet").add_run("Приоритет 2: снижение концентрации выручки")
        table = document.add_table(rows=3, cols=2)
        table.cell(0, 0).text = "Показатель"
        table.cell(0, 1).text = "Значение"
        table.cell(1, 0).text = "GMV"
        table.cell(1, 1).text = "125"
        table.cell(2, 0).text = "NPS"
        table.cell(2, 1).text = "61"
        document.add_heading("2. Следующий раздел", level=1)
        document.add_paragraph("Дополнительный narrative-блок для устойчивой классификации документа.")

        buffer = BytesIO()
        document.save(buffer)
        content = buffer.getvalue()

        text, tables, blocks = self.extractor.extract("mixed-contract.docx", content)
        plan = self.planner.build_plan("corp_light_v1", text, None, tables, blocks)

        self.assertGreaterEqual(len(plan.slides), 3)
        self.assertTrue(any(slide.kind == SlideKind.TABLE for slide in plan.slides))
        self.assertTrue(any(slide.kind in {SlideKind.TEXT, SlideKind.BULLETS} for slide in plan.slides[1:]))

        manifest = self.registry.get_template("corp_light_v1")
        template_path = self.registry.get_template_pptx_path("corp_light_v1")
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = self.generator.generate(
                template_path=template_path,
                manifest=manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            presentation = Presentation(str(output_path))
            self.assertEqual(len(presentation.slides), len(plan.slides))

    def test_api_roundtrip_contract_keeps_extract_plan_generate_download_in_sync(self) -> None:
        document = Document()
        document.add_paragraph("Стратегический отчёт")
        document.add_heading("Рынок", level=1)
        document.add_paragraph("Компания усиливает позиции за счет платформенных интеграций.")
        table = document.add_table(rows=3, cols=2)
        table.cell(0, 0).text = "Сегмент"
        table.cell(0, 1).text = "Выручка"
        table.cell(1, 0).text = "SMB"
        table.cell(1, 1).text = "120"
        table.cell(2, 0).text = "Enterprise"
        table.cell(2, 1).text = "250"

        buffer = BytesIO()
        document.save(buffer)
        upload = UploadFile(filename="api-contract.docx", file=BytesIO(buffer.getvalue()))

        extracted = asyncio.run(routes_module.extract_document_text(upload))
        self.assertEqual(len(extracted.tables), 1)
        self.assertEqual(len(extracted.chart_assessments), 1)

        plan = routes_module.plan_from_text(
            TextPlanRequest(
                template_id="corp_light_v1",
                title="API Contract",
                raw_text=extracted.text,
                tables=extracted.tables,
                blocks=extracted.blocks,
            )
        )
        self.assertGreaterEqual(len(plan.slides), 2)

        generated = routes_module.generate_presentation(plan)
        downloaded = routes_module.download_presentation(generated.file_name)
        self.assertEqual(Path(downloaded.path).name, generated.file_name)

    def test_layout_capacity_profiles_are_consistent_with_planner_contract(self) -> None:
        self.assertEqual(profile_for_layout("list_full_width"), LIST_FULL_WIDTH_PROFILE)
        self.assertEqual(profile_for_layout("text_full_width"), TEXT_FULL_WIDTH_PROFILE)
        self.assertEqual(self.planner.list_profile.max_items, self.planner.LIST_BATCH_SIZE)
        self.assertEqual(self.planner.list_profile.max_weight, self.planner.LIST_SLIDE_MAX_WEIGHT)
        self.assertEqual(self.planner.text_profile.max_chars, self.planner.TEXT_SLIDE_MAX_CHARS)
        self.assertEqual(self.planner.text_profile.max_primary_chars, self.planner.TEXT_PRIMARY_MAX_CHARS)
        self.assertLessEqual(self.planner.text_profile.min_font_pt, self.planner.text_profile.max_font_pt)
        self.assertLessEqual(self.planner.list_profile.min_font_pt, self.planner.list_profile.max_font_pt)

    def test_deck_audit_reports_body_font_sizes_within_layout_profile_bounds(self) -> None:
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="Audit Contract",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="Audit Contract", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.BULLETS,
                    title="Dense bullets",
                    bullets=[
                        "Первый длинный пункт объясняет стратегическую логику и ограничения.",
                        "Второй длинный пункт добавляет риски, сроки и организационные последствия.",
                        "Третий длинный пункт описывает KPI, unit-экономику и инфраструктурные требования.",
                        "Четвёртый длинный пункт связывает выводы с дорожной картой внедрения.",
                    ],
                    preferred_layout_key="list_full_width",
                ),
                SlideSpec(
                    kind=SlideKind.TEXT,
                    title="Dense text",
                    text=(
                        "Первый длинный абзац описывает контекст, ограничения, допущения и критерии принятия решения. "
                        "Второй длинный абзац добавляет финансовые ориентиры и риски реализации. "
                        "Третий длинный абзац связывает выводы с KPI и дорожной картой."
                    ),
                    preferred_layout_key="text_full_width",
                ),
            ],
        )

        manifest = self.registry.get_template("corp_light_v1")
        template_path = self.registry.get_template_pptx_path("corp_light_v1")
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = self.generator.generate(
                template_path=template_path,
                manifest=manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            audits = audit_generated_presentation(output_path, plan)

        self.assertGreaterEqual(len(audits), 2)
        for audit in audits:
            with self.subTest(slide=audit.slide_index):
                self.assertTrue(audit.body_font_sizes)
                self.assertGreaterEqual(min(audit.body_font_sizes), audit.profile.min_font_pt)
                self.assertLessEqual(max(audit.body_font_sizes), audit.profile.max_font_pt)

    def test_deck_audit_detects_continuation_groups_for_multislide_sections(self) -> None:
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="Audit Continuations",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="Audit Continuations", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.BULLETS,
                    title="Раздел",
                    bullets=["Пункт 1", "Пункт 2", "Пункт 3", "Пункт 4", "Пункт 5", "Пункт 6"],
                    preferred_layout_key="list_full_width",
                ),
                SlideSpec(
                    kind=SlideKind.BULLETS,
                    title="Раздел (2)",
                    bullets=["Пункт 7", "Пункт 8"],
                    preferred_layout_key="list_full_width",
                ),
            ],
        )

        manifest = self.registry.get_template("corp_light_v1")
        template_path = self.registry.get_template_pptx_path("corp_light_v1")
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = self.generator.generate(
                template_path=template_path,
                manifest=manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            audits = audit_generated_presentation(output_path, plan)

        groups = continuation_groups(audits)
        self.assertIn("Раздел", groups)
        self.assertEqual(len(groups["Раздел"]), 2)

    def test_deck_audit_flags_underfilled_continuation_pairs(self) -> None:
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="Audit Violations",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="Audit Violations", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.BULLETS,
                    title="Раздел",
                    bullets=[
                        "Первый длинный пункт подробно описывает стратегическую инициативу и контекст принятия решения.",
                        "Второй длинный пункт раскрывает риски, ресурсы, ограничения и организационные зависимости.",
                        "Третий длинный пункт связывает инициативу с метриками, сроками и критериями качества.",
                        "Четвёртый длинный пункт объясняет изменения процесса и требования к операционной модели.",
                        "Пятый длинный пункт добавляет детали по продукту, рынку и коммерческому эффекту.",
                        "Шестой длинный пункт завершает блок ожидаемыми результатами и контрольными точками.",
                    ],
                    preferred_layout_key="list_full_width",
                ),
                SlideSpec(
                    kind=SlideKind.BULLETS,
                    title="Раздел (2)",
                    bullets=["Короткий хвост."],
                    preferred_layout_key="list_full_width",
                ),
            ],
        )

        manifest = self.registry.get_template("corp_light_v1")
        template_path = self.registry.get_template_pptx_path("corp_light_v1")
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = self.generator.generate(
                template_path=template_path,
                manifest=manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            audits = audit_generated_presentation(output_path, plan)

        violations = find_capacity_violations(audits)
        violation_rules = {violation.rule for violation in violations}
        self.assertIn("continuation_balance", violation_rules)
        self.assertIn("underfilled_continuation", violation_rules)

    def test_deck_audit_accepts_balanced_dense_slides_without_capacity_violations(self) -> None:
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="Audit Healthy",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="Audit Healthy", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.BULLETS,
                    title="Сбалансированный список",
                    bullets=[
                        "Пункт 1 подробно описывает цель и ограничения.",
                        "Пункт 2 раскрывает ресурсы и зависимости.",
                        "Пункт 3 фиксирует KPI и ожидаемый эффект.",
                        "Пункт 4 связывает решение с дорожной картой.",
                    ],
                    preferred_layout_key="list_full_width",
                ),
                SlideSpec(
                    kind=SlideKind.TEXT,
                    title="Сбалансированный текст",
                    text=(
                        "Первый абзац задаёт контекст и основные допущения. "
                        "Второй абзац описывает ожидаемый эффект и критерии контроля. "
                        "Третий абзац связывает решение с финансовыми и операционными метриками."
                    ),
                    preferred_layout_key="text_full_width",
                ),
            ],
        )

        manifest = self.registry.get_template("corp_light_v1")
        template_path = self.registry.get_template_pptx_path("corp_light_v1")
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = self.generator.generate(
                template_path=template_path,
                manifest=manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            audits = audit_generated_presentation(output_path, plan)

        violations = find_capacity_violations(audits)
        self.assertEqual(violations, [])

    def test_deck_audit_validates_table_layout_geometry(self) -> None:
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="Audit Table Geometry",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="Audit Table Geometry", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.TABLE,
                    title="Ключевые показатели",
                    subtitle="Таблица должна занимать рабочую ширину layout",
                    table=TableBlock(
                        headers=["Показатель", "Значение"],
                        rows=[["Выручка", "120"], ["Маржа", "24%"], ["NPS", "61"]],
                    ),
                    preferred_layout_key="table",
                ),
            ],
        )

        manifest = self.registry.get_template("corp_light_v1")
        template_path = self.registry.get_template_pptx_path("corp_light_v1")
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = self.generator.generate(
                template_path=template_path,
                manifest=manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            audits = audit_generated_presentation(output_path, plan)

        table_audit = next(audit for audit in audits if audit.kind == SlideKind.TABLE.value)
        self.assertTrue(table_audit.has_table)
        self.assertGreaterEqual(table_audit.content_width_ratio, 0.9)
        self.assertGreaterEqual(table_audit.footer_width_ratio, 0.9)
        violations = find_capacity_violations(audits)
        self.assertEqual(violations, [])

    def test_deck_audit_validates_chart_layout_geometry(self) -> None:
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="Audit Chart Geometry",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="Audit Chart Geometry", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.CHART,
                    title="Выручка по каналам",
                    subtitle="График должен занимать рабочую ширину layout",
                    chart=ChartSpec(
                        chart_id="chart_geometry",
                        source_table_id="table_1",
                        chart_type=ChartType.COLUMN,
                        title="Выручка",
                        categories=["SEO", "Ads", "Referral"],
                        series=[ChartSeries(name="Выручка", values=[120.0, 200.0, 90.0])],
                        confidence=ChartConfidence.HIGH,
                    ),
                    preferred_layout_key="table",
                ),
            ],
        )

        manifest = self.registry.get_template("corp_light_v1")
        template_path = self.registry.get_template_pptx_path("corp_light_v1")
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = self.generator.generate(
                template_path=template_path,
                manifest=manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            audits = audit_generated_presentation(output_path, plan)

        chart_audit = next(audit for audit in audits if audit.kind == SlideKind.CHART.value)
        self.assertTrue(chart_audit.has_chart)
        self.assertGreaterEqual(chart_audit.content_width_ratio, 0.9)
        self.assertGreaterEqual(chart_audit.footer_width_ratio, 0.9)
        violations = find_capacity_violations(audits)
        self.assertEqual(violations, [])

    def test_deck_audit_validates_image_layout_geometry(self) -> None:
        small_png_base64 = (
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO9W6i8AAAAASUVORK5CYII="
        )
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="Audit Image Geometry",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="Audit Image Geometry", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.IMAGE,
                    title="Схема процесса",
                    text="Изображение должно рендериться как picture shape и сохранять рабочую геометрию layout.",
                    preferred_layout_key="image_text",
                    image_base64=small_png_base64,
                    image_content_type="image/png",
                ),
            ],
        )

        manifest = self.registry.get_template("corp_light_v1")
        template_path = self.registry.get_template_pptx_path("corp_light_v1")
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = self.generator.generate(
                template_path=template_path,
                manifest=manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            audits = audit_generated_presentation(output_path, plan)

        image_audit = next(audit for audit in audits if audit.kind == SlideKind.IMAGE.value)
        self.assertTrue(image_audit.has_image)
        self.assertGreaterEqual(image_audit.content_width_ratio, 0.35)
        violations = find_capacity_violations(audits)
        self.assertEqual(violations, [])

    def _smoke_slides_for_manifest(self, manifest) -> list[SlideSpec]:
        slides = [SlideSpec(kind=SlideKind.TITLE, title=f"{manifest.display_name} Smoke", preferred_layout_key="cover")]
        supported_kinds = self._supported_kinds(manifest)

        if SlideKind.TEXT.value in supported_kinds:
            slides.append(
                SlideSpec(
                    kind=SlideKind.TEXT,
                    title="Smoke text",
                    text="Smoke generation validates runtime compatibility between manifest and generator.",
                )
            )
            return slides

        if SlideKind.BULLETS.value in supported_kinds:
            slides.append(
                SlideSpec(
                    kind=SlideKind.BULLETS,
                    title="Smoke bullets",
                    bullets=["Smoke generation", "Manifest compatibility", "Generator compatibility"],
                )
            )
            return slides

        if SlideKind.TABLE.value in supported_kinds:
            slides.append(
                SlideSpec(
                    kind=SlideKind.TABLE,
                    title="Smoke table",
                    table=TableBlock(headers=["Metric", "Value"], rows=[["GMV", "125"]]),
                )
            )
            return slides

        slides.append(
            SlideSpec(
                kind=SlideKind.TEXT,
                title="Fallback smoke",
                text="Fallback slide for templates without explicit text support.",
            )
        )
        return slides

    def _supported_kinds(self, manifest) -> set[str]:
        supported: set[str] = set()
        for layout in manifest.layouts:
            supported.update(layout.supported_slide_kinds)
        for prototype in manifest.prototype_slides:
            supported.update(prototype.supported_slide_kinds)
        return supported


if __name__ == "__main__":
    unittest.main()
