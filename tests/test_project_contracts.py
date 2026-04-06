from __future__ import annotations

import asyncio
import tempfile
import unittest
from io import BytesIO
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pptx import Presentation
from pptx.dml.color import RGBColor
from starlette.datastructures import UploadFile

from a3presentation.api import routes as routes_module
from a3presentation.domain.api import TextPlanRequest
from a3presentation.domain.chart import ChartConfidence, ChartSeries, ChartSpec, ChartType
from a3presentation.domain.presentation import (
    PresentationPlan,
    SlideContentBlock,
    SlideContentBlockKind,
    SlideKind,
    SlideSpec,
    TableBlock,
)
from a3presentation.domain.template import GenerationMode, PlaceholderKind
from a3presentation.services.deck_audit import (
    SlideAudit,
    audit_generated_presentation,
    continuation_groups,
    find_capacity_violations,
)
from a3presentation.services.document_text_extractor import DocumentTextExtractor
from a3presentation.services.layout_capacity import (
    LIST_FULL_WIDTH_PROFILE,
    TEXT_FULL_WIDTH_PROFILE,
    derive_capacity_profile_for_geometry,
    geometry_policy_for_layout,
    profile_for_layout,
    spacing_policy_for_layout,
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

    def test_template_analyzer_extracts_placeholder_geometry_for_uploaded_layouts(self) -> None:
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

                analyzed_placeholders = [
                    placeholder
                    for layout in analyzed_manifest.layouts
                    for placeholder in layout.placeholders
                ]
                self.assertTrue(analyzed_placeholders)
                self.assertTrue(
                    any(self._has_geometry_metadata(placeholder) for placeholder in analyzed_placeholders)
                )
                self.assertTrue(
                    any(
                        placeholder.kind in {PlaceholderKind.TITLE, PlaceholderKind.SUBTITLE, PlaceholderKind.BODY}
                        and self._has_text_margin_metadata(placeholder)
                        for placeholder in analyzed_placeholders
                    )
                )

    def test_template_analyzer_extracts_shape_geometry_for_uploaded_prototype_tokens(self) -> None:
        manifests = self.registry.list_templates()

        for manifest in manifests:
            template_path = self.settings.templates_dir / manifest.template_id / manifest.source_pptx
            if not template_path.exists():
                continue

            analyzed_manifest = self.analyzer.analyze(
                template_id=manifest.template_id,
                template_path=template_path,
                display_name=manifest.display_name,
            )
            if analyzed_manifest.generation_mode.value != "prototype":
                continue

            with self.subTest(template_id=manifest.template_id):
                tokens = [
                    token
                    for prototype in analyzed_manifest.prototype_slides
                    for token in prototype.tokens
                ]
                self.assertTrue(tokens)
                self.assertTrue(any(self._has_geometry_metadata(token) for token in tokens))
                self.assertTrue(
                    any(
                        token.binding in {"title", "subtitle", "text", "body", "main_text", "secondary_text", "cover_title"}
                        and self._has_text_margin_metadata(token)
                        for token in tokens
                    )
                )

    def test_generator_applies_analyzer_geometry_metadata_for_uploaded_layout_templates(self) -> None:
        template_id = "razmeshchenie_soglasiy"
        template_path = self.settings.templates_dir / template_id / "template.pptx"
        analyzed_manifest = self.analyzer.analyze(
            template_id=template_id,
            template_path=template_path,
            display_name="Размещение согласий",
        )
        analyzed_manifest.generation_mode = GenerationMode.LAYOUT
        layout = next(layout for layout in analyzed_manifest.layouts if "text" in layout.supported_slide_kinds)
        body_placeholder = next(placeholder for placeholder in layout.placeholders if placeholder.kind == PlaceholderKind.BODY)
        body_placeholder.left_emu = 777777
        body_placeholder.top_emu = 1888888
        body_placeholder.width_emu = 5555555
        body_placeholder.height_emu = 2222222
        body_placeholder.margin_left_emu = 11111
        body_placeholder.margin_right_emu = 22222
        body_placeholder.margin_top_emu = 33333
        body_placeholder.margin_bottom_emu = 44444

        plan = PresentationPlan(
            template_id=template_id,
            title="Uploaded Layout Geometry",
            slides=[
                SlideSpec(
                    kind=SlideKind.TEXT,
                    title="",
                    text="Текст должен следовать analyzer-derived placeholder metadata.",
                    preferred_layout_key=layout.key,
                )
            ],
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = self.generator.generate(
                template_path=template_path,
                manifest=analyzed_manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            presentation = Presentation(str(output_path))

        slide = presentation.slides[0]
        shape = next(placeholder for placeholder in slide.placeholders if placeholder.placeholder_format.idx == body_placeholder.idx)
        self.assertEqual(shape.left, body_placeholder.left_emu)
        self.assertEqual(shape.top, body_placeholder.top_emu)
        self.assertEqual(shape.width, body_placeholder.width_emu)
        self.assertEqual(shape.height, body_placeholder.height_emu)
        self.assertEqual(shape.text_frame.margin_left, body_placeholder.margin_left_emu)
        self.assertEqual(shape.text_frame.margin_right, body_placeholder.margin_right_emu)
        self.assertEqual(shape.text_frame.margin_top, body_placeholder.margin_top_emu)
        self.assertEqual(shape.text_frame.margin_bottom, body_placeholder.margin_bottom_emu)

    def test_deck_audit_uses_analyzer_geometry_metadata_for_uploaded_layout_templates(self) -> None:
        template_id = "razmeshchenie_soglasiy"
        template_path = self.settings.templates_dir / template_id / "template.pptx"
        analyzed_manifest = self.analyzer.analyze(
            template_id=template_id,
            template_path=template_path,
            display_name="Размещение согласий",
        )
        analyzed_manifest.generation_mode = GenerationMode.LAYOUT
        layout = next(layout for layout in analyzed_manifest.layouts if "text" in layout.supported_slide_kinds)
        body_placeholder = next(placeholder for placeholder in layout.placeholders if placeholder.kind == PlaceholderKind.BODY)
        body_placeholder.left_emu = 777777
        body_placeholder.top_emu = 1888888
        body_placeholder.width_emu = 5555555
        body_placeholder.height_emu = 2222222
        body_placeholder.margin_left_emu = 11111
        body_placeholder.margin_right_emu = 22222
        body_placeholder.margin_top_emu = 33333
        body_placeholder.margin_bottom_emu = 44444

        plan = PresentationPlan(
            template_id=template_id,
            title="Uploaded Layout Geometry",
            slides=[
                SlideSpec(
                    kind=SlideKind.TEXT,
                    title="",
                    text="Текст должен проходить template-aware deck audit на analyzer-derived layout metadata.",
                    preferred_layout_key=layout.key,
                )
            ],
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = self.generator.generate(
                template_path=template_path,
                manifest=analyzed_manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            audits = audit_generated_presentation(output_path, plan, analyzed_manifest)
            violations = find_capacity_violations(audits)

        self.assertTrue(audits)
        self.assertFalse(any(item.rule == "body_left_misalignment" for item in violations))
        self.assertFalse(any(item.rule == "body_margin_mismatch" for item in violations))
        self.assertLess(audits[0].profile.max_chars, TEXT_FULL_WIDTH_PROFILE.max_chars)

    def test_generator_applies_manifest_geometry_metadata_for_uploaded_prototype_templates(self) -> None:
        template_id = "razmeshchenie_soglasiy"
        template_path = self.settings.templates_dir / template_id / "template.pptx"
        manifest = self.registry.get_template(template_id)
        prototype = manifest.prototype_slides[0]
        bound_text_token = next(
            token
            for token in prototype.tokens
            if token.binding in {"cover_meta", "subtitle", "text", "body", "main_text", "secondary_text", "presentation_name"}
        )
        bound_text_token.left_emu = 999999
        bound_text_token.top_emu = 2111111
        bound_text_token.width_emu = 4444444
        bound_text_token.height_emu = 1234567
        bound_text_token.margin_left_emu = 21000
        bound_text_token.margin_right_emu = 22000
        bound_text_token.margin_top_emu = 23000
        bound_text_token.margin_bottom_emu = 24000

        plan = PresentationPlan(
            template_id=template_id,
            title="Uploaded Prototype Geometry",
            slides=[
                SlideSpec(
                    kind=SlideKind.TITLE,
                    title="Заголовок",
                    notes="Метаданные prototype token shape должны примениться к cloned slide.",
                    preferred_layout_key=prototype.key,
                )
            ],
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = self.generator.generate(
                template_path=template_path,
                manifest=manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            presentation = Presentation(str(output_path))

        slide = presentation.slides[0]
        shape = next(shape for shape in slide.shapes if shape.name == bound_text_token.shape_name)
        self.assertEqual(shape.left, bound_text_token.left_emu)
        self.assertEqual(shape.top, bound_text_token.top_emu)
        self.assertEqual(shape.width, bound_text_token.width_emu)
        self.assertEqual(shape.height, bound_text_token.height_emu)
        self.assertEqual(shape.text_frame.margin_left, bound_text_token.margin_left_emu)
        self.assertEqual(shape.text_frame.margin_right, bound_text_token.margin_right_emu)
        self.assertEqual(shape.text_frame.margin_top, bound_text_token.margin_top_emu)
        self.assertEqual(shape.text_frame.margin_bottom, bound_text_token.margin_bottom_emu)

    def test_deck_audit_uses_manifest_geometry_metadata_for_uploaded_prototype_templates(self) -> None:
        template_id = "razmeshchenie_soglasiy"
        template_path = self.settings.templates_dir / template_id / "template.pptx"
        manifest = self.registry.get_template(template_id)
        prototype = next(
            item for item in manifest.prototype_slides if any(token.binding == "main_text" for token in item.tokens)
        )
        body_token = next(token for token in prototype.tokens if token.binding == "main_text")
        body_token.left_emu = 999999
        body_token.top_emu = 2111111
        body_token.width_emu = 4444444
        body_token.height_emu = 1234567
        body_token.margin_left_emu = 21000
        body_token.margin_right_emu = 22000
        body_token.margin_top_emu = 23000
        body_token.margin_bottom_emu = 24000

        plan = PresentationPlan(
            template_id=template_id,
            title="Uploaded Prototype Audit Geometry",
            slides=[
                SlideSpec(
                    kind=SlideKind.TEXT,
                    title="Заголовок",
                    subtitle="Подзаголовок",
                    text="Template-aware deck audit должен читать body geometry и margins из prototype token metadata.",
                    notes="Служебная строка для prototype footer.",
                    preferred_layout_key=prototype.key,
                )
            ],
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = self.generator.generate(
                template_path=template_path,
                manifest=manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            audits = audit_generated_presentation(output_path, plan, manifest)
            violations = find_capacity_violations(audits)

        self.assertTrue(audits)
        self.assertFalse(any(item.rule == "body_left_misalignment" for item in violations))
        self.assertFalse(any(item.rule == "body_margin_mismatch" for item in violations))
        self.assertLess(audits[0].profile.max_chars, TEXT_FULL_WIDTH_PROFILE.max_chars)

    def test_capacity_profile_derivation_reduces_limits_for_narrower_body_geometry(self) -> None:
        reference_body = geometry_policy_for_layout("text_full_width").placeholders[14]
        derived = derive_capacity_profile_for_geometry(
            "text_full_width",
            width_emu=int(reference_body.width_emu * 0.6),
            height_emu=int(reference_body.height_emu * 0.7),
        )

        self.assertLess(derived.max_chars, TEXT_FULL_WIDTH_PROFILE.max_chars)
        self.assertLessEqual(derived.max_font_pt, TEXT_FULL_WIDTH_PROFILE.max_font_pt)

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

    def test_full_pipeline_contract_for_uploaded_prototype_template_keeps_template_aware_audit_green(self) -> None:
        document = Document()
        document.add_paragraph("A3")
        document.add_paragraph("Стратегический документ для пользовательского шаблона")
        document.add_heading("1. Основные выводы", level=1)
        document.add_paragraph(
            "Рост выручки обеспечивается платформенными интеграциями, улучшением конверсии и управляемым контуром сопровождения."
        )
        document.add_paragraph(
            "Следующий абзац нужен для text-flow сценария и проверки того, что template-aware audit читает реальную геометрию пользовательского prototype template."
        )
        document.add_heading("2. Таблица метрик", level=1)
        table = document.add_table(rows=3, cols=2)
        table.cell(0, 0).text = "Метрика"
        table.cell(0, 1).text = "Значение"
        table.cell(1, 0).text = "GMV"
        table.cell(1, 1).text = "125"
        table.cell(2, 0).text = "NPS"
        table.cell(2, 1).text = "61"

        buffer = BytesIO()
        document.save(buffer)
        content = buffer.getvalue()

        template_id = "razmeshchenie_soglasiy"
        manifest = self.registry.get_template(template_id)
        template_path = self.settings.templates_dir / template_id / manifest.source_pptx

        text, tables, blocks = self.extractor.extract("uploaded-prototype-contract.docx", content)
        plan = self.planner.build_plan(template_id, text, None, tables, blocks)

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = self.generator.generate(
                template_path=template_path,
                manifest=manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            audits = audit_generated_presentation(output_path, plan, manifest)
            violations = find_capacity_violations(audits)

        self.assertEqual(len(audits), len([slide for slide in plan.slides if slide.kind in {
            SlideKind.TEXT,
            SlideKind.BULLETS,
            SlideKind.TWO_COLUMN,
            SlideKind.TABLE,
            SlideKind.CHART,
            SlideKind.IMAGE,
        }]))
        self.assertFalse(any(item.rule == "body_margin_mismatch" for item in violations))
        self.assertFalse(any(item.rule == "body_left_misalignment" for item in violations))

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

    def test_layout_geometry_policies_keep_full_width_contracts(self) -> None:
        text_policy = geometry_policy_for_layout("text_full_width")
        list_policy = geometry_policy_for_layout("list_full_width")
        table_policy = geometry_policy_for_layout("table")
        image_policy = geometry_policy_for_layout("image_text")
        cards_policy = geometry_policy_for_layout("cards_3")
        icons_policy = geometry_policy_for_layout("list_with_icons")
        contacts_policy = geometry_policy_for_layout("contacts")

        self.assertEqual(text_policy.placeholders[0].width_emu, 11198224)
        self.assertEqual(text_policy.placeholders[14].width_emu, 11198224)
        self.assertEqual(text_policy.placeholders[17].top_emu, 6384626)
        self.assertEqual(list_policy.placeholders[14].width_emu, text_policy.placeholders[14].width_emu)
        self.assertEqual(table_policy.placeholders[15].width_emu, 11198224)
        self.assertEqual(image_policy.placeholders[16].width_emu, 4990840)
        self.assertEqual(cards_policy.placeholders[11].width_emu, cards_policy.placeholders[12].width_emu)
        self.assertEqual(icons_policy.placeholders[21].top_emu, 6384626)
        self.assertEqual(contacts_policy.placeholders[10].width_emu, 3724275)

    def test_layout_spacing_policies_keep_bullet_indent_contracts(self) -> None:
        text_spacing = spacing_policy_for_layout("text_full_width")
        list_spacing = spacing_policy_for_layout("list_full_width")

        self.assertEqual(text_spacing.bullet.margin_left_emu, 342900)
        self.assertEqual(text_spacing.bullet.indent_emu, -171450)
        self.assertEqual(text_spacing.body.line_spacing, 1.1)
        self.assertEqual(text_spacing.body.space_after_pt, 6.0)
        self.assertGreater(list_spacing.bullet.margin_left_emu, text_spacing.bullet.margin_left_emu)
        self.assertLess(list_spacing.bullet.indent_emu, text_spacing.bullet.indent_emu)

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

    def test_deck_audit_does_not_flag_borderline_underfilled_continuation(self) -> None:
        audits = [
            SlideAudit(
                slide_index=1,
                title="Раздел",
                kind=SlideKind.TEXT.value,
                layout_key="text_full_width",
                body_char_count=458,
                body_font_sizes=(14.0,),
                profile=TEXT_FULL_WIDTH_PROFILE,
            ),
            SlideAudit(
                slide_index=2,
                title="Раздел (2)",
                kind=SlideKind.TEXT.value,
                layout_key="text_full_width",
                body_char_count=291,
                body_font_sizes=(14.0,),
                profile=TEXT_FULL_WIDTH_PROFILE,
            ),
        ]

        violations = find_capacity_violations(audits)

        violation_rules = {violation.rule for violation in violations}
        self.assertNotIn("underfilled_continuation", violation_rules)
        self.assertNotIn("continuation_balance", violation_rules)

    def test_deck_audit_keeps_balance_violation_for_materially_underfilled_group(self) -> None:
        audits = [
            SlideAudit(
                slide_index=1,
                title="Раздел",
                kind=SlideKind.BULLETS.value,
                layout_key="list_full_width",
                body_char_count=720,
                body_font_sizes=(14.0,),
                profile=LIST_FULL_WIDTH_PROFILE,
            ),
            SlideAudit(
                slide_index=2,
                title="Раздел (2)",
                kind=SlideKind.BULLETS.value,
                layout_key="list_full_width",
                body_char_count=270,
                body_font_sizes=(14.0,),
                profile=LIST_FULL_WIDTH_PROFILE,
            ),
        ]

        violations = find_capacity_violations(audits)

        violation_rules = {violation.rule for violation in violations}
        self.assertIn("underfilled_continuation", violation_rules)
        self.assertIn("continuation_balance", violation_rules)

    def test_deck_audit_does_not_flag_borderline_bullet_tail(self) -> None:
        audits = [
            SlideAudit(
                slide_index=1,
                title="Раздел",
                kind=SlideKind.BULLETS.value,
                layout_key="list_full_width",
                body_char_count=507,
                body_font_sizes=(14.0,),
                profile=LIST_FULL_WIDTH_PROFILE,
            ),
            SlideAudit(
                slide_index=2,
                title="Раздел (2)",
                kind=SlideKind.BULLETS.value,
                layout_key="list_full_width",
                body_char_count=474,
                body_font_sizes=(14.0,),
                profile=LIST_FULL_WIDTH_PROFILE,
            ),
        ]

        violations = find_capacity_violations(audits)

        self.assertNotIn("underfilled_continuation", {violation.rule for violation in violations})

    def test_deck_audit_keeps_expected_bullet_order_for_mixed_slide(self) -> None:
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="Audit Order",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="Audit Order", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.BULLETS,
                    title="Смешанный блок",
                    bullets=[
                        "Вводный абзац задает контекст.",
                        "Первый тезис фиксирует массовый каталог.",
                        "Второй тезис фиксирует SLA-контур.",
                        "Финальный абзац завершает аргументацию.",
                    ],
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

        mixed_audit = next(audit for audit in audits if audit.title == "Смешанный блок")
        self.assertEqual(
            tuple(item.strip() for item in mixed_audit.rendered_items if item.strip()),
            tuple(item.strip() for item in mixed_audit.expected_items if item.strip()),
        )
        violations = find_capacity_violations(audits)
        self.assertNotIn("content_order_mismatch", {violation.rule for violation in violations})

    def test_deck_audit_uses_content_blocks_for_mixed_paragraph_and_list_order(self) -> None:
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="Audit Mixed Blocks",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="Audit Mixed Blocks", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.BULLETS,
                    title="Смешанный контейнер",
                    bullets=[
                        "Вводный абзац задает контекст.",
                        "Первый тезис фиксирует массовый каталог.",
                        "Второй тезис фиксирует SLA-контур.",
                        "Финальный абзац завершает аргументацию.",
                    ],
                    content_blocks=[
                        SlideContentBlock(kind=SlideContentBlockKind.PARAGRAPH, text="Вводный абзац задает контекст."),
                        SlideContentBlock(
                            kind=SlideContentBlockKind.BULLET_LIST,
                            items=[
                                "Первый тезис фиксирует массовый каталог.",
                                "Второй тезис фиксирует SLA-контур.",
                            ],
                        ),
                        SlideContentBlock(kind=SlideContentBlockKind.PARAGRAPH, text="Финальный абзац завершает аргументацию."),
                    ],
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

        mixed_audit = next(audit for audit in audits if audit.title == "Смешанный контейнер")
        self.assertEqual(
            tuple(item.strip() for item in mixed_audit.rendered_items if item.strip()),
            (
                "Вводный абзац задает контекст.",
                "Первый тезис фиксирует массовый каталог.",
                "Второй тезис фиксирует SLA-контур.",
                "Финальный абзац завершает аргументацию.",
            ),
        )
        violations = find_capacity_violations(audits)
        self.assertNotIn("content_order_mismatch", {violation.rule for violation in violations})

    def test_template_binding_keeps_content_blocks_for_text_slide_with_list_layout(self) -> None:
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="Binding Content Blocks",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="Binding Content Blocks", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.TEXT,
                    title="Смешанный narrative",
                    text="Вводный абзац задает контекст.",
                    content_blocks=[
                        SlideContentBlock(kind=SlideContentBlockKind.PARAGRAPH, text="Вводный абзац задает контекст."),
                        SlideContentBlock(
                            kind=SlideContentBlockKind.BULLET_LIST,
                            items=[
                                "Первый тезис описывает охват каталога.",
                                "Второй тезис описывает SLA-контур.",
                            ],
                        ),
                    ],
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

        target_audit = next(audit for audit in audits if audit.title == "Смешанный narrative")
        self.assertEqual(
            tuple(item.strip() for item in target_audit.rendered_items if item.strip()),
            (
                "Вводный абзац задает контекст.",
                "Первый тезис описывает охват каталога.",
                "Второй тезис описывает SLA-контур.",
            ),
        )
        violations = find_capacity_violations(audits)
        self.assertNotIn("content_order_mismatch", {violation.rule for violation in violations})

    def test_template_binding_does_not_duplicate_notes_when_content_blocks_fill_body(self) -> None:
        manifest = self.registry.get_template("corp_light_v1")
        template_path = self.registry.get_template_pptx_path("corp_light_v1")
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="No Duplicate Tail",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="No Duplicate Tail", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.TEXT,
                    title="Narrative",
                    text="Первый абзац задает контекст.",
                    notes="Хвост не должен дублироваться отдельным placeholder.",
                    content_blocks=[
                        SlideContentBlock(kind=SlideContentBlockKind.PARAGRAPH, text="Первый абзац задает контекст."),
                        SlideContentBlock(
                            kind=SlideContentBlockKind.PARAGRAPH,
                            text="Хвост не должен дублироваться отдельным placeholder.",
                        ),
                    ],
                    preferred_layout_key="text_full_width",
                ),
            ],
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = self.generator.generate(
                template_path=template_path,
                manifest=manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            presentation = Presentation(str(output_path))

        slide = presentation.slides[1]
        placeholders = {
            shape.placeholder_format.idx: shape
            for shape in slide.placeholders
            if getattr(shape, "is_placeholder", False)
        }
        self.assertIn(14, placeholders)
        self.assertIn("Хвост не должен дублироваться", placeholders[14].text)
        self.assertNotIn(15, placeholders)

    def test_template_binding_clears_duplicate_subtitle_when_body_already_starts_with_it(self) -> None:
        manifest = self.registry.get_template("corp_light_v1")
        template_path = self.registry.get_template_pptx_path("corp_light_v1")
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="No Duplicate Subtitle",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="No Duplicate Subtitle", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.TEXT,
                    title="Narrative",
                    subtitle="Первый абзац задает контекст.",
                    text="Первый абзац задает контекст.",
                    content_blocks=[
                        SlideContentBlock(
                            kind=SlideContentBlockKind.PARAGRAPH,
                            text="Первый абзац задает контекст. Дальше идет основной текст без отдельного subtitle-повтора.",
                        ),
                    ],
                    preferred_layout_key="text_full_width",
                ),
            ],
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = self.generator.generate(
                template_path=template_path,
                manifest=manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            presentation = Presentation(str(output_path))

        slide = presentation.slides[1]
        placeholders = {
            shape.placeholder_format.idx: shape
            for shape in slide.placeholders
            if getattr(shape, "is_placeholder", False)
        } 
        self.assertNotIn(13, placeholders)

    def test_generator_clears_duplicate_subtitle_for_plain_text_slide(self) -> None:
        manifest = self.registry.get_template("corp_light_v1")
        template_path = self.registry.get_template_pptx_path("corp_light_v1")
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="No Plain Duplicate Subtitle",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="No Plain Duplicate Subtitle", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.TEXT,
                    title="Narrative",
                    subtitle="Первый абзац задает контекст и открывает раздел.",
                    text="Первый абзац задает контекст и открывает раздел. Дальше идет основной narrative без отдельного подзаголовка.",
                    preferred_layout_key="text_full_width",
                ),
            ],
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = self.generator.generate(
                template_path=template_path,
                manifest=manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            presentation = Presentation(str(output_path))

        slide = presentation.slides[1]
        placeholders = {
            shape.placeholder_format.idx: shape
            for shape in slide.placeholders
            if getattr(shape, "is_placeholder", False)
        }
        self.assertNotIn(13, placeholders)
        self.assertIn(14, placeholders)

    def test_deck_audit_ignores_text_only_continuations_when_tracking_group_order(self) -> None:
        audits = [
            SlideAudit(
                slide_index=2,
                title="Группа",
                kind=SlideKind.TEXT.value,
                layout_key="text_full_width",
                body_char_count=320,
                body_font_sizes=(16.0,),
                profile=profile_for_layout("text_full_width"),
                expected_items=(),
                rendered_items=("Вводный narrative блок.",),
            ),
            SlideAudit(
                slide_index=3,
                title="Группа (2)",
                kind=SlideKind.BULLETS.value,
                layout_key="list_full_width",
                body_char_count=220,
                body_font_sizes=(14.0,),
                profile=profile_for_layout("list_full_width"),
                expected_items=("Первый тезис.", "Второй тезис."),
                rendered_items=("Первый тезис.", "Второй тезис."),
            ),
            SlideAudit(
                slide_index=4,
                title="Группа (3)",
                kind=SlideKind.TEXT.value,
                layout_key="text_full_width",
                body_char_count=280,
                body_font_sizes=(16.0,),
                profile=profile_for_layout("text_full_width"),
                expected_items=(),
                rendered_items=("Финальный narrative блок.",),
            ),
        ]

        violations = find_capacity_violations(audits)
        self.assertNotIn("continuation_order_mismatch", {violation.rule for violation in violations})

    def test_deck_audit_accepts_question_and_appendix_style_slide_without_order_violations(self) -> None:
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="Audit Question Appendix",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="Audit Question Appendix", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.BULLETS,
                    title="Частые вопросы и выводы",
                    bullets=[
                        "Почему нужен второй контур инфраструктуры?",
                        "Важно: без резервного контура SLA для критичных платежей останется хрупким.",
                        "Итог: резервный контур снижает риск простоя и упрощает масштабирование.",
                    ],
                    content_blocks=[
                        SlideContentBlock(kind=SlideContentBlockKind.QA_ITEM, text="Почему нужен второй контур инфраструктуры?"),
                        SlideContentBlock(
                            kind=SlideContentBlockKind.CALLOUT,
                            text="Важно: без резервного контура SLA для критичных платежей останется хрупким.",
                        ),
                        SlideContentBlock(
                            kind=SlideContentBlockKind.CALLOUT,
                            text="Итог: резервный контур снижает риск простоя и упрощает масштабирование.",
                        ),
                    ],
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
        self.assertEqual(violations, [])

    def test_deck_audit_accepts_semantic_text_layout_for_question_and_callout_blocks(self) -> None:
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="Semantic Text Layout",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="Semantic Text Layout", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.TEXT,
                    title="FAQ и выводы",
                    text="Вопрос: почему нужен второй контур инфраструктуры?",
                    notes=(
                        "Важно: резервный контур удерживает SLA в часы пиковой нагрузки. "
                        "Итог: архитектура снижает риск простоя без усложнения операционной модели."
                    ),
                    content_blocks=[
                        SlideContentBlock(
                            kind=SlideContentBlockKind.QA_ITEM,
                            text="Вопрос: почему нужен второй контур инфраструктуры?",
                        ),
                        SlideContentBlock(
                            kind=SlideContentBlockKind.CALLOUT,
                            text="Важно: резервный контур удерживает SLA в часы пиковой нагрузки.",
                        ),
                        SlideContentBlock(
                            kind=SlideContentBlockKind.CALLOUT,
                            text="Итог: архитектура снижает риск простоя без усложнения операционной модели.",
                        ),
                    ],
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

        self.assertEqual(len(audits), 1)
        self.assertTrue(audits[0].body_font_sizes)
        violations = find_capacity_violations(audits)
        self.assertNotIn("content_order_mismatch", {violation.rule for violation in violations})

    def test_manifests_keep_generic_text_or_bullet_layout_for_semantic_blocks(self) -> None:
        manifests = self.registry.list_templates()
        self.assertGreaterEqual(len(manifests), 1)

        for manifest in manifests:
            layout_keys = {layout.key for layout in manifest.layouts}
            supported_slide_kinds = {
                supported_kind
                for layout in manifest.layouts
                for supported_kind in layout.supported_slide_kinds
            }
            with self.subTest(template_id=manifest.template_id):
                if {"text", "bullets"} & supported_slide_kinds:
                    self.assertTrue(
                        any(
                            layout_key in layout_keys
                            for layout_key in {"text_full_width", "list_full_width", "content", "слайд_с_перечислением"}
                        )
                    )

    def test_deck_audit_accepts_long_title_and_dense_body_without_capacity_regression(self) -> None:
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="Audit Long Title Dense Body",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="Audit Long Title Dense Body", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.TEXT,
                    title=(
                        "Очень длинный заголовок, который одновременно фиксирует стратегическую развилку, "
                        "операционные ограничения и требования к инфраструктуре платежной платформы"
                    ),
                    text=(
                        "Первый плотный абзац объясняет, почему устойчивость SLA зависит от второго контура, "
                        "какие ограничения есть у текущей архитектуры и как это влияет на масштабирование продукта. "
                        "Второй плотный абзац связывает инфраструктурные решения с коммерческой логикой, "
                        "переговорной позицией с банками и требованиями к государственным интеграциям."
                    ),
                    notes=(
                        "Третий плотный абзац фиксирует риски внедрения, требования к операционной модели и "
                        "критерии, по которым можно считать архитектурное решение успешным после запуска."
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

    def test_deck_audit_flags_continuation_order_mismatch(self) -> None:
        audits = [
            SlideAudit(
                slide_index=2,
                title="Раздел",
                kind=SlideKind.BULLETS.value,
                layout_key="list_full_width",
                body_char_count=120,
                body_font_sizes=(14.0,),
                profile=profile_for_layout("list_full_width"),
                expected_items=("Первый", "Второй"),
                rendered_items=("Второй", "Первый"),
            ),
            SlideAudit(
                slide_index=3,
                title="Раздел (2)",
                kind=SlideKind.BULLETS.value,
                layout_key="list_full_width",
                body_char_count=120,
                body_font_sizes=(14.0,),
                profile=profile_for_layout("list_full_width"),
                expected_items=("Третий",),
                rendered_items=("Третий",),
            ),
        ]

        violations = find_capacity_violations(audits)
        rules = {violation.rule for violation in violations}
        self.assertIn("content_order_mismatch", rules)
        self.assertIn("continuation_order_mismatch", rules)

    def test_deck_audit_flags_large_font_jump_between_continuation_slides(self) -> None:
        audits = [
            SlideAudit(
                slide_index=2,
                title="Раздел",
                kind=SlideKind.TEXT.value,
                layout_key="text_full_width",
                body_char_count=320,
                body_font_sizes=(16.0,),
                profile=profile_for_layout("text_full_width"),
                expected_items=("Первый блок",),
                rendered_items=("Первый блок",),
            ),
            SlideAudit(
                slide_index=3,
                title="Раздел (2)",
                kind=SlideKind.TEXT.value,
                layout_key="text_full_width",
                body_char_count=300,
                body_font_sizes=(12.0,),
                profile=profile_for_layout("text_full_width"),
                expected_items=("Второй блок",),
                rendered_items=("Второй блок",),
            ),
        ]

        violations = find_capacity_violations(audits)
        self.assertIn("continuation_font_delta", {violation.rule for violation in violations})

    def test_deck_audit_accepts_small_font_delta_between_continuation_slides(self) -> None:
        audits = [
            SlideAudit(
                slide_index=2,
                title="Раздел",
                kind=SlideKind.TEXT.value,
                layout_key="text_full_width",
                body_char_count=320,
                body_font_sizes=(14.0,),
                profile=profile_for_layout("text_full_width"),
                expected_items=("Первый блок",),
                rendered_items=("Первый блок",),
            ),
            SlideAudit(
                slide_index=3,
                title="Раздел (2)",
                kind=SlideKind.TEXT.value,
                layout_key="text_full_width",
                body_char_count=300,
                body_font_sizes=(13.0,),
                profile=profile_for_layout("text_full_width"),
                expected_items=("Второй блок",),
                rendered_items=("Второй блок",),
            ),
        ]

        violations = find_capacity_violations(audits)
        self.assertNotIn("continuation_font_delta", {violation.rule for violation in violations})

    def test_deck_audit_flags_title_body_overlap_for_text_slide(self) -> None:
        audit = SlideAudit(
            slide_index=2,
            title="Раздел",
            kind=SlideKind.TEXT.value,
            layout_key="text_full_width",
            body_char_count=140,
            body_font_sizes=(14.0,),
            profile=profile_for_layout("text_full_width"),
            title_top=671247,
            title_height=900000,
            body_top=1700000,
            body_height=2800000,
            footer_top=6384626,
            footer_width=11198224,
            body_left=442913,
        )

        violations = find_capacity_violations([audit])
        self.assertIn("title_body_overlap", {violation.rule for violation in violations})

    def test_deck_audit_flags_subtitle_body_overlap_for_text_slide(self) -> None:
        audit = SlideAudit(
            slide_index=2,
            title="Раздел",
            kind=SlideKind.TEXT.value,
            layout_key="text_full_width",
            body_char_count=140,
            body_font_sizes=(14.0,),
            profile=profile_for_layout("text_full_width"),
            title_top=671247,
            title_height=800000,
            subtitle_top=1600000,
            subtitle_height=400000,
            body_top=2050000,
            body_height=2500000,
            footer_top=6384626,
            footer_width=11198224,
            body_left=442913,
        )

        violations = find_capacity_violations([audit])
        self.assertIn("subtitle_body_overlap", {violation.rule for violation in violations})

    def test_deck_audit_flags_narrow_footer_for_text_layout(self) -> None:
        audit = SlideAudit(
            slide_index=2,
            title="Раздел",
            kind=SlideKind.TEXT.value,
            layout_key="text_full_width",
            body_char_count=320,
            body_font_sizes=(14.0,),
            profile=profile_for_layout("text_full_width"),
            title_top=671247,
            title_height=800000,
            body_top=1900000,
            body_height=3600000,
            footer_top=6384626,
            footer_width=int(PptxGenerator.FULL_CONTENT_WIDTH_EMU * 0.75),
            body_left=442913,
        )

        violations = find_capacity_violations([audit])
        self.assertIn("narrow_text_footer", {violation.rule for violation in violations})

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

    def test_generator_preserves_table_cell_fill_colors_from_table_block(self) -> None:
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="Audit Table Fill Colors",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="Audit Table Fill Colors", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.TABLE,
                    title="Цветная таблица",
                    table=TableBlock(
                        headers=["Показатель", "Значение"],
                        header_fill_colors=["1F4E79", "1F4E79"],
                        rows=[["Выручка", "120"], ["Маржа", "24%"]],
                        row_fill_colors=[[None, "D9EAF7"], ["FDE7D7", None]],
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
            presentation = Presentation(str(output_path))

        slide = presentation.slides[1]
        table = next(shape.table for shape in slide.shapes if getattr(shape, "has_table", False))
        self.assertEqual(table.cell(0, 0).fill.fore_color.rgb, RGBColor(0x1F, 0x4E, 0x79))
        self.assertEqual(table.cell(1, 1).fill.fore_color.rgb, RGBColor(0xD9, 0xEA, 0xF7))
        self.assertEqual(table.cell(2, 0).fill.fore_color.rgb, RGBColor(0xFD, 0xE7, 0xD7))

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

    def test_deck_audit_validates_cards_layout_geometry(self) -> None:
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="Audit Cards Geometry",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="Audit Cards Geometry", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.BULLETS,
                    title="Три направления роста",
                    bullets=["Усилить ядро платформы", "Ускорить интеграции", "Повысить retention"],
                    preferred_layout_key="cards_3",
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

        card_audit = next(audit for audit in audits if audit.layout_key == "cards_3")
        violations = find_capacity_violations(audits)
        self.assertTrue(card_audit.auxiliary_widths)
        self.assertEqual([v.rule for v in violations if v.slide_index == card_audit.slide_index], [])

    def test_deck_audit_validates_two_column_layout_geometry(self) -> None:
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="Audit Two Column Geometry",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="Audit Two Column Geometry", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.TWO_COLUMN,
                    title="Модель взаимодействия",
                    subtitle="Колонки и иконки должны сохранять рабочую сетку",
                    left_bullets=["Контекст проекта", "Ограничения", "Допущения"],
                    right_bullets=["Шаг 1", "Шаг 2", "Шаг 3", "Шаг 4"],
                    preferred_layout_key="list_with_icons",
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

        two_column_audit = next(audit for audit in audits if audit.layout_key == "list_with_icons")
        violations = find_capacity_violations(audits)
        self.assertTrue(two_column_audit.auxiliary_widths)
        self.assertEqual([v.rule for v in violations if v.slide_index == two_column_audit.slide_index], [])

    def test_deck_audit_validates_contacts_layout_geometry(self) -> None:
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="Audit Contacts Geometry",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="Audit Contacts Geometry", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.TEXT,
                    title="Иван Иванов",
                    subtitle="CEO",
                    left_bullets=["+7 999 123-45-67"],
                    right_bullets=["ivan@example.com"],
                    preferred_layout_key="contacts",
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

        contacts_audit = next(audit for audit in audits if audit.layout_key == "contacts")
        violations = find_capacity_violations(audits)
        self.assertTrue(contacts_audit.auxiliary_widths)
        self.assertEqual([v.rule for v in violations if v.slide_index == contacts_audit.slide_index], [])

    def test_deck_audit_flags_underfilled_body_height_for_full_width_text(self) -> None:
        audit = SlideAudit(
            slide_index=2,
            title="Раздел",
            kind=SlideKind.TEXT.value,
            layout_key="text_full_width",
            body_char_count=140,
            body_font_sizes=(14.0,),
            profile=profile_for_layout("text_full_width"),
            title_top=671247,
            title_height=800000,
            body_top=1900000,
            body_height=1200000,
            footer_top=6384626,
            footer_width=11198224,
            body_left=442913,
        )

        violations = find_capacity_violations([audit])
        self.assertIn("underfilled_body_height", {violation.rule for violation in violations})

    def test_deck_audit_keeps_body_text_frame_margin_contract_for_full_width_text(self) -> None:
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="Audit Body Margins",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="Audit Body Margins", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.TEXT,
                    title="Раздел",
                    text="Основной текст должен рендериться с явными внутренними margin-параметрами.",
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

        text_audit = next(audit for audit in audits if audit.layout_key == "text_full_width")
        self.assertEqual(text_audit.body_margin_left, PptxGenerator.DEFAULT_TEXT_MARGIN_X_EMU)
        self.assertEqual(text_audit.body_margin_right, PptxGenerator.DEFAULT_TEXT_MARGIN_X_EMU)
        self.assertEqual(text_audit.body_margin_top, PptxGenerator.DEFAULT_TEXT_MARGIN_Y_EMU)
        self.assertEqual(text_audit.body_margin_bottom, PptxGenerator.DEFAULT_TEXT_MARGIN_Y_EMU)
        violations = find_capacity_violations(audits)
        self.assertNotIn("body_margin_mismatch", {violation.rule for violation in violations})

    def test_full_width_text_body_keeps_paragraph_spacing_contract(self) -> None:
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="Audit Paragraph Spacing",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="Audit Paragraph Spacing", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.TEXT,
                    title="Раздел",
                    text="Первый абзац основного текста.\nВторой абзац основного текста.",
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
            presentation = Presentation(str(output_path))

        slide = presentation.slides[1]
        body = next(shape for shape in slide.placeholders if shape.placeholder_format.idx == 14)
        paragraphs = [paragraph for paragraph in body.text_frame.paragraphs if paragraph.text.strip()]
        self.assertTrue(paragraphs)
        for paragraph in paragraphs:
            with self.subTest(text=paragraph.text):
                self.assertEqual(paragraph.line_spacing, 1.1)
                self.assertEqual(paragraph.space_after.pt, 6.0)

    def test_full_width_bullets_keep_paragraph_spacing_contract(self) -> None:
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="Audit Bullet Paragraph Spacing",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="Audit Bullet Paragraph Spacing", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.BULLETS,
                    title="Раздел",
                    bullets=["Первый пункт", "Второй пункт", "Третий пункт"],
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
            presentation = Presentation(str(output_path))

        slide = presentation.slides[1]
        body = next(shape for shape in slide.placeholders if shape.placeholder_format.idx == 14)
        paragraphs = [paragraph for paragraph in body.text_frame.paragraphs if paragraph.text.strip()]
        self.assertTrue(paragraphs)
        for paragraph in paragraphs:
            with self.subTest(text=paragraph.text):
                self.assertEqual(paragraph.line_spacing, 1.05)
                self.assertEqual(paragraph.space_after.pt, 5.0)

    def test_cover_text_frames_keep_margin_contract(self) -> None:
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="Стратегия А3",
            slides=[
                SlideSpec(
                    kind=SlideKind.TITLE,
                    title="Стратегия А3",
                    notes="Горизонт планирования: 2026-2030\nМарт 2026",
                    preferred_layout_key="cover",
                )
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
            presentation = Presentation(str(output_path))

        slide = presentation.slides[0]
        text_shapes = [shape for shape in slide.shapes if getattr(shape, "has_text_frame", False)]
        self.assertTrue(text_shapes)
        for shape in text_shapes:
            with self.subTest(shape=getattr(shape, "name", "")):
                text_frame = shape.text_frame
                self.assertEqual(text_frame.margin_left, PptxGenerator.DEFAULT_TEXT_MARGIN_X_EMU)
                self.assertEqual(text_frame.margin_right, PptxGenerator.DEFAULT_TEXT_MARGIN_X_EMU)
                self.assertEqual(text_frame.margin_top, PptxGenerator.DEFAULT_TEXT_MARGIN_Y_EMU)
                self.assertEqual(text_frame.margin_bottom, PptxGenerator.DEFAULT_TEXT_MARGIN_Y_EMU)

    def test_prototype_templates_keep_margin_contract_on_nonempty_text_shapes(self) -> None:
        manifests = self.registry.list_templates()

        for manifest in manifests:
            template_path = self.settings.templates_dir / manifest.template_id / manifest.source_pptx
            if not template_path.exists() or manifest.generation_mode.value != "prototype":
                continue

            with self.subTest(template_id=manifest.template_id):
                plan = PresentationPlan(
                    template_id=manifest.template_id,
                    title=f"{manifest.display_name} Prototype Margin Contract",
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

                nonempty_text_shapes = [
                    shape
                    for slide in presentation.slides
                    for shape in slide.shapes
                    if getattr(shape, "has_text_frame", False) and (shape.text or "").strip()
                ]
                self.assertTrue(nonempty_text_shapes)
                for shape in nonempty_text_shapes:
                    with self.subTest(template_id=manifest.template_id, shape=getattr(shape, "name", "")):
                        text_frame = shape.text_frame
                        self.assertEqual(text_frame.margin_left, PptxGenerator.DEFAULT_TEXT_MARGIN_X_EMU)
                        self.assertEqual(text_frame.margin_right, PptxGenerator.DEFAULT_TEXT_MARGIN_X_EMU)
                        self.assertEqual(text_frame.margin_top, PptxGenerator.DEFAULT_TEXT_MARGIN_Y_EMU)
                        self.assertEqual(text_frame.margin_bottom, PptxGenerator.DEFAULT_TEXT_MARGIN_Y_EMU)

    def test_prototype_templates_render_nonempty_bound_content(self) -> None:
        manifests = self.registry.list_templates()

        for manifest in manifests:
            template_path = self.settings.templates_dir / manifest.template_id / manifest.source_pptx
            if not template_path.exists() or manifest.generation_mode.value != "prototype":
                continue

            with self.subTest(template_id=manifest.template_id):
                plan = PresentationPlan(
                    template_id=manifest.template_id,
                    title=f"{manifest.display_name} Prototype Content Contract",
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

                rendered_payload = "\n".join(
                    shape.text.strip()
                    for slide in presentation.slides
                    for shape in slide.shapes
                    if getattr(shape, "has_text_frame", False) and (shape.text or "").strip()
                )
                has_visual_payload = any(
                    getattr(shape, "has_table", False)
                    or getattr(shape, "has_chart", False)
                    or hasattr(shape, "image")
                    for slide in presentation.slides
                    for shape in slide.shapes
                )
                self.assertTrue(rendered_payload.strip() or has_visual_payload)

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

    def _has_geometry_metadata(self, spec) -> bool:
        return all(
            isinstance(value, int) and value > 0
            for value in (spec.left_emu, spec.top_emu, spec.width_emu, spec.height_emu)
        )

    def _has_text_margin_metadata(self, spec) -> bool:
        return all(
            isinstance(value, int) and value >= 0
            for value in (
                spec.margin_left_emu,
                spec.margin_right_emu,
                spec.margin_top_emu,
                spec.margin_bottom_emu,
            )
        )


if __name__ == "__main__":
    unittest.main()
