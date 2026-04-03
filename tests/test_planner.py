from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_MARKER_STYLE

from a3presentation.domain.api import ChartOverride, DocumentBlock
from a3presentation.domain.chart import ChartConfidence, ChartSeries, ChartSpec, ChartType
from a3presentation.domain.presentation import PresentationPlan, SlideKind, SlideSpec, TableBlock
from a3presentation.services.pptx_generator import PptxGenerator
from a3presentation.services.planner import TextToPlanService
from a3presentation.services.template_registry import TemplateRegistry
from a3presentation.settings import get_settings


class TextToPlanServiceTests(unittest.TestCase):
    def test_build_plan_handles_large_text_sections_without_crashing(self) -> None:
        service = TextToPlanService()
        blocks = [
            DocumentBlock(kind="title", text="A3 Presentation", level=0),
            DocumentBlock(kind="heading", text="Большой аналитический раздел", level=1),
            DocumentBlock(
                kind="paragraph",
                text=(
                    "Первое длинное предложение объясняет контекст задачи и описывает ограничения модели. "
                    "Второе предложение добавляет параметры для расчета юнит-экономики по каждому сегменту. "
                    "Третье предложение вводит сравнение с конкурентами и связывает выводы с рекомендацией для CEO. "
                    "Четвертое предложение завершает блок и гарантирует, что раздел не помещается в один слайд."
                ),
            ),
            DocumentBlock(
                kind="paragraph",
                text=(
                    "Дополнительный абзац нужен, чтобы планировщик переходил к логике разбиения большого раздела. "
                    "Именно этот путь раньше падал из-за обращения к отсутствующей константе LIST_BULLET_MAX_CHARS."
                ),
            ),
            DocumentBlock(kind="heading", text="Вторая секция", level=1),
            DocumentBlock(
                kind="paragraph",
                text=(
                    "Эта секция нужна только для того, чтобы первая точно не была воспринята как титульная и skipped при сборке плана."
                ),
            ),
        ]

        plan = service.build_plan(
            template_id="corp_light_v1",
            raw_text="\n".join(block.text or "" for block in blocks),
            title="A3 Presentation",
            blocks=blocks,
        )

        self.assertGreaterEqual(len(plan.slides), 2)
        self.assertEqual(plan.slides[0].preferred_layout_key, "cover")
        self.assertTrue(any(slide.preferred_layout_key in {"text_full_width", "list_full_width"} for slide in plan.slides[1:]))

    def test_planner_splits_dense_bullets_by_content_weight(self) -> None:
        service = TextToPlanService()
        bullets = [
            "Первый очень длинный пункт объясняет стратегическую развилку, перечисляет ограничения, содержит несколько зависимых условий для принятия решения, раскрывает влияние на экономику сегмента, описывает организационные последствия и дополняется уточнением по срокам внедрения.",
            "Второй очень длинный пункт продолжает тему, добавляет экономические критерии, риски реализации, ожидаемый эффект для нескольких сегментов, перечисляет технологические ограничения, требования к команде сопровождения и допущения по инфраструктуре.",
            "Третий очень длинный пункт суммирует логику, связывает рекомендации с KPI, инвестиционным циклом и влиянием на портфель поставщиков, а также отдельно фиксирует ограничения по данным, срокам и допустимой нагрузке на операционный контур.",
        ]
        chunks = service._chunk_bullets_for_slides(bullets)
        self.assertGreaterEqual(len(chunks), 2)

    def test_cover_meta_stays_compact_for_long_first_section(self) -> None:
        service = TextToPlanService()
        blocks = [
            DocumentBlock(kind="title", text="Большая стратегическая презентация", level=0),
            DocumentBlock(kind="heading", text="Контекст задачи и что именно считать поставщиком", level=1),
            DocumentBlock(
                kind="paragraph",
                text=(
                    'В регулярных платежах метрика "количество поставщиков" почти всегда многозначна и требует отдельной '
                    "декомпозиции по юридическим лицам, договорам, биллинговым витринам и маршрутам списаний."
                ),
            ),
            DocumentBlock(
                kind="paragraph",
                text=(
                    "На рынке обычно существуют как минимум три уровня поставщиков, и каждый по-разному влияет на экономику, "
                    "операционный контур и переговорную позицию компании."
                ),
            ),
            DocumentBlock(kind="heading", text="Следующий раздел", level=1),
            DocumentBlock(kind="paragraph", text="Короткий текст для второй секции."),
        ]

        plan = service.build_plan(
            template_id="corp_light_v1",
            raw_text="\n".join(block.text or "" for block in blocks),
            title="Большая стратегическая презентация",
            blocks=blocks,
        )

        cover = plan.slides[0]
        self.assertEqual(cover.preferred_layout_key, "cover")
        self.assertLessEqual(len((cover.notes or "").splitlines()), 2)
        self.assertNotIn("многозначна", cover.notes or "")
        self.assertNotIn("как минимум три уровня", cover.notes or "")

    def test_cover_keeps_short_meta_lines_when_present(self) -> None:
        service = TextToPlanService()
        blocks = [
            DocumentBlock(kind="title", text="Стратегия развития А3", level=0),
            DocumentBlock(kind="paragraph", text="Горизонт планирования: 2026-2030"),
            DocumentBlock(kind="paragraph", text="Март 2026"),
            DocumentBlock(kind="paragraph", text="Конфиденциальный документ"),
            DocumentBlock(kind="heading", text="Первый раздел", level=1),
            DocumentBlock(kind="paragraph", text="Основной контент презентации."),
        ]

        plan = service.build_plan(
            template_id="corp_light_v1",
            raw_text="\n".join(block.text or "" for block in blocks),
            title="Стратегия развития А3",
            blocks=blocks,
        )

        cover = plan.slides[0]
        self.assertEqual(cover.notes, "Горизонт планирования: 2026-2030\nМарт 2026")

    def test_first_section_with_tables_is_not_swallowed_by_cover(self) -> None:
        service = TextToPlanService()
        blocks = [
            DocumentBlock(kind="paragraph", text="АНКЕТА КАНДИДАТА"),
            DocumentBlock(
                kind="table",
                table=TableBlock(
                    headers=["Поле", "Значение"],
                    rows=[["ФИО", "Игнатов Иван"], ["Телефон", "+7 900 000-00-00"]],
                ),
            ),
            DocumentBlock(kind="paragraph", text="ОБРАЗОВАНИЕ"),
            DocumentBlock(
                kind="table",
                table=TableBlock(
                    headers=["Период", "Учебное заведение", "Статус"],
                    rows=[["2018-2022", "Университет", "Высшее"]],
                ),
            ),
        ]

        plan = service.build_plan(
            template_id="corp_light_v1",
            raw_text="\n".join(block.text or "" for block in blocks),
            title=None,
            tables=[],
            blocks=blocks,
        )

        self.assertGreaterEqual(len(plan.slides), 3)
        self.assertEqual(plan.slides[0].preferred_layout_key, "cover")
        self.assertTrue(any(slide.table is not None for slide in plan.slides[1:]))

    def test_non_narrative_document_uses_safe_fallback_planner(self) -> None:
        service = TextToPlanService()
        blocks = [
            DocumentBlock(kind="paragraph", text="АНКЕТА КАНДИДАТА"),
            DocumentBlock(
                kind="table",
                table=TableBlock(
                    headers=["Поле", "Значение"],
                    rows=[["ФИО", "Игнатов Иван"], ["Телефон", "+7 900 000-00-00"]],
                ),
            ),
            DocumentBlock(kind="paragraph", text="ОБРАЗОВАНИЕ"),
            DocumentBlock(
                kind="table",
                table=TableBlock(
                    headers=["Период", "Учебное заведение", "Статус"],
                    rows=[["2018-2022", "Университет", "Высшее"]],
                ),
            ),
            DocumentBlock(kind="paragraph", text="ДОПОЛНИТЕЛЬНЫЕ СВЕДЕНИЯ"),
            DocumentBlock(
                kind="paragraph",
                text="Подписывая настоящую Анкету, кандидат выражает согласие на обработку персональных данных в целях трудоустройства.",
            ),
        ]

        plan = service.build_plan(
            template_id="corp_light_v1",
            raw_text="\n".join(block.text or "" for block in blocks),
            title=None,
            tables=[],
            blocks=blocks,
        )

        self.assertGreaterEqual(len(plan.slides), 4)
        self.assertEqual(plan.slides[1].preferred_layout_key, "list_full_width")
        self.assertTrue(any(slide.preferred_layout_key == "text_full_width" for slide in plan.slides[1:]))
        self.assertTrue(any(slide.preferred_layout_key == "table" for slide in plan.slides[1:]))

    def test_resume_document_uses_resume_fallback(self) -> None:
        service = TextToPlanService()
        blocks = [
            DocumentBlock(kind="paragraph", text="Иван Игнатов"),
            DocumentBlock(kind="paragraph", text="ivan@example.com"),
            DocumentBlock(kind="paragraph", text="+7 999 000-00-00"),
            DocumentBlock(kind="paragraph", text="ОПЫТ РАБОТЫ"),
            DocumentBlock(
                kind="paragraph",
                text="Руководил командой из десяти человек, запускал внутренние продукты, выстраивал процессы аналитики и координировал кросс-функциональные проекты.",
            ),
            DocumentBlock(kind="paragraph", text="ОБРАЗОВАНИЕ"),
            DocumentBlock(
                kind="paragraph",
                text="Высшее техническое образование, дополнительная программа по продуктовому менеджменту и регулярные курсы по аналитике данных.",
            ),
            DocumentBlock(kind="paragraph", text="НАВЫКИ"),
            DocumentBlock(kind="list", items=["Product management", "SQL", "Презентации"]),
        ]

        plan = service.build_plan(
            template_id="corp_light_v1",
            raw_text="\n".join(block.text or "" for block in blocks),
            title=None,
            tables=[],
            blocks=blocks,
        )

        self.assertGreaterEqual(len(plan.slides), 3)
        self.assertEqual(plan.slides[1].preferred_layout_key, "list_full_width")
        self.assertTrue(any(slide.preferred_layout_key == "text_full_width" for slide in plan.slides[1:]))
        self.assertFalse(any(slide.preferred_layout_key == "table" for slide in plan.slides[1:]))

    def test_table_heavy_document_adds_table_count_summary(self) -> None:
        service = TextToPlanService()
        blocks = [
            DocumentBlock(kind="paragraph", text="Сводный операционный отчёт"),
            DocumentBlock(kind="table", table=TableBlock(headers=["Показатель", "Значение"], rows=[["GMV", "100"]])),
            DocumentBlock(kind="table", table=TableBlock(headers=["Показатель", "Значение"], rows=[["MAU", "200"]])),
            DocumentBlock(kind="table", table=TableBlock(headers=["Показатель", "Значение"], rows=[["NPS", "65"]])),
        ]

        plan = service.build_plan(
            template_id="corp_light_v1",
            raw_text="\n".join(block.text or "" for block in blocks),
            title=None,
            tables=[],
            blocks=blocks,
        )

        self.assertGreaterEqual(len(plan.slides), 4)
        self.assertEqual(plan.slides[1].preferred_layout_key, "list_full_width")
        self.assertTrue(any("Таблиц в документе" in bullet for bullet in plan.slides[1].bullets))

    def test_planner_replaces_selected_table_with_chart_slide(self) -> None:
        service = TextToPlanService()
        table = TableBlock(
            headers=["Канал", "Лиды"],
            rows=[["SEO", "120"], ["Ads", "200"], ["Referral", "90"]],
        )

        plan = service.build_plan(
            template_id="corp_light_v1",
            raw_text="Отчет по каналам",
            title="Отчет по каналам",
            tables=[table],
            chart_overrides=[
                ChartOverride(
                    table_id="table_1",
                    mode="chart",
                    selected_chart=ChartSpec(
                        chart_id="chart_1",
                        source_table_id="table_1",
                        chart_type=ChartType.COLUMN,
                        title="Лиды по каналам",
                        categories=["SEO", "Ads", "Referral"],
                        series=[ChartSeries(name="Лиды", values=[120.0, 200.0, 90.0])],
                        confidence=ChartConfidence.HIGH,
                    ),
                )
            ],
        )

        chart_slides = [slide for slide in plan.slides if slide.kind == SlideKind.CHART]
        self.assertEqual(len(chart_slides), 1)
        self.assertEqual(chart_slides[0].source_table_id, "table_1")
        self.assertIsNotNone(chart_slides[0].chart)
        self.assertIsNone(chart_slides[0].table)

    def test_generator_expands_long_title_and_pushes_content_down(self) -> None:
        settings = get_settings()
        registry = TemplateRegistry(settings.templates_dir)
        manifest = registry.get_template("corp_light_v1")
        template_path = registry.get_template_pptx_path("corp_light_v1")

        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="A3 Presentation",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="A3 Presentation", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.TEXT,
                    title="Очень длинный заголовок слайда для проверки автоматического увеличения высоты заголовка и сдвига основного текста вниз без наложения на контент",
                    subtitle="Короткий подзаголовок",
                    text="Основной текст слайда должен оказаться ниже заголовка после перерасчета вертикального потока.",
                    preferred_layout_key="text_full_width",
                ),
            ],
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = PptxGenerator().generate(
                template_path=template_path,
                manifest=manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            presentation = Presentation(str(output_path))
            slide = presentation.slides[1]
            placeholders = {shape.placeholder_format.idx: shape for shape in slide.placeholders}

            self.assertGreater(placeholders[0].height, 600000)
            self.assertEqual(placeholders[0].width, 11198224)
            self.assertGreater(placeholders[13].top, placeholders[0].top + placeholders[0].height)
            self.assertGreater(placeholders[14].top, 1791494)
            title_runs = [run for paragraph in placeholders[0].text_frame.paragraphs for run in paragraph.runs]
            self.assertTrue(title_runs)
            self.assertLessEqual(title_runs[0].font.size.pt, 30)

    def test_generator_keeps_short_title_compact_and_readable(self) -> None:
        settings = get_settings()
        registry = TemplateRegistry(settings.templates_dir)
        manifest = registry.get_template("corp_light_v1")
        template_path = registry.get_template_pptx_path("corp_light_v1")

        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="A3 Presentation",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="A3 Presentation", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.TEXT,
                    title="Короткий заголовок",
                    text="Основной текст без подзаголовка.",
                    preferred_layout_key="text_full_width",
                ),
            ],
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = PptxGenerator().generate(
                template_path=template_path,
                manifest=manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            presentation = Presentation(str(output_path))
            slide = presentation.slides[1]
            placeholders = {shape.placeholder_format.idx: shape for shape in slide.placeholders}
            title_runs = [run for paragraph in placeholders[0].text_frame.paragraphs for run in paragraph.runs]

            self.assertTrue(title_runs)
            self.assertGreaterEqual(title_runs[0].font.size.pt, 28)
            self.assertLess(placeholders[0].height, 1000000)
            self.assertGreater(placeholders[14].top, placeholders[0].top + placeholders[0].height)

    def test_generator_reduces_body_font_for_dense_text(self) -> None:
        settings = get_settings()
        registry = TemplateRegistry(settings.templates_dir)
        manifest = registry.get_template("corp_light_v1")
        template_path = registry.get_template_pptx_path("corp_light_v1")

        dense_text = (
            "Первый длинный абзац описывает контекст, ограничения, допущения и критерии принятия решения по продуктовой стратегии. "
            "Второй длинный абзац добавляет финансовые ориентиры, риски реализации и требования к инфраструктуре. "
            "Третий длинный абзац связывает выводы с целевыми показателями, KPI и дорожной картой внедрения."
        )

        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="A3 Presentation",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="A3 Presentation", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.TEXT,
                    title="Плотный текстовый слайд",
                    text=dense_text,
                    preferred_layout_key="text_full_width",
                ),
            ],
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = PptxGenerator().generate(
                template_path=template_path,
                manifest=manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            presentation = Presentation(str(output_path))
            slide = presentation.slides[1]
            placeholders = {shape.placeholder_format.idx: shape for shape in slide.placeholders}
            body_runs = [run for paragraph in placeholders[14].text_frame.paragraphs for run in paragraph.runs]

            self.assertTrue(body_runs)
            self.assertIsNotNone(body_runs[0].font.size)
            self.assertLessEqual(body_runs[0].font.size.pt, 14)

    def test_generator_keeps_footer_in_bottom_zone(self) -> None:
        settings = get_settings()
        registry = TemplateRegistry(settings.templates_dir)
        manifest = registry.get_template("corp_light_v1")
        template_path = registry.get_template_pptx_path("corp_light_v1")

        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="Очень длинное название презентации для проверки нижнего блока и корректного положения footer",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="A3 Presentation", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.TEXT,
                    title="Тест footer",
                    text="Основной текст слайда.",
                    preferred_layout_key="text_full_width",
                ),
                SlideSpec(
                    kind=SlideKind.TABLE,
                    title="Тест footer на табличном слайде",
                    subtitle="Подзаголовок",
                    table=TableBlock(headers=["Показатель", "Значение"], rows=[["GMV", "125"]]),
                    preferred_layout_key="table",
                ),
            ],
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = PptxGenerator().generate(
                template_path=template_path,
                manifest=manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            presentation = Presentation(str(output_path))

            text_slide_placeholders = {shape.placeholder_format.idx: shape for shape in presentation.slides[1].placeholders}
            table_slide_placeholders = {shape.placeholder_format.idx: shape for shape in presentation.slides[2].placeholders}

            self.assertGreaterEqual(text_slide_placeholders[17].top, 6200000)
            self.assertGreaterEqual(table_slide_placeholders[15].top, 6200000)

    def test_generator_renders_chart_slide_into_pptx(self) -> None:
        settings = get_settings()
        registry = TemplateRegistry(settings.templates_dir)
        manifest = registry.get_template("corp_light_v1")
        template_path = registry.get_template_pptx_path("corp_light_v1")

        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="A3 Presentation",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="A3 Presentation", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.CHART,
                    title="Лиды по каналам",
                    subtitle="Тест chart render",
                    chart=ChartSpec(
                        chart_id="chart_1",
                        source_table_id="table_1",
                        chart_type=ChartType.COLUMN,
                        title="Лиды",
                        categories=["SEO", "Ads", "Referral"],
                        series=[ChartSeries(name="Лиды", values=[120.0, 200.0, 90.0])],
                        confidence=ChartConfidence.HIGH,
                    ),
                    preferred_layout_key="table",
                ),
            ],
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = PptxGenerator().generate(
                template_path=template_path,
                manifest=manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            presentation = Presentation(str(output_path))
            slide = presentation.slides[1]
            chart_shapes = [shape for shape in slide.shapes if getattr(shape, "has_chart", False)]

            self.assertEqual(len(chart_shapes), 1)
            self.assertEqual(chart_shapes[0].chart.series[0].name, "Лиды")
            self.assertEqual(chart_shapes[0].chart.series[0].format.fill.fore_color.rgb, RGBColor(0x67, 0x9A, 0xEA))
            self.assertEqual(
                chart_shapes[0].chart.chart_title.text_frame.paragraphs[0].runs[0].font.color.rgb,
                RGBColor(0x18, 0x20, 0x33),
            )

    def test_generator_styles_line_chart_markers_and_percent_format(self) -> None:
        settings = get_settings()
        registry = TemplateRegistry(settings.templates_dir)
        manifest = registry.get_template("corp_light_v1")
        template_path = registry.get_template_pptx_path("corp_light_v1")

        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="A3 Presentation",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="A3 Presentation", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.CHART,
                    title="Конверсия",
                    chart=ChartSpec(
                        chart_id="chart_line",
                        source_table_id="table_1",
                        chart_type=ChartType.LINE,
                        title="Конверсия",
                        categories=["Q1", "Q2", "Q3"],
                        series=[ChartSeries(name="CR", values=[18.0, 22.0, 27.0])],
                        confidence=ChartConfidence.HIGH,
                        data_labels_visible=True,
                        value_format="percent",
                    ),
                    preferred_layout_key="table",
                ),
            ],
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = PptxGenerator().generate(
                template_path=template_path,
                manifest=manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            presentation = Presentation(str(output_path))
            chart = next(shape.chart for shape in presentation.slides[1].shapes if getattr(shape, "has_chart", False))
            series = chart.series[0]

            self.assertEqual(series.marker.style, XL_MARKER_STYLE.CIRCLE)
            self.assertEqual(series.data_labels.number_format, '0"%"')

    def test_generator_renders_real_combo_chart_with_bar_and_line_plots(self) -> None:
        settings = get_settings()
        registry = TemplateRegistry(settings.templates_dir)
        manifest = registry.get_template("corp_light_v1")
        template_path = registry.get_template_pptx_path("corp_light_v1")

        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="A3 Presentation",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="A3 Presentation", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.CHART,
                    title="Combo",
                    chart=ChartSpec(
                        chart_id="chart_combo",
                        source_table_id="table_1",
                        chart_type=ChartType.COMBO,
                        title="План и тренд",
                        categories=["Q1", "Q2", "Q3"],
                        series=[
                            ChartSeries(name="План", values=[100.0, 130.0, 150.0]),
                            ChartSeries(name="Факт", values=[90.0, 120.0, 160.0]),
                            ChartSeries(name="Маржа", values=[18.0, 22.0, 27.0]),
                        ],
                        confidence=ChartConfidence.HIGH,
                    ),
                    preferred_layout_key="table",
                ),
            ],
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = PptxGenerator().generate(
                template_path=template_path,
                manifest=manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            presentation = Presentation(str(output_path))
            chart = next(shape.chart for shape in presentation.slides[1].shapes if getattr(shape, "has_chart", False))

            bar_charts = chart._chartSpace.xpath(".//c:barChart")
            line_charts = chart._chartSpace.xpath(".//c:lineChart")

            self.assertEqual(len(bar_charts), 1)
            self.assertEqual(len(line_charts), 1)
            self.assertEqual(len(bar_charts[0].xpath("./c:ser")), 2)
            self.assertEqual(len(line_charts[0].xpath("./c:ser")), 1)
            self.assertEqual(chart.series[-1].marker.style, XL_MARKER_STYLE.CIRCLE)

    def test_generator_formats_value_axis_in_millions_for_large_currency_values(self) -> None:
        settings = get_settings()
        registry = TemplateRegistry(settings.templates_dir)
        manifest = registry.get_template("corp_light_v1")
        template_path = registry.get_template_pptx_path("corp_light_v1")

        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="A3 Presentation",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="A3 Presentation", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.CHART,
                    title="Доход",
                    chart=ChartSpec(
                        chart_id="chart_money",
                        source_table_id="table_1",
                        chart_type=ChartType.COLUMN,
                        title="Доход",
                        categories=["Q1", "Q2", "Q3"],
                        series=[ChartSeries(name="Доход", values=[104_300_000.0, 111_300_000.0, 135_700_000.0])],
                        confidence=ChartConfidence.HIGH,
                        value_format="currency",
                    ),
                    preferred_layout_key="table",
                ),
            ],
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = PptxGenerator().generate(
                template_path=template_path,
                manifest=manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            presentation = Presentation(str(output_path))
            chart = next(shape.chart for shape in presentation.slides[1].shapes if getattr(shape, "has_chart", False))

            self.assertEqual(chart.value_axis.tick_labels.number_format, '0.0,," млн ₽"')

    def test_generator_styles_pie_points_with_distinct_brand_colors(self) -> None:
        settings = get_settings()
        registry = TemplateRegistry(settings.templates_dir)
        manifest = registry.get_template("corp_light_v1")
        template_path = registry.get_template_pptx_path("corp_light_v1")

        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="A3 Presentation",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="A3 Presentation", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.CHART,
                    title="Структура выручки",
                    chart=ChartSpec(
                        chart_id="chart_pie",
                        source_table_id="table_1",
                        chart_type=ChartType.PIE,
                        title="Структура",
                        categories=["A", "B", "C"],
                        series=[ChartSeries(name="Доли", values=[40.0, 35.0, 25.0])],
                        confidence=ChartConfidence.HIGH,
                    ),
                    preferred_layout_key="table",
                ),
            ],
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = PptxGenerator().generate(
                template_path=template_path,
                manifest=manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            presentation = Presentation(str(output_path))
            chart = next(shape.chart for shape in presentation.slides[1].shapes if getattr(shape, "has_chart", False))
            points = chart.series[0].points

            self.assertEqual(points[0].format.fill.fore_color.rgb, RGBColor(0x67, 0x9A, 0xEA))
            self.assertEqual(points[1].format.fill.fore_color.rgb, RGBColor(0x5A, 0xB2, 0x9C))

    def test_generator_resolves_supported_chart_types_and_combo_fallback(self) -> None:
        generator = PptxGenerator()

        def make_spec(chart_type: ChartType) -> ChartSpec:
            return ChartSpec(
                chart_id=f"chart_{chart_type.value}",
                source_table_id="table_1",
                chart_type=chart_type,
                title="Тест",
                categories=["A", "B"],
                series=[ChartSeries(name="Series", values=[10.0, 20.0])],
                confidence=ChartConfidence.HIGH,
            )

        self.assertEqual(generator._resolve_chart_type(make_spec(ChartType.BAR)), XL_CHART_TYPE.BAR_CLUSTERED)
        self.assertEqual(generator._resolve_chart_type(make_spec(ChartType.COLUMN)), XL_CHART_TYPE.COLUMN_CLUSTERED)
        self.assertEqual(generator._resolve_chart_type(make_spec(ChartType.LINE)), XL_CHART_TYPE.LINE)
        self.assertEqual(generator._resolve_chart_type(make_spec(ChartType.STACKED_BAR)), XL_CHART_TYPE.BAR_STACKED)
        self.assertEqual(generator._resolve_chart_type(make_spec(ChartType.STACKED_COLUMN)), XL_CHART_TYPE.COLUMN_STACKED)
        self.assertEqual(generator._resolve_chart_type(make_spec(ChartType.PIE)), XL_CHART_TYPE.PIE)
        self.assertEqual(generator._resolve_chart_type(make_spec(ChartType.COMBO)), XL_CHART_TYPE.COLUMN_CLUSTERED)

    def test_generator_adapts_cover_title_height_and_meta_spacing(self) -> None:
        settings = get_settings()
        registry = TemplateRegistry(settings.templates_dir)
        manifest = registry.get_template("corp_light_v1")
        template_path = registry.get_template_pptx_path("corp_light_v1")

        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="A3 Presentation",
            slides=[
                SlideSpec(
                    kind=SlideKind.TITLE,
                    title="Оптимальное количество поставщиков для А3 в регулярных платежах - методология, сценарии и рекомендации для CEO",
                    notes='Контекст задачи и что именно считать "поставщиком"',
                    preferred_layout_key="cover",
                ),
            ],
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = PptxGenerator().generate(
                template_path=template_path,
                manifest=manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            presentation = Presentation(str(output_path))
            slide = presentation.slides[0]
            placeholders = {shape.placeholder_format.idx: shape for shape in slide.placeholders}
            title = placeholders[0]
            meta = placeholders[15]
            title_runs = [run for paragraph in title.text_frame.paragraphs for run in paragraph.runs if run.text]

            self.assertTrue(title_runs)
            self.assertLessEqual(title_runs[0].font.size.pt, 46)
            self.assertGreater(meta.top, title.top + title.height)
            self.assertGreaterEqual(title.height, 1422646)


if __name__ == "__main__":
    unittest.main()
