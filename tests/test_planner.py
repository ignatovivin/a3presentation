from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_MARKER_STYLE
from pptx.enum.text import MSO_AUTO_SIZE

from a3presentation.domain.api import ChartOverride, DocumentBlock
from a3presentation.domain.chart import ChartConfidence, ChartSeries, ChartSpec, ChartType
from a3presentation.domain.presentation import PresentationPlan, SlideKind, SlideSpec, TableBlock
from a3presentation.services.pptx_generator import PptxGenerator
from a3presentation.services.planner import TextToPlanService
from a3presentation.services.template_registry import TemplateRegistry
from a3presentation.settings import get_settings


class TextToPlanServiceTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        settings = get_settings()
        registry = TemplateRegistry(settings.templates_dir)
        cls.manifest = registry.get_template("corp_light_v1")
        cls.template_path = registry.get_template_pptx_path("corp_light_v1")

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

    def test_cover_with_five_leading_lines_does_not_create_duplicate_content_slide(self) -> None:
        service = TextToPlanService()
        blocks = [
            DocumentBlock(kind="paragraph", text="А3"),
            DocumentBlock(kind="paragraph", text="Бизнес-стратегия 2026"),
            DocumentBlock(kind="paragraph", text="Горизонт планирования: 2026-2030"),
            DocumentBlock(kind="paragraph", text="Март 2026"),
            DocumentBlock(kind="paragraph", text="Конфиденциальный документ"),
            DocumentBlock(kind="heading", text="1. Vision и стратегические цели", level=1),
            DocumentBlock(kind="paragraph", text="Основной контент раздела."),
        ]

        plan = service.build_plan(
            template_id="corp_light_v1",
            raw_text="\n".join(block.text or "" for block in blocks),
            title=None,
            tables=[],
            blocks=blocks,
        )

        self.assertEqual(plan.slides[0].title, "А3 Бизнес-стратегия 2026")
        self.assertFalse(any(slide.title == "А3" for slide in plan.slides[1:]))

    def test_mixed_section_preserves_paragraphs_when_bullets_force_split(self) -> None:
        service = TextToPlanService()
        blocks = [
            DocumentBlock(kind="paragraph", text="A3"),
            DocumentBlock(kind="heading", text="5.3 R&D новые продукты", level=1),
            DocumentBlock(
                kind="paragraph",
                text="Цель: построить системный пайплайн проверки гипотез и новых продуктов для роста выручки и диверсификации бизнеса А3.",
            ),
            DocumentBlock(
                kind="list",
                items=[
                    "Процесс Double Diamond запущен, первые инициативы в работе",
                    "Банк идей = единый бэклог, синхронизирован с Jira",
                    "ICE-приоритизация со стейкхолдерами раз в 2 недели",
                    "Юнит-экономика по поставщикам и партнёрам",
                    "Инициатива: Инфраструктура цифрового рубля",
                ],
            ),
            DocumentBlock(
                kind="paragraph",
                text="Q2 2026 (Discovery): Анализ архитектуры платформы цифрового рубля ЦБ и определение роли А3.",
            ),
            DocumentBlock(
                kind="list",
                items=[
                    "Q3 2026 (Proof of Concept): Техническое прототипирование с командой Payments",
                    "Q4 2026 — Пилот: Пилотный запуск с ограниченным объемом",
                ],
            ),
        ]

        plan = service.build_plan(
            template_id="corp_light_v1",
            raw_text="\n".join(block.text or "" for block in blocks),
            title=None,
            tables=[],
            blocks=blocks,
        )

        bullets_text = "\n".join("\n".join(slide.bullets) for slide in plan.slides if slide.kind == SlideKind.BULLETS)
        self.assertIn("Цель: построить системный пайплайн проверки гипотез", bullets_text)
        self.assertIn("Q2 2026 (Discovery): Анализ архитектуры платформы цифрового рубля ЦБ", bullets_text)

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

    def test_split_table_keeps_compact_two_column_ten_row_table_on_single_slide(self) -> None:
        service = TextToPlanService()
        table = TableBlock(
            headers=["Показатель", "Значение"],
            rows=[
                ["Выручка с НДС", "2 433 млн ₽"],
                ["Выручка без НДС", "2 122 млн ₽"],
                ["Нетранзакционный доход", "470 млн ₽ и выше"],
                ["Доля Альфа-Банка", "< 40% выручки"],
                ["НДС от выручки", "%"],
                ["РНКО", "Лицензия ЦБ и запуск у 1 партнёра"],
                ["А3 Лаб", "Создание продукта для формирования п/п"],
                ["Poseidon", "Промышленное использование у 6+ партнёров"],
                ["Банки-партнёры", "3-4"],
                ["Поставщики", "+1 000 новых подключений, итого 3 000+"],
            ],
        )

        chunks = service._split_table_for_slides(table)

        self.assertEqual(len(chunks), 1)

    def test_split_table_keeps_compact_five_column_market_map_on_single_slide(self) -> None:
        service = TextToPlanService()
        table = TableBlock(
            headers=["Игрок", "Тип", "Фокус", "Риск", "Комментарий"],
            rows=[
                ["Сбер", "Банк-экосистема", "ЖКХ, подписки, платежи", "Высокий", "Есть уникальный сервис, но слабый аккаунтинг и медлительность"],
                ["ВТБ", "Банк-экосистема", "Платежи, подписки", "Высокий", "Партнёр А3, пока нет мультибанковой сети"],
                ["Т Банк", "Банк-экосистема", "Платежи, персонализация", "Средний", "Нет мультибанковой сети и планов на её создание"],
                ["Монета", "Платёжные сервисы", "Онлайн-платежи, ЖКХ", "Средний", "Есть B2C и B2B модель, слабый аккаунтинг"],
                ["Элекснет", "Платёжный сервис", "ЖКХ, интернет, связь", "Низкий", "Есть мультибанковская сеть, нет доступа к ГИС"],
                ["Система Город", "Региональный агрегатор", "Оплата ЖКХ, локальные поставщики", "Низкий", "Локальный охват без масштаба и технологий"],
                ["Бурмистр", "SaaS для ЖКХ", "Управление УК и оплата ЖКХ", "Низкий", "Другой слой цепочки, нет банковской сети"],
            ],
        )

        chunks = service._split_table_for_slides(table)

        self.assertEqual(len(chunks), 1)

    def test_split_table_keeps_seven_row_three_column_journey_on_single_slide(self) -> None:
        service = TextToPlanService()
        table = TableBlock(
            headers=["Этап", "Статус", "Комментарий"],
            rows=[
                ["01. Лид", "Работает", "CRM / ГИС ЖКХ / холодный поиск / портал ЖКХ"],
                ["02. Договор", "Работает", ""],
                ["03. Интеграция", "Конкур. преим.", "AI: 5 минут, аналитики и PM"],
                [
                    "04. Открытие банками",
                    "Самостоятельное открытие банками",
                    "Банки сами решают срок открытия, внедряем систему для активного открытия и автоуведомлений",
                ],
                ["05. Выход на объём", "6 мес. лаг", "Около 6 месяцев до нормального платёжного потока"],
                ["06. Поддержка", "Улучшить", ""],
                ["07. Метрики", "В процессе", "NPS/CSAT внедряется"],
            ],
        )

        chunks = service._split_table_for_slides(table)

        self.assertEqual(len(chunks), 1)

    def test_split_table_reduces_over_fragmentation_for_risk_tables(self) -> None:
        service = TextToPlanService()
        table = TableBlock(
            headers=["Риск", "Вероятность", "Митигация"],
            rows=[
                ["Сбер/ВТБ строят собственную инфраструктуру", "Средняя", "Развитие уникальных технологий и партнёрств"],
                ["Регуляторные изменения ЦБ", "Средняя", "Проактивный мониторинг и РНКО как инструмент независимости"],
                ["Уход ключевых партнёров", "Низкая", "Диверсификация долей и рост новых банков"],
                ["Отмена реестра российского ПО", "Средняя", "Альтернативные модели налоговой оптимизации"],
                ["AI у конкурентов", "Средняя", "Инвестиции в AI и R&D фабрику гипотез"],
                ["Ужесточение налоговой политики", "Средняя", "Сценарий уже заложен в бюджет"],
                ["Задержка РНКО на год", "Средняя", "Скорректировать стратегию без смены курса"],
                ["Рост негатива из-за СМЭВ-согласий", "Средняя", "Коммуникационная стратегия и работа с брендом"],
                ["Потеря доступа к данным ГИС", "Низкая", "Прямые договоры и альтернативные источники данных"],
                [
                    "Макроэкономические факторы: высокая ключевая ставка, инфляция и санкционное давление",
                    "Высокая",
                    "Даже при сохранении текущего давления регулярные обязательные платежи остаются антицикличным сегментом и поддерживают рост GMV",
                ],
            ],
        )

        chunks = service._split_table_for_slides(table)

        self.assertLessEqual(len(chunks), 2)

    def test_planner_and_generator_keep_compact_table_on_single_slide(self) -> None:
        service = TextToPlanService()
        blocks = [
            DocumentBlock(kind="title", text="A3 Presentation", level=0),
            DocumentBlock(kind="heading", text="1.3 Цели 2026", level=1),
            DocumentBlock(
                kind="table",
                table=TableBlock(
                    headers=["Показатель", "Значение"],
                    rows=[
                        ["Выручка с НДС", "2 433 млн ₽"],
                        ["Выручка без НДС", "2 122 млн ₽"],
                        ["Нетранзакционный доход", "470 млн ₽ и выше"],
                        ["Доля Альфа-Банка", "< 40% выручки"],
                        ["НДС от выручки", "%"],
                        ["РНКО", "Лицензия ЦБ и запуск у 1 партнёра"],
                        ["А3 Лаб", "Создание продукта для формирования п/п"],
                        ["Poseidon", "Промышленное использование у 6+ партнёров"],
                        ["Банки-партнёры", "3-4"],
                        ["Поставщики", "+1 000 новых подключений, итого 3 000+"],
                    ],
                ),
            ),
            DocumentBlock(kind="heading", text="2. Следующий раздел", level=1),
            DocumentBlock(kind="paragraph", text="Текст для предотвращения fallback-сценария."),
        ]

        plan = service.build_plan(
            template_id="corp_light_v1",
            raw_text="\n".join(block.text or "" for block in blocks),
            title="A3 Presentation",
            blocks=blocks,
        )

        table_slides = [slide for slide in plan.slides if slide.kind == SlideKind.TABLE and slide.title == "1.3 Цели 2026"]
        self.assertEqual(len(table_slides), 1)
        self.assertEqual(len(table_slides[0].table.rows), 10)

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = PptxGenerator().generate(
                template_path=self.template_path,
                manifest=self.manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            presentation = Presentation(str(output_path))
            self.assertEqual(len(presentation.slides), len(plan.slides))
            table_shapes = [shape for shape in presentation.slides[1].shapes if getattr(shape, "has_table", False)]
            self.assertEqual(len(table_shapes), 1)
            rendered_table = table_shapes[0].table
            self.assertEqual(len(rendered_table.rows), 11)
            self.assertEqual(rendered_table.cell(1, 0).text, "Выручка с НДС")

    def test_planner_and_generator_keep_competitor_map_in_single_rendered_table(self) -> None:
        service = TextToPlanService()
        section = [
            DocumentBlock(kind="title", text="A3 Presentation", level=0),
            DocumentBlock(kind="heading", text="3.3 Карта конкурентов", level=1),
            DocumentBlock(kind="paragraph", text="Контекст конкурентной карты с вводным описанием рынка."),
            DocumentBlock(
                kind="table",
                table=TableBlock(
                    headers=["Игрок", "Тип", "Фокус", "Риск", "Комментарий"],
                    rows=[
                        ["Сбер", "Банк-экосистема", "ЖКХ, подписки, платежи", "Высокий", "Есть уникальный сервис, но слабый аккаунтинг и медлительность"],
                        ["ВТБ", "Банк-экосистема", "Платежи, подписки", "Высокий", "Партнёр А3, пока нет мультибанковой сети"],
                        ["Т Банк", "Банк-экосистема", "Платежи, персонализация", "Средний", "Нет мультибанковой сети и планов на её создание"],
                        ["Монета", "Платёжные сервисы", "Онлайн-платежи, ЖКХ", "Средний", "Есть B2C и B2B модель, слабый аккаунтинг"],
                        ["Элекснет", "Платёжный сервис", "ЖКХ, интернет, связь", "Низкий", "Есть мультибанковская сеть, нет доступа к ГИС"],
                        ["Система Город", "Региональный агрегатор", "Оплата ЖКХ, локальные поставщики", "Низкий", "Локальный охват без масштаба и технологий"],
                        ["Бурмистр", "SaaS для ЖКХ", "Управление УК и оплата ЖКХ", "Низкий", "Другой слой цепочки, нет банковской сети"],
                    ],
                ),
            ),
            DocumentBlock(kind="heading", text="4. Следующий раздел", level=1),
            DocumentBlock(kind="paragraph", text="Текст для устойчивой narrative-классификации."),
        ]

        plan = service.build_plan(
            template_id="corp_light_v1",
            raw_text="\n".join(block.text or "" for block in section),
            title="A3 Presentation",
            blocks=section,
        )

        table_slides = [slide for slide in plan.slides if slide.kind == SlideKind.TABLE and "3.3 Карта конкурентов" in (slide.title or "")]
        self.assertEqual(len(table_slides), 1)
        self.assertEqual(len(table_slides[0].table.rows), 7)

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = PptxGenerator().generate(
                template_path=self.template_path,
                manifest=self.manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            presentation = Presentation(str(output_path))
            matching_slides = []
            for slide in presentation.slides:
                slide_text = " ".join(shape.text for shape in slide.shapes if getattr(shape, "has_text_frame", False))
                if "3.3 Карта конкурентов" in slide_text:
                    matching_slides.append(slide)
            rendered_table_slides = [slide for slide in matching_slides if any(getattr(shape, "has_table", False) for shape in slide.shapes)]
            self.assertEqual(len(rendered_table_slides), 1)
            table_shape = next(shape for shape in rendered_table_slides[0].shapes if getattr(shape, "has_table", False))
            self.assertEqual(len(table_shape.table.rows), 8)

    def test_generator_expands_long_title_and_pushes_content_down(self) -> None:
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
                template_path=self.template_path,
                manifest=self.manifest,
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
                template_path=self.template_path,
                manifest=self.manifest,
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
                template_path=self.template_path,
                manifest=self.manifest,
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
            self.assertEqual(placeholders[14].text_frame.auto_size, MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE)

    def test_generator_applies_explicit_font_size_to_sparse_bullet_slide(self) -> None:
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="A3 Presentation",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="A3 Presentation", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.BULLETS,
                    title="R&D continuation",
                    bullets=[
                        "Юнит-экономика: по поставщикам и интеграциям",
                        "Инициатива: Инфраструктура цифрового рубля",
                        "Q3 2026: Техническое прототипирование",
                        "Q4 2026: Пилотный запуск и оценка unit-экономики",
                    ],
                    preferred_layout_key="list_full_width",
                ),
            ],
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = PptxGenerator().generate(
                template_path=self.template_path,
                manifest=self.manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            presentation = Presentation(str(output_path))
            placeholders = {shape.placeholder_format.idx: shape for shape in presentation.slides[1].placeholders}
            body = placeholders[14]
            body_runs = [run for paragraph in body.text_frame.paragraphs for run in paragraph.runs]

            self.assertTrue(body_runs)
            self.assertTrue(all(run.font.size is not None for run in body_runs))
            self.assertGreaterEqual(body_runs[0].font.size.pt, 14)
            self.assertEqual(body.text_frame.auto_size, MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE)

    def test_generator_shrinks_dense_bullet_container_to_avoid_overflow(self) -> None:
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="A3 Presentation",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="A3 Presentation", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.BULLETS,
                    title="Плотный appendix slide",
                    bullets=[
                        "Q10: Зачем собственная инфраструктура Poseidon и второй ЦОД для критической финансовой платформы с жёсткими SLA и требованиями по контролю среды.",
                        "Q11: Почему целевыми считаются малые банки, несмотря на концентрацию выручки в верхнем сегменте и более скромный текущий оборот.",
                        "Q12: Что будет происходить после подключения 3000 поставщиков и какие операционные ограничения нужно закрыть заранее.",
                        "Контроль SLA: внешние решения не дают достаточного уровня гарантии для финансовой инфраструктуры и государственных интеграций.",
                        "Независимость от вендоров: критична для скорости изменений, контроля инцидентов и стоимости владения платформой.",
                        "Защита данных: госинтеграции требуют более высокого уровня контроля над инфраструктурой и доступом.",
                        "Инфраструктурное проникновение как барьер для конкурентов и основа для новых продуктов.",
                        "Дополнительный заработок на минимальных контрактах и последующем upsell поверх инфраструктурного контура.",
                    ],
                    preferred_layout_key="list_full_width",
                ),
            ],
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = PptxGenerator().generate(
                template_path=self.template_path,
                manifest=self.manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            presentation = Presentation(str(output_path))
            placeholders = {shape.placeholder_format.idx: shape for shape in presentation.slides[1].placeholders}
            body = placeholders[14]
            body_runs = [run for paragraph in body.text_frame.paragraphs for run in paragraph.runs]

            self.assertTrue(body_runs)
            self.assertLessEqual(body_runs[0].font.size.pt, 13)
            self.assertEqual(body.text_frame.auto_size, MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE)

    def test_generator_keeps_footer_in_bottom_zone(self) -> None:
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
                template_path=self.template_path,
                manifest=self.manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            presentation = Presentation(str(output_path))

            text_slide_placeholders = {shape.placeholder_format.idx: shape for shape in presentation.slides[1].placeholders}
            table_slide_placeholders = {shape.placeholder_format.idx: shape for shape in presentation.slides[2].placeholders}

            self.assertGreaterEqual(text_slide_placeholders[17].top, 6200000)
            self.assertGreaterEqual(table_slide_placeholders[15].top, 6200000)

    def test_generator_expands_table_footer_to_full_width_without_subtitle(self) -> None:
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="A3 Presentation",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="A3 Presentation", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.TABLE,
                    title="Таблица без подзаголовка",
                    table=TableBlock(headers=["Показатель", "Значение"], rows=[["GMV", "125"]]),
                    preferred_layout_key="table",
                ),
            ],
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = PptxGenerator().generate(
                template_path=self.template_path,
                manifest=self.manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            presentation = Presentation(str(output_path))
            placeholders = {shape.placeholder_format.idx: shape for shape in presentation.slides[1].placeholders}

            self.assertEqual(placeholders[15].left, 442913)
            self.assertEqual(placeholders[15].width, 11198224)
            self.assertGreaterEqual(placeholders[15].top, 6200000)

    def test_generator_keeps_table_subtitle_readable_instead_of_tiny_font(self) -> None:
        plan = PresentationPlan(
            template_id="corp_light_v1",
            title="A3 Presentation",
            slides=[
                SlideSpec(kind=SlideKind.TITLE, title="A3 Presentation", preferred_layout_key="cover"),
                SlideSpec(
                    kind=SlideKind.TABLE,
                    title="Раздел",
                    subtitle="6.1 РНКО — получение лицензии",
                    table=TableBlock(
                        headers=["Показатель", "Значение"],
                        rows=[["Статус", "В работе"], ["Срок", "Q4 2026"]],
                    ),
                    preferred_layout_key="table",
                ),
            ],
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = PptxGenerator().generate(
                template_path=self.template_path,
                manifest=self.manifest,
                plan=plan,
                output_dir=Path(temp_dir),
            )
            presentation = Presentation(str(output_path))
            subtitle = next(
                placeholder for placeholder in presentation.slides[1].placeholders if placeholder.placeholder_format.idx == 13
            )
            font_sizes = sorted(
                {
                    run.font.size.pt
                    for paragraph in subtitle.text_frame.paragraphs
                    for run in paragraph.runs
                    if run.font.size is not None
                }
            )

            self.assertTrue(font_sizes)
            self.assertGreaterEqual(min(font_sizes), 13.0)

    def test_generator_renders_chart_slide_into_pptx(self) -> None:
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
                template_path=self.template_path,
                manifest=self.manifest,
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
