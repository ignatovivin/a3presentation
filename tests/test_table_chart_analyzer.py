from __future__ import annotations

import unittest
from pathlib import Path

from a3presentation.domain.chart import ChartTableClassification, ChartType
from a3presentation.domain.presentation import TableBlock
from a3presentation.services.document_text_extractor import DocumentTextExtractor
from a3presentation.services.table_chart_analyzer import TableChartAnalyzer


class TableChartAnalyzerTests(unittest.TestCase):
    def setUp(self) -> None:
        self.analyzer = TableChartAnalyzer()

    def test_single_series_category_table_is_chartable(self) -> None:
        table = TableBlock(
            headers=["Канал", "Лиды"],
            rows=[
                ["SEO", "120"],
                ["Ads", "95"],
                ["Partners", "80"],
            ],
        )

        assessment = self.analyzer.analyze(table, table_id="marketing_leads")

        self.assertTrue(assessment.chartable)
        self.assertEqual(assessment.classification, ChartTableClassification.SINGLE_SERIES_CATEGORY)
        self.assertEqual(assessment.candidate_specs[0].chart_type, ChartType.COLUMN)
        self.assertEqual(assessment.candidate_specs[0].categories, ["SEO", "Ads", "Partners"])

    def test_text_dominant_table_is_not_chartable(self) -> None:
        table = TableBlock(
            headers=["Этап", "Комментарий"],
            rows=[
                ["Q1", "Аудит текущего позиционирования и сбор команды"],
                ["Q2", "Создание бренд-стратегии и айдентики"],
                ["Q3", "Запуск сайта и тестирование гипотез"],
            ],
        )

        assessment = self.analyzer.analyze(table, table_id="roadmap")

        self.assertFalse(assessment.chartable)
        self.assertEqual(assessment.classification, ChartTableClassification.TEXT_DOMINANT)

    def test_ordinal_status_table_is_not_chartable(self) -> None:
        table = TableBlock(
            headers=["Статус", "Этап"],
            rows=[
                ["Исследование", "1"],
                ["Концепция", "2"],
                ["Прототип", "3"],
                ["Тестирование", "4"],
                ["Запуск", "5"],
            ],
        )

        assessment = self.analyzer.analyze(table, table_id="customer_journey")

        self.assertFalse(assessment.chartable)
        self.assertEqual(assessment.classification, ChartTableClassification.NOT_CHARTABLE)
        self.assertTrue(any("ordinal index" in warning for warning in assessment.warnings))

    def test_quarterly_financial_table_from_real_docx_is_chartable(self) -> None:
        path = Path(r"C:\Users\mrfra\Desktop\A3_Strategy_2026_review_1.docx")
        if not path.exists():
            self.skipTest("real fixture document is not available on this machine")

        _, tables, _ = DocumentTextExtractor().extract(path.name, path.read_bytes())
        source_table = next(
            table
            for table in tables
            if table.headers == ["Квартал", "Выручка б/НДС", "Чистая прибыль (от выручки)"]
        )

        assessment = self.analyzer.analyze(source_table, table_id="quarter_profile")

        self.assertTrue(assessment.chartable)
        self.assertEqual(assessment.classification, ChartTableClassification.TIME_SERIES)
        self.assertGreaterEqual(len(assessment.candidate_specs), 2)
        self.assertEqual(assessment.candidate_specs[0].series[0].name, "Выручка б/НДС")
        self.assertEqual(assessment.candidate_specs[0].categories, ["Q1 (янв–мар)", "Q2 (апр–июн)", "Q3 (июл–сен)", "Q4 (окт–дек)"])

    def test_composition_table_suggests_pie_chart_first(self) -> None:
        table = TableBlock(
            headers=["Канал", "Доля"],
            rows=[
                ["SEO", "45%"],
                ["Ads", "35%"],
                ["Partners", "20%"],
            ],
        )

        assessment = self.analyzer.analyze(table, table_id="channel_mix")

        self.assertTrue(assessment.chartable)
        self.assertEqual(assessment.classification, ChartTableClassification.COMPOSITION)
        self.assertEqual([spec.chart_type for spec in assessment.candidate_specs[:3]], [ChartType.PIE, ChartType.COLUMN, ChartType.BAR])
        self.assertTrue(assessment.candidate_specs[0].data_labels_visible)

    def test_multi_series_same_unit_table_suggests_stacked_chart(self) -> None:
        table = TableBlock(
            headers=["Сегмент", "Новые", "Повторные"],
            rows=[
                ["SMB", "120", "80"],
                ["Mid", "90", "70"],
                ["Enterprise", "45", "55"],
            ],
        )

        assessment = self.analyzer.analyze(table, table_id="segment_clients")

        self.assertTrue(assessment.chartable)
        self.assertEqual(assessment.classification, ChartTableClassification.MULTI_SERIES_CATEGORY)
        self.assertEqual(assessment.candidate_specs[0].chart_type, ChartType.STACKED_COLUMN)
        self.assertEqual(assessment.candidate_specs[0].stacking, "stacked")

    def test_mixed_unit_table_suggests_column_and_line_without_combo_by_default(self) -> None:
        table = TableBlock(
            headers=["Квартал", "Выручка", "Маржа"],
            rows=[
                ["Q1", "120 млн руб", "18%"],
                ["Q2", "150 млн руб", "22%"],
                ["Q3", "190 млн руб", "27%"],
            ],
        )

        assessment = self.analyzer.analyze(table, table_id="revenue_margin")

        self.assertTrue(assessment.chartable)
        self.assertEqual(assessment.classification, ChartTableClassification.TIME_SERIES)
        self.assertEqual(
            [spec.chart_type for spec in assessment.candidate_specs],
            [ChartType.COLUMN, ChartType.LINE],
        )
        self.assertTrue(all(spec.chart_type != ChartType.COMBO for spec in assessment.candidate_specs))
        self.assertTrue(any("single value axis" in warning for warning in assessment.candidate_specs[0].warnings))

    def test_table_with_too_many_mixed_units_is_not_chartable(self) -> None:
        table = TableBlock(
            headers=["Метрика", "Деньги", "Доля", "Количество"],
            rows=[
                ["A", "120 млн руб", "18%", "25"],
                ["B", "150 млн руб", "22%", "31"],
                ["C", "190 млн руб", "27%", "44"],
            ],
        )

        assessment = self.analyzer.analyze(table, table_id="mixed_metrics")

        self.assertFalse(assessment.chartable)
        self.assertEqual(assessment.classification, ChartTableClassification.NOT_CHARTABLE)
        self.assertTrue(any("mixed units are too ambiguous" in warning for warning in assessment.warnings))

    def test_summary_row_is_filtered_from_chart_series(self) -> None:
        table = TableBlock(
            headers=["Канал", "Лиды"],
            rows=[
                ["SEO", "120"],
                ["Ads", "95"],
                ["Partners", "80"],
                ["Итого", "295"],
            ],
        )

        assessment = self.analyzer.analyze(table, table_id="marketing_leads_total")

        self.assertTrue(assessment.chartable)
        self.assertEqual(assessment.structured_table.summary_rows, [4])
        self.assertEqual(assessment.candidate_specs[0].categories, ["SEO", "Ads", "Partners"])
        self.assertEqual(assessment.candidate_specs[0].series[0].values, [120.0, 95.0, 80.0])
        self.assertTrue(any("summary row filtered" in warning for warning in assessment.warnings))


if __name__ == "__main__":
    unittest.main()
