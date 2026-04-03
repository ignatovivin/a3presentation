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


if __name__ == "__main__":
    unittest.main()
