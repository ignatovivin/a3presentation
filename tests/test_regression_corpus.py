from __future__ import annotations

import base64
import tempfile
import unittest
from io import BytesIO
from pathlib import Path

from docx import Document
from pptx import Presentation

from a3presentation.services.document_text_extractor import DocumentTextExtractor
from a3presentation.services.planner import TextToPlanService
from a3presentation.services.pptx_generator import PptxGenerator
from a3presentation.services.template_registry import TemplateRegistry
from a3presentation.settings import get_settings


SMALL_PNG_BYTES = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO9W6i8AAAAASUVORK5CYII="
)


class RegressionCorpusTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.fixtures_dir = Path(__file__).parent / "fixtures" / "regression"
        settings = get_settings()
        registry = TemplateRegistry(settings.templates_dir)
        cls.manifest = registry.get_template("corp_light_v1")
        cls.template_path = registry.get_template_pptx_path("corp_light_v1")

    def test_fixture_text_documents_generate_non_empty_presentations(self) -> None:
        cases = [
            ("strategy_report.md", "Стратегический markdown-документ"),
            ("mixed_notes.txt", "Смешанный текстовый документ"),
        ]
        extractor = DocumentTextExtractor()
        planner = TextToPlanService()

        for fixture_name, label in cases:
            with self.subTest(case=label):
                content = (self.fixtures_dir / fixture_name).read_bytes()
                text, tables, blocks = extractor.extract(fixture_name, content)
                plan = planner.build_plan("corp_light_v1", text, None, tables, blocks)

                self.assertGreaterEqual(len(plan.slides), 2)
                self.assertTrue(any((slide.title or "").strip() for slide in plan.slides[1:]))

                with tempfile.TemporaryDirectory() as temp_dir:
                    output_path = PptxGenerator().generate(
                        template_path=self.template_path,
                        manifest=self.manifest,
                        plan=plan,
                        output_dir=Path(temp_dir),
                    )
                    presentation = Presentation(str(output_path))
                    self.assertEqual(len(presentation.slides), len(plan.slides))

    def test_generated_docx_corpus_covers_report_form_resume_and_table_heavy(self) -> None:
        cases = [
            ("report.docx", self._build_report_docx()),
            ("form.docx", self._build_form_docx()),
            ("resume.docx", self._build_resume_docx()),
            ("table-heavy.docx", self._build_table_heavy_docx()),
        ]
        extractor = DocumentTextExtractor()
        planner = TextToPlanService()

        for filename, content in cases:
            with self.subTest(case=filename):
                text, tables, blocks = extractor.extract(filename, content)
                plan = planner.build_plan("corp_light_v1", text, None, tables, blocks)

                self.assertGreaterEqual(len(plan.slides), 2)
                self.assertFalse(all(slide.kind.value == "title" for slide in plan.slides))

                with tempfile.TemporaryDirectory() as temp_dir:
                    output_path = PptxGenerator().generate(
                        template_path=self.template_path,
                        manifest=self.manifest,
                        plan=plan,
                        output_dir=Path(temp_dir),
                    )
                    presentation = Presentation(str(output_path))
                    self.assertEqual(len(presentation.slides), len(plan.slides))

    def test_strategy_like_docx_skips_cover_lines_as_first_content_section(self) -> None:
        extractor = DocumentTextExtractor()
        planner = TextToPlanService()

        text, tables, blocks = extractor.extract("report.docx", self._build_report_docx())
        plan = planner.build_plan("corp_light_v1", text, None, tables, blocks)

        self.assertGreaterEqual(len(plan.slides), 2)
        self.assertNotEqual(plan.slides[1].title, "A3")
        self.assertIn("Vision", plan.slides[1].title)

    def _build_report_docx(self) -> bytes:
        document = Document()
        document.add_paragraph("A3")
        document.add_paragraph("Бизнес-стратегия 2026")
        document.add_heading("1. Vision и стратегические цели", level=1)
        document.add_paragraph(
            "Платформа должна расти за счет новых сегментов, устойчивой экономики продукта и масштабируемой архитектуры."
        )
        document.add_heading("2. Позиционирование", level=1)
        document.add_paragraph(
            "Компания должна восприниматься как инфраструктурный игрок, ускоряющий финансовые процессы клиентов."
        )
        return self._save_document(document)

    def _build_form_docx(self) -> bytes:
        document = Document()
        document.add_paragraph("Анкета нового сотрудника")
        document.add_paragraph("ФИО: Иван Игнатов")
        document.add_paragraph("Дата рождения: 05.09.1994")
        document.add_paragraph("Телефон: +7 999 000-00-00")
        document.add_paragraph("Email: ivan@example.com")
        table = document.add_table(rows=3, cols=2)
        table.cell(0, 0).text = "Поле"
        table.cell(0, 1).text = "Значение"
        table.cell(1, 0).text = "Город"
        table.cell(1, 1).text = "Екатеринбург"
        table.cell(2, 0).text = "Отдел"
        table.cell(2, 1).text = "Разработка"
        return self._save_document(document)

    def _build_resume_docx(self) -> bytes:
        document = Document()
        document.add_paragraph("Иван Игнатов")
        document.add_paragraph("ivan@example.com")
        document.add_paragraph("+7 999 000-00-00")
        document.add_paragraph("ОПЫТ РАБОТЫ")
        document.add_paragraph("Руководил развитием внутренних платформ и продуктовой аналитики.")
        document.add_paragraph("НАВЫКИ")
        document.add_paragraph("Стратегия, аналитика, управление продуктом, интеграции.")
        return self._save_document(document)

    def _build_table_heavy_docx(self) -> bytes:
        document = Document()
        document.add_paragraph("Сводка KPI по сегментам")
        table = document.add_table(rows=5, cols=3)
        headers = ["Сегмент", "Выручка", "Маржа"]
        for col_index, header in enumerate(headers):
            table.cell(0, col_index).text = header
        rows = [
            ("SMB", "120", "18%"),
            ("Enterprise", "260", "24%"),
            ("PSP", "180", "21%"),
            ("Acquiring", "140", "17%"),
        ]
        for row_index, row in enumerate(rows, start=1):
            for col_index, value in enumerate(row):
                table.cell(row_index, col_index).text = value
        document.add_picture(BytesIO(SMALL_PNG_BYTES))
        return self._save_document(document)

    def _save_document(self, document: Document) -> bytes:
        buffer = BytesIO()
        document.save(buffer)
        return buffer.getvalue()


if __name__ == "__main__":
    unittest.main()
