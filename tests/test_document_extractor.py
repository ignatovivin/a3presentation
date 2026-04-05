from __future__ import annotations

from io import BytesIO
import unittest

from docx import Document

from a3presentation.services.document_text_extractor import DocumentTextExtractor


class DocumentTextExtractorTests(unittest.TestCase):
    def test_markdown_extracts_headings_and_lists(self) -> None:
        extractor = DocumentTextExtractor()
        markdown = (
            "# Рынок медицины для A3\n\n"
            "## Основные сегменты\n"
            "- Частные клиники\n"
            "- Телемедицина\n\n"
            "Рынок быстро растет за счет цифровизации и частного спроса.\n"
        ).encode("utf-8")

        text, tables, blocks = extractor.extract("sample.md", markdown)

        self.assertTrue(text)
        self.assertEqual(tables, [])
        self.assertGreaterEqual(len(blocks), 3)
        self.assertEqual(blocks[0].kind, "heading")
        self.assertEqual(blocks[0].text, "Рынок медицины для A3")
        self.assertEqual(blocks[1].kind, "subheading")
        self.assertEqual(blocks[2].kind, "list")
        self.assertEqual(blocks[2].items, ["Частные клиники", "Телемедицина"])

    def test_docx_detects_outline_and_quarter_headings_without_explicit_heading_style(self) -> None:
        extractor = DocumentTextExtractor()
        document = Document()
        document.add_paragraph("A3")
        document.add_paragraph("1.4 Показатели 2025")
        document.add_paragraph("Ключевые финансовые результаты.")
        document.add_paragraph("Q2 2026")
        document.add_paragraph("Переход к следующему этапу.")

        buffer = BytesIO()
        document.save(buffer)

        _, _, blocks = extractor.extract("sample.docx", buffer.getvalue())

        self.assertEqual(blocks[1].kind, "subheading")
        self.assertEqual(blocks[1].level, 2)
        self.assertEqual(blocks[3].kind, "subheading")
        self.assertEqual(blocks[3].level, 2)


if __name__ == "__main__":
    unittest.main()
