from __future__ import annotations

import unittest

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


if __name__ == "__main__":
    unittest.main()
