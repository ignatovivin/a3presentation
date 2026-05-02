from __future__ import annotations

import unittest

from a3presentation.services.text_fit import (
    FitBox,
    TextBoxMetrics,
    best_fit_font_size_pt,
    estimate_text_height_emu,
    split_oversized_chunks_for_box,
    text_fits_box,
)


class TextFitTests(unittest.TestCase):
    def test_estimated_height_grows_with_font_size(self) -> None:
        metrics = TextBoxMetrics(width_emu=3_000_000, height_emu=1_200_000, font_size_pt=18.0)
        text = ["Текстовая строка для проверки оценки высоты."]

        small = estimate_text_height_emu(text, metrics, font_size_pt=12)
        large = estimate_text_height_emu(text, metrics, font_size_pt=24)

        self.assertGreater(large, small)

    def test_best_fit_font_size_respects_bounds(self) -> None:
        metrics = TextBoxMetrics(width_emu=3_400_000, height_emu=900_000, font_size_pt=18.0)
        text = ["Короткий текст помещается в блок."]

        best = best_fit_font_size_pt(text, metrics, min_font_pt=10, max_font_pt=24)

        self.assertIsNotNone(best)
        assert best is not None
        self.assertGreaterEqual(best, 10)
        self.assertLessEqual(best, 24)
        self.assertTrue(text_fits_box(text, metrics, font_size_pt=best))

    def test_split_oversized_chunks_produces_fitting_parts(self) -> None:
        box = FitBox(
            metrics=TextBoxMetrics(width_emu=1_600_000, height_emu=420_000, font_size_pt=18.0),
            min_font_pt=10,
            max_font_pt=18,
        )
        text = " ".join([f"слово{index}" for index in range(70)])

        parts = split_oversized_chunks_for_box([text], box)

        self.assertGreater(len(parts), 1)
        self.assertTrue(all(text_fits_box([part], box.metrics, font_size_pt=box.min_font_pt) for part in parts))


if __name__ == "__main__":
    unittest.main()
