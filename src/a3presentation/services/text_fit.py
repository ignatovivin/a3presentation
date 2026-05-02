from __future__ import annotations

import math
from dataclasses import dataclass

from a3presentation.services.layout_capacity import spacing_policy_for_layout


EMU_PER_PT = 12700
DEFAULT_TEXT_MARGIN_X_EMU = 91440
DEFAULT_TEXT_MARGIN_Y_EMU = 45720
DEFAULT_BODY_FONT_PT = 18.0
DEFAULT_LINE_HEIGHT_FACTOR = 1.18


@dataclass(frozen=True)
class TextBoxMetrics:
    width_emu: int
    height_emu: int
    margin_left_emu: int = DEFAULT_TEXT_MARGIN_X_EMU
    margin_right_emu: int = DEFAULT_TEXT_MARGIN_X_EMU
    margin_top_emu: int = DEFAULT_TEXT_MARGIN_Y_EMU
    margin_bottom_emu: int = DEFAULT_TEXT_MARGIN_Y_EMU
    font_size_pt: float = DEFAULT_BODY_FONT_PT
    line_spacing: float = DEFAULT_LINE_HEIGHT_FACTOR
    space_after_pt: float = 0.0
    role: str | None = None

    @property
    def available_width_emu(self) -> int:
        return max(self.width_emu - self.margin_left_emu - self.margin_right_emu, int(self.width_emu * 0.45))

    @property
    def available_height_emu(self) -> int:
        return max(self.height_emu - self.margin_top_emu - self.margin_bottom_emu, int(self.height_emu * 0.45))


@dataclass(frozen=True)
class FitBox:
    metrics: TextBoxMetrics
    min_font_pt: int
    max_font_pt: int


@dataclass(frozen=True)
class TextPackResult:
    batches: list[list[str]]
    fits: bool
    failed_box_index: int | None = None
    failed_chunk_index: int | None = None
    reason: str | None = None


def metrics_from_slot(slot, *, layout_key: str, fallback_font_size_pt: float) -> TextBoxMetrics | None:
    width = getattr(slot, "width_emu", None)
    height = getattr(slot, "height_emu", None)
    if not isinstance(width, int) or not isinstance(height, int) or width <= 0 or height <= 0:
        return None
    return TextBoxMetrics(
        width_emu=width,
        height_emu=height,
        margin_left_emu=_slot_margin_emu(slot, "left"),
        margin_right_emu=_slot_margin_emu(slot, "right"),
        margin_top_emu=_slot_margin_emu(slot, "top"),
        margin_bottom_emu=_slot_margin_emu(slot, "bottom"),
        font_size_pt=_slot_font_size_pt(slot, fallback_font_size_pt),
        line_spacing=_slot_line_height_factor(slot, layout_key),
        space_after_pt=_slot_space_after_pt(slot, layout_key),
        role=getattr(slot, "editable_role", None),
    )


def metrics_from_shape(shape, *, layout_key: str, fallback_font_size_pt: float) -> TextBoxMetrics | None:
    width = getattr(shape, "width", None)
    height = getattr(shape, "height", None)
    if not isinstance(width, int) or not isinstance(height, int) or width <= 0 or height <= 0:
        return None
    text_frame = getattr(shape, "text_frame", None) if getattr(shape, "has_text_frame", False) else None
    spacing = spacing_policy_for_layout(layout_key).body
    return TextBoxMetrics(
        width_emu=width,
        height_emu=height,
        margin_left_emu=getattr(text_frame, "margin_left", DEFAULT_TEXT_MARGIN_X_EMU) or DEFAULT_TEXT_MARGIN_X_EMU,
        margin_right_emu=getattr(text_frame, "margin_right", DEFAULT_TEXT_MARGIN_X_EMU) or DEFAULT_TEXT_MARGIN_X_EMU,
        margin_top_emu=getattr(text_frame, "margin_top", DEFAULT_TEXT_MARGIN_Y_EMU) or DEFAULT_TEXT_MARGIN_Y_EMU,
        margin_bottom_emu=getattr(text_frame, "margin_bottom", DEFAULT_TEXT_MARGIN_Y_EMU) or DEFAULT_TEXT_MARGIN_Y_EMU,
        font_size_pt=fallback_font_size_pt,
        line_spacing=max(spacing.line_spacing, DEFAULT_LINE_HEIGHT_FACTOR),
        space_after_pt=float(spacing.space_after_pt),
    )


def estimate_text_height_emu(texts: list[str], metrics: TextBoxMetrics, *, font_size_pt: float | None = None) -> int:
    active_font_size = float(font_size_pt or metrics.font_size_pt or DEFAULT_BODY_FONT_PT)
    paragraphs = [text.strip() for text in texts if text and text.strip()]
    if not paragraphs:
        return 0

    width_pt = metrics.available_width_emu / EMU_PER_PT
    average_char_width_pt = max(active_font_size * 0.52, 1.0)
    line_height_pt = max(active_font_size * max(metrics.line_spacing, 0.85), active_font_size)
    total_lines = sum(
        estimate_wrapped_line_count(text, width_pt, average_char_width_pt, min_chars_per_line=8)
        for text in paragraphs
    )
    paragraph_gap_pt = max(metrics.space_after_pt, 0.0) * max(len(paragraphs) - 1, 0)
    vertical_padding_pt = active_font_size * 0.7
    return int((total_lines * line_height_pt + paragraph_gap_pt + vertical_padding_pt) * EMU_PER_PT)


def text_fits_box(texts: list[str], metrics: TextBoxMetrics, *, font_size_pt: float | None = None, tolerance: float = 1.04) -> bool:
    return estimate_text_height_emu(texts, metrics, font_size_pt=font_size_pt) <= int(metrics.available_height_emu * tolerance)


def best_fit_font_size_pt(texts: list[str], metrics: TextBoxMetrics, *, min_font_pt: int, max_font_pt: int) -> int | None:
    low = max(1, int(min_font_pt))
    high = max(low, int(max_font_pt))
    if not text_fits_box(texts, metrics, font_size_pt=low):
        return None

    best = low
    while low <= high:
        candidate = (low + high) // 2
        if text_fits_box(texts, metrics, font_size_pt=candidate):
            best = candidate
            low = candidate + 1
        else:
            high = candidate - 1
    return best


def text_fits_within_font_bounds(texts: list[str], metrics: TextBoxMetrics, *, min_font_pt: int, max_font_pt: int) -> bool:
    return best_fit_font_size_pt(texts, metrics, min_font_pt=min_font_pt, max_font_pt=max_font_pt) is not None


def texts_fit_box(texts: list[str], box: FitBox) -> bool:
    return text_fits_within_font_bounds(
        texts,
        box.metrics,
        min_font_pt=box.min_font_pt,
        max_font_pt=box.max_font_pt,
    )


def pack_text_chunks_for_boxes(chunks: list[str], boxes: list[FitBox]) -> list[list[str]] | None:
    result = assign_text_chunks_to_boxes(chunks, boxes)
    return result.batches if result.fits else None


def assign_text_chunks_to_boxes(chunks: list[str], boxes: list[FitBox]) -> TextPackResult:
    if not boxes:
        return TextPackResult([], fits=False, reason="no_boxes")
    batches: list[list[str]] = [[] for _ in boxes]
    box_index = 0
    for chunk_index, chunk in enumerate(chunks):
        candidate = [*batches[box_index], chunk]
        if batches[box_index] and not texts_fit_box(candidate, boxes[box_index]):
            box_index += 1
            if box_index >= len(boxes):
                return TextPackResult(
                    batches,
                    fits=False,
                    failed_box_index=len(boxes) - 1,
                    failed_chunk_index=chunk_index,
                    reason="no_remaining_box",
                )
            candidate = [chunk]
        if not texts_fit_box(candidate, boxes[box_index]):
            return TextPackResult(
                batches,
                fits=False,
                failed_box_index=box_index,
                failed_chunk_index=chunk_index,
                reason="chunk_too_large",
            )
        batches[box_index] = candidate
    return TextPackResult(batches, fits=True)


def split_oversized_chunks_for_box(chunks: list[str], box: FitBox) -> list[str]:
    normalized: list[str] = []
    for chunk in chunks:
        if texts_fit_box([chunk], box):
            normalized.append(chunk)
            continue
        words = chunk.split()
        if not words:
            continue
        normalized.extend(_split_words_to_fit_box(words, box))
    return normalized


def estimate_capacity_chars(metrics: TextBoxMetrics, *, density: float | None = None) -> int:
    font_size_pt = max(float(metrics.font_size_pt or DEFAULT_BODY_FONT_PT), 1.0)
    width_pt = metrics.available_width_emu / EMU_PER_PT
    height_pt = metrics.available_height_emu / EMU_PER_PT
    chars_per_line = max(int(width_pt / max(font_size_pt * 0.52, 1.0)), 8)
    line_height_pt = max(font_size_pt * max(metrics.line_spacing, 0.85) + max(metrics.space_after_pt, 0.0) * 0.35, font_size_pt)
    vertical_padding_pt = font_size_pt * 0.7
    available_lines = max(int(max(height_pt - vertical_padding_pt, 0) / line_height_pt), 0)
    if available_lines <= 0:
        return 0
    active_density = density if density is not None else (0.74 if metrics.role in {"bullet_list", "bullet_item"} else 0.82)
    return int(chars_per_line * available_lines * active_density)


def estimate_wrapped_line_count(
    text: str,
    width_pt: float,
    average_char_width_pt: float,
    *,
    min_chars_per_line: int,
) -> int:
    chars_per_line = max(int(width_pt / average_char_width_pt), min_chars_per_line)
    wrapped_lines = 0

    for paragraph in text.splitlines() or [text]:
        normalized = paragraph.strip()
        if not normalized:
            wrapped_lines += 1
            continue

        words = normalized.split()
        if len(words) <= 1:
            wrapped_lines += max(1, math.ceil(len(normalized) / chars_per_line))
            continue

        current_line_len = 0
        paragraph_lines = 1
        for word in words:
            word_len = len(word)
            projected = word_len if current_line_len == 0 else current_line_len + 1 + word_len
            if projected <= chars_per_line:
                current_line_len = projected
                continue
            if current_line_len == 0:
                paragraph_lines += max(math.ceil(word_len / chars_per_line) - 1, 0)
                current_line_len = word_len % chars_per_line or chars_per_line
                continue
            paragraph_lines += 1
            current_line_len = word_len

        wrapped_lines += paragraph_lines

    return wrapped_lines


def _split_words_to_fit_box(words: list[str], box: FitBox) -> list[str]:
    parts: list[str] = []
    start = 0
    while start < len(words):
        if not texts_fit_box([words[start]], box):
            parts.extend(_split_unbreakable_text_to_fit_box(words[start], box))
            start += 1
            continue

        low = start
        high = len(words) - 1
        best = start
        while low <= high:
            middle = (low + high) // 2
            candidate = " ".join(words[start : middle + 1])
            if texts_fit_box([candidate], box):
                best = middle
                low = middle + 1
            else:
                high = middle - 1
        parts.append(" ".join(words[start : best + 1]))
        start = best + 1
    return parts


def _split_unbreakable_text_to_fit_box(text: str, box: FitBox) -> list[str]:
    parts: list[str] = []
    start = 0
    while start < len(text):
        low = start + 1
        high = len(text)
        best = start + 1
        while low <= high:
            middle = (low + high) // 2
            candidate = text[start:middle]
            if texts_fit_box([candidate], box):
                best = middle
                low = middle + 1
            else:
                high = middle - 1
        parts.append(text[start:best])
        start = best
    return parts


def _slot_margin_emu(slot, side: str) -> int:
    value = getattr(slot, f"margin_{side}_emu", None)
    if isinstance(value, int) and value >= 0:
        return value
    shape_style = getattr(slot, "shape_style", None)
    inset = getattr(shape_style, f"inset_{side}_emu", None) if shape_style is not None else None
    if isinstance(inset, int) and inset >= 0:
        return inset
    return DEFAULT_TEXT_MARGIN_X_EMU if side in {"left", "right"} else DEFAULT_TEXT_MARGIN_Y_EMU


def _slot_font_size_pt(slot, fallback_font_size_pt: float) -> float:
    text_style = getattr(slot, "text_style", None)
    if text_style is not None and text_style.font_size_pt:
        return float(text_style.font_size_pt)
    paragraph_styles = getattr(slot, "paragraph_styles", None)
    level_styles = getattr(paragraph_styles, "level_styles", None) if paragraph_styles is not None else None
    if level_styles:
        style = level_styles.get("0") or next(iter(level_styles.values()), None)
        if style is not None and style.font_size_pt:
            return float(style.font_size_pt)
    return float(fallback_font_size_pt or DEFAULT_BODY_FONT_PT)


def _slot_line_height_factor(slot, layout_key: str) -> float:
    text_style = getattr(slot, "text_style", None)
    if text_style is not None and text_style.line_spacing:
        return max(float(text_style.line_spacing), 0.85)
    return max(spacing_policy_for_layout(layout_key).body.line_spacing, DEFAULT_LINE_HEIGHT_FACTOR)


def _slot_space_after_pt(slot, layout_key: str) -> float:
    text_style = getattr(slot, "text_style", None)
    if text_style is not None and text_style.space_after_pt is not None:
        return max(float(text_style.space_after_pt), 0.0)
    return float(spacing_policy_for_layout(layout_key).body.space_after_pt)
