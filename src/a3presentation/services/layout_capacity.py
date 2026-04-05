from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class LayoutCapacityProfile:
    layout_key: str
    max_items: int
    max_weight: float
    max_chars: int
    max_primary_chars: int
    min_font_pt: int
    max_font_pt: int
    target_fill_ratio: float
    max_fill_ratio: float
    continuation_balance_tolerance: float


TEXT_FULL_WIDTH_PROFILE = LayoutCapacityProfile(
    layout_key="text_full_width",
    max_items=8,
    max_weight=8.5,
    max_chars=520,
    max_primary_chars=320,
    min_font_pt=12,
    max_font_pt=16,
    target_fill_ratio=0.78,
    max_fill_ratio=0.93,
    continuation_balance_tolerance=0.18,
)

LIST_FULL_WIDTH_PROFILE = LayoutCapacityProfile(
    layout_key="list_full_width",
    max_items=8,
    max_weight=11.0,
    max_chars=900,
    max_primary_chars=0,
    min_font_pt=12,
    max_font_pt=16,
    target_fill_ratio=0.8,
    max_fill_ratio=0.94,
    continuation_balance_tolerance=0.18,
)

TABLE_PROFILE = LayoutCapacityProfile(
    layout_key="table",
    max_items=10,
    max_weight=10.0,
    max_chars=0,
    max_primary_chars=0,
    min_font_pt=8,
    max_font_pt=11,
    target_fill_ratio=0.84,
    max_fill_ratio=0.96,
    continuation_balance_tolerance=0.12,
)


def profile_for_layout(layout_key: str) -> LayoutCapacityProfile:
    if layout_key == "table":
        return TABLE_PROFILE
    if layout_key == "list_full_width":
        return LIST_FULL_WIDTH_PROFILE
    return TEXT_FULL_WIDTH_PROFILE
