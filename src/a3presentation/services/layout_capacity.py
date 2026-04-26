from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class PlaceholderGeometryPolicy:
    placeholder_idx: int
    left_emu: int
    top_emu: int
    width_emu: int
    height_emu: int


@dataclass(frozen=True)
class BulletSpacingPolicy:
    margin_left_emu: int
    indent_emu: int


@dataclass(frozen=True)
class ParagraphSpacingPolicy:
    line_spacing: float
    space_after_pt: float


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


@dataclass(frozen=True)
class LayoutGeometryPolicy:
    layout_key: str
    placeholders: dict[int, PlaceholderGeometryPolicy]
    title_content_gap_emu: int
    title_body_gap_no_subtitle_emu: int
    content_footer_gap_emu: int


@dataclass(frozen=True)
class LayoutSpacingPolicy:
    layout_key: str
    bullet: BulletSpacingPolicy
    body: ParagraphSpacingPolicy
    title: ParagraphSpacingPolicy
    subtitle: ParagraphSpacingPolicy
    cover: ParagraphSpacingPolicy


TEXT_FULL_WIDTH_PROFILE = LayoutCapacityProfile(
    layout_key="text_full_width",
    max_items=8,
    max_weight=8.5,
    max_chars=520,
    max_primary_chars=320,
    min_font_pt=12,
    max_font_pt=20,
    target_fill_ratio=0.78,
    max_fill_ratio=0.93,
    continuation_balance_tolerance=0.18,
)

DENSE_TEXT_FULL_WIDTH_PROFILE = LayoutCapacityProfile(
    layout_key="dense_text_full_width",
    max_items=9,
    max_weight=11.5,
    max_chars=900,
    max_primary_chars=520,
    min_font_pt=11,
    max_font_pt=20,
    target_fill_ratio=0.8,
    max_fill_ratio=0.98,
    continuation_balance_tolerance=0.16,
)

LIST_FULL_WIDTH_PROFILE = LayoutCapacityProfile(
    layout_key="list_full_width",
    max_items=8,
    max_weight=11.0,
    max_chars=900,
    max_primary_chars=0,
    min_font_pt=12,
    max_font_pt=20,
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

IMAGE_TEXT_PROFILE = LayoutCapacityProfile(
    layout_key="image_text",
    max_items=8,
    max_weight=8.0,
    max_chars=480,
    max_primary_chars=260,
    min_font_pt=12,
    max_font_pt=20,
    target_fill_ratio=0.76,
    max_fill_ratio=0.92,
    continuation_balance_tolerance=0.18,
)

CARDS_3_PROFILE = LayoutCapacityProfile(
    layout_key="cards_3",
    max_items=3,
    max_weight=6.0,
    max_chars=360,
    max_primary_chars=0,
    min_font_pt=12,
    max_font_pt=24,
    target_fill_ratio=0.74,
    max_fill_ratio=0.9,
    continuation_balance_tolerance=0.16,
)

CARDS_KPI_PROFILE = LayoutCapacityProfile(
    layout_key="cards_kpi",
    max_items=4,
    max_weight=5.5,
    max_chars=220,
    max_primary_chars=0,
    min_font_pt=14,
    max_font_pt=44,
    target_fill_ratio=0.68,
    max_fill_ratio=0.88,
    continuation_balance_tolerance=0.14,
)

LIST_WITH_ICONS_PROFILE = LayoutCapacityProfile(
    layout_key="list_with_icons",
    max_items=8,
    max_weight=10.5,
    max_chars=720,
    max_primary_chars=220,
    min_font_pt=12,
    max_font_pt=20,
    target_fill_ratio=0.78,
    max_fill_ratio=0.93,
    continuation_balance_tolerance=0.18,
)

CONTACTS_PROFILE = LayoutCapacityProfile(
    layout_key="contacts",
    max_items=4,
    max_weight=4.5,
    max_chars=220,
    max_primary_chars=120,
    min_font_pt=12,
    max_font_pt=20,
    target_fill_ratio=0.72,
    max_fill_ratio=0.9,
    continuation_balance_tolerance=0.16,
)


TEXT_LAYOUT_GEOMETRY_POLICY = LayoutGeometryPolicy(
    layout_key="text_full_width",
    placeholders={
        0: PlaceholderGeometryPolicy(placeholder_idx=0, left_emu=442913, top_emu=671247, width_emu=11198224, height_emu=1120247),
        13: PlaceholderGeometryPolicy(placeholder_idx=13, left_emu=442913, top_emu=1228230, width_emu=11198224, height_emu=552402),
        14: PlaceholderGeometryPolicy(placeholder_idx=14, left_emu=442913, top_emu=1791494, width_emu=11198224, height_emu=3550000),
        15: PlaceholderGeometryPolicy(placeholder_idx=15, left_emu=442913, top_emu=3800000, width_emu=11198224, height_emu=1850000),
        17: PlaceholderGeometryPolicy(placeholder_idx=17, left_emu=442913, top_emu=6384626, width_emu=11198224, height_emu=260000),
    },
    title_content_gap_emu=180000,
    title_body_gap_no_subtitle_emu=300000,
    content_footer_gap_emu=180000,
)

LIST_LAYOUT_GEOMETRY_POLICY = LayoutGeometryPolicy(
    layout_key="list_full_width",
    placeholders=TEXT_LAYOUT_GEOMETRY_POLICY.placeholders,
    title_content_gap_emu=TEXT_LAYOUT_GEOMETRY_POLICY.title_content_gap_emu,
    title_body_gap_no_subtitle_emu=TEXT_LAYOUT_GEOMETRY_POLICY.title_body_gap_no_subtitle_emu,
    content_footer_gap_emu=TEXT_LAYOUT_GEOMETRY_POLICY.content_footer_gap_emu,
)

TABLE_LAYOUT_GEOMETRY_POLICY = LayoutGeometryPolicy(
    layout_key="table",
    placeholders={
        0: PlaceholderGeometryPolicy(placeholder_idx=0, left_emu=442913, top_emu=671247, width_emu=11198224, height_emu=584960),
        13: PlaceholderGeometryPolicy(placeholder_idx=13, left_emu=442913, top_emu=1228230, width_emu=11198224, height_emu=700000),
        15: PlaceholderGeometryPolicy(placeholder_idx=15, left_emu=442913, top_emu=6384626, width_emu=11198224, height_emu=260000),
    },
    title_content_gap_emu=180000,
    title_body_gap_no_subtitle_emu=300000,
    content_footer_gap_emu=180000,
)

IMAGE_TEXT_LAYOUT_GEOMETRY_POLICY = LayoutGeometryPolicy(
    layout_key="image_text",
    placeholders={
        0: PlaceholderGeometryPolicy(placeholder_idx=0, left_emu=442913, top_emu=671247, width_emu=11198224, height_emu=1120247),
        13: PlaceholderGeometryPolicy(placeholder_idx=13, left_emu=442913, top_emu=1228230, width_emu=5653087, height_emu=552402),
        14: PlaceholderGeometryPolicy(placeholder_idx=14, left_emu=442913, top_emu=1791494, width_emu=4370387, height_emu=1932608),
        15: PlaceholderGeometryPolicy(placeholder_idx=15, left_emu=442913, top_emu=3831706, width_emu=4370387, height_emu=2464622),
        16: PlaceholderGeometryPolicy(placeholder_idx=16, left_emu=6650297, top_emu=1228230, width_emu=4990840, height_emu=4532490),
        17: PlaceholderGeometryPolicy(placeholder_idx=17, left_emu=442913, top_emu=6384626, width_emu=3371850, height_emu=277813),
    },
    title_content_gap_emu=180000,
    title_body_gap_no_subtitle_emu=300000,
    content_footer_gap_emu=180000,
)

CARDS_3_LAYOUT_GEOMETRY_POLICY = LayoutGeometryPolicy(
    layout_key="cards_3",
    placeholders={
        0: PlaceholderGeometryPolicy(placeholder_idx=0, left_emu=442913, top_emu=671247, width_emu=11198224, height_emu=1325563),
        11: PlaceholderGeometryPolicy(placeholder_idx=11, left_emu=739775, top_emu=1723633, width_emu=3259138, height_emu=4164013),
        12: PlaceholderGeometryPolicy(placeholder_idx=12, left_emu=4412456, top_emu=1723633, width_emu=3259138, height_emu=4164013),
        13: PlaceholderGeometryPolicy(placeholder_idx=13, left_emu=8193087, top_emu=1723633, width_emu=3259138, height_emu=4164013),
        15: PlaceholderGeometryPolicy(placeholder_idx=15, left_emu=442913, top_emu=6384626, width_emu=3371850, height_emu=277813),
    },
    title_content_gap_emu=220000,
    title_body_gap_no_subtitle_emu=320000,
    content_footer_gap_emu=180000,
)

CARDS_KPI_LAYOUT_GEOMETRY_POLICY = LayoutGeometryPolicy(
    layout_key="cards_kpi",
    placeholders={
        0: PlaceholderGeometryPolicy(placeholder_idx=0, left_emu=828675, top_emu=610000, width_emu=10300000, height_emu=1400000),
        11: PlaceholderGeometryPolicy(placeholder_idx=11, left_emu=828675, top_emu=3170000, width_emu=3600000, height_emu=1350000),
        12: PlaceholderGeometryPolicy(placeholder_idx=12, left_emu=6980000, top_emu=3170000, width_emu=3600000, height_emu=1350000),
        13: PlaceholderGeometryPolicy(placeholder_idx=13, left_emu=828675, top_emu=4950000, width_emu=3600000, height_emu=1350000),
    },
    title_content_gap_emu=220000,
    title_body_gap_no_subtitle_emu=420000,
    content_footer_gap_emu=0,
)

LIST_WITH_ICONS_LAYOUT_GEOMETRY_POLICY = LayoutGeometryPolicy(
    layout_key="list_with_icons",
    placeholders={
        0: PlaceholderGeometryPolicy(placeholder_idx=0, left_emu=442913, top_emu=671247, width_emu=11198224, height_emu=1109385),
        13: PlaceholderGeometryPolicy(placeholder_idx=13, left_emu=442913, top_emu=1228230, width_emu=5653087, height_emu=552402),
        12: PlaceholderGeometryPolicy(placeholder_idx=12, left_emu=550352, top_emu=1720850, width_emu=3221037, height_emu=2393950),
        14: PlaceholderGeometryPolicy(placeholder_idx=14, left_emu=5219700, top_emu=1690688, width_emu=6421438, height_emu=4291012),
        15: PlaceholderGeometryPolicy(placeholder_idx=15, left_emu=4546770, top_emu=1690688, width_emu=507658, height_emu=507658),
        16: PlaceholderGeometryPolicy(placeholder_idx=16, left_emu=4546770, top_emu=2337615, width_emu=507658, height_emu=507658),
        17: PlaceholderGeometryPolicy(placeholder_idx=17, left_emu=4546770, top_emu=2984542, width_emu=507658, height_emu=507658),
        18: PlaceholderGeometryPolicy(placeholder_idx=18, left_emu=4546770, top_emu=3627337, width_emu=507658, height_emu=507658),
        19: PlaceholderGeometryPolicy(placeholder_idx=19, left_emu=4546770, top_emu=4270132, width_emu=507658, height_emu=507658),
        20: PlaceholderGeometryPolicy(placeholder_idx=20, left_emu=4546770, top_emu=4913483, width_emu=507658, height_emu=507658),
        21: PlaceholderGeometryPolicy(placeholder_idx=21, left_emu=442913, top_emu=6384626, width_emu=3371850, height_emu=277813),
    },
    title_content_gap_emu=180000,
    title_body_gap_no_subtitle_emu=300000,
    content_footer_gap_emu=180000,
)

CONTACTS_LAYOUT_GEOMETRY_POLICY = LayoutGeometryPolicy(
    layout_key="contacts",
    placeholders={
        10: PlaceholderGeometryPolicy(placeholder_idx=10, left_emu=7486650, top_emu=2305374, width_emu=3724275, height_emu=1037901),
        11: PlaceholderGeometryPolicy(placeholder_idx=11, left_emu=7486650, top_emu=3429000, width_emu=3724275, height_emu=361950),
        12: PlaceholderGeometryPolicy(placeholder_idx=12, left_emu=7486650, top_emu=3826037, width_emu=3724275, height_emu=361950),
        13: PlaceholderGeometryPolicy(placeholder_idx=13, left_emu=7486649, top_emu=4290481, width_emu=3724275, height_emu=361950),
    },
    title_content_gap_emu=180000,
    title_body_gap_no_subtitle_emu=220000,
    content_footer_gap_emu=180000,
)


TEXT_LAYOUT_SPACING_POLICY = LayoutSpacingPolicy(
    layout_key="text_full_width",
    bullet=BulletSpacingPolicy(margin_left_emu=342900, indent_emu=-171450),
    body=ParagraphSpacingPolicy(line_spacing=1.1, space_after_pt=6.0),
    title=ParagraphSpacingPolicy(line_spacing=1.0, space_after_pt=0.0),
    subtitle=ParagraphSpacingPolicy(line_spacing=1.0, space_after_pt=4.0),
    cover=ParagraphSpacingPolicy(line_spacing=1.0, space_after_pt=6.0),
)

LIST_LAYOUT_SPACING_POLICY = LayoutSpacingPolicy(
    layout_key="list_full_width",
    bullet=BulletSpacingPolicy(margin_left_emu=400000, indent_emu=-200000),
    body=ParagraphSpacingPolicy(line_spacing=1.05, space_after_pt=5.0),
    title=ParagraphSpacingPolicy(line_spacing=1.0, space_after_pt=0.0),
    subtitle=ParagraphSpacingPolicy(line_spacing=1.0, space_after_pt=4.0),
    cover=ParagraphSpacingPolicy(line_spacing=1.0, space_after_pt=6.0),
)

TABLE_LAYOUT_SPACING_POLICY = LayoutSpacingPolicy(
    layout_key="table",
    bullet=TEXT_LAYOUT_SPACING_POLICY.bullet,
    body=ParagraphSpacingPolicy(line_spacing=1.0, space_after_pt=4.0),
    title=ParagraphSpacingPolicy(line_spacing=1.0, space_after_pt=0.0),
    subtitle=ParagraphSpacingPolicy(line_spacing=1.0, space_after_pt=4.0),
    cover=ParagraphSpacingPolicy(line_spacing=1.0, space_after_pt=6.0),
)

IMAGE_TEXT_LAYOUT_SPACING_POLICY = LayoutSpacingPolicy(
    layout_key="image_text",
    bullet=TEXT_LAYOUT_SPACING_POLICY.bullet,
    body=ParagraphSpacingPolicy(line_spacing=1.08, space_after_pt=5.0),
    title=ParagraphSpacingPolicy(line_spacing=1.0, space_after_pt=0.0),
    subtitle=ParagraphSpacingPolicy(line_spacing=1.0, space_after_pt=4.0),
    cover=ParagraphSpacingPolicy(line_spacing=1.0, space_after_pt=6.0),
)

CARDS_3_LAYOUT_SPACING_POLICY = LayoutSpacingPolicy(
    layout_key="cards_3",
    bullet=TEXT_LAYOUT_SPACING_POLICY.bullet,
    body=ParagraphSpacingPolicy(line_spacing=1.0, space_after_pt=4.0),
    title=ParagraphSpacingPolicy(line_spacing=1.0, space_after_pt=0.0),
    subtitle=ParagraphSpacingPolicy(line_spacing=1.0, space_after_pt=4.0),
    cover=ParagraphSpacingPolicy(line_spacing=1.0, space_after_pt=6.0),
)

LIST_WITH_ICONS_LAYOUT_SPACING_POLICY = LayoutSpacingPolicy(
    layout_key="list_with_icons",
    bullet=LIST_LAYOUT_SPACING_POLICY.bullet,
    body=ParagraphSpacingPolicy(line_spacing=1.05, space_after_pt=5.0),
    title=ParagraphSpacingPolicy(line_spacing=1.0, space_after_pt=0.0),
    subtitle=ParagraphSpacingPolicy(line_spacing=1.0, space_after_pt=4.0),
    cover=ParagraphSpacingPolicy(line_spacing=1.0, space_after_pt=6.0),
)

CONTACTS_LAYOUT_SPACING_POLICY = LayoutSpacingPolicy(
    layout_key="contacts",
    bullet=TEXT_LAYOUT_SPACING_POLICY.bullet,
    body=ParagraphSpacingPolicy(line_spacing=1.0, space_after_pt=3.0),
    title=ParagraphSpacingPolicy(line_spacing=1.0, space_after_pt=0.0),
    subtitle=ParagraphSpacingPolicy(line_spacing=1.0, space_after_pt=2.0),
    cover=ParagraphSpacingPolicy(line_spacing=1.0, space_after_pt=6.0),
)

BUILTIN_RUNTIME_PROFILE_KEYS = {
    "text_full_width",
    "dense_text_full_width",
    "list_full_width",
    "table",
    "image_text",
    "cards_3",
    "cards_kpi",
    "list_with_icons",
    "contacts",
}


def runtime_profile_key_for_target(
    target,
    *,
    fallback_layout_key: str | None = None,
    slide_kind: str | None = None,
) -> str:
    target_key = getattr(target, "key", None) or fallback_layout_key or ""
    if target_key in BUILTIN_RUNTIME_PROFILE_KEYS:
        return target_key

    slots = getattr(target, "placeholders", None) or getattr(target, "tokens", None) or ()
    supported_slide_kinds = {
        str(kind)
        for kind in (getattr(target, "supported_slide_kinds", None) or ())
        if kind
    }
    representation_hints = {
        str(hint)
        for hint in (getattr(target, "representation_hints", None) or ())
        if hint
    }
    editable_roles = {
        str(role)
        for role in (
            getattr(slot, "editable_role", None)
            for slot in slots
        )
        if role
    }
    editable_slot_count = sum(
        1
        for slot in slots
        if getattr(slot, "editable_role", None) or getattr(slot, "editable_capabilities", None)
    )

    normalized_slide_kind = slide_kind or ""
    if hasattr(normalized_slide_kind, "value"):
        normalized_slide_kind = normalized_slide_kind.value
    normalized_slide_kind = str(normalized_slide_kind)

    if "contacts" in representation_hints:
        return "contacts"
    if "cards" in representation_hints:
        if target_key == "cards_kpi" or editable_slot_count >= 4:
            return "cards_kpi"
        return "cards_3"
    if "table" in representation_hints or normalized_slide_kind == "table":
        return "table"
    if "image" in representation_hints or normalized_slide_kind == "image":
        return "image_text"
    if "two_column" in representation_hints:
        return "list_with_icons"
    if (
        "bullet_list" in editable_roles
        or "bullet_item" in editable_roles
        or "bullets" in supported_slide_kinds
        or normalized_slide_kind == "bullets"
    ):
        return "list_full_width"
    return "text_full_width"


def profile_for_layout(layout_key: str) -> LayoutCapacityProfile:
    if layout_key == "dense_text_full_width":
        return DENSE_TEXT_FULL_WIDTH_PROFILE
    if layout_key == "table":
        return TABLE_PROFILE
    if layout_key == "image_text":
        return IMAGE_TEXT_PROFILE
    if layout_key == "cards_3":
        return CARDS_3_PROFILE
    if layout_key == "cards_kpi":
        return CARDS_KPI_PROFILE
    if layout_key == "list_with_icons":
        return LIST_WITH_ICONS_PROFILE
    if layout_key == "contacts":
        return CONTACTS_PROFILE
    if layout_key == "list_full_width":
        return LIST_FULL_WIDTH_PROFILE
    return TEXT_FULL_WIDTH_PROFILE


def derive_capacity_profile_for_geometry(
    layout_key: str,
    *,
    width_emu: int | None = None,
    height_emu: int | None = None,
    reference_width_emu: int | None = None,
    reference_height_emu: int | None = None,
) -> LayoutCapacityProfile:
    base_profile = profile_for_layout(layout_key)
    if width_emu is None or height_emu is None or base_profile.max_chars <= 0:
        return base_profile

    if reference_width_emu is None or reference_height_emu is None:
        reference_geometry = geometry_policy_for_layout(layout_key)
        reference_body = reference_geometry.placeholders.get(14)
        if reference_body is None or reference_body.width_emu <= 0 or reference_body.height_emu <= 0:
            return base_profile
        reference_width_emu = reference_body.width_emu
        reference_height_emu = reference_body.height_emu

    if reference_width_emu <= 0 or reference_height_emu <= 0:
        return base_profile

    width_ratio = max(0.55, min(width_emu / reference_width_emu, 1.45))
    height_ratio = max(0.55, min(height_emu / reference_height_emu, 1.45))
    area_ratio = max(0.45, min(width_ratio * height_ratio, 1.6))
    item_ratio = max(0.75, min(height_ratio, 1.2))
    min_ratio = min(width_ratio, height_ratio)

    max_font_pt = base_profile.max_font_pt
    if min_ratio < 0.9:
        max_font_pt -= 1
    if min_ratio < 0.75:
        max_font_pt -= 1
    if min_ratio > 1.15:
        max_font_pt += 1
    max_font_pt = max(base_profile.min_font_pt, max_font_pt)

    return LayoutCapacityProfile(
        layout_key=base_profile.layout_key,
        max_items=max(1, int(round(base_profile.max_items * item_ratio))),
        max_weight=round(base_profile.max_weight * max(0.75, min(area_ratio, 1.25)), 2),
        max_chars=max(80, int(round(base_profile.max_chars * area_ratio))),
        max_primary_chars=max(0, int(round(base_profile.max_primary_chars * area_ratio))),
        min_font_pt=base_profile.min_font_pt,
        max_font_pt=max_font_pt,
        target_fill_ratio=base_profile.target_fill_ratio,
        max_fill_ratio=base_profile.max_fill_ratio,
        continuation_balance_tolerance=base_profile.continuation_balance_tolerance,
    )


def geometry_policy_for_layout(layout_key: str) -> LayoutGeometryPolicy:
    if layout_key == "dense_text_full_width":
        return TEXT_LAYOUT_GEOMETRY_POLICY
    if layout_key == "table":
        return TABLE_LAYOUT_GEOMETRY_POLICY
    if layout_key == "image_text":
        return IMAGE_TEXT_LAYOUT_GEOMETRY_POLICY
    if layout_key == "cards_3":
        return CARDS_3_LAYOUT_GEOMETRY_POLICY
    if layout_key == "cards_kpi":
        return CARDS_KPI_LAYOUT_GEOMETRY_POLICY
    if layout_key == "list_with_icons":
        return LIST_WITH_ICONS_LAYOUT_GEOMETRY_POLICY
    if layout_key == "contacts":
        return CONTACTS_LAYOUT_GEOMETRY_POLICY
    if layout_key == "list_full_width":
        return LIST_LAYOUT_GEOMETRY_POLICY
    return TEXT_LAYOUT_GEOMETRY_POLICY


def spacing_policy_for_layout(layout_key: str) -> LayoutSpacingPolicy:
    if layout_key == "dense_text_full_width":
        return TEXT_LAYOUT_SPACING_POLICY
    if layout_key == "table":
        return TABLE_LAYOUT_SPACING_POLICY
    if layout_key == "image_text":
        return IMAGE_TEXT_LAYOUT_SPACING_POLICY
    if layout_key == "cards_3":
        return CARDS_3_LAYOUT_SPACING_POLICY
    if layout_key == "cards_kpi":
        return CARDS_3_LAYOUT_SPACING_POLICY
    if layout_key == "list_with_icons":
        return LIST_WITH_ICONS_LAYOUT_SPACING_POLICY
    if layout_key == "contacts":
        return CONTACTS_LAYOUT_SPACING_POLICY
    if layout_key == "list_full_width":
        return LIST_LAYOUT_SPACING_POLICY
    return TEXT_LAYOUT_SPACING_POLICY
