from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path, PurePosixPath

from a3presentation.domain.api import (
    DetectedComponentSummary,
    EditableTargetSummary,
    InventoryTargetSummary,
    SlideLayoutOption,
    SlideLayoutReview,
    TemplateInventorySummary,
)
from a3presentation.domain.presentation import (
    PresentationPlan,
    RenderTargetType,
    SlideContentBlock,
    SlideContentBlockKind,
    SlideKind,
    SlideRenderTarget,
    SlideSpec,
)
from a3presentation.domain.template import (
    ComponentEditability,
    PlaceholderKind,
    TemplateComponentStyleSpec,
    TemplateManifest,
    TemplateShapeStyleSpec,
)
from a3presentation.services.layout_capacity import derive_capacity_profile_for_geometry, runtime_profile_key_for_target


@dataclass(frozen=True)
class _InventoryTarget:
    key: str
    name: str
    source: str
    source_label: str
    supported_slide_kinds: list[str]
    representation_hints: list[str]
    editable_roles: list[str]
    editable_slot_count: int
    slots: list


class TemplateRegistry:
    def __init__(self, templates_dir: Path) -> None:
        self._templates_dir = templates_dir.resolve()

    def list_templates(self) -> list[TemplateManifest]:
        manifests: list[TemplateManifest] = []
        if not self._templates_dir.exists():
            return manifests

        for template_dir in sorted(path for path in self._templates_dir.iterdir() if path.is_dir()):
            manifest_path = template_dir / "manifest.json"
            if not manifest_path.exists():
                continue
            manifests.append(self._load_manifest(manifest_path))
        return manifests

    def get_template(self, template_id: str) -> TemplateManifest:
        manifest_path = self._template_dir(template_id) / "manifest.json"
        if not manifest_path.exists():
            raise FileNotFoundError(f"Template '{template_id}' not found")
        return self._load_manifest(manifest_path)

    def get_template_pptx_path(self, template_id: str) -> Path:
        manifest = self.get_template(template_id)
        template_dir = self._template_dir(template_id)
        pptx_path = template_dir / manifest.source_pptx
        if not pptx_path.exists():
            raise FileNotFoundError(f"Template PPTX not found for '{template_id}'")
        return pptx_path

    def save_manifest(self, manifest: TemplateManifest) -> Path:
        template_dir = self._template_dir(manifest.template_id)
        template_dir.mkdir(parents=True, exist_ok=True)
        manifest_path = template_dir / "manifest.json"
        manifest_path.write_text(
            manifest.model_dump_json(indent=2),
            encoding="utf-8",
        )
        return manifest_path

    def save_template_file(self, template_id: str, filename: str, content: bytes) -> Path:
        template_dir = self._template_dir(template_id)
        template_dir.mkdir(parents=True, exist_ok=True)
        target_path = self._safe_child_path(template_dir, filename)
        target_path.write_bytes(content)
        return target_path

    def normalize_manifest(self, manifest: TemplateManifest) -> TemplateManifest:
        return self._normalize_manifest(manifest)

    def build_inventory_summary(self, manifest: TemplateManifest) -> TemplateInventorySummary:
        targets = self._inventory_targets(manifest)
        return TemplateInventorySummary(
            generation_mode=manifest.generation_mode.value,
            usability_status=self.usability_status(manifest),
            has_usable_layout_inventory=manifest.inventory.has_usable_layout_inventory,
            has_prototype_inventory=manifest.inventory.has_prototype_inventory,
            degradation_mode=manifest.inventory.degradation_mode,
            warnings=list(manifest.inventory.warnings),
            layout_target_count=sum(1 for target in targets if target.source == "layout"),
            prototype_target_count=sum(1 for target in targets if target.source == "prototype"),
            targets=[
                InventoryTargetSummary(
                    key=target.key,
                    name=target.name,
                    source=target.source,
                    source_label=target.source_label,
                    supported_slide_kinds=list(target.supported_slide_kinds),
                    representation_hints=list(target.representation_hints),
                    editable_slot_count=target.editable_slot_count,
                    editable_roles=list(target.editable_roles),
                )
                for target in targets
            ],
        )

    def build_editable_targets(self, manifest: TemplateManifest) -> list[EditableTargetSummary]:
        editable_targets: list[EditableTargetSummary] = []
        for target in self._inventory_targets(manifest):
            if target.editable_slot_count <= 0 and not target.editable_roles:
                continue
            runtime_profile_key = runtime_profile_key_for_target(
                self._target_by_key(manifest, target.key),
                fallback_layout_key=target.key,
            )
            editable_targets.append(
                EditableTargetSummary(
                    key=target.key,
                    name=target.name,
                    source=target.source,
                    source_label=target.source_label,
                    runtime_profile_key=runtime_profile_key,
                    supported_slide_kinds=list(target.supported_slide_kinds),
                    representation_hints=list(target.representation_hints),
                    editable_slot_count=target.editable_slot_count,
                    editable_roles=list(target.editable_roles),
                )
            )
        return editable_targets

    def build_detected_components(self, manifest: TemplateManifest) -> list[DetectedComponentSummary]:
        return [
            DetectedComponentSummary(
                component_id=component.component_id,
                source_kind=component.source_kind.value,
                source_index=component.source_index,
                source_name=component.source_name,
                shape_name=component.shape_name,
                component_type=component.component_type.value,
                role=component.role.value,
                binding=component.binding,
                confidence=component.confidence.value,
                editability=component.editability.value,
                capabilities=list(component.capabilities),
                geometry=component.geometry.model_copy(deep=True),
                text_excerpt=component.text_excerpt,
                child_component_ids=list(component.child_component_ids),
            )
            for component in manifest.inventory.components
        ]

    def usability_status(self, manifest: TemplateManifest) -> str:
        editable_targets = self.build_editable_targets(manifest)
        if not editable_targets:
            return "not_safely_editable"
        if manifest.generation_mode.value == "prototype" and manifest.inventory.has_prototype_inventory:
            return "usable_with_degradation"
        if manifest.inventory.degradation_mode:
            return "usable_with_degradation"
        if manifest.inventory.has_prototype_inventory and not manifest.inventory.has_usable_layout_inventory:
            return "usable_with_degradation"
        if manifest.inventory.has_usable_layout_inventory:
            return "usable"
        if any(component.editability in {ComponentEditability.EDITABLE, ComponentEditability.SEMI_EDITABLE} for component in manifest.inventory.components):
            return "usable_with_degradation"
        return "not_safely_editable"

    def apply_layout_inventory_to_plan(self, manifest: TemplateManifest, plan: PresentationPlan) -> PresentationPlan:
        adapted = plan.model_copy(deep=True)
        for slide in adapted.slides:
            ranked_options = self.layout_options_for_slide(manifest, slide)
            resolved = ranked_options[0].key if ranked_options else self.resolve_layout_key_for_slide(manifest, slide)
            if not resolved:
                slide.render_target = self._auto_layout_render_target_for_slide(slide)
                continue
            target = self._target_by_key(manifest, resolved)
            if target is not None:
                normalized_slide = self._normalize_slide_for_target(slide, target)
                slide.kind = normalized_slide.kind
                slide.title = normalized_slide.title
                slide.subtitle = normalized_slide.subtitle
                slide.text = normalized_slide.text
                slide.bullets = normalized_slide.bullets
                slide.content_blocks = normalized_slide.content_blocks
                slide.left_bullets = normalized_slide.left_bullets
                slide.right_bullets = normalized_slide.right_bullets
                slide.notes = normalized_slide.notes
                slide.runtime_profile_key = normalized_slide.runtime_profile_key
                slide.render_target = self._render_target_for_inventory_target(target)
            else:
                slide.render_target = self._auto_layout_render_target_for_slide(slide)
            slide.preferred_layout_key = resolved
        return adapted

    def build_slide_layout_reviews(self, manifest: TemplateManifest, plan: PresentationPlan) -> list[SlideLayoutReview]:
        return [
            SlideLayoutReview(
                slide_index=index,
                current_layout_key=slide.preferred_layout_key,
                current_target_key=slide.render_target.key if slide.render_target is not None else slide.preferred_layout_key,
                current_target_type=slide.render_target.type.value if slide.render_target is not None else None,
                current_runtime_profile_key=slide.runtime_profile_key,
                available_layouts=self.layout_options_for_slide(manifest, slide),
            )
            for index, slide in enumerate(plan.slides)
        ]

    def _load_manifest(self, manifest_path: Path) -> TemplateManifest:
        payload = json.loads(manifest_path.read_text(encoding="utf-8"))
        manifest = TemplateManifest.model_validate(payload)
        return self._normalize_manifest(manifest)

    def _normalize_manifest(self, manifest: TemplateManifest) -> TemplateManifest:
        if not manifest.component_styles:
            manifest.component_styles = self._default_component_styles(manifest)
        if not manifest.layouts:
            return manifest

        for layout in manifest.layouts:
            is_table_layout = self._layout_looks_like_table(layout)
            is_contacts_layout = self._layout_looks_like_contacts(layout)
            contact_binding_map = {
                10: "contact_name_or_title",
                11: "contact_role",
                12: "contact_phone",
                13: "contact_email",
            }
            for placeholder in layout.placeholders:
                if is_table_layout and placeholder.idx == 14 and placeholder.binding is None:
                    placeholder.binding = "table"
                if is_contacts_layout and placeholder.idx in contact_binding_map and placeholder.binding is None:
                    placeholder.binding = contact_binding_map[placeholder.idx]
                if placeholder.idx == 17 and placeholder.kind == PlaceholderKind.UNKNOWN:
                    placeholder.kind = PlaceholderKind.FOOTER
                self._sync_placeholder_editable_metadata(placeholder)
            layout.representation_hints = self._representation_hints_for_layout(layout)

        if manifest.default_layout_key and manifest.default_layout_key not in {layout.key for layout in manifest.layouts}:
            default_layout = next(
                (layout.key for layout in manifest.layouts if "text" in layout.supported_slide_kinds),
                manifest.layouts[0].key,
            )
            manifest.default_layout_key = default_layout
        return manifest

    def _sync_placeholder_editable_metadata(self, placeholder) -> None:
        binding = placeholder.binding
        if binding:
            placeholder.editable_role = self._editable_role_for_binding(binding)
            placeholder.editable_capabilities = self._editable_capabilities_for_binding(binding)
            return

        role = self._editable_role_for_kind(placeholder.kind)
        placeholder.editable_role = role
        placeholder.editable_capabilities = self._editable_capabilities_for_role(role)

    def _editable_role_for_kind(self, kind: PlaceholderKind) -> str | None:
        if kind == PlaceholderKind.TITLE:
            return "title"
        if kind == PlaceholderKind.SUBTITLE:
            return "subtitle"
        if kind == PlaceholderKind.BODY:
            return "body"
        if kind == PlaceholderKind.IMAGE:
            return "image"
        if kind == PlaceholderKind.TABLE:
            return "table"
        if kind == PlaceholderKind.CHART:
            return "chart"
        if kind == PlaceholderKind.FOOTER:
            return "body"
        return None

    def _editable_role_for_binding(self, binding: str) -> str:
        binding_name = binding.lower()
        if binding_name in {"title", "cover_title"}:
            return "title"
        if binding_name in {"subtitle", "cover_meta", "presentation_name"}:
            return "subtitle"
        if binding_name in {"table"}:
            return "table"
        if binding_name in {"chart"}:
            return "chart"
        if binding_name in {"image", "chart_image"} or binding_name.startswith("image_"):
            return "image"
        if binding_name in {"bullets", "left_bullets", "right_bullets"}:
            return "bullet_list"
        if binding_name.startswith("bullet_") or binding_name.startswith("left_bullet_") or binding_name.startswith("right_bullet_"):
            return "bullet_item"
        return "body"

    def _editable_capabilities_for_binding(self, binding: str) -> list[str]:
        return self._editable_capabilities_for_role(self._editable_role_for_binding(binding))

    def _editable_capabilities_for_role(self, role: str | None) -> list[str]:
        if role in {"title", "subtitle", "body"}:
            return ["text"]
        if role == "bullet_list":
            return ["bullet_list", "text"]
        if role == "bullet_item":
            return ["text", "list_item"]
        if role == "image":
            return ["image"]
        if role == "table":
            return ["table"]
        if role == "chart":
            return ["chart"]
        return []

    def resolve_layout_key_for_slide(self, manifest: TemplateManifest, slide: SlideSpec) -> str | None:
        targets = self._inventory_targets(manifest)
        if not targets:
            return self._preferred_target_key(slide)

        existing_keys = {target.key for target in targets}
        preferred_key = self._preferred_target_key(slide) or ""
        default_key = manifest.default_layout_key if manifest.default_layout_key in existing_keys else (targets[0].key if targets else None)
        if preferred_key and preferred_key in existing_keys:
            return preferred_key

        ranked_options = self.layout_options_for_slide(manifest, slide)
        if ranked_options:
            return ranked_options[0].key

        return preferred_key or default_key

    def layout_options_for_slide(self, manifest: TemplateManifest, slide: SlideSpec) -> list[SlideLayoutOption]:
        preferred_key = self._preferred_target_key(slide)
        options: list[tuple[int, SlideLayoutOption]] = []

        for target in self._inventory_targets(manifest):
            runtime_profile_key = self._logical_capacity_layout_key(
                slide=slide,
                key=target.key,
                representation_hints=target.representation_hints,
                editable_roles=target.editable_roles,
            )
            score, estimated_capacity, match_summary = self._score_layout_candidate(
                slide=slide,
                key=target.key,
                supported_slide_kinds=target.supported_slide_kinds,
                representation_hints=target.representation_hints,
                editable_roles=target.editable_roles,
                editable_slot_count=target.editable_slot_count,
                slots=target.slots,
                preferred_key=preferred_key,
            )
            options.append((
                score,
                SlideLayoutOption(
                    key=target.key,
                    name=target.name,
                    source=target.source,
                    source_label=target.source_label,
                    runtime_profile_key=runtime_profile_key,
                    supported_slide_kinds=list(target.supported_slide_kinds),
                    representation_hints=list(target.representation_hints),
                    editable_slot_count=target.editable_slot_count,
                    editable_roles=list(target.editable_roles),
                    supports_current_slide_kind=slide.kind in target.supported_slide_kinds,
                    estimated_text_capacity_chars=estimated_capacity,
                    match_summary=match_summary,
                    recommendation_label=self._recommendation_label_for_score(score),
                    recommendation_reasons=self._recommendation_reasons(
                        slide=slide,
                        key=target.key,
                        supported_slide_kinds=target.supported_slide_kinds,
                        editable_roles=target.editable_roles,
                        representation_hints=target.representation_hints,
                        estimated_capacity=estimated_capacity,
                        preferred_key=preferred_key,
                    ),
                ),
            ))

        options.sort(key=lambda item: (-item[0], item[1].name.lower(), item[1].key))
        return [option for _, option in options]

    def _preferred_target_key(self, slide: SlideSpec) -> str | None:
        if slide.render_target is not None and slide.render_target.key:
            return slide.render_target.key
        return slide.preferred_layout_key

    def _score_layout_candidate(
        self,
        *,
        slide: SlideSpec,
        key: str,
        supported_slide_kinds: list[str],
        representation_hints: list[str],
        editable_roles: list[str],
        editable_slot_count: int,
        slots: list,
        preferred_key: str | None,
    ) -> tuple[int, int | None, str | None]:
        supports_current_slide_kind = slide.kind in supported_slide_kinds
        score = editable_slot_count
        summary_parts: list[str] = []
        logical_runtime_profile = self._logical_capacity_layout_key(
            slide=slide,
            key=key,
            representation_hints=representation_hints,
            editable_roles=editable_roles,
        )

        if supports_current_slide_kind:
            score += 100
            summary_parts.append(slide.kind.value)
        if preferred_key and key == preferred_key:
            score += 250
            summary_parts.append("selected")
        if slide.runtime_profile_key and logical_runtime_profile == slide.runtime_profile_key:
            score += 40
            summary_parts.append(slide.runtime_profile_key)

        semantic_score, semantic_label = self._slot_semantic_score(
            slide=slide,
            representation_hints=representation_hints,
            editable_roles=editable_roles,
        )
        score += semantic_score
        if semantic_label:
            summary_parts.append(semantic_label)

        estimated_capacity = self._estimated_text_capacity_chars(
            slide=slide,
            key=key,
            representation_hints=representation_hints,
            editable_roles=editable_roles,
            slots=slots,
        )
        capacity_score = self._capacity_fit_score(slide, estimated_capacity)
        score += capacity_score
        if estimated_capacity:
            summary_parts.append(f"~{estimated_capacity} chars")

        summary = " · ".join(dict.fromkeys(summary_parts))
        if not summary:
            summary = "reserve option"
        return score, estimated_capacity, summary

    def _recommendation_label_for_score(self, score: int) -> str:
        if score >= 120:
            return "Рекомендуем"
        if score >= 60:
            return "Подходит"
        return "Запасной вариант"

    def _recommendation_reasons(
        self,
        *,
        slide: SlideSpec,
        key: str,
        supported_slide_kinds: list[str],
        editable_roles: list[str],
        representation_hints: list[str],
        estimated_capacity: int | None,
        preferred_key: str | None,
    ) -> list[str]:
        reasons: list[str] = []
        hints = set(representation_hints)
        roles = set(editable_roles)

        if preferred_key and key == preferred_key:
            reasons.append("Этот вариант уже выбран для текущего слайда.")
        if slide.kind in supported_slide_kinds:
            reasons.append("Совпадает с текущим типом слайда.")

        if slide.kind == SlideKind.TABLE and ("table" in hints or "table" in roles):
            reasons.append("Подходит для табличного представления данных.")
        elif slide.kind == SlideKind.CHART:
            if "chart" in hints or "chart" in roles:
                reasons.append("Подходит для графика и не требует табличного fallback.")
            elif "table" in hints or "table" in roles:
                reasons.append("Можно использовать как запасной вариант для слайда с данными.")
        elif slide.kind == SlideKind.IMAGE and ("image" in hints or "image" in roles):
            reasons.append("Подходит для слайда с изображением и подписью.")
        elif slide.kind == SlideKind.TWO_COLUMN and "two_column" in hints:
            reasons.append("Подходит для двухколоночной компоновки.")
        elif slide.kind in {SlideKind.TEXT, SlideKind.BULLETS}:
            if "contacts" in hints:
                reasons.append("Подходит для контактного или справочного блока.")
            elif "cards" in hints:
                reasons.append("Подходит для карточной подачи коротких тезисов.")
            elif "bullet_list" in roles or "bullet_item" in roles:
                reasons.append("Удобен для списка и коротких пунктов.")
            elif "body" in roles:
                reasons.append("Подходит для основного текста без ручной перестройки.")

        if estimated_capacity:
            reasons.append(f"Ориентировочно вмещает до {estimated_capacity} символов основного текста.")

        if not reasons:
            reasons.append("Можно использовать как запасной вариант для этого слайда.")

        return list(dict.fromkeys(reasons))

    def _slot_semantic_score(
        self,
        *,
        slide: SlideSpec,
        representation_hints: list[str],
        editable_roles: list[str],
    ) -> tuple[int, str | None]:
        roles = set(editable_roles)
        hints = set(representation_hints)
        score = 0
        labels: list[str] = []

        def reward(points: int, label: str) -> None:
            nonlocal score
            score += points
            labels.append(label)

        if slide.kind == SlideKind.TABLE:
            if "table" in hints or "table" in roles:
                reward(45, "table")
            return score, ", ".join(dict.fromkeys(labels)) or None
        if slide.kind == SlideKind.CHART:
            if "chart" in hints or "chart" in roles:
                reward(45, "chart")
            elif "table" in hints or "table" in roles:
                reward(20, "table fallback")
            return score, ", ".join(dict.fromkeys(labels)) or None
        if slide.kind == SlideKind.IMAGE:
            if "image" in hints or "image" in roles:
                reward(45, "image")
            return score, ", ".join(dict.fromkeys(labels)) or None

        if slide.kind == SlideKind.TWO_COLUMN and "two_column" in hints:
            reward(35, "two-column")
        if slide.kind in {SlideKind.TEXT, SlideKind.BULLETS} and "cards" in hints:
            reward(10, "cards")
        if slide.kind == SlideKind.TEXT and self._looks_like_contacts_slide(slide) and "contacts" in hints:
            reward(35, "contacts")

        if "title" in roles:
            reward(6, "title")
        if slide.subtitle and "subtitle" in roles:
            reward(6, "subtitle")

        bullet_count = self._slide_bullet_count(slide)
        if bullet_count:
            if "bullet_list" in roles:
                reward(28, "bullet slots")
            elif "bullet_item" in roles:
                reward(20, "bullet items")
            elif "body" in roles:
                reward(8, "body text")
        elif slide.kind in {SlideKind.TEXT, SlideKind.TWO_COLUMN}:
            if "body" in roles:
                reward(28, "body text")
            elif "bullet_list" in roles:
                reward(10, "bullet fallback")

        if slide.kind == SlideKind.TWO_COLUMN:
            if roles & {"bullet_list", "bullet_item"}:
                reward(8, "columns")
        return score, ", ".join(dict.fromkeys(labels)) or None

    def _estimated_text_capacity_chars(
        self,
        *,
        slide: SlideSpec,
        key: str,
        representation_hints: list[str],
        editable_roles: list[str],
        slots: list,
    ) -> int | None:
        if slide.kind not in {SlideKind.TEXT, SlideKind.BULLETS, SlideKind.TWO_COLUMN}:
            return None

        text_slots = [
            slot
            for slot in slots
            if getattr(slot, "editable_role", None) in {"body", "bullet_list", "bullet_item", "subtitle"}
            or "text" in getattr(slot, "editable_capabilities", [])
            or "list_item" in getattr(slot, "editable_capabilities", [])
        ]
        if not text_slots:
            return None

        lefts = [getattr(slot, "left_emu", None) for slot in text_slots if isinstance(getattr(slot, "left_emu", None), int)]
        tops = [getattr(slot, "top_emu", None) for slot in text_slots if isinstance(getattr(slot, "top_emu", None), int)]
        rights = [
            getattr(slot, "left_emu", 0) + getattr(slot, "width_emu", 0)
            for slot in text_slots
            if isinstance(getattr(slot, "left_emu", None), int) and isinstance(getattr(slot, "width_emu", None), int)
        ]
        bottoms = [
            getattr(slot, "top_emu", 0) + getattr(slot, "height_emu", 0)
            for slot in text_slots
            if isinstance(getattr(slot, "top_emu", None), int) and isinstance(getattr(slot, "height_emu", None), int)
        ]
        if not lefts or not tops or not rights or not bottoms:
            return None

        logical_layout_key = self._logical_capacity_layout_key(
            slide=slide,
            key=key,
            representation_hints=representation_hints,
            editable_roles=editable_roles,
        )
        profile = derive_capacity_profile_for_geometry(
            logical_layout_key,
            width_emu=max(rights) - min(lefts),
            height_emu=max(bottoms) - min(tops),
        )
        return profile.max_chars

    def _logical_capacity_layout_key(
        self,
        *,
        slide: SlideSpec,
        key: str,
        representation_hints: list[str],
        editable_roles: list[str],
    ) -> str:
        if slide.runtime_profile_key:
            return slide.runtime_profile_key

        @dataclass(frozen=True)
        class _TargetView:
            key: str
            representation_hints: list[str]
            supported_slide_kinds: list[str]
            editable_roles: list[str]

        return runtime_profile_key_for_target(
            _TargetView(
                key=key,
                representation_hints=list(representation_hints),
                supported_slide_kinds=[slide.kind.value],
                editable_roles=list(editable_roles),
            ),
            fallback_layout_key=slide.runtime_profile_key or key,
            slide_kind=slide.kind.value,
        )

    def _capacity_fit_score(self, slide: SlideSpec, estimated_capacity_chars: int | None) -> int:
        if estimated_capacity_chars is None or estimated_capacity_chars <= 0:
            return 0
        text_demand = self._slide_text_demand_chars(slide)
        if text_demand <= 0:
            return 0
        ratio = text_demand / estimated_capacity_chars
        if 0.45 <= ratio <= 1.02:
            return 35
        if 0.25 <= ratio < 0.45:
            return 12
        if 1.02 < ratio <= 1.18:
            return 10
        if 1.18 < ratio <= 1.45:
            return -12
        return -36

    def _slide_text_demand_chars(self, slide: SlideSpec) -> int:
        text_parts: list[str] = []
        text_parts.extend(item for item in slide.bullets if item)
        text_parts.extend(item for item in slide.left_bullets if item)
        text_parts.extend(item for item in slide.right_bullets if item)
        if slide.text:
            text_parts.append(slide.text)
        for block in slide.content_blocks:
            if block.text:
                text_parts.append(block.text)
            text_parts.extend(item for item in block.items if item)
        if slide.notes:
            text_parts.append(slide.notes)
        return sum(len(part.strip()) for part in text_parts if part and part.strip())

    def _slide_bullet_count(self, slide: SlideSpec) -> int:
        if slide.bullets:
            return len([item for item in slide.bullets if item.strip()])
        if slide.left_bullets or slide.right_bullets:
            return len([item for item in [*slide.left_bullets, *slide.right_bullets] if item.strip()])
        return sum(len([item for item in block.items if item.strip()]) for block in slide.content_blocks)

    def _slide_looks_like_kpi_cards(self, slide: SlideSpec) -> bool:
        items = [*slide.bullets, *slide.left_bullets, *slide.right_bullets]
        if not items and slide.text:
            items = [slide.text]
        numeric_items = 0
        for item in items[:4]:
            if any(char.isdigit() for char in item):
                numeric_items += 1
        return numeric_items >= 2

    def _inventory_targets(self, manifest: TemplateManifest) -> list[_InventoryTarget]:
        targets: list[_InventoryTarget] = []
        seen_keys: set[str] = set()

        for layout in manifest.layouts:
            if layout.key in seen_keys:
                continue
            seen_keys.add(layout.key)
            editable_roles = sorted({
                placeholder.editable_role
                for placeholder in layout.placeholders
                if placeholder.editable_role
            })
            editable_slot_count = sum(
                1
                for placeholder in layout.placeholders
                if placeholder.editable_role or placeholder.editable_capabilities
            )
            targets.append(
                _InventoryTarget(
                    key=layout.key,
                    name=layout.name,
                    source="layout",
                    source_label=f"layout {layout.slide_layout_index + 1}",
                    supported_slide_kinds=list(layout.supported_slide_kinds),
                    representation_hints=list(layout.representation_hints),
                    editable_roles=editable_roles,
                    editable_slot_count=editable_slot_count,
                    slots=layout.placeholders,
                )
            )

        for prototype in manifest.prototype_slides:
            if prototype.key in seen_keys:
                continue
            seen_keys.add(prototype.key)
            editable_roles = sorted({
                token.editable_role
                for token in prototype.tokens
                if token.editable_role
            })
            editable_slot_count = sum(
                1
                for token in prototype.tokens
                if token.editable_role or token.editable_capabilities
            )
            targets.append(
                _InventoryTarget(
                    key=prototype.key,
                    name=prototype.name,
                    source="prototype",
                    source_label=f"prototype slide {prototype.source_slide_index + 1}",
                    supported_slide_kinds=list(prototype.supported_slide_kinds),
                    representation_hints=list(prototype.representation_hints),
                    editable_roles=editable_roles,
                    editable_slot_count=editable_slot_count,
                    slots=prototype.tokens,
                )
            )

        return targets

    def _target_by_key(self, manifest: TemplateManifest, key: str) -> _InventoryTarget | None:
        return next((target for target in self._inventory_targets(manifest) if target.key == key), None)

    def _render_target_for_inventory_target(self, target: _InventoryTarget) -> SlideRenderTarget:
        binding_keys = sorted(
            {
                getattr(slot, "binding", None)
                for slot in target.slots
                if getattr(slot, "binding", None)
            }
        )
        return SlideRenderTarget(
            type=RenderTargetType(target.source),
            key=target.key,
            label=target.name,
            source=target.source_label,
            binding_keys=binding_keys,
            confidence="high",
            degradation_reasons=[] if target.source == RenderTargetType.LAYOUT.value else ["inventory_fallback"],
        )

    def _auto_layout_render_target_for_slide(self, slide: SlideSpec) -> SlideRenderTarget:
        return SlideRenderTarget(
            type=RenderTargetType.AUTO_LAYOUT,
            key=slide.runtime_profile_key or slide.preferred_layout_key,
            label="Auto layout fallback",
            source="runtime fallback",
            binding_keys=[],
            confidence="medium",
            degradation_reasons=["inventory_unresolved"],
        )

    def _normalize_slide_for_target(self, slide: SlideSpec, target: _InventoryTarget) -> SlideSpec:
        hints = set(target.representation_hints)
        supported = set(target.supported_slide_kinds)
        runtime_profile_key = slide.runtime_profile_key or runtime_profile_key_for_target(
            target,
            fallback_layout_key=slide.preferred_layout_key,
            slide_kind=slide.kind.value,
        )

        if "contacts" not in hints and (slide.left_bullets or slide.right_bullets):
            text_parts = [part for part in [slide.text or "", slide.notes or ""] if part.strip()]
            contact_parts = [*slide.left_bullets, *slide.right_bullets]
            if supported == {"bullets"} or ("bullets" in supported and "text" not in supported):
                merged_bullets = [*slide.bullets, *[item.strip() for item in text_parts if item.strip()], *[item.strip() for item in contact_parts if item.strip()]]
                return slide.model_copy(
                    update={
                        "kind": SlideKind.BULLETS,
                        "text": None,
                        "bullets": merged_bullets,
                        "content_blocks": [self._list_block(merged_bullets)] if merged_bullets else [],
                        "left_bullets": [],
                        "right_bullets": [],
                        "notes": None,
                        "runtime_profile_key": runtime_profile_key if runtime_profile_key != "contacts" else "list_full_width",
                    },
                    deep=True,
                )

            merged_text = self._merge_text_parts([*text_parts, *[item.strip() for item in contact_parts if item.strip()]])
            return slide.model_copy(
                update={
                    "kind": SlideKind.TEXT,
                    "text": merged_text or slide.text,
                        "content_blocks": self._paragraph_blocks_from_parts(merged_text) if merged_text else slide.content_blocks,
                        "left_bullets": [],
                        "right_bullets": [],
                        "notes": None if merged_text else slide.notes,
                        "runtime_profile_key": "text_full_width",
                    },
                    deep=True,
                )

        if slide.kind == SlideKind.TWO_COLUMN and "two_column" not in hints:
            merged_bullets = [*slide.left_bullets, *slide.right_bullets]
            if "bullets" in supported:
                return slide.model_copy(
                    update={
                        "kind": SlideKind.BULLETS,
                        "bullets": merged_bullets,
                        "content_blocks": [self._list_block(merged_bullets)] if merged_bullets else [],
                        "left_bullets": [],
                        "right_bullets": [],
                        "runtime_profile_key": "list_with_icons" if slide.runtime_profile_key == "list_with_icons" else "list_full_width",
                    },
                    deep=True,
                )
            merged_text = self._merge_text_parts(merged_bullets)
            return slide.model_copy(
                update={
                    "kind": SlideKind.TEXT,
                        "text": merged_text,
                        "content_blocks": self._paragraph_blocks_from_parts(merged_text),
                        "left_bullets": [],
                        "right_bullets": [],
                        "runtime_profile_key": "text_full_width",
                    },
                    deep=True,
                )

        if slide.kind == SlideKind.TEXT and "text" not in supported and "bullets" in supported:
            merged_bullets = self._text_to_bullets(slide)
            return slide.model_copy(
                update={
                    "kind": SlideKind.BULLETS,
                        "text": None,
                        "bullets": merged_bullets,
                        "content_blocks": [self._list_block(merged_bullets)] if merged_bullets else [],
                        "notes": None,
                        "runtime_profile_key": "list_full_width",
                    },
                    deep=True,
                )

        if slide.kind == SlideKind.BULLETS and "bullets" not in supported and "text" in supported:
            merged_text = self._merge_text_parts(slide.bullets)
            return slide.model_copy(
                update={
                    "kind": SlideKind.TEXT,
                        "text": merged_text,
                        "bullets": [],
                        "content_blocks": self._paragraph_blocks_from_parts(merged_text),
                        "runtime_profile_key": "text_full_width",
                    },
                    deep=True,
                )

        if slide.runtime_profile_key == runtime_profile_key:
            return slide
        return slide.model_copy(update={"runtime_profile_key": runtime_profile_key}, deep=True)

    def _text_to_bullets(self, slide: SlideSpec) -> list[str]:
        bullets = [item.strip() for item in slide.bullets if item.strip()]
        if bullets:
            return bullets
        merged = self._merge_text_parts([slide.text or "", slide.notes or ""])
        if not merged:
            return []
        parts = [part.strip() for part in re.split(r"(?<=[.!?;])\s+", merged) if part.strip()]
        return parts or [merged]

    def _merge_text_parts(self, parts: list[str]) -> str:
        normalized = [part.strip() for part in parts if part and part.strip()]
        return "\n".join(normalized)

    def _paragraph_blocks_from_parts(self, *parts: str) -> list[SlideContentBlock]:
        return [
            SlideContentBlock(kind=SlideContentBlockKind.PARAGRAPH, text=part.strip())
            for part in parts
            if part and part.strip()
        ]

    def _list_block(self, items: list[str]) -> SlideContentBlock:
        return SlideContentBlock(
            kind=SlideContentBlockKind.BULLET_LIST,
            items=[item.strip() for item in items if item and item.strip()],
        )

    def _first_target_with_hint(self, targets: list[_InventoryTarget], hint: str):
        for target in targets:
            if hint in target.representation_hints or (hint == "title" and "title" in target.supported_slide_kinds):
                return target
        return None

    def _best_text_target(self, targets: list[_InventoryTarget]):
        for target in targets:
            if "text" not in target.supported_slide_kinds:
                continue
            if any(hint in target.representation_hints for hint in {"table", "chart", "image", "contacts", "cards", "two_column"}):
                continue
            return target
        return next((target for target in targets if "text" in target.supported_slide_kinds), None)

    def _best_bullets_target(self, targets: list[_InventoryTarget]):
        for target in targets:
            if "bullets" not in target.supported_slide_kinds:
                continue
            if any(hint in target.representation_hints for hint in {"table", "chart", "image", "contacts", "cards"}):
                continue
            return target
        return next((target for target in targets if "bullets" in target.supported_slide_kinds), None)

    def _looks_like_contacts_slide(self, slide: SlideSpec) -> bool:
        text_parts = [slide.title or "", slide.subtitle or "", slide.text or "", slide.notes or "", *slide.bullets]
        combined = " ".join(part for part in text_parts if part).lower()
        return "@" in combined or "тел" in combined or "phone" in combined or "email" in combined

    def _representation_hints_for_layout(self, layout) -> list[str]:
        hints = list(layout.representation_hints)
        capabilities = {capability for placeholder in layout.placeholders for capability in placeholder.editable_capabilities}
        bindings = {placeholder.binding for placeholder in layout.placeholders if placeholder.binding}

        if self._layout_looks_like_cards(layout) and "cards" not in hints:
            hints.append("cards")
        if self._layout_looks_like_contacts(layout) or {"contact_name_or_title", "contact_role", "contact_phone", "contact_email"} & bindings:
            hints.append("contacts")
        if "table" in capabilities or self._layout_looks_like_table(layout):
            hints.append("table")
        if "chart" in capabilities or any(placeholder.kind == PlaceholderKind.CHART for placeholder in layout.placeholders):
            hints.append("chart")
        if "image" in capabilities or any(
            placeholder.kind == PlaceholderKind.IMAGE or placeholder.idx == 16 for placeholder in layout.placeholders
        ):
            hints.append("image")
        return list(dict.fromkeys(hints))

    def _layout_looks_like_cards(self, layout) -> bool:
        return (
            "карточ" in layout.name.lower()
            or "cards" in layout.name.lower()
            or sum(1 for placeholder in layout.placeholders if placeholder.idx in {11, 12, 13}) >= 3
        )

    def _layout_looks_like_contacts(self, layout) -> bool:
        contact_slots = {placeholder.idx for placeholder in layout.placeholders if placeholder.idx in {10, 11, 12, 13}}
        bindings = {placeholder.binding for placeholder in layout.placeholders if placeholder.binding}
        return (
            "конт" in layout.name.lower()
            or len(contact_slots) >= 2
            or bool({"contact_name_or_title", "contact_role", "contact_phone", "contact_email"} & bindings)
        )

    def _layout_looks_like_table(self, layout) -> bool:
        bindings = {placeholder.binding for placeholder in layout.placeholders if placeholder.binding}
        return (
            "табл" in layout.name.lower()
            or any(placeholder.kind == PlaceholderKind.TABLE for placeholder in layout.placeholders)
            or "table" in layout.supported_slide_kinds
            or "table" in bindings
        )

    def _default_component_styles(self, manifest: TemplateManifest) -> dict[str, TemplateComponentStyleSpec]:
        theme = manifest.theme
        title_style = theme.master_text_styles.get("title")
        body_style = theme.master_text_styles.get("body")
        other_style = theme.master_text_styles.get("other")
        cards_body_font = manifest.design_tokens.get("cards_body_font_size_pt")
        margin_x = 91440
        margin_y = 45720
        cards_layout = next((layout for layout in manifest.layouts if layout.key in {"cards_3", "cards_kpi"}), None)
        if cards_layout is not None:
            body_placeholder = next((item for item in cards_layout.placeholders if item.idx in {11, 12, 13}), None)
            if body_placeholder is not None:
                margin_x = body_placeholder.margin_left_emu or margin_x
                margin_y = body_placeholder.margin_top_emu or margin_y
        cards_text_styles = {}
        if title_style is not None:
            cards_text_styles["title"] = title_style.model_copy(update={"font_size_pt": min(title_style.font_size_pt or 20.0, 20.0), "color": "#FFFFFF"})
        if body_style is not None:
            cards_text_styles["body"] = body_style.model_copy(update={"font_size_pt": float(cards_body_font) if isinstance(cards_body_font, (int, float)) else 16.0, "color": "#FFFFFF"})
            cards_text_styles["kpi_value"] = body_style.model_copy(update={"font_size_pt": 22.0, "bold": True, "color": "#FFFFFF"})
        if other_style is not None:
            cards_text_styles["kpi_label"] = other_style.model_copy(update={"font_size_pt": 12.0, "color": "#E4F1FF"})
        text_text_styles = {}
        if title_style is not None:
            text_text_styles["title"] = title_style
        if body_style is not None:
            text_text_styles["body"] = body_style
        if other_style is not None:
            text_text_styles["footer"] = other_style
        table_text_styles = {}
        if body_style is not None:
            table_text_styles["header"] = body_style.model_copy(update={"bold": True, "color": "#091E38"})
            table_text_styles["body"] = body_style.model_copy(update={"color": "#081C4F"})
        chart_text_styles = {}
        if title_style is not None:
            chart_text_styles["title"] = title_style
        if body_style is not None:
            chart_text_styles["subtitle"] = body_style
        image_text_styles = {}
        if title_style is not None:
            image_text_styles["title"] = title_style
        if body_style is not None:
            image_text_styles["body"] = body_style
            image_text_styles["subtitle"] = body_style.model_copy(update={"font_size_pt": 18.0})
        if other_style is not None:
            image_text_styles["footer"] = other_style
        cover_text_styles = {}
        if title_style is not None:
            cover_text_styles["title"] = title_style.model_copy(update={"font_size_pt": 46.0, "color": "#F5F9FE", "bold": True})
        if body_style is not None:
            cover_text_styles["meta"] = body_style.model_copy(update={"font_size_pt": 22.0, "color": "#F5F9FE"})
        if other_style is not None:
            cover_text_styles["footer"] = other_style.model_copy(update={"font_size_pt": 14.0, "color": "#F5F9FE"})
        list_icons_text_styles = {}
        if title_style is not None:
            list_icons_text_styles["title"] = title_style
        if body_style is not None:
            list_icons_text_styles["subtitle"] = body_style.model_copy(update={"font_size_pt": 18.0})
            list_icons_text_styles["left"] = body_style
            list_icons_text_styles["right"] = body_style
        if other_style is not None:
            list_icons_text_styles["footer"] = other_style
        contacts_text_styles = {}
        if body_style is not None:
            contacts_text_styles["primary"] = body_style.model_copy(update={"font_size_pt": 18.0})
            contacts_text_styles["secondary"] = body_style.model_copy(update={"font_size_pt": 14.0})
        return {
            "cards": TemplateComponentStyleSpec(
                text_styles=cards_text_styles,
                spacing_tokens={
                    "content_margin_x_emu": margin_x,
                    "content_margin_y_emu": margin_y,
                    "title_body_gap_emu": 100000,
                    "body_metrics_gap_emu": 180000,
                    "metrics_gap_x_emu": 180000,
                    "metrics_gap_y_emu": 160000,
                },
                behavior_tokens={
                    "kpi_max_metrics": 4,
                    "kpi_value_compact_font_pt": 20,
                    "kpi_value_regular_font_pt": 22,
                },
            ),
            "text": TemplateComponentStyleSpec(
                text_styles=text_text_styles,
                spacing_tokens={
                    "content_margin_x_emu": 91440,
                    "content_margin_y_emu": 45720,
                },
            ),
            "table": TemplateComponentStyleSpec(
                text_styles=table_text_styles,
                shape_style=theme.master_shape_styles.get("table"),
                spacing_tokens={
                    "cell_margin_left_emu": 80000,
                    "cell_margin_right_emu": 80000,
                    "cell_margin_top_emu": 40000,
                    "cell_margin_bottom_emu": 40000,
                },
                behavior_tokens={
                    "header_fill_color": manifest.design_tokens.get("table_header_fill_color"),
                    "header_text_color": manifest.design_tokens.get("table_header_text_color"),
                    "border_color": manifest.design_tokens.get("table_border_color"),
                    "body_fill_transparent": manifest.design_tokens.get("table_body_fill_transparent"),
                    "preserve_source_fill_colors": manifest.design_tokens.get("table_preserve_source_fill_colors"),
                    "render_as_shapes": manifest.design_tokens.get("table_render_as_shapes"),
                },
            ),
            "chart": TemplateComponentStyleSpec(
                text_styles=chart_text_styles,
                shape_style=theme.master_shape_styles.get("chart") or TemplateShapeStyleSpec(
                    role="chart",
                    chart_plot_left_factor=0.0,
                    chart_plot_top_factor=0.0,
                    chart_plot_width_factor=1.0,
                    chart_plot_height_factor=1.0,
                ),
                behavior_tokens={
                    "rank_color_1": manifest.design_tokens.get("chart_rank_color_1"),
                    "rank_color_2": manifest.design_tokens.get("chart_rank_color_2"),
                    "rank_color_3": manifest.design_tokens.get("chart_rank_color_3"),
                    "rank_color_4": manifest.design_tokens.get("chart_rank_color_4"),
                },
            ),
            "image": TemplateComponentStyleSpec(
                text_styles=image_text_styles,
                shape_style=theme.master_shape_styles.get("image"),
                spacing_tokens={
                    "content_margin_x_emu": 91440,
                    "content_margin_y_emu": 45720,
                    "title_content_gap_emu": 180000,
                    "title_body_gap_no_subtitle_emu": 300000,
                    "content_footer_gap_emu": 180000,
                    "min_image_height_emu": 1200000,
                    "secondary_min_height_emu": 700000,
                    "secondary_reserved_image_gap_emu": 900000,
                },
            ),
            "cover": TemplateComponentStyleSpec(
                text_styles=cover_text_styles,
                spacing_tokens={
                    "title_top_emu": 651176,
                    "title_left_emu": 444249,
                    "title_width_emu": 10693901,
                    "title_min_height_emu": 1422646,
                    "meta_left_emu": 444249,
                    "meta_top_emu": 2438400,
                    "meta_width_emu": 8200000,
                    "meta_min_height_emu": 700000,
                    "meta_gap_emu": 220000,
                    "bottom_limit_emu": 6200000,
                },
            ),
            "list_with_icons": TemplateComponentStyleSpec(
                text_styles=list_icons_text_styles,
                spacing_tokens={
                    "title_content_gap_emu": 180000,
                    "title_body_gap_no_subtitle_emu": 300000,
                    "content_footer_gap_emu": 180000,
                },
            ),
            "contacts": TemplateComponentStyleSpec(
                text_styles=contacts_text_styles,
                behavior_tokens={
                    "primary_threshold_chars": 60,
                    "secondary_threshold_chars": 40,
                    "font_decrement_pt": 2.0,
                },
            ),
        }

    def _template_dir(self, template_id: str) -> Path:
        return self._safe_child_path(self._templates_dir, template_id)

    def _safe_child_path(self, base_dir: Path, child_name: str) -> Path:
        if not child_name or not child_name.strip():
            raise ValueError("Path segment must not be empty")
        normalized = child_name.replace("\\", "/").strip()
        if Path(normalized).is_absolute() or normalized.startswith("/"):
            raise ValueError("Absolute paths are not allowed")
        parts = PurePosixPath(normalized).parts
        if not parts:
            raise ValueError("Path segment must not be empty")
        if len(parts) != 1 or parts[0] in {".", ".."}:
            raise ValueError(f"Path '{child_name}' escapes the storage root")

        candidate = (base_dir / parts[0]).resolve()
        try:
            candidate.relative_to(base_dir)
        except ValueError as exc:
            raise ValueError(f"Path '{child_name}' escapes the storage root") from exc
        return candidate
