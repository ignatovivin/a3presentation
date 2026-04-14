from __future__ import annotations

import json
import re
from pathlib import Path

from pptx import Presentation

from a3presentation.domain.template import (
    GenerationMode,
    LayoutSpec,
    PlaceholderKind,
    PlaceholderSpec,
    PrototypeSlideSpec,
    PrototypeTokenSpec,
    TemplateManifest,
)


class TemplateAnalyzer:
    TOKEN_PATTERN = re.compile(r"{{\s*([a-zA-Z0-9_]+)\s*}}")

    def analyze(self, template_id: str, template_path: Path, display_name: str | None = None) -> TemplateManifest:
        manifest_path = template_path.with_name("manifest.json")
        if manifest_path.exists():
            payload = json.loads(manifest_path.read_text(encoding="utf-8"))
            payload["template_id"] = template_id
            if display_name:
                payload["display_name"] = display_name
            payload["source_pptx"] = template_path.name
            manifest = TemplateManifest.model_validate(payload)
            analyzed = self._analyze_presentation(template_path, template_id, display_name)
            self._backfill_geometry(manifest, analyzed)
            return manifest

        return self._analyze_presentation(template_path, template_id, display_name)

    def _analyze_presentation(self, template_path: Path, template_id: str, display_name: str | None = None) -> TemplateManifest:
        presentation = Presentation(str(template_path))
        layouts: list[LayoutSpec] = []
        prototype_slides: list[PrototypeSlideSpec] = []
        used_layout_keys: set[str] = set()

        for master_index, slide_master in enumerate(presentation.slide_masters):
            for layout_index, slide_layout in enumerate(slide_master.slide_layouts):
                placeholders: list[PlaceholderSpec] = []
                for shape in slide_layout.placeholders:
                    placeholder_format = shape.placeholder_format
                    placeholders.append(
                        PlaceholderSpec(
                            name=shape.name,
                            kind=self._map_placeholder_kind(placeholder_format.type),
                            idx=placeholder_format.idx,
                            shape_name=shape.name,
                            left_emu=int(shape.left),
                            top_emu=int(shape.top),
                            width_emu=int(shape.width),
                            height_emu=int(shape.height),
                            margin_left_emu=self._text_frame_margin(shape, "left"),
                            margin_right_emu=self._text_frame_margin(shape, "right"),
                            margin_top_emu=self._text_frame_margin(shape, "top"),
                            margin_bottom_emu=self._text_frame_margin(shape, "bottom"),
                        )
                    )

                supported_slide_kinds = self._infer_slide_kinds(placeholders, slide_layout.name)
                layout_key = self._make_unique_layout_key(
                    base_key=self._slugify(slide_layout.name, fallback=f"layout_{master_index}_{layout_index}"),
                    master_index=master_index,
                    layout_index=layout_index,
                    used_keys=used_layout_keys,
                )
                layouts.append(
                    LayoutSpec(
                        key=layout_key,
                        name=slide_layout.name,
                        slide_master_index=master_index,
                        slide_layout_index=layout_index,
                        supported_slide_kinds=supported_slide_kinds,
                        placeholders=placeholders,
                    )
                )

        for index, slide in enumerate(presentation.slides):
            tokens: list[PrototypeTokenSpec] = []
            seen_tokens: set[tuple[str, str | None]] = set()

            for shape in slide.shapes:
                if not getattr(shape, "has_text_frame", False):
                    continue
                text = shape.text or ""
                for match in self.TOKEN_PATTERN.finditer(text):
                    token = match.group(1)
                    key = (token, shape.name)
                    if key in seen_tokens:
                        continue
                    seen_tokens.add(key)
                    tokens.append(
                        PrototypeTokenSpec(
                            token=token,
                            binding=self._infer_binding(token),
                            shape_name=shape.name,
                            left_emu=int(shape.left),
                            top_emu=int(shape.top),
                            width_emu=int(shape.width),
                            height_emu=int(shape.height),
                            margin_left_emu=self._text_frame_margin(shape, "left"),
                            margin_right_emu=self._text_frame_margin(shape, "right"),
                            margin_top_emu=self._text_frame_margin(shape, "top"),
                            margin_bottom_emu=self._text_frame_margin(shape, "bottom"),
                        )
                    )

            if not tokens:
                continue

            prototype_slides.append(
                PrototypeSlideSpec(
                    key=f"slide_{index + 1}",
                    name=slide.name or f"Slide {index + 1}",
                    source_slide_index=index,
                    supported_slide_kinds=self._infer_slide_kinds_from_tokens(tokens),
                    tokens=tokens,
                )
            )

        default_layout_key = layouts[1].key if len(layouts) > 1 else (layouts[0].key if layouts else None)
        return TemplateManifest(
            template_id=template_id,
            display_name=display_name or template_id,
            source_pptx=template_path.name,
            generation_mode=GenerationMode.PROTOTYPE if prototype_slides else GenerationMode.LAYOUT,
            default_layout_key=default_layout_key,
            layouts=layouts,
            prototype_slides=prototype_slides,
        )

    def _backfill_geometry(self, manifest: TemplateManifest, analyzed: TemplateManifest) -> None:
        analyzed_layouts = {(layout.name, layout.slide_layout_index): layout for layout in analyzed.layouts}
        for layout in manifest.layouts:
            source_layout = analyzed_layouts.get((layout.name, layout.slide_layout_index))
            if source_layout is None:
                continue
            for placeholder in layout.placeholders:
                source_placeholder = next(
                    (
                        item
                        for item in source_layout.placeholders
                        if item.shape_name == placeholder.shape_name
                        or (item.idx is not None and item.idx == placeholder.idx)
                    ),
                    None,
                )
                if source_placeholder is None:
                    continue
                for field in (
                    "left_emu",
                    "top_emu",
                    "width_emu",
                    "height_emu",
                    "margin_left_emu",
                    "margin_right_emu",
                    "margin_top_emu",
                    "margin_bottom_emu",
                ):
                    if getattr(placeholder, field, None) is None:
                        setattr(placeholder, field, getattr(source_placeholder, field, None))

        analyzed_prototypes = {slide.name: slide for slide in analyzed.prototype_slides}
        for prototype in manifest.prototype_slides:
            source_slide = analyzed_prototypes.get(prototype.name)
            if source_slide is not None:
                for index, token in enumerate(prototype.tokens):
                    source_token = next(
                        (item for item in source_slide.tokens if item.shape_name == token.shape_name or item.token == token.token),
                        source_slide.tokens[index] if index < len(source_slide.tokens) else None,
                    )
                    if source_token is None:
                        continue
                    for field in (
                        "left_emu",
                        "top_emu",
                        "width_emu",
                        "height_emu",
                        "margin_left_emu",
                        "margin_right_emu",
                        "margin_top_emu",
                        "margin_bottom_emu",
                    ):
                        if getattr(token, field, None) is None:
                            setattr(token, field, getattr(source_token, field, None))
            text_token = next(
                (
                    item
                    for item in prototype.tokens
                    if item.binding in {"title", "subtitle", "text", "body", "main_text", "secondary_text", "cover_title"}
                ),
                prototype.tokens[0] if prototype.tokens else None,
            )
            if text_token is None:
                continue
            if not any(
                all(isinstance(v, int) and v > 0 for v in [token.left_emu, token.top_emu, token.width_emu, token.height_emu])
                for token in prototype.tokens
            ):
                token = text_token
                token.left_emu = token.left_emu or 1
                token.top_emu = token.top_emu or 1
                token.width_emu = token.width_emu or 1
                token.height_emu = token.height_emu or 1
            if not any(
                item.binding in {"title", "subtitle", "text", "body", "main_text", "secondary_text", "cover_title"}
                and any(getattr(item, field, None) is not None for field in (
                    "margin_left_emu",
                    "margin_right_emu",
                    "margin_top_emu",
                    "margin_bottom_emu",
                ))
                for item in prototype.tokens
            ):
                token = text_token
                token.margin_left_emu = token.margin_left_emu if token.margin_left_emu is not None else 0
                token.margin_right_emu = token.margin_right_emu if token.margin_right_emu is not None else 0
                token.margin_top_emu = token.margin_top_emu if token.margin_top_emu is not None else 0
                token.margin_bottom_emu = token.margin_bottom_emu if token.margin_bottom_emu is not None else 0

    def _map_placeholder_kind(self, placeholder_type) -> PlaceholderKind:
        name = getattr(placeholder_type, "name", str(placeholder_type)).lower()
        if "title" in name and "ctr" not in name:
            return PlaceholderKind.TITLE
        if "sub" in name:
            return PlaceholderKind.SUBTITLE
        if "pic" in name or "media" in name or "obj" in name:
            return PlaceholderKind.IMAGE
        if "tbl" in name:
            return PlaceholderKind.TABLE
        if "chart" in name:
            return PlaceholderKind.CHART
        if "footer" in name or "dt" in name or "sld_num" in name:
            return PlaceholderKind.FOOTER
        if "body" in name or "content" in name or "text" in name:
            return PlaceholderKind.BODY
        return PlaceholderKind.UNKNOWN

    def _infer_slide_kinds(self, placeholders: list[PlaceholderSpec], layout_name: str = "") -> list[str]:
        kinds = {item.kind for item in placeholders}
        supported: list[str] = []
        layout_name_lower = layout_name.lower()

        if PlaceholderKind.TITLE in kinds and PlaceholderKind.SUBTITLE in kinds and len(placeholders) <= 3:
            supported.append("title")
        elif PlaceholderKind.TITLE in kinds and ("тит" in layout_name_lower or "title" in layout_name_lower):
            supported.append("title")
        if PlaceholderKind.BODY in kinds:
            supported.extend(["bullets", "text", "table"])
        if PlaceholderKind.IMAGE in kinds:
            supported.append("image")
        return list(dict.fromkeys(supported))

    def _infer_slide_kinds_from_tokens(self, tokens: list[PrototypeTokenSpec]) -> list[str]:
        token_names = {token.token.lower() for token in tokens}
        supported: list[str] = []

        if "title" in token_names and ("subtitle" in token_names or len(token_names) <= 2):
            supported.append("title")
        if "bullets" in token_names or any(name.startswith("bullet_") for name in token_names):
            supported.append("bullets")
        if "text" in token_names or "body" in token_names or "summary" in token_names:
            supported.append("text")
        if "left_bullets" in token_names or "right_bullets" in token_names:
            supported.append("two_column")
        if "image" in token_names or any(name.startswith("image_") for name in token_names):
            supported.append("image")
        if not supported:
            supported.append("text")
        return list(dict.fromkeys(supported))

    def _infer_binding(self, token: str) -> str:
        token_name = token.lower()
        if token_name in {"title", "subtitle", "text", "body", "summary", "notes"}:
            return token_name
        if token_name in {"bullets", "left_bullets", "right_bullets"}:
            return token_name
        if token_name.startswith("bullet_"):
            return token_name
        if token_name.startswith("left_bullet_") or token_name.startswith("right_bullet_"):
            return token_name
        return "text"

    def _slugify(self, value: str, fallback: str) -> str:
        cleaned = "".join(char.lower() if char.isalnum() else "_" for char in value).strip("_")
        while "__" in cleaned:
            cleaned = cleaned.replace("__", "_")
        return cleaned or fallback

    def _make_unique_layout_key(self, base_key: str, master_index: int, layout_index: int, used_keys: set[str]) -> str:
        candidate = base_key
        if candidate in used_keys:
            candidate = f"{base_key}_m{master_index}_l{layout_index}"
        used_keys.add(candidate)
        return candidate

    def _text_frame_margin(self, shape, side: str) -> int | None:
        if not getattr(shape, "has_text_frame", False):
            return None
        text_frame = shape.text_frame
        value = getattr(text_frame, f"margin_{side}", None)
        return int(value) if value is not None else None
