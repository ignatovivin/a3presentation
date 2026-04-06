from __future__ import annotations

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
