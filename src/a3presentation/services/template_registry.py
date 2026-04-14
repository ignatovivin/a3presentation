from __future__ import annotations

import json
from pathlib import Path, PurePosixPath

from a3presentation.domain.template import PlaceholderKind, TemplateManifest


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

    def _load_manifest(self, manifest_path: Path) -> TemplateManifest:
        payload = json.loads(manifest_path.read_text(encoding="utf-8"))
        manifest = TemplateManifest.model_validate(payload)
        return self._normalize_manifest(manifest)

    def _normalize_manifest(self, manifest: TemplateManifest) -> TemplateManifest:
        if not manifest.layouts:
            return manifest

        logical_keys = {
            "cover": lambda layout: "тит" in layout.name.lower() or ("title" in layout.supported_slide_kinds and len(layout.placeholders) <= 3),
            "text_full_width": lambda layout: (
                "text" in layout.supported_slide_kinds
                and any(placeholder.idx == 14 for placeholder in layout.placeholders)
                and any(placeholder.idx == 17 for placeholder in layout.placeholders)
                and not any(placeholder.idx == 16 for placeholder in layout.placeholders)
                and "конт" not in layout.name.lower()
                and "карточ" not in layout.name.lower()
                and "переч" not in layout.name.lower()
                and "фон" not in layout.name.lower()
            ),
            "list_full_width": lambda layout: (
                ("переч" in layout.name.lower() or "list" in layout.name.lower() or "list_full_width" == layout.key)
                and "bullets" in layout.supported_slide_kinds
                and any(placeholder.idx == 14 for placeholder in layout.placeholders)
                and any(placeholder.idx == 17 for placeholder in layout.placeholders)
                and not any(placeholder.idx == 16 for placeholder in layout.placeholders)
            ),
            "table": lambda layout: ("табл" in layout.name.lower() or "table" == layout.key) and any(
                placeholder.idx == 14 for placeholder in layout.placeholders
            ),
            "image_text": lambda layout: "image" in layout.supported_slide_kinds and any(
                placeholder.idx == 16 or placeholder.kind == PlaceholderKind.IMAGE for placeholder in layout.placeholders
            ),
            "cards_3": lambda layout: (
                ("карточ" in layout.name.lower() or "cards" in layout.name.lower())
                or (
                    sum(1 for placeholder in layout.placeholders if placeholder.idx in {11, 12, 13}) >= 3
                    and not any(placeholder.idx == 14 for placeholder in layout.placeholders)
                )
            ),
            "list_with_icons": lambda layout: any(placeholder.idx == 21 for placeholder in layout.placeholders),
            "contacts": lambda layout: (
                "конт" in layout.name.lower()
                or any(placeholder.idx == 10 for placeholder in layout.placeholders)
            ),
        }
        assigned = {layout.key for layout in manifest.layouts}
        for logical_key, predicate in logical_keys.items():
            if logical_key in assigned:
                continue
            candidate = next((layout for layout in manifest.layouts if predicate(layout)), None)
            if candidate is not None:
                manifest.layouts.append(candidate.model_copy(update={"key": logical_key}, deep=True))
                assigned.add(logical_key)
        if "text_full_width" not in assigned:
            candidate = next((layout for layout in manifest.layouts if layout.key == "list_full_width"), None)
            if candidate is not None:
                manifest.layouts.append(candidate.model_copy(update={"key": "text_full_width"}, deep=True))
                assigned.add("text_full_width")

        for layout in manifest.layouts:
            if layout.key == "table" or "табл" in layout.name.lower():
                for placeholder in layout.placeholders:
                    if placeholder.idx == 14 and placeholder.binding is None:
                        placeholder.binding = "table"
            elif layout.key == "contacts":
                binding_map = {
                    10: "contact_name_or_title",
                    11: "contact_role",
                    12: "contact_phone",
                    13: "contact_email",
                }
                for placeholder in layout.placeholders:
                    if placeholder.idx in binding_map and placeholder.binding is None:
                        placeholder.binding = binding_map[placeholder.idx]
            elif layout.key in {"text_full_width", "list_full_width"}:
                for placeholder in layout.placeholders:
                    if placeholder.idx == 17 and placeholder.kind == PlaceholderKind.UNKNOWN:
                        placeholder.kind = PlaceholderKind.FOOTER

        if manifest.default_layout_key and manifest.default_layout_key not in {layout.key for layout in manifest.layouts}:
            default_layout = next(
                (layout.key for layout in manifest.layouts if "text" in layout.supported_slide_kinds),
                manifest.layouts[0].key,
            )
            manifest.default_layout_key = default_layout
        return manifest

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
