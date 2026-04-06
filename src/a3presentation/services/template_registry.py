from __future__ import annotations

import json
from pathlib import Path

from a3presentation.domain.template import TemplateManifest


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
        return TemplateManifest.model_validate(payload)

    def _template_dir(self, template_id: str) -> Path:
        return self._safe_child_path(self._templates_dir, template_id)

    def _safe_child_path(self, base_dir: Path, child_name: str) -> Path:
        if not child_name or not child_name.strip():
            raise ValueError("Path segment must not be empty")
        if Path(child_name).is_absolute():
            raise ValueError("Absolute paths are not allowed")

        candidate = (base_dir / child_name).resolve()
        try:
            candidate.relative_to(base_dir)
        except ValueError as exc:
            raise ValueError(f"Path '{child_name}' escapes the storage root") from exc
        return candidate
