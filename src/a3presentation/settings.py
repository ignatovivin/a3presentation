from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class Settings:
    project_root: Path
    storage_dir: Path
    templates_dir: Path
    outputs_dir: Path


def get_settings() -> Settings:
    project_root = Path(__file__).resolve().parents[2]
    storage_dir = project_root / "storage"
    templates_dir = storage_dir / "templates"
    outputs_dir = storage_dir / "outputs"
    return Settings(
        project_root=project_root,
        storage_dir=storage_dir,
        templates_dir=templates_dir,
        outputs_dir=outputs_dir,
    )
