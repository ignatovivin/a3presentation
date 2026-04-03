from __future__ import annotations

from dataclasses import dataclass
import os
from pathlib import Path


@dataclass(frozen=True)
class Settings:
    project_root: Path
    storage_dir: Path
    templates_dir: Path
    outputs_dir: Path
    bundled_templates_dir: Path
    cors_origins: tuple[str, ...]


def _parse_cors_origins() -> tuple[str, ...]:
    raw_value = os.getenv("CORS_ORIGINS", "").strip()
    if raw_value:
        origins = tuple(origin.strip() for origin in raw_value.split(",") if origin.strip())
        if origins:
            return origins

    return (
        "http://127.0.0.1:5173",
        "http://localhost:5173",
    )


def get_settings() -> Settings:
    project_root = Path(__file__).resolve().parents[2]
    storage_dir = Path(os.getenv("STORAGE_DIR", project_root / "storage")).resolve()
    templates_dir = Path(os.getenv("TEMPLATES_DIR", storage_dir / "templates")).resolve()
    outputs_dir = Path(os.getenv("OUTPUTS_DIR", storage_dir / "outputs")).resolve()
    return Settings(
        project_root=project_root,
        storage_dir=storage_dir,
        templates_dir=templates_dir,
        outputs_dir=outputs_dir,
        bundled_templates_dir=(project_root / "storage" / "templates").resolve(),
        cors_origins=_parse_cors_origins(),
    )
