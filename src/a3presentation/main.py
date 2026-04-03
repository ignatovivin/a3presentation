from __future__ import annotations

import shutil

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from a3presentation.api.routes import router
from a3presentation.settings import get_settings


def _seed_templates_if_needed() -> None:
    settings = get_settings()
    source_root = settings.bundled_templates_dir
    destination_root = settings.templates_dir

    if source_root == destination_root or not source_root.exists():
        return

    for template_dir in source_root.iterdir():
        if not template_dir.is_dir():
            continue
        destination_dir = destination_root / template_dir.name
        if destination_dir.exists():
            continue
        shutil.copytree(template_dir, destination_dir)


def create_app() -> FastAPI:
    settings = get_settings()
    settings.templates_dir.mkdir(parents=True, exist_ok=True)
    settings.outputs_dir.mkdir(parents=True, exist_ok=True)
    _seed_templates_if_needed()

    app = FastAPI(title="A3 Presentation API", version="0.1.0")
    app.add_middleware(
        CORSMiddleware,
        allow_origins=list(settings.cors_origins),
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )
    app.include_router(router)
    return app


app = create_app()
