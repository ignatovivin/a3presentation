# A3 Presentation

Service for turning text, markdown, and DOCX documents into branded PowerPoint presentations.

## Branches

- `dev`: active development branch
- `test`: pre-release verification branch
- `main`: stable branch

Recommended flow:

1. Work in `dev`
2. Move verified changes into `test`
3. Promote stable changes into `main`

## What is implemented

- FastAPI backend for extraction, planning, and PPTX generation
- React + Vite frontend for template upload, analysis, generation, and download
- Semantic pipeline: extract -> classify -> normalize -> plan -> render
- Fallback logic for weakly structured, form-like, resume, and table-heavy documents
- Template registry in `storage/templates`
- Regression and end-to-end backend tests

## Project structure

```text
frontend/              React 19 + Vite UI
src/a3presentation/
  api/                 HTTP routes
  domain/              Pydantic models
  services/            Extractor, normalizer, planner, PPTX generator
  settings.py          Paths and app configuration
storage/
  templates/           Versioned template manifests and source PPTX files
  generated/           Local generated artifacts, not tracked by git
  outputs/             Runtime output files, not tracked by git
tests/                 Unit, regression, and end-to-end tests
examples/              Example manifests and presentation plans
```

Storage rules are documented in [storage/README.md](/C:/Project/a3presentation/storage/README.md).

## Local run

Backend:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -e .
uvicorn a3presentation.main:app --reload
```

Frontend:

```bash
cd frontend
yarn install
yarn dev --host 127.0.0.1 --port 5173
```

Useful frontend commands:

```bash
yarn typecheck
yarn build
yarn verify
```

## Environment requirements

- Python 3.11+
- Node.js 20.19+ or 22.12+
- Yarn 1.x

## Vercel deployment

This repository is prepared for deploying the frontend to Vercel from the repo root.

Required Vercel environment variable:

```bash
VITE_API_BASE_URL=https://your-backend-host.example.com
```

Important:

- the current FastAPI backend is not Vercel-ready as-is
- backend code writes templates and generated files to local `storage/`
- for production, host the backend separately on a VM, container, or another platform with persistent filesystem access
- Vercel should be used for the frontend in the current architecture

Files added for Vercel:

- [vercel.json](/C:/Project/a3presentation/vercel.json)
- [frontend/.env.example](/C:/Project/a3presentation/frontend/.env.example)

## Test and verification

Backend tests:

```bash
python -m unittest discover -s tests -v
```

Frontend verification:

```bash
cd frontend
yarn verify
```

GitHub Actions runs both checks on pushes and pull requests for `dev`, `test`, and `main`.

## API entry points

- Backend: `http://127.0.0.1:8000`
- Health: `http://127.0.0.1:8000/health`
- Frontend: `http://127.0.0.1:5173`

Example calls:

```bash
curl http://127.0.0.1:8000/templates
```

```bash
curl -X POST http://127.0.0.1:8000/plans/from-text ^
  -H "Content-Type: application/json" ^
  -d "{\"template_id\":\"demo_business\",\"title\":\"Demo\",\"raw_text\":\"Intro\n- point 1\n- point 2\"}"
```

```bash
curl -X POST http://127.0.0.1:8000/presentations/generate ^
  -H "Content-Type: application/json" ^
  -d @examples/sample_presentation_plan.json
```

## Template workflow

1. Put a template in `storage/templates/<template_id>/template.pptx`
2. Store or regenerate `manifest.json` for that template
3. Use real prototype slides with tags such as `{{title}}`, `{{subtitle}}`, `{{text}}`, `{{bullets}}`
4. Let the analyzer build layout metadata from the PowerPoint file
5. Generate a plan from source text or send a prepared plan directly
6. Render the final `.pptx`

## Current focus

- hardening the semantic and planning pipeline
- expanding regression coverage for real documents
- improving layout safety for dense content
- keeping generated artifacts out of the repository
