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
- Chart and image slide support in the planning and rendering pipeline
- Template registry in `storage/templates`
- Regression, contract, and quality backend tests
- Frontend smoke and visual checks with Playwright

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

## Railway deployment

Recommended production target: Railway.

Deploy the project as two services inside one Railway project:

- backend service from repo root `/`
- frontend service from `/frontend`

Deployment guide:

- [docs/railway-deploy.md](/C:/Project/a3presentation/docs/railway-deploy.md)

Important production variables:

```bash
TEMPLATES_DIR=/data/templates
OUTPUTS_DIR=/data/outputs
CORS_ORIGINS=https://your-frontend-domain.up.railway.app
VITE_API_BASE_URL=https://your-backend-domain.up.railway.app
```

Railway-specific notes:

- attach a persistent volume to the backend at `/data`
- keep frontend and backend as separate Railway services
- this is the cleanest way to keep the whole product in one Railway project with persistent file storage

## Timeweb Cloud Server

If you want the whole system on one VPS, use the server deployment mode.

Guide:

- [docs/timeweb-server-deploy.md](/C:/Project/a3presentation/docs/timeweb-server-deploy.md)

Main command on the server:

```bash
docker compose -f docker-compose.server.yml up -d --build
```

## Test and verification

Backend tests:

```bash
python -m unittest discover -s tests -v
```

Quality contracts:

```bash
.venv\Scripts\python.exe scripts/run_quality_contracts.py
```

Quality-contract layer is documented in [docs/quality-contracts.md](/C:/Project/a3presentation/docs/quality-contracts.md).

Frontend verification:

```bash
cd frontend
yarn verify
```

Frontend smoke and visual checks:

```bash
cd frontend
yarn test:smoke
yarn test:visual
```

GitHub Actions should run backend tests, frontend verification, and the dedicated quality-contract gate on pushes and pull requests for `dev`, `test`, and `main`.
Frontend smoke tests are also suitable as a separate CI gate.

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
- keeping planner, generator, and quality contracts aligned
- keeping generated artifacts out of the repository
