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
- Deck-level quality contracts for capacity, geometry, and mixed-content order

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

Storage rules are documented in [storage/README.md](storage/README.md).
Analysis and troubleshooting workflow is documented in [docs/analysis_rule.md](docs/analysis_rule.md).
Document-class quality expectations are documented in [docs/document_class_matrix.md](docs/document_class_matrix.md).

## Process Map

The current process links are:

1. `frontend/src/App.tsx` calls API helpers from `frontend/src/api.ts`.
2. Template flow: `/templates`, `/templates/auto`, and `/templates/{template_id}/analyze` use `TemplateRegistry` and `TemplateAnalyzer` to persist or refresh `TemplateManifest`.
3. Extraction flow: `/documents/extract-text` uses `DocumentTextExtractor` to produce plain text, ordered `DocumentBlock` values, tables, images, and table metadata.
4. Chart-assessment flow: extracted tables pass through `TableChartAnalyzer`; the UI can keep a table slide or send a chart override. Chart candidates now reject unsafe ordinal/text tables, avoid default mixed-unit combo suggestions, and are covered by API, preview, generator XML, and deck-audit checks.
5. Planning flow: `/plans/from-text` validates `template_id` through `TemplateRegistry`, then `TextToPlanService` uses `SemanticDocumentNormalizer`, source blocks, tables, and chart overrides to build `PresentationPlan`.
6. Rendering flow: `/presentations/generate` resolves the active `TemplateManifest` and source `.pptx`, then `PptxGenerator` renders the deck.
7. Quality flow: planner, generator, and `deck_audit` share capacity/order expectations through `layout_capacity`; `scripts/run_quality_contracts.py` runs the curated deck-level gate.
   Chart audit also validates rendered chart type, series count, combo structure, title/subtitle font sizes, and compact value-axis number formats.
8. Delivery flow: generated files are saved under runtime outputs and downloaded through `/presentations/files/{file_name}`.

Working standard:

- before meaningful analysis or implementation, project rule documents must be checked and applied without waiting for an explicit reminder
- fixes must be global, not tuned to one file or one slide
- after identifying one broken manifestation, similar manifestations across related components and layers must also be checked and fixed if they share the same mechanism
- current template must not be treated as a constant because users and companies can upload their own templates
- after analysis, the next safe implementation step should be executed automatically without asking for redundant confirmation

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

## Timeweb Cloud Server

The current production setup runs on a single Timeweb VPS.

Guide:

- [docs/timeweb-server-deploy.md](docs/timeweb-server-deploy.md)

Main command on the server:

```bash
bash scripts/deploy_server.sh
```

On Timeweb server deploys, backend reads versioned templates directly from `/app/storage/templates` inside the image.
Only runtime outputs stay on the persistent host volume.
The server compose publishes docker nginx on `127.0.0.1:8080`, because public `80/443` are handled by host nginx with Let's Encrypt.
Production domain:

- `https://a3presentation.ru`
- `https://www.a3presentation.ru`

Current production deployment flow:

1. push verified code into `dev`
2. GitHub Actions runs backend, quality-contracts, frontend verify, and frontend smoke
3. if all checks pass, `deploy-timeweb` connects to the server over SSH
4. server checkout is hard-reset to `origin/dev`
5. `bash scripts/deploy_server.sh` recreates the full docker app stack

Current production checks:

- `https://a3presentation.ru/api/health`
- `https://a3presentation.ru/api/templates`

## Test and verification

Backend tests:

```bash
python -m unittest discover -s tests -v
```

Quality contracts:

```bash
.venv\Scripts\python.exe scripts/run_quality_contracts.py
```

Quality-contract layer is documented in [docs/quality-contracts.md](docs/quality-contracts.md).
It now validates not only capacity and geometry, but also rendered content order for mixed text/list continuations.

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

GitHub Actions runs backend tests, frontend verification, and the dedicated quality-contract gate on pushes and pull requests for `dev`, `test`, and `main`.
Frontend smoke tests are also suitable as a separate CI gate.
For the actual production server, pushes to `dev` trigger an SSH deploy job to Timeweb after all checks pass.

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
- expanding regression coverage for real documents and document classes
- improving layout safety for dense content and generated decks
- keeping planner, generator, deck audit, and production deploy contracts aligned
- preserving mixed-content order across extractor, planner, generator, and deck audit
- keeping generated artifacts out of the repository
