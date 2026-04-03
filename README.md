# A3 Presentation

MVP backend for generating `.pptx` presentations from text using reusable PowerPoint templates.

## What is implemented

- FastAPI service with health and generation endpoints
- Domain models for template manifests and presentation plans
- Local template registry stored in `storage/templates`
- Template upload endpoint for `.pptx` plus manifest
- Auto-upload endpoint that stores `.pptx` and analyzes it into a draft manifest
- Basic `.pptx` generator built on `python-pptx`
- Example template manifest and example presentation plan
- React + Vite + shadcn/ui frontend for template upload, analysis, plan generation and download

## Project structure

```text
frontend/              React 19.2 + Vite 8 UI
src/a3presentation/
  api/                HTTP routes
  domain/             Pydantic models
  services/           Template registry, planning, PPTX generation
  settings.py         Paths and app configuration
storage/
  templates/          Uploaded templates and manifests
  outputs/            Generated presentations
examples/
  sample_template_manifest.json
  sample_presentation_plan.json
```

## MVP flow

1. Add a template folder into `storage/templates/<template_id>/`
2. Place original `.pptx` as `template.pptx`
3. Put prototype slides into the template PPTX with text tags like `{{title}}`, `{{subtitle}}`, `{{bullets}}`, `{{bullet_1}}`
4. Analyze the template to build `manifest.json`
5. Send text or a structured plan to the API
6. Service clones matching prototype slides and replaces tags inside real designer blocks

## Run

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
yarn dev
```

Useful yarn commands:

```bash
yarn typecheck
yarn build
yarn verify
```

Requirements for frontend:

- Node.js 20.19+ or 22.12+ because Vite 8 requires it according to the official migration guide: https://vite.dev/guide/migration.html
- `shadcn/ui` is configured manually for the existing Vite app using the official setup approach: https://ui.shadcn.com/docs/installation/vite

## First API calls

List templates:

```bash
curl http://127.0.0.1:8000/templates
```

Open the UI:

```text
http://127.0.0.1:5173
```

UI flow now includes:

- upload `.pptx` template from the browser
- analyze the selected template in a dialog
- inspect layouts in a table
- generate and download the final `.pptx`

## Prototype-based templates

The recommended way to build real templates is no longer "empty layout + placeholders".
Use real slides inside your `.pptx` as prototypes and put tags directly into text blocks:

```text
{{title}}
{{subtitle}}
{{text}}
{{bullets}}
{{bullet_1}}
{{bullet_2}}
{{left_bullets}}
{{right_bullets}}
```

How it works now:

1. You create real styled slides in PowerPoint.
2. You insert tags into the existing text boxes.
3. The analyzer reads those slides and stores them as `prototype_slides`.
4. The generator clones the styled source slide instead of creating an empty layout slide.
5. Tags are replaced with actual content.

This is the mode that preserves the visual composition of the original template.

Create a plan from raw text:

```bash
curl -X POST http://127.0.0.1:8000/plans/from-text ^
  -H "Content-Type: application/json" ^
  -d "{\"template_id\":\"demo_business\",\"title\":\"Demo\",\"raw_text\":\"Intro\n- point 1\n- point 2\"}"
```

Generate from a structured plan:

```bash
curl -X POST http://127.0.0.1:8000/presentations/generate ^
  -H "Content-Type: application/json" ^
  -d @examples/sample_presentation_plan.json
```

Upload a real template:

```bash
curl -X POST http://127.0.0.1:8000/templates ^
  -F "manifest_json={\"template_id\":\"corp_blue_v1\",\"display_name\":\"Corporate Blue\",\"source_pptx\":\"template.pptx\",\"default_layout_key\":\"content\",\"layouts\":[{\"key\":\"title\",\"name\":\"Title Slide\",\"slide_layout_index\":0,\"supported_slide_kinds\":[\"title\"],\"placeholders\":[{\"name\":\"title\",\"kind\":\"title\",\"idx\":0},{\"name\":\"subtitle\",\"kind\":\"subtitle\",\"idx\":1}]},{\"key\":\"content\",\"name\":\"Content\",\"slide_layout_index\":1,\"supported_slide_kinds\":[\"bullets\",\"text\",\"table\"],\"placeholders\":[{\"name\":\"title\",\"kind\":\"title\",\"idx\":0},{\"name\":\"body\",\"kind\":\"body\",\"idx\":1}]}]}" ^
  -F "template_file=@C:\path\to\template.pptx"
```

Analyze the uploaded `.pptx` and rebuild `manifest.json` from its layouts:

```bash
curl -X POST "http://127.0.0.1:8000/templates/corp_blue_v1/analyze?display_name=Corporate%20Blue"
```

## Recommended next implementation order

1. Add preview generation for layouts so the user can compare template variants visually.
2. Replace the current simple text splitter with an LLM pipeline that emits `PresentationPlan`.
3. Extend generator support for images, tables and charts with layout-specific fill rules.
4. Add manual editing for analyzed manifests because placeholder classification from `.pptx` is only a first-pass heuristic.

## Next steps

- Add layout preview generation from real `.pptx`
- Add LLM-based text to slide-plan transformation
- Add image/table/chart support beyond text placeholders
- Add preview rendering to PDF or PNG
