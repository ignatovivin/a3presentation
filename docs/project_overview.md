# A3 Presentation Project Overview

## Purpose

This project is a focused service for turning structured business documents into branded PowerPoint presentations.

The target workflow is:

1. User uploads a `docx` or other text document.
2. The system reads the document structure.
3. The system understands where the document has:
   - title
   - section headings
   - subheadings
   - paragraphs
   - bullet lists
   - tables
   - images
4. The planner converts that structure into a slide plan.
5. The PowerPoint generator renders the plan into a branded `.pptx` using the corporate PowerPoint template.

This is not a generic "any presentation style" tool.
It is intentionally focused on structured business decks, a controlled set of slide types, and a template-aware rendering pipeline.
The current built-in corporate template is important, but it must not be treated as a permanent constant because users and companies can upload their own templates.

## Stack

### Backend

- Python
- FastAPI
- Pydantic
- `python-pptx`
- `python-docx`
- `pypdf`

### Frontend

- React 19
- Vite
- TypeScript
- `shadcn/ui`
- Tailwind-based styling approach through `shadcn/ui`

### Storage

- Local filesystem
- Templates stored in `storage/templates`
- Generated presentations stored in `storage/outputs`
- Production outputs on Timeweb stored in `data/outputs`

## Main Project Structure

```text
src/a3presentation/
  api/                  FastAPI routes
  domain/               Pydantic models for API, templates, presentation plan
  services/             Core business logic
    chart_style.py
    deck_audit.py
    document_text_extractor.py
    layout_capacity.py
    planner.py
    pptx_generator.py
    semantic_normalizer.py
    table_chart_analyzer.py
    template_analyzer.py
    template_registry.py
  settings.py           Project paths

frontend/               React UI
docs/                   Internal documentation
storage/
  templates/            PowerPoint templates + manifests
  outputs/              Generated presentations
```

## Core Pipeline

### Current process links

The current backend and frontend process map is:

1. UI flow starts in [App.tsx](../frontend/src/App.tsx) and calls [api.ts](../frontend/src/api.ts).
2. Template registry and analysis are handled by [template_registry.py](../src/a3presentation/services/template_registry.py) and [template_analyzer.py](../src/a3presentation/services/template_analyzer.py).
3. Source documents are extracted by [document_text_extractor.py](../src/a3presentation/services/document_text_extractor.py) into text, ordered document blocks, tables, hyperlinks, and image blocks.
4. Extracted tables are assessed by [table_chart_analyzer.py](../src/a3presentation/services/table_chart_analyzer.py); chartable candidates can be passed back as chart overrides.
5. [planner.py](../src/a3presentation/services/planner.py) builds a `PresentationPlan` using source blocks, tables, chart overrides, and [semantic_normalizer.py](../src/a3presentation/services/semantic_normalizer.py).
6. [pptx_generator.py](../src/a3presentation/services/pptx_generator.py) renders the plan against the active `TemplateManifest` and source `.pptx`.
7. [deck_audit.py](../src/a3presentation/services/deck_audit.py) and [layout_capacity.py](../src/a3presentation/services/layout_capacity.py) keep planner/generator capacity, geometry, and mixed-content order contracts aligned.
8. [run_quality_contracts.py](../scripts/run_quality_contracts.py) is the curated quality gate for generated decks.
9. Production delivery is handled by GitHub Actions plus [deploy_server.sh](../scripts/deploy_server.sh) and [docker-compose.server.yml](../docker-compose.server.yml).

### 1. Document extraction

Main file:
- [document_text_extractor.py](../src/a3presentation/services/document_text_extractor.py)

The extractor reads input documents and converts them into structured blocks.

For `docx`, it preserves document order and emits:

- `title`
- `heading`
- `subheading`
- `paragraph`
- `list`
- `table`
- `image`

The important part is that the extractor does not just read plain text.
It tries to preserve document semantics:

- heading detection from Word styles
- list detection from:
  - list styles
  - numbering properties (`numPr`)
  - numbering inherited from style
  - hanging indent typical for Word lists
  - explicit list-like text prefixes
- tables extracted as structured rows and headers
- image blocks preserved for semantic planning

This means the system can distinguish:

- simple text paragraph
- actual bullet list
- actual table

instead of treating everything as raw text.

### 2. Planning slides

Main file:
- [planner.py](../src/a3presentation/services/planner.py)

Related files:
- [semantic_normalizer.py](../src/a3presentation/services/semantic_normalizer.py)
- [layout_capacity.py](../src/a3presentation/services/layout_capacity.py)
- [table_chart_analyzer.py](../src/a3presentation/services/table_chart_analyzer.py)

The planner converts document blocks into a `PresentationPlan`.

Key idea:
- not `block -> slide`
- but `section -> 1..N slides`

Sections are built primarily from headings and subheadings.
The planner tries to keep related content together on one slide when it fits.

Supported logical slide outcomes include:

- `cover`
- `text_full_width`
- `list_full_width`
- `table`
- `chart`
- `image_text`
- `cards_3`
- `contacts`

The API contract for this stage is [PresentationPlan](../src/a3presentation/domain/presentation.py).
The frontend mirror of that contract is [types.ts](../frontend/src/types.ts).

Important planner behavior that was implemented:

- first page of the document becomes the cover slide
- cover title is built from leading lines before the first real section
- normal lists now go to `list_full_width`
- blue-card list layouts are not used for ordinary bullet lists
- `cards_3` is restricted to very short, label-like items only
- long sections are split only when needed
- tiny text tails are not split into separate meaningless fragments
- mixed `paragraph -> list -> paragraph` order is preserved on the main planning path
- top-level headings are preserved instead of getting lost
- compact tables stay on one slide when possible
- larger tables are paginated across slides
- chart overrides can replace selected table slides with chart slides
- semantic image blocks can become image slides
- first real numbered section is protected from being swallowed by cover heuristics

### 3. Template resolution

Main files:
- [template_registry.py](../src/a3presentation/services/template_registry.py)
- [template_analyzer.py](../src/a3presentation/services/template_analyzer.py)
- [manifest.json](../storage/templates/corp_light_v1/manifest.json)

The project currently ships with one main built-in template:

- `corp_light_v1`

But the architecture is moving toward template-aware generation rather than hard-coding one fixed template forever.
Users and companies may upload their own templates, and analyzer/manifest metadata should drive generator and audit behavior for those templates too.

The user does not need to think in terms of internal layout mapping.
Instead, the system chooses among logical slide types and then resolves them against the active template.

The manifest defines which PowerPoint layouts correspond to which logical slide type.

For example:

- `cover`
- `text_full_width`
- `list_full_width`
- `table`
- `image_text`
- `cards_3`
- `list_with_icons`
- `contacts`

## Working rules for analysis and implementation

This project should be evolved under these constraints:

- fixes must target the general mechanism, not one document or one slide
- every significant planner/generator/audit change must be checked against document classes, not one regression case
- the active template must not be assumed to be permanently fixed
- if the next implementation step is safe and follows directly from the current task, it should be executed without stopping for redundant confirmation

## PowerPoint Generation

Main file:
- [pptx_generator.py](../src/a3presentation/services/pptx_generator.py)

Related files:
- [deck_audit.py](../src/a3presentation/services/deck_audit.py)
- [chart_style.py](../src/a3presentation/services/chart_style.py)
- [layout_capacity.py](../src/a3presentation/services/layout_capacity.py)

The generator renders the `PresentationPlan` into a real `.pptx`.

### Important architectural decision

The implementation moved away from unsafe manual slide cloning as the primary path.

The current stable mode is layout-based generation using PowerPoint masters and layouts.

This is important because `.pptx` is an Open XML package, and naive slide cloning can corrupt:

- relationships
- media references
- layout links
- internal package structure

The generator now validates output after save:

- checks ZIP integrity
- checks duplicate entries
- reopens the generated file through `python-pptx`

If the file is invalid, it is rejected instead of being silently returned.

## What Was Improved

### 1. Cover slide

Problem before:
- cover title was positioned incorrectly
- title and meta were treated as technical text
- cover typography did not match the branded style

Current behavior:
- first document page is mapped to the cover
- title is built from leading document lines
- cover title is rendered large
- cover body/meta is rendered as the main supporting block, not as a footer
- text color is adjusted for the dark blue cover background

### 2. Tables

Problem before:
- tables were detached from document context
- table titles were wrong
- tables often overflowed
- rows were stretched too much
- some tables were split unnecessarily

Current behavior:
- tables are extracted in document order
- tables stay attached to the nearest logical section
- table widths are computed more intelligently
- row heights are reduced instead of stretching to full placeholder height
- compact tables stay on one slide
- larger tables split across multiple slides when truly necessary
- table colors are aligned with the corporate palette
- footer and content width on table slides are normalized to the active layout contract

### 3. Charts

Current behavior:

- chartable tables can be promoted to real chart slides
- column, line, stacked, pie, and combo scenarios are covered
- rendered charts are validated in regression and contract tests
- chart slides use the same layout-quality contract as the rest of the deck

### 4. Bullet lists

Problem before:
- real lists were sometimes treated as plain text
- some short lists turned into blue cards unexpectedly
- PowerPoint slides showed separate lines but not real bullet markers

Current behavior:
- extractor identifies lists from Word structure
- planner routes ordinary lists to `list_full_width`
- planner uses ordered block sections from source documents when available instead of flattening mixed content too early
- `cards_3` is reserved for very short card-like content only
- generator now writes real PowerPoint bullet markers (`buChar`) into paragraph XML
- continuation slides are rebalanced to avoid obviously underfilled tails
- deck audit checks rendered bullet/text order against the planned content order

### 5. Text slides

Problem before:
- normal text sometimes went into image-based layouts
- narrow text columns looked bad
- some slides contained placeholder garbage from PowerPoint

Current behavior:
- text-only content uses dedicated wide layouts
- unused placeholders are cleared
- the system avoids leaving PowerPoint instructional placeholder text
- planner and generator now share a common capacity-contract layer for text and bullets

### 6. Images

Current behavior:

- semantic image blocks can be rendered as image slides
- cover-skip heuristics no longer swallow numbered first sections that include images
- image layouts are now covered by deck-level quality contracts

## Frontend

Main frontend files:
- [App.tsx](../frontend/src/App.tsx)
- [api.ts](../frontend/src/api.ts)

Frontend responsibilities:

- upload source document
- upload/refresh template if needed
- trigger plan creation
- trigger PPTX generation
- download result
- review extracted structure before generation
- switch chartable tables between `table` and `chart`

UI was simplified from a technical admin-like screen into a more user-oriented flow:

1. upload template or use the system template
2. enter or upload source text/document
3. generate presentation
4. download `.pptx`

## Current State

The project is now in a working state for the main scenarios:

- input: business `docx`, markdown, and mixed text
- output: branded `.pptx`
- style: `corp_light_v1`

The most important end-to-end capabilities are working:

- cover generation
- heading-based section planning
- bullet list recognition
- table extraction and rendering
- chart generation
- image slide generation
- PowerPoint file validity
- frontend smoke and visual checks
- backend quality-contract suite for deck-level layout safety
- production deployment on Timeweb with Let's Encrypt HTTPS
- GitHub Actions SSH auto-deploy for `dev`

## Current Verification Layers

- unit and regression backend tests
- API contract tests
- planner/generator compatibility tests
- deck-level quality-contract tests
- dedicated `quality-contracts` runner
- frontend `playwright` smoke and visual checks

## Production Runtime

Current production topology:

- host nginx on Ubuntu terminates HTTPS for `a3presentation.ru`
- host nginx proxies to docker nginx on `127.0.0.1:8080`
- docker nginx routes `/` to frontend and `/api/*` to backend
- backend reads bundled templates from `/app/storage/templates`
- runtime outputs are persisted in `data/outputs`
- pushes to `dev` can auto-deploy through GitHub Actions after all checks pass

## Remaining Improvements

These are no longer core blockers, but they can improve quality further:

- extend deck-audit further for more layout-specific geometry rules
- expand visual snapshots for more frontend and generated-slide scenarios
- introduce template-specific typography rules per layout
- add export previews

## Summary

The project is a document-to-presentation engine built around one branded PowerPoint style.

Its main architecture is:

`docx -> structured blocks -> slide plan -> branded pptx`

The key work completed during implementation was not just "generate slides", but to make the system understand:

- what is a title
- what is a section
- what is a paragraph
- what is a real list
- what is a table
- which corporate slide type should be used for each case

That is why the current result is much closer to a usable product than the initial MVP.
