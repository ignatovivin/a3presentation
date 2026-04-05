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
It is intentionally narrow and optimized for a fixed corporate presentation style and a fixed set of slide types.

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

## Main Project Structure

```text
src/a3presentation/
  api/                  FastAPI routes
  domain/               Pydantic models for API, templates, presentation plan
  services/             Core business logic
    document_text_extractor.py
    planner.py
    pptx_generator.py
    template_registry.py
  settings.py           Project paths

frontend/               React UI
docs/                   Internal documentation
storage/
  templates/            PowerPoint templates + manifests
  outputs/              Generated presentations
```

## Core Pipeline

### 1. Document extraction

Main file:
- [document_text_extractor.py](C:\Project\a3presentation\src\a3presentation\services\document_text_extractor.py)

The extractor reads input documents and converts them into structured blocks.

For `docx`, it preserves document order and emits:

- `title`
- `heading`
- `subheading`
- `paragraph`
- `list`
- `table`

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
- [planner.py](C:\Project\a3presentation\src\a3presentation\services\planner.py)

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

Important planner behavior that was implemented:

- first page of the document becomes the cover slide
- cover title is built from leading lines before the first real section
- normal lists now go to `list_full_width`
- blue-card list layouts are not used for ordinary bullet lists
- `cards_3` is restricted to very short, label-like items only
- long sections are split only when needed
- tiny text tails are not split into separate meaningless fragments
- top-level headings are preserved instead of getting lost
- compact tables stay on one slide when possible
- larger tables are paginated across slides
- chart overrides can replace selected table slides with chart slides
- semantic image blocks can become image slides
- first real numbered section is protected from being swallowed by cover heuristics

### 3. Template resolution

Main files:
- [template_registry.py](C:\Project\a3presentation\src\a3presentation\services\template_registry.py)
- [manifest.json](C:\Project\a3presentation\storage\templates\corp_light_v1\manifest.json)

The project now works around a single main corporate template:

- `corp_light_v1`

The user does not need to think in terms of multiple template choices.
Instead, the system chooses among slide types inside one presentation style.

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

## PowerPoint Generation

Main file:
- [pptx_generator.py](C:\Project\a3presentation\src\a3presentation\services\pptx_generator.py)

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
- `cards_3` is reserved for very short card-like content only
- generator now writes real PowerPoint bullet markers (`buChar`) into paragraph XML
- continuation slides are rebalanced to avoid obviously underfilled tails

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
- [App.tsx](C:\Project\a3presentation\frontend\src\App.tsx)
- [api.ts](C:\Project\a3presentation\frontend\src\api.ts)

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

## Current Verification Layers

- unit and regression backend tests
- API contract tests
- planner/generator compatibility tests
- deck-level quality-contract tests
- dedicated `quality-contracts` runner
- frontend `playwright` smoke and visual checks

## Remaining Improvements

These are no longer core blockers, but they can improve quality further:

- move `quality-contracts` into a dedicated CI gate
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
