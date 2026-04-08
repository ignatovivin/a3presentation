# Quality Contracts

`quality-contracts` is a dedicated verification layer for generated decks.

It is narrower than the full backend test suite and focuses on end-to-end layout quality contracts:

- text and bullets capacity
- continuation balance
- mixed-content order preservation
- table layout geometry
- chart layout geometry
- chart semantic contracts:
  - rendered chart type
  - rendered series count
  - combo bar/line structure
  - title/subtitle font sizes
  - value-axis number format for compact large-number display
- image layout geometry
- template-aware behavior for uploaded or analyzer-derived templates
- representative document classes:
  - text-only
  - mixed text
  - report-like
  - strategy-heavy
  - form-like
  - resume-like
  - table-heavy
  - chart-heavy
  - image-heavy
  - fact-only

Current runner:

- [run_quality_contracts.py](../scripts/run_quality_contracts.py)

## Why a dedicated runner

The project already uses `unittest` everywhere.

A dedicated runner is the lowest-friction option because it:

- reuses the current test stack
- avoids introducing another task runner or marker system
- gives a stable command for local checks and CI gates
- keeps the quality gate intentionally curated instead of coupling it to every backend test

## Run

From the repo root:

```bash
.venv\Scripts\python.exe scripts/run_quality_contracts.py
```

## When to run

Recommended cases:

- after changes in `planner`
- after changes in `pptx_generator`
- after layout/template changes
- after changes in template analyzer or template-aware geometry logic
- before moving changes from `dev` to `test`
- before release verification

## Scope Boundary

This layer is intended to answer:

- did the deck render with the expected layout contracts?
- did planner and generator stay aligned on slide capacity rules?
- did mixed paragraph/list content keep its expected order after rendering?
- did representative document classes remain stable?
- did the change remain valid beyond one built-in template or one regression document?

It is not intended to replace:

- the full backend suite
- API contract tests
- frontend smoke and visual tests

Those layers should continue to run alongside `quality-contracts`, not instead of it.

## Chart-Specific Contracts

Current chart behavior:

- chartable tables can be promoted to real chart slides through explicit chart overrides
- default chart candidates include column, line, bar, stacked column/bar, and pie where appropriate
- combo remains supported by the generator for explicit specs and legacy plans, but is not offered as the default mixed-unit UI option
- unsafe mixed-unit tables with too many unit families are rejected as not chartable
- ordinal/status tables with `1..N` index-like values are rejected as not chartable
- summary rows such as `Итого` are filtered out before chart series are built
- column, bar, line, stacked, pie, and explicit combo scenarios are covered by generator XML tests
- chart slides use the same layout-quality contract as the rest of the deck
- chart value axes are audited for compact number formats such as `млн`, `млрд`, `%`, and `₽`
- chart title/subtitle sizes are audited against the shared `28 pt` / `18 pt` contract

Frontend-adjacent checks:

- structure drawer smoke verifies chart/table mode switching
- chart type select no longer exposes combo for default mixed-unit candidates
- hidden series are preserved in the `chart_overrides` payload
- preview smoke covers column, bar, line, stacked column, stacked bar, pie, and explicit combo render paths
- line preview smoke checks marker/line/label counts and guards against invalid `NaN` coordinates

## Execution rule

This verification layer should support the general working standard of the project:

- fixes are accepted as global only if they survive document-class and template-aware checks
- one problematic file is a regression case, not proof that a local tweak is acceptable
- after fixing one manifestation, related manifestations in neighboring components, layouts, or render paths should also be checked and covered by verification where applicable
- verification after changes should include targeted regressions, representative document classes, and the broader backend suite rather than a single happy-path rerun
- generation-related fixes must include a fresh `.pptx` generation after the code change and direct inspection of the newly generated artifact; inspecting an old deck is diagnostic only
- after identifying a safe next verification or implementation step, work should continue automatically without redundant stop-and-ask behavior
