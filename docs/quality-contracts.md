# Quality Contracts

`quality-contracts` is a dedicated verification layer for generated decks.

It is narrower than the full backend test suite and focuses on end-to-end layout quality contracts:

- text and bullets capacity
- continuation balance
- table layout geometry
- chart layout geometry
- image layout geometry
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

- [run_quality_contracts.py](/C:/Project/a3presentation/scripts/run_quality_contracts.py)

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
- before moving changes from `dev` to `test`
- before release verification

## Scope Boundary

This layer is intended to answer:

- did the deck render with the expected layout contracts?
- did planner and generator stay aligned on slide capacity rules?
- did representative document classes remain stable?

It is not intended to replace:

- the full backend suite
- API contract tests
- frontend smoke and visual tests

Those layers should continue to run alongside `quality-contracts`, not instead of it.
