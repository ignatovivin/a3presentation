# Release Gate Roadmap

## Goal

Bring the project to a state where users can reliably:

1. upload a `.docx`
2. select a registry template or upload a `.pptx`
3. review chart and layout choices
4. generate a `.pptx`
5. download the result

without needing engineering intervention for normal flows.

## Template Policy

- the release flow must not assume that one registry template defines the correct behavior of the system
- a reference template may still exist for stable smoke checks, but it is not the product contract
- the real contract is generic PowerPoint extraction and generation for arbitrary `.pptx` files
- release hardening should move from template-id specific checks toward invariant-based checks:
  text-capable path, chart/table/image-capable path, geometry metadata, safe degradation, no fatal missing-shape failures

## Priorities

### P1

- Keep a stable release gate with mandatory checks.
- Show user-friendly frontend errors for extraction, planning, generation, and quality gate failures.
- Keep runtime smoke green without manual backend startup.
- Bring visual regression back under control.

### P2

- Reduce technical wording in template analysis and slide review UI.
- Improve explanation of why a layout/prototype option is recommended.
- Review mobile and narrow-screen behavior.

### P3

- Expand regression corpus with more real-world user files.
- Add more mixed-template and mixed-layout ranking cases.
- Shift template coverage from reference-template checks to generic PowerPoint component invariants.
- Further harden test/runtime infrastructure for CI parallelism.
- Document Playwright/runtime limitations and fixed port policy.
- Keep dedicated runtime smoke serial and isolated from shared backends.

## Mandatory Release Gate

Run these checks before shipping:

1. `.\scripts\release-gate.ps1`
2. If UI changed, run `.\scripts\release-gate.ps1 -IncludeVisual` and review snapshot diffs explicitly.
3. If runtime/backend config changed, `.\scripts\release-gate.ps1` already reruns the dedicated runtime smoke; do not skip it.

## Current Status

### Done

- Functional upload/review/generate frontend flow is covered.
- Runtime Playwright smoke can start backend automatically.
- API review metadata is covered by tests.
- Layout vs prototype UI labeling is clearer.

### In Progress

- Visual regression baseline approval.
- Broader optional user-template regression sweep.

### Remaining

- Expand corpus and edge-case ranking coverage with more real user files when they are available in-repo.
- Reduce reliance of mandatory checks on a single reference template as generic uploaded-template corpus grows.

## Runtime/Test Guardrails

- Manual frontend dev uses `127.0.0.1:5173`; Playwright uses `127.0.0.1:4173` by default.
- Playwright backend uses `127.0.0.1:8000` by default.
- Override ports only through `PLAYWRIGHT_FRONTEND_PORT` and `PLAYWRIGHT_BACKEND_PORT`.
- Frontend/backend Playwright ports must be different; config should fail fast otherwise.
- Runtime smoke must stay local-only and must not reuse production/staging API origins.
- Use `cd frontend && yarn test:runtime` for runtime-specific reruns; keep it serial.
- `.\scripts\release-gate.ps1` is the mandatory local pre-release gate.
- Visual approval is opt-in and separate: `.\scripts\release-gate.ps1 -IncludeVisual`.
