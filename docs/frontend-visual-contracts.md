# Frontend And Visual Contracts

## Current State

- Frontend runtime contract includes `yarn verify`
- The project includes `playwright` smoke and visual tests
- UI and visual validation are partially automated and still need broader scenario coverage
- backend deck-level quality checks are documented separately in [quality-contracts.md](quality-contracts.md)
- backend quality layer now also validates mixed-content order inside generated deck body containers
- `frontend smoke` is suitable for CI today
- `frontend visual` should stay a separate gate until stable cross-platform baselines are introduced
- chart preview smoke now covers the supported chart layout matrix, including line marker/label geometry and invalid-coordinate guards

## Required Frontend Smoke Flow

The minimal automated UI flow should verify:

1. app bootstraps and loads templates from backend
2. user can upload a `.docx` document
3. extracted text, tables, and chart assessments appear in UI state
4. user can open the structure drawer
5. user can switch a chartable table between `table` and `chart`
6. user can generate a presentation
7. user can receive a downloadable result link
8. user can dismiss success and error panels
9. chart type select stays aligned with backend candidates and does not expose unsupported defaults such as mixed-unit combo
10. hidden chart series are preserved in the `chart_overrides` request payload
11. drawer switch/select controls remain accessible by role and visible to Playwright

## Required Visual Regression Set

The minimal visual regression layer should cover:

1. cover slide
2. dense text slide
3. bullets slide
4. compact table slide
5. wide table slide
6. chart slide
7. image slide
8. footer positioning on long-title slides

## Chart Preview Smoke Matrix

The current automated chart preview smoke should cover:

- `column`
- `bar`
- `line`
- `stacked_column`
- `stacked_bar`
- `pie`
- explicit `combo` specs for legacy/generator parity
- dense line categories with large compact values
- negative values without `NaN` marker coordinates
- hidden-series behavior in the plan payload

## Suggested Tooling

- UI smoke: `playwright`
- visual snapshots: `playwright` screenshots or equivalent browser snapshot layer
- CI entry points:
  - `yarn verify`
  - `yarn test:smoke`
  - `yarn test:visual` after cross-platform snapshot policy is fixed

## Readiness Gate

Frontend and visual contract layer can be considered ready when:

- smoke flow is automated end-to-end against local backend
- at least one golden snapshot exists for each required visual scenario
- regressions fail CI before manual QA
- frontend gate stays aligned with backend quality-contract gate
