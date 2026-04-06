# System Backlog

## Purpose

This backlog tracks the major unfinished system tasks that still matter after the current `content_blocks`, planner, generator, and deck-audit improvements.

It must be read together with:

- [analysis_rule.md](/C:/Project/a3presentation/docs/analysis_rule.md)
- [document_class_matrix.md](/C:/Project/a3presentation/docs/document_class_matrix.md)

The backlog is ordered by architectural importance, not by one problematic document.
It also assumes that the current built-in template is not a permanent constant and that future company templates must be supported without local tuning.

## 1. Placeholder-level capacity model

Status:

- partially done

Current state:

- `LayoutCapacityProfile` exists in [layout_capacity.py](/C:/Project/a3presentation/src/a3presentation/services/layout_capacity.py)
- profiles are still coarse and keyed only by `layout_key`
- `TemplateAnalyzer` now extracts placeholder geometry and text-frame margins from real `.pptx` templates
- uploaded templates can carry shape metadata through analyzer/manifests, but generator/audit still rely mostly on curated layout policies

Why it matters:

- different placeholders inside one layout may need different fill targets and font bounds
- coarse layout-level limits force planner and generator to rely on broad heuristics
- future company templates cannot depend on `corp_light_v1` geometry being treated as a constant

Target:

- capacity policy at least at `layout + placeholder kind`, and preferably `layout + placeholder idx`
- template-aware derivation path where analyzer metadata can override or bootstrap generator/audit geometry for uploaded templates

## 2. Continuation packing stability

Status:

- partially done

Current state:

- continuation rebalance exists in [planner.py](/C:/Project/a3presentation/src/a3presentation/services/planner.py)
- dense narrative sections still show underfilled continuation tails

Why it matters:

- this is the main remaining source of quality regressions in narrative and mixed documents

Target:

- stable balancing for text-only and mixed continuation groups across document classes

## 3. Deck-audit invariant expansion

Status:

- partially done

Current state:

- [deck_audit.py](/C:/Project/a3presentation/src/a3presentation/services/deck_audit.py) already checks:
  - font bounds
  - overflow risk
  - continuation balance
  - underfilled continuation
  - content order
  - table/chart/image geometry

Missing invariants:

- font-size delta between neighboring continuation slides
- title/body gap consistency
- broader footer-width checks outside table/chart cases
- placeholder-level fill targets

## 4. Semantic false-positive control

Status:

- partially done

Current state:

- false facts and false contacts were reduced in [semantic_normalizer.py](/C:/Project/a3presentation/src/a3presentation/services/semantic_normalizer.py)

Why it matters:

- report documents must not turn appendix/fallback logic into garbage slides
- form/resume documents must still keep useful fact extraction

Target:

- stronger separation between narrative prose and real field-value content

## 5. Corpus expansion

Status:

- partially done

Current state:

- regression corpus already covers report, mixed, form, resume, table-heavy, chart-heavy, image-heavy, markdown, fact-only

Missing:

- more dense narrative reports
- more appendix/source-heavy reports
- more question/callout-heavy documents
- more layout stress cases with long titles and compact body pairs

## 6. PPTX quality-layer expansion

Status:

- not done

Current state:

- quality contracts exist
- coverage is good, but not yet equivalent to a broader PPTX quality gate

Target:

- strengthen golden checks around:
  - dense bullets
  - continuation pairs
  - appendix/question slides
  - mixed paragraph+bullets
  - long title + dense body
  - compact vs underfilled neighbor slides

## Execution rule

When picking the next task:

1. prefer architecture-wide improvements over one-document tuning
2. add or update tests first when possible
3. validate against document classes, not just one regression file
4. continue automatically to the next safe step instead of stopping for redundant confirmation inside an already agreed task
