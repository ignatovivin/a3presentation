# Document Class Matrix

## Purpose

This matrix defines the main input document classes the project must support.
It exists to prevent overfitting the planner and generator to one problematic document.

The target is not "works for one file".
The target is "stable across document classes".

This document must be used together with [analysis_rule.md](/C:/Project/a3presentation/docs/analysis_rule.md).
It must also be used with the assumption that the active template may change, because users and companies can upload their own templates.

## Core rule

Every planner or generator change must be evaluated against document classes, not a single example.
The same applies to template behavior: one built-in template is not enough evidence for a globally correct fix.

For each class we care about:

- expected extractor behavior
- expected planner behavior
- expected PPTX rendering behavior
- expected deck-audit behavior
- expected behavior on non-default uploaded templates
- current regression coverage
- current risk

## Classes

### 1. Narrative report

Examples:

- strategy report
- long CEO memo
- market analysis
- methodology write-up

Expected behavior:

- headings become sections
- long prose stays paragraph flow, not bullet spam
- continuation slides stay balanced
- appendix/reference tail does not pollute the main deck

Current coverage:

- `tests/test_regression_corpus.py`
  - `test_report_docx_generates_deck_without_capacity_violations`
  - `test_report_docx_prefers_text_flow_for_narrative_sections`
  - `test_report_docx_skips_reference_tail_from_main_deck`
  - `test_report_docx_does_not_add_appendix_from_false_semantic_facts`

Current risk:

- continuation underfill on dense prose
- unstable split between text and bullets in mixed analytical sections

### 2. Mixed narrative + bullets

Examples:

- sections with paragraph -> list -> paragraph
- analytical documents with operational bullet inserts

Expected behavior:

- semantic order is preserved
- paragraph blocks stay paragraph blocks
- bullet lists stay bullet lists
- mixed sections render through `content_blocks`

Current coverage:

- `tests/test_planner.py`
  - `test_mixed_section_preserves_paragraph_list_paragraph_order`
  - `test_mixed_continuation_with_paragraph_dominance_stays_text_slide`
- `tests/test_regression_corpus.py`
  - `test_generic_mixed_section_rebalances_continuation_series_without_underfilled_tail`
  - `test_generic_mixed_section_keeps_block_order_through_planner_and_generator`

Current risk:

- mixed continuations can still flip between text and bullet layouts
- continuation rebalance can affect ordering if not anchored to `content_blocks`

### 3. Reference-tail report

Examples:

- report with bibliography
- report with raw URLs at the end

Expected behavior:

- raw reference lines do not enter the main deck
- URL-only lines are ignored or isolated from main narrative
- appendix is not auto-created from false facts

Current coverage:

- `tests/test_regression_corpus.py`
  - `test_report_docx_skips_reference_tail_from_main_deck`
  - `test_report_docx_does_not_add_appendix_from_false_semantic_facts`

Current risk:

- semantic false positives from `label: value`
- URL fragments mistaken for facts

### 4. Form-like document

Examples:

- анкета
- form with short fields and values

Expected behavior:

- safe fallback path is allowed
- compact fact-style content may produce appendix/summary
- deck should not break on sparse structure

Current coverage:

- `tests/test_regression_corpus.py`
  - `test_form_like_docx_generates_deck_without_capacity_violations`
- `tests/test_semantic_pipeline.py`
  - fact/contact/date extraction baseline

Current risk:

- overly aggressive anti-fact heuristics could damage true field/value extraction

### 5. Resume / CV

Examples:

- candidate profile
- CV with contacts, education, experience

Expected behavior:

- document classified as `resume`
- contacts/facts stay useful
- fallback planner produces compact profile slides

Current coverage:

- `tests/test_regression_corpus.py`
  - `test_resume_like_docx_generates_deck_without_capacity_violations`
- `tests/test_semantic_pipeline.py`
  - `test_normalizer_extracts_facts_contacts_dates_and_kind`

Current risk:

- report-oriented heuristics must not break resume summary extraction

### 6. Table-heavy document

Examples:

- KPI tables
- operational matrices

Expected behavior:

- tables stay attached to section context
- compact tables stay on one slide when possible
- split tables preserve geometry and footer width contracts

Current coverage:

- `tests/test_regression_corpus.py`
  - `test_table_heavy_docx_generates_deck_without_text_capacity_violations`
- `tests/test_project_contracts.py`
  - table geometry audit checks

Current risk:

- planner changes aimed at prose must not destabilize table flow

### 7. Chart-heavy document

Examples:

- table that should become chart
- business metrics deck

Expected behavior:

- chartable tables may become chart slides
- chart layout respects content width contracts
- chart style remains deterministic

Current coverage:

- `tests/test_regression_corpus.py`
  - `test_chart_heavy_docx_generates_chart_slide_and_preserves_text_capacity_contract`
- `tests/test_project_contracts.py`
  - chart geometry audit checks

Current risk:

- planner/content-block changes should not affect chart overrides or table-to-chart mapping

### 8. Image-heavy document

Examples:

- report with embedded scheme or illustration

Expected behavior:

- image blocks are preserved semantically
- image slides render without breaking surrounding narrative

Current coverage:

- `tests/test_regression_corpus.py`
  - `test_image_heavy_docx_generates_image_slide_and_preserves_text_capacity_contract`
  - `test_cover_skip_does_not_swallow_numbered_first_section_with_image`
- `tests/test_project_contracts.py`
  - image geometry audit checks

Current risk:

- cover heuristics and section-skip heuristics can swallow the first real image section

### 9. Markdown / plain text

Examples:

- markdown strategy memo
- mixed notes text file

Expected behavior:

- markdown headings and lists are preserved
- numbered markdown lists are not misclassified as headings
- text-only documents still produce stable decks

Current coverage:

- `tests/test_document_extractor.py`
  - markdown numbered-list classification
- `tests/test_regression_corpus.py`
  - markdown and mixed text fixtures

Current risk:

- plain-text heuristics are more fragile than DOCX structure and need dedicated regression coverage

## Deck-level quality criteria

These are the current universal quality criteria across classes.

- no invalid PPTX output
- no missing required rendered content
- content order must remain stable for mixed sections
- continuation groups should avoid severe underfill/overflow imbalance
- text font sizes must stay within layout profile
- table/chart/image slides must satisfy basic layout geometry

Current automated layer:

- `tests/test_project_contracts.py`
- `tests/test_regression_corpus.py`
- `src/a3presentation/services/deck_audit.py`

## Current gaps

The most important remaining gaps are:

- placeholder-specific capacity rules are still coarse
- continuation packing is still not stable enough on dense narrative documents
- quality checks do not yet cover the full set of desired invariants like:
  - font delta between neighbor continuation slides
  - title/body gap
  - footer width consistency for all layouts
  - placeholder-level fill targets

## Change policy

When changing extractor, semantic normalizer, planner, generator, or audit:

1. Identify which document classes are affected.
2. Check the relevant external technical documentation first.
3. Run the regression corpus for the affected classes.
4. Reject a change that improves one class but regresses another without an explicit design decision.
