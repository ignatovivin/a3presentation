from __future__ import annotations

import sys
import unittest
from pathlib import Path


QUALITY_TEST_NAMES = [
    "tests.test_quality_contracts",
    "tests.test_project_contracts.ProjectContractTests.test_deck_audit_reports_body_font_sizes_within_layout_profile_bounds",
    "tests.test_project_contracts.ProjectContractTests.test_deck_audit_detects_continuation_groups_for_multislide_sections",
    "tests.test_project_contracts.ProjectContractTests.test_deck_audit_flags_underfilled_continuation_pairs",
    "tests.test_project_contracts.ProjectContractTests.test_deck_audit_accepts_balanced_dense_slides_without_capacity_violations",
    "tests.test_project_contracts.ProjectContractTests.test_deck_audit_validates_table_layout_geometry",
    "tests.test_project_contracts.ProjectContractTests.test_deck_audit_validates_chart_layout_geometry",
    "tests.test_project_contracts.ProjectContractTests.test_deck_audit_validates_image_layout_geometry",
    "tests.test_regression_corpus.RegressionCorpusTests.test_text_only_markdown_fixture_generates_deck_without_capacity_violations",
    "tests.test_regression_corpus.RegressionCorpusTests.test_mixed_text_fixture_generates_deck_without_capacity_violations",
    "tests.test_regression_corpus.RegressionCorpusTests.test_report_docx_generates_deck_without_capacity_violations",
    "tests.test_regression_corpus.RegressionCorpusTests.test_source_heavy_report_docx_skips_reference_tail_from_main_deck",
    "tests.test_regression_corpus.RegressionCorpusTests.test_question_callout_heavy_docx_preserves_semantic_blocks_without_capacity_violations",
    "tests.test_regression_corpus.RegressionCorpusTests.test_long_title_layout_stress_docx_generates_deck_without_capacity_violations",
    "tests.test_regression_corpus.RegressionCorpusTests.test_appendix_heavy_docx_generates_appendix_without_capacity_violations",
    "tests.test_regression_corpus.RegressionCorpusTests.test_long_title_with_subtitle_layout_stress_docx_generates_deck_without_capacity_violations",
    "tests.test_regression_corpus.RegressionCorpusTests.test_long_title_with_subtitle_dense_continuation_docx_keeps_continuation_balanced",
    "tests.test_regression_corpus.RegressionCorpusTests.test_long_title_with_subtitle_reference_tail_docx_skips_sources_and_keeps_layout_green",
    "tests.test_regression_corpus.RegressionCorpusTests.test_strategy_edge_case_docx_generates_deck_without_capacity_violations",
    "tests.test_regression_corpus.RegressionCorpusTests.test_form_like_docx_generates_deck_without_capacity_violations",
    "tests.test_regression_corpus.RegressionCorpusTests.test_resume_like_docx_generates_deck_without_capacity_violations",
    "tests.test_regression_corpus.RegressionCorpusTests.test_table_heavy_docx_generates_deck_without_text_capacity_violations",
    "tests.test_regression_corpus.RegressionCorpusTests.test_chart_heavy_docx_generates_chart_slide_and_preserves_text_capacity_contract",
    "tests.test_regression_corpus.RegressionCorpusTests.test_image_heavy_docx_generates_image_slide_and_preserves_text_capacity_contract",
    "tests.test_regression_corpus.RegressionCorpusTests.test_fact_only_docx_generates_appendix_without_capacity_violations",
    "tests.test_project_contracts.ProjectContractTests.test_deck_audit_uses_analyzer_geometry_metadata_for_uploaded_layout_templates",
    "tests.test_project_contracts.ProjectContractTests.test_deck_audit_uses_manifest_geometry_metadata_for_uploaded_prototype_templates",
    "tests.test_project_contracts.ProjectContractTests.test_full_pipeline_contract_for_uploaded_prototype_template_keeps_template_aware_audit_green",
]


def main() -> int:
    project_root = Path(__file__).resolve().parents[1]
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))

    loader = unittest.defaultTestLoader
    suite = unittest.TestSuite()
    for test_name in QUALITY_TEST_NAMES:
        suite.addTests(loader.loadTestsFromName(test_name))

    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    return 0 if result.wasSuccessful() else 1


if __name__ == "__main__":
    raise SystemExit(main())
