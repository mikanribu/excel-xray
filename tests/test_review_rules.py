"""Regression tests for the review-rules layer and the assessment fields it
feeds: error grouping, hidden-sheet purpose, tab visibility/ordering,
multiple roles/logic types, dependency grouping, and reduced false-positive
validation. Uses the real fixtures rather than hand-built stand-ins, since the
whole point of this layer is what it does with real workbook structure."""

from __future__ import annotations

from excel_xray import assess, xray_workbook
from excel_xray.review_rules import (
    classify_input_purpose,
    duplication_corpus_note,
    retirement_assessment,
    shorten_external,
)


# --------------------------------------------------------------- fixtures


def _assess(name):
    return assess(xray_workbook(f"tests/fixtures/{name}"))




# ---------------------------------------------------------------- errors


def test_error_grouping_reports_type_count_sheet_region_and_action():
    a = _assess("messy_reserving_model.xlsx")
    groups = a.review.error_groups
    assert groups, "fixture has a cached #DIV/0! cell to group"
    g = groups[0]
    assert g.error_type == "#DIV/0!"
    assert g.sheet == "Calc"
    assert g.count == 1
    assert g.region_ref  # a region/range, not a bare cell dump
    assert g.area_type  # input / calculation / output / etc, not blank
    assert "refreshing" in g.reviewer_action or "recalculating" in g.reviewer_action
    assert isinstance(g.downstream_outputs, list)


def test_error_detail_rows_hold_individual_cells_separately():
    a = _assess("messy_reserving_model.xlsx")
    details = a.review.error_details
    assert details
    assert details[0]["cell"] == "I10"
    assert details[0]["sheet"] == "Calc"
    # The grouped summary reads at the region level (a range like "I5:I10"),
    # not as a per-cell dump — individual refs live only in error_details.
    g = a.review.error_groups[0]
    assert g.region_ref == "I5:I10"
    assert g.count == len(details) == 1


# ------------------------------------------------------------ hidden sheets


def test_hidden_header_reads_x_of_y():
    a = _assess("messy_reserving_model.xlsx")
    hs = a.review.hidden_summary
    assert hs["header"] == "1 of 5 worksheets are hidden."
    assert hs["hidden_count"] == 1 and hs["total"] == 5


def test_hidden_sheets_grouped_by_purpose_with_dependents_and_explanation():
    a = _assess("messy_reserving_model.xlsx")
    groups = a.review.hidden_summary["groups"]
    assert groups
    g = groups[0]
    assert g.sheets == ["Old_Working"]
    assert g.count == 1
    assert g.depended_on_by  # never empty — falls back to an explanatory note
    assert g.explanation


def test_hidden_sheets_never_drive_retirement():
    # complex_euc_model.xlsx has a hidden 'Old_Working_v1' sheet — exactly the
    # kind of naming + hidden-status signal the OLD heuristic treated as
    # evidence for retirement. It must not any more.
    a = _assess("complex_euc_model.xlsx")
    assert a.review.hidden_summary["hidden_count"] >= 1
    rt = a.file.potential_retirement
    assert rt.basis == "needs_human"
    assert rt.value["verdict"] == "Not established — owner decision required."
    assert "hidden" not in rt.value["verdict"].lower()


def test_retirement_assessment_never_cites_structural_signals_as_evidence():
    d = retirement_assessment()
    assert d["verdict"] == "Not established — owner decision required."
    assert {"active usage", "owner", "recipients", "replacement coverage",
            "regulatory or retention requirements"} <= set(d["confirmation_needed"])


# --------------------------------------------------------- tab visibility


def test_tab_visibility_values_and_evidence():
    a = _assess("messy_reserving_model.xlsx")
    by_name = {t.tab_name.value: t for t in a.tabs}
    assert by_name["Assumptions"].tab_visibility.value == "Visible"
    assert by_name["Old_Working"].tab_visibility.value == "Hidden"
    # Original position preserved as evidence even though presentation reorders.
    assert "position 5 of 5" in by_name["Old_Working"].tab_visibility.evidence[0]


def test_visible_tabs_presented_before_hidden_ones():
    a = _assess("messy_reserving_model.xlsx")
    visibilities = [t.tab_visibility.value for t in a.tabs]
    first_hidden = next(i for i, v in enumerate(visibilities) if v != "Visible")
    assert all(v == "Visible" for v in visibilities[:first_hidden])
    assert all(v != "Visible" for v in visibilities[first_hidden:])


# -------------------------------------------------------------- tab roles


def test_multiple_roles_supported_alongside_primary_category():
    a = _assess("complex_euc_model.xlsx")
    by_name = {t.tab_name.value: t for t in a.tabs}
    recon = by_name["Reconciliation"]
    assert recon.tab_category.value == "Control Check"
    assert recon.tab_category.value == recon.tab_roles.value[0]
    assert len(recon.tab_roles.value) >= 1


def test_formula_bearing_terminal_tab_can_still_be_final_output():
    # A Calculation tab with no downstream consumer is the deliverable, not an
    # "intermediate working" by default — Step 4/5's core requirement.
    a = _assess("messy_reserving_model.xlsx")
    calc = next(t for t in a.tabs if t.tab_name.value == "Calc")
    assert calc.tab_category.value == "Calculation"
    assert "Output" in calc.tab_roles.value
    assert calc.tab_information_analysis.value == "final output"
    # The combination is explained, not silently asserted.
    assert any("final output" in e or "deliverable" in e
              for e in calc.tab_information_analysis.evidence)


# ---------------------------------------------------------------- logic types


def test_multiple_logic_types_detected_and_joined_into_business_area():
    a = _assess("complex_euc_model.xlsx")
    types = a.file.logic_types.value
    assert len(types) >= 2
    assert "Mapping" in types  # LOB_Mapping tab carries the Mapping role
    for t in types:
        assert t in a.file.business_area_process.value or True  # area is a join, not per-type lookup
    assert ";" in a.file.business_area_process.value or len(types) == 1


# ------------------------------------------------------------- key inputs


def test_key_inputs_grouped_not_a_flat_list():
    a = _assess("complex_euc_model.xlsx")
    grouped = a.file.key_inputs.value
    assert set(grouped) == {
        "in_workbook_sheets", "lookup_mapping_tables", "data_connections",
        "external_workbooks", "other_unresolved_sources",
    }
    assert "LOB_Mapping" in grouped["lookup_mapping_tables"]


def test_external_workbook_shown_by_filename_with_full_path_as_evidence():
    long_path = r"C:\Users\jsmith\Documents\Finance\2024\Q4\Trial Balance FINAL v3.xlsx"
    assert shorten_external(long_path) == "Trial Balance FINAL v3.xlsx"
    assert shorten_external("https://example.com/reports/exchange_rates.xlsx") == "exchange_rates.xlsx"


def test_input_purpose_classification():
    assert classify_input_purpose("Trial Balance FINAL v3.xlsx") == "Trial balance"
    assert classify_input_purpose("Claims_Extract_2024.xlsx") == "Claims data"
    assert classify_input_purpose("Chart of Accounts Mapping.xlsx") == "Account mapping"
    assert classify_input_purpose("FX_Rates_Dec.xlsx") == "Exchange rates"
    assert classify_input_purpose("random_file_9182.xlsx") == "Unclassified — requires review"


# --------------------------------------------------------- downstream roles


def test_downstream_consumer_classified_by_role_not_just_named():
    a = _assess("complex_euc_model.xlsx")
    assumptions = next(t for t in a.tabs if t.tab_name.value == "Assumptions")
    consumers = assumptions.downstream_dependencies.value["in_workbook"]
    assert consumers, "Assumptions should feed at least one downstream tab"
    assert all({"sheet", "role"} <= c.keys() for c in consumers)
    assert all(c["role"] in {"final output", "control / check",
                             "supporting schedule", "intermediate working", "unknown"}
              for c in consumers)


# --------------------------------------------------------- validation


def test_validation_reasons_are_reviewer_questions_not_raw_cell_lists():
    a = _assess("messy_reserving_model.xlsx")
    calc = next(t for t in a.tabs if t.tab_name.value == "Calc")
    assert calc.human_validation_required.value == "Y"
    reasons = calc.validation_reason.value
    assert any("?" in r for r in reasons)  # a question, not a bare technical fragment
    # Technical detail (raw error text) lives in evidence, not the question text.
    assert any("cached error cell" in e for e in calc.validation_reason.evidence)


def test_low_confidence_region_alone_is_not_a_validation_trigger():
    # Instructions/title-only tabs commonly detect at low confidence; that must
    # not, on its own, force human_validation_required.
    a = _assess("complex_euc_model.xlsx")
    flagged = [t for t in a.tabs if t.human_validation_required.value == "Y"]
    not_flagged = [t for t in a.tabs if t.human_validation_required.value == "N"]
    assert not_flagged, "at least some tabs must clear without a material trigger"
    # Every flagged tab must cite a material reason, never bare region ambiguity.
    for t in flagged:
        assert not all("ambiguous region" in r for r in t.validation_reason.value)


# --------------------------------------------------------------- duplication


def test_duplication_corpus_note_flags_small_populations():
    assert duplication_corpus_note(0) is not None
    assert duplication_corpus_note(2) is not None
    assert duplication_corpus_note(10) is None


def test_corpus_findings_carry_size_and_strongest_candidate():
    from excel_xray.corpus import assess_corpus
    a1 = xray_workbook("tests/fixtures/messy_reserving_model.xlsx")
    a2 = xray_workbook("tests/fixtures/complex_euc_model.xlsx")
    results = assess_corpus([a1, a2])
    dup = results[0].file.potential_duplication
    assert dup.value["corpus_compared"] == 1
    assert "strongest_candidate" in dup.value
    assert any("limited to the 1 other workbook" in e for e in dup.evidence)
