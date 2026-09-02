"""EUC assessment fields (Steps 1-4)."""

from __future__ import annotations

import json

from excel_xray import assess


VALID_BASES = {"extracted", "derived", "drafted", "inferred",
               "needs_llm", "needs_human", "needs_corpus"}


def test_every_field_declares_a_valid_basis(xray):
    fa = assess(xray).file
    for name, fld in vars(fa).items():
        assert fld.basis in VALID_BASES, f"{name} has bad basis {fld.basis!r}"


def test_extracted_facts(xray):
    fa = assess(xray).file
    assert fa.file_name.value == "messy_reserving_model.xlsx"
    assert fa.file_id.basis == "extracted" and len(fa.file_id.value) == 12
    assert fa.macros_vba_external_links.value == ["none detected"]


def test_complexity_is_low_for_small_fixture(xray):
    fa = assess(xray).file
    assert fa.complexity.basis == "derived"
    assert fa.complexity.value in {"Low", "Medium", "High"}
    assert fa.complexity.value == "Low"  # 375 cells, 105 formulas, no links/VBA
    assert fa.complexity.evidence


def test_logic_type_is_calculation(xray):
    fa = assess(xray).file
    assert fa.logic_type.value == "Calculation"


def test_manual_intervention_flags_hardcoded_inputs(xray):
    fa = assess(xray).file
    assert fa.manual_intervention.value in {"Low", "Medium", "High"}
    # 66 formulas carry a hardcoded number in this fixture.
    assert any("hardcoded" in e for e in fa.manual_intervention.evidence)


def test_deferred_fields_carry_intended_basis(xray):
    fa = assess(xray).file
    # These stay genuinely deferred after Step 4.
    assert fa.usage_frequency.basis == "needs_human"
    assert fa.completion_timeline.basis == "needs_human"
    assert fa.potential_duplication.basis == "needs_corpus"
    # Deferred fields have no fabricated value.
    assert fa.usage_frequency.value is None


def test_key_calculations_extracted(xray):
    fa = assess(xray).file
    assert fa.key_calculations_logic.basis == "extracted"
    assert fa.key_calculations_logic.value["top_functions"]


# --------------------------------------------------------------- tab level


def _tab(xray, name):
    return next(t for t in assess(xray).tabs if t.tab_name.value == name)


def test_one_tab_per_sheet(xray):
    a = assess(xray)
    assert len(a.tabs) == len(xray.sheets)


def test_tab_category_valid_or_uncertain(xray):
    for t in assess(xray).tabs:
        cat = t.tab_category
        assert cat.value in {
            "Input", "Calculation", "Mapping", "Control Check",
            "Validation", "Output", "Uncertain",
        }
        # Below-threshold categories must be flagged for the model, not asserted.
        if cat.value == "Uncertain":
            assert cat.basis == "needs_llm"
        else:
            assert cat.basis == "derived" and cat.confidence >= 0.55


def test_assumptions_is_input_feeding_calc(xray):
    t = _tab(xray, "Assumptions")
    assert t.tab_category.value == "Input"
    assert t.downstream_dependencies.value["in_workbook"] == ["sheet: Calc"]


def test_calc_is_calculation_reading_assumptions(xray):
    t = _tab(xray, "Calc")
    assert t.tab_category.value == "Calculation"
    assert "sheet: Assumptions" in t.upstream_dependencies.value


def test_calc_needs_validation_for_error_and_hardcodes(xray):
    t = _tab(xray, "Calc")
    assert t.human_validation_required.value == "Y"
    assert any("#DIV/0!" in r for r in t.validation_reason.value)


def test_downstream_cross_file_left_to_corpus(xray):
    t = _tab(xray, "Calc")
    assert t.downstream_dependencies.value["other_files"] is None


# ------------------------------------------------- AI findings (Step 3)


def test_simplification_flags_hardcodes_and_hidden_sheet(xray):
    fa = assess(xray).file
    sf = fa.potential_simplification
    assert sf.basis == "derived"
    assert sf.value["verdict"] == "Yes"
    ops = " ".join(sf.value["opportunities"])
    assert "hardcoded" in ops and "hidden" in ops


def test_automation_is_a_candidate(xray):
    fa = assess(xray).file
    au = fa.potential_automation
    assert au.basis == "derived"
    assert au.value["verdict"] in {"Yes", "Possibly"}
    assert au.value["drivers"]


def test_retirement_defers_when_no_signal(xray):
    # A freshly written fixture with no backup-naming has no retirement signal.
    fa = assess(xray).file
    rt = fa.potential_retirement
    assert rt.basis in {"derived", "needs_human"}
    if rt.basis == "needs_human":
        assert rt.value["verdict"] == "No automated retirement signal"


def test_reconciliation_absent_on_calc_model(xray):
    fa = assess(xray).file
    rl = fa.reconciliation_logic
    assert rl.basis == "derived"
    assert rl.value == "No reconciliation pattern detected"


def test_corpus_findings_still_deferred(xray):
    fa = assess(xray).file
    assert fa.potential_duplication.basis == "needs_corpus"
    assert fa.potential_consolidation.basis == "needs_corpus"


def test_months_since_helper():
    from excel_xray.assessment import _months_since
    assert _months_since(None) is None
    assert _months_since("not-a-date") is None
    assert _months_since("2020-01-01T00:00:00") >= 12
    assert _months_since("2020-01-01T00:00:00Z") >= 12  # tz-aware parses too


# ------------------------------------------------- narrative (Step 4)


def test_business_area_derived_from_logic(xray):
    fa = assess(xray).file
    assert fa.business_area_process.basis == "derived"
    assert fa.business_area_process.value == "Calculation / modelling"


def test_offline_narrative_is_drafted_and_grounded(xray):
    # Default assessor is the offline stub — no network, basis "drafted".
    fa = assess(xray).file
    assert fa.purpose_of_file.basis == "drafted"
    assert fa.purpose_of_file.value and "calculation" in fa.purpose_of_file.value.lower()
    assert fa.key_outputs.basis == "drafted"
    assert isinstance(fa.key_outputs.value, list) and fa.key_outputs.value


def test_every_tab_gets_a_purpose(xray):
    for t in assess(xray).tabs:
        assert t.tab_purpose_description.basis == "drafted"
        assert t.tab_purpose_description.value


def test_custom_assessor_marks_fields_inferred(xray):
    from excel_xray.narrative import Narrative

    class FakeLLM:
        basis = "inferred"
        label = "fake"

        def narrate(self, bundle):
            names = [t["name"] for t in bundle["tabs"]]
            return Narrative(
                purpose_of_file="A considered purpose.",
                key_output_outcome="A considered outcome.",
                key_outputs=["out"],
                tabs={n: f"purpose of {n}" for n in names},
            )

    a = assess(xray, assessor=FakeLLM())
    assert a.file.purpose_of_file.basis == "inferred"
    assert a.file.purpose_of_file.value == "A considered purpose."
    assert all(t.tab_purpose_description.basis == "inferred" for t in a.tabs)


def test_parse_json_tolerates_fenced_output():
    from excel_xray.narrative import _parse_json
    assert _parse_json('```json\n{"a": 1}\n```') == {"a": 1}
    assert _parse_json('here you go: {"a": 2} thanks') == {"a": 2}


def test_bundle_carries_no_cell_values(xray):
    # Privacy: the evidence bundle sent to a model must not embed cell values.
    from excel_xray.narrative import build_bundle
    a = assess(xray)
    bundle = build_bundle(a, xray)
    blob = json.dumps(bundle)
    # A distinctive cached figure from Calc!C5 must not appear.
    assert "59937" not in blob
