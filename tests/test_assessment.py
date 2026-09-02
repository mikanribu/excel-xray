"""Step 1 — file-level EUC assessment fields."""

from __future__ import annotations

from excel_xray import assess


VALID_BASES = {"extracted", "derived", "needs_llm", "needs_human", "needs_corpus"}


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
    assert fa.purpose_of_file.basis == "needs_llm"
    assert fa.usage_frequency.basis == "needs_human"
    assert fa.potential_duplication.basis == "needs_corpus"
    # Deferred fields have no fabricated value.
    assert fa.purpose_of_file.value is None


def test_key_calculations_extracted(xray):
    fa = assess(xray).file
    assert fa.key_calculations_logic.basis == "extracted"
    assert fa.key_calculations_logic.value["top_functions"]
