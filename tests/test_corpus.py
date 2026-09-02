"""Step 5 — corpus duplication / consolidation."""

from __future__ import annotations

from excel_xray import xray_workbook
from excel_xray.corpus import (
    DUP_THRESHOLD,
    Fingerprint,
    _apply_corpus,
    assess_corpus,
    similarity,
)


def _fp(name, skeletons, headers, logic="Calculation", functions=("SUM",)):
    return Fingerprint(
        file_id=name, file_name=name, logic_type=logic,
        skeletons=set(skeletons), headers=set(headers),
        functions=set(functions), sheet_names={name.lower()},
    )


# ---------------------------------------------------------------- similarity


def test_identical_fingerprints_are_maximally_similar():
    a = _fp("a.xlsx", ["RC[-1]*N", "SUM(R[-3]C:R[-1]C)"], ["policy id", "premium"])
    b = _fp("b.xlsx", ["RC[-1]*N", "SUM(R[-3]C:R[-1]C)"], ["policy id", "premium"])
    s = similarity(a, b)
    assert s["skeleton"] == 1.0 and s["header"] == 1.0
    assert s["overall"] >= DUP_THRESHOLD


def test_disjoint_fingerprints_are_dissimilar():
    a = _fp("a.xlsx", ["RC[-1]*N"], ["policy id"])
    b = _fp("b.xlsx", ["SHEET!R1C1+N"], ["gl account"], logic="Reporting", functions=("VLOOKUP",))
    assert similarity(a, b)["overall"] < 0.35


def test_no_formulas_falls_back_to_headers():
    a = _fp("a.xlsx", [], ["region", "q1", "q2"], functions=())
    b = _fp("b.xlsx", [], ["region", "q1", "q2"], functions=())
    assert similarity(a, b)["overall"] >= 0.65


# ------------------------------------------------------------- apply findings


class _Fld:
    def __init__(self):
        self.value = None
        self.basis = None
        self.confidence = None
        self.evidence = []


class _Assess:
    """Minimal stand-in exposing just the three corpus fields."""

    class file:
        potential_duplication = _Fld()
        similar_duplicate_files = _Fld()
        potential_consolidation = _Fld()


def test_single_file_corpus_defers_gracefully():
    a = _Assess()
    a.file.potential_duplication = _Fld()
    a.file.similar_duplicate_files = _Fld()
    a.file.potential_consolidation = _Fld()
    fp = _fp("only.xlsx", ["RC[-1]*N"], ["x"])
    _apply_corpus(a, fp, [])
    assert a.file.potential_duplication.value["verdict"] == "No"
    assert "only one workbook" in a.file.potential_duplication.evidence[0]


def test_duplicate_pair_flags_duplication_and_consolidation():
    a = _Assess()
    a.file.potential_duplication = _Fld()
    a.file.similar_duplicate_files = _Fld()
    a.file.potential_consolidation = _Fld()
    fp = _fp("mine.xlsx", ["RC[-1]*N", "SUM(R[-3]C:R[-1]C)"], ["policy id", "premium"])
    twin = _fp("twin.xlsx", ["RC[-1]*N", "SUM(R[-3]C:R[-1]C)"], ["policy id", "premium"])
    _apply_corpus(a, fp, [twin])
    assert a.file.potential_duplication.value["verdict"] == "Yes"
    assert "twin.xlsx" in a.file.similar_duplicate_files.value
    assert a.file.potential_consolidation.value["verdict"] == "Yes"


# --------------------------------------------------------------- end to end


def test_assess_corpus_detects_identical_workbooks(fixture_path):
    # Two references to the same workbook must see each other as duplicates.
    results = assess_corpus([xray_workbook(fixture_path), xray_workbook(fixture_path)])
    assert len(results) == 2
    dup = results[0].file.potential_duplication
    assert dup.basis == "derived"
    assert dup.value["verdict"] == "Yes"
    assert results[0].file.similar_duplicate_files.value  # non-empty
