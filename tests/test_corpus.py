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
    assert a.file.potential_duplication.value["verdict"] == "Not established — functional comparison required"
    assert "only one workbook" in a.file.potential_duplication.evidence[0]


def test_structural_similarity_only_shortlists_candidates():
    a = _Assess()
    a.file.potential_duplication = _Fld()
    a.file.similar_duplicate_files = _Fld()
    a.file.potential_consolidation = _Fld()
    fp = _fp("mine.xlsx", ["RC[-1]*N", "SUM(R[-3]C:R[-1]C)"], ["policy id", "premium"])
    twin = _fp("twin.xlsx", ["RC[-1]*N", "SUM(R[-3]C:R[-1]C)"], ["policy id", "premium"])
    _apply_corpus(a, fp, [twin])
    assert a.file.potential_duplication.value["verdict"].startswith("Not established")
    assert a.file.potential_duplication.value["matches"][0]["file"] == "twin.xlsx"
    assert a.file.similar_duplicate_files.value.startswith("Not established")
    assert a.file.potential_consolidation.value["verdict"].startswith("Not established")


def test_same_process_or_structure_does_not_make_duplicate_without_evidence():
    from types import SimpleNamespace
    from excel_xray.assessment import FileAssessment
    from excel_xray.corpus import _normalise_rationalization

    assessment = SimpleNamespace(file=FileAssessment())
    current = {
        "comparison_id": "current", "file_id": "hash-a", "file_name": "current.xlsx",
        "process": "Claims", "sub_process": "Monthly review", "evidence": [],
    }
    peer = {
        "comparison_id": "peer", "file_id": "hash-b", "file_name": "peer.xlsx",
        "process": "Claims", "sub_process": "Monthly review", "evidence": [],
    }
    result = {
        "comparison_id": "current",
        "potential_duplication": {"verdict": "duplicate", "matches": [{
            "comparison_id": "peer", "functional_similarities": ["same process"],
            "material_differences": ["none listed"], "evidence": [],
        }]},
        "similar_duplicate_files": [],
        "potential_simplification": {"candidates": [{"recommendation": "remove a step", "evidence": []}]},
        "potential_consolidation": {"candidates": []},
        "potential_automation": {"candidates": []},
        "potential_retirement": {"verdict": "retire", "replacement_comparison_ids": [], "evidence": []},
    }
    assert _normalise_rationalization(assessment, current, [peer], {"results": [result]})
    assert assessment.file.potential_duplication.value["matches"] == []
    assert assessment.file.potential_duplication.value["verdict"].startswith(
        "No evidence-backed functional duplicate")
    assert assessment.file.potential_simplification.value["candidates"] == []
    assert assessment.file.potential_retirement.value["replacement_file_ids"] == []


def test_evidence_backed_duplicate_requires_both_business_summaries():
    from types import SimpleNamespace
    from excel_xray.assessment import FileAssessment
    from excel_xray.corpus import _normalise_rationalization

    assessment = SimpleNamespace(file=FileAssessment())
    current = {"comparison_id": "current", "file_id": "hash-a", "file_name": "current.xlsx",
               "evidence": [{"comparison_id": "current", "file_id": "hash-a",
                             "field": "Purpose of File", "text": "Monthly claims reserve"},
                            {"comparison_id": "current", "file_id": "hash-a",
                             "field": "Key Outputs", "text": "Reserve movement report"}]}
    peer = {"comparison_id": "peer", "file_id": "hash-b", "file_name": "peer.xlsx",
            "evidence": [{"comparison_id": "peer", "file_id": "hash-b",
                          "field": "Purpose of File", "text": "Monthly claims reserve"},
                         {"comparison_id": "peer", "file_id": "hash-b",
                          "field": "Key Outputs", "text": "Reserve movement report"}]}
    citations = [
        {**e, "comparison_id": e["comparison_id"]}
        for e in [*current["evidence"], *peer["evidence"]]
    ]
    result = {"comparison_id": "current",
              "potential_duplication": {"matches": [{
                  "comparison_id": "peer", "functional_similarities": ["same business output"],
                  "material_differences": ["different source extract"], "evidence": citations,
              }]},
              "similar_duplicate_files": [{"comparison_id": "peer",
                  "functional_similarities": ["same business output"],
                  "material_differences": ["different source extract"], "evidence": citations}],
              "potential_simplification": {"candidates": []},
              "potential_consolidation": {"candidates": []},
              "potential_automation": {"candidates": []},
              "potential_retirement": {"retirement_basis": "duplicated purpose",
                  "replacement_comparison_ids": ["peer"], "evidence": citations}}
    assert _normalise_rationalization(assessment, current, [peer], {"results": [result]})
    assert len(assessment.file.potential_duplication.value["matches"]) == 1
    assert assessment.file.potential_duplication.value["matches"][0]["file"] == "peer.xlsx"
    assert assessment.file.potential_retirement.value["replacement_file_ids"] == ["hash-b"]


def test_assess_corpus_runs_functional_review_for_structural_shortlist(fixture_path):
    from excel_xray.narrative import Narrative

    class FakeLLM:
        basis = "inferred"
        label = "fake"
        calls = 0

        def narrate(self, bundle):
            return Narrative(
                purpose_of_file="Monthly claims reserve analysis",
                key_output_outcome="A monthly reserve movement report",
                key_outputs=["Monthly reserve movement report"],
                tabs={tab["name"]: "Supports review" for tab in bundle["tabs"]},
            )

        def reconcile(self, bundle):
            return {"reconciliation_count": 0, "reconciliations": []}

        def rationalize(self, payload):
            self.calls += 1
            current = payload["current"]
            peer = payload["candidates"][0]
            citations = []
            for context in (current, peer):
                for field in ("Purpose of File", "Key Output / Outcome"):
                    item = next(e for e in context["evidence"] if e["field"] == field)
                    citations.append({"comparison_id": context["comparison_id"], **item})
            match = {"comparison_id": peer["comparison_id"],
                     "functional_similarities": ["same claims reserve purpose and report"],
                     "material_differences": ["source workbooks require owner review"],
                     "evidence": citations}
            return {"results": [{
                "comparison_id": current["comparison_id"],
                "potential_duplication": {"matches": [match]},
                "similar_duplicate_files": [match],
                "potential_simplification": {"candidates": []},
                "potential_consolidation": {"candidates": []},
                "potential_automation": {"candidates": []},
                "potential_retirement": {"replacement_comparison_ids": [], "evidence": []},
            }]}

    assessor = FakeLLM()
    workbook = xray_workbook(fixture_path)
    results = assess_corpus([workbook, xray_workbook(fixture_path)], assessor=assessor)
    assert assessor.calls == 2
    assert results[0].file.potential_duplication.basis == "inferred"
    assert results[0].file.potential_duplication.value["matches"][0]["file"] == workbook.filename


# --------------------------------------------------------------- end to end


def test_assess_corpus_keeps_identical_workbooks_as_review_candidates(fixture_path):
    # Identical structure is a shortlist signal, not proof of same business use.
    results = assess_corpus([xray_workbook(fixture_path), xray_workbook(fixture_path)])
    assert len(results) == 2
    dup = results[0].file.potential_duplication
    assert dup.basis == "needs_corpus"
    assert dup.value["verdict"].startswith("Not established")
    assert dup.value["matches"]
