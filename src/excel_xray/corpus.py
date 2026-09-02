"""Corpus layer (Step 5): duplication and consolidation across a folder.

The three remaining EUC findings — *Potential Duplication*, *Similar / Duplicate
File(s)* and *Potential Consolidation* — cannot be answered from one workbook;
they need the set. This module fingerprints each workbook and compares them
pairwise, so a review of a folder can flag look-alikes and merge candidates.

The fingerprint is deliberately layout-invariant where it can be: normalised
R1C1 formula *shapes* (the same fill-down logic produces the same shape
regardless of where it sits) are the strongest "same activity" signal, backed by
column headers and the function profile. No cell values are used.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .assessment import Field, assess
from .scan import WorkbookXray

# A pair at/above this overall similarity is treated as materially duplicative.
DUP_THRESHOLD = 0.65
# …and at/above this, as similar enough to list and to consider consolidating.
SIMILAR_THRESHOLD = 0.35
# A very high formula-shape overlap alone is enough to call duplication.
SKELETON_DUP = 0.80


@dataclass
class Fingerprint:
    file_id: str
    file_name: str
    logic_type: str
    skeletons: set = field(default_factory=set)
    headers: set = field(default_factory=set)
    functions: set = field(default_factory=set)
    sheet_names: set = field(default_factory=set)


def fingerprint(wx: WorkbookXray, assessment) -> Fingerprint:
    """Build a comparable fingerprint from an already-computed assessment."""
    skeletons: set = set()
    headers: set = set()
    functions: set = set()
    sheet_names: set = set()
    for s in wx.sheets:
        sheet_names.add(s.name.strip().lower())
        for sk, _ in s.formula_profile.get("top_skeletons", []):
            skeletons.add(sk)
        for fn, _ in s.formula_profile.get("top_functions", []):
            functions.add(fn)
        for r in s.regions:
            for h in r.headers:
                for part in (h or "").split("::"):
                    p = part.strip().lower()
                    if p:
                        headers.add(p)
    return Fingerprint(
        file_id=assessment.file.file_id.value,
        file_name=wx.filename,
        logic_type=assessment.file.logic_type.value,
        skeletons=skeletons,
        headers=headers,
        functions=functions,
        sheet_names=sheet_names,
    )


def _jaccard(a: set, b: set) -> float:
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)


def similarity(a: Fingerprint, b: Fingerprint) -> dict:
    """Component and overall similarity in [0, 1]."""
    sk = _jaccard(a.skeletons, b.skeletons)
    hd = _jaccard(a.headers, b.headers)
    fn = _jaccard(a.functions, b.functions)
    if a.skeletons or b.skeletons:
        overall = 0.6 * sk + 0.3 * hd + 0.1 * fn
    else:  # value/data workbooks with no formulas: lean on structure
        overall = 0.7 * hd + 0.3 * fn
    return {
        "skeleton": round(sk, 3),
        "header": round(hd, 3),
        "function": round(fn, 3),
        "overall": round(overall, 3),
    }


def _apply_corpus(assessment, fp: Fingerprint, others: list[Fingerprint]) -> None:
    """Fill the three corpus findings on ``assessment`` from the comparisons."""
    fa = assessment.file

    if not others:
        note = "only one workbook in the corpus — no comparison possible"
        fa.potential_duplication = Field.derived({"verdict": "No", "matches": []}, 0.4, [note])
        fa.similar_duplicate_files = Field.derived([], 0.4, [note])
        fa.potential_consolidation = Field.derived({"verdict": "No"}, 0.4, [note])
        return

    scored = sorted(
        ((o, similarity(fp, o)) for o in others),
        key=lambda x: x[1]["overall"], reverse=True,
    )
    similar = [
        {"file": o.file_name, "file_id": o.file_id,
         "similarity": s["overall"], "signals": s,
         "same_logic": o.logic_type == fp.logic_type}
        for o, s in scored if s["overall"] >= SIMILAR_THRESHOLD
    ]
    dups = [
        m for m, (_, s) in zip(similar, [(o, s) for o, s in scored if s["overall"] >= SIMILAR_THRESHOLD])
        if s["overall"] >= DUP_THRESHOLD or s["skeleton"] >= SKELETON_DUP
    ]

    top = scored[0][1]["overall"]
    fa.similar_duplicate_files = Field.derived(
        [m["file"] for m in similar] or ["none above similarity threshold"],
        0.6, [f"{len(others)} other workbook(s) compared; top overall {top}"],
    )
    fa.potential_duplication = Field.derived(
        {"verdict": "Yes" if dups else "No", "matches": dups},
        0.6 + (0.2 if dups else 0.0),
        [f"{len(dups)} match(es) at/above duplication threshold {DUP_THRESHOLD}"],
    )

    # Consolidation: related workbooks sharing the same logic type. Exact
    # duplicates are the strongest merge candidates; moderate look-alikes are
    # "Possibly".
    consol = [m for m in similar if m["same_logic"]]
    if dups:
        verdict = "Yes"
    elif consol:
        verdict = "Possibly"
    else:
        verdict = "No"
    fa.potential_consolidation = Field.derived(
        {"verdict": verdict, "candidates": [m["file"] for m in consol]},
        0.55,
        [f"{len(consol)} candidate(s) share logic type '{fp.logic_type}'"],
    )


def assess_corpus(workbooks: list[WorkbookXray], assessor=None) -> list:
    """Assess every workbook, then fill each one's corpus findings by comparing
    its fingerprint against all the others."""
    results = [assess(wx, assessor) for wx in workbooks]
    fps = [fingerprint(wx, a) for wx, a in zip(workbooks, results)]
    for i, (a, fp) in enumerate(zip(results, fps)):
        others = [fps[j] for j in range(len(fps)) if j != i]
        _apply_corpus(a, fp, others)
    return results
