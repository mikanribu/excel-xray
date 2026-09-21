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
from .review_rules import shorten_external
from .scan import WorkbookXray

# Structural thresholds shortlist pairs for a later functional review; they do
# not determine whether EUCs are business duplicates.
DUP_THRESHOLD = 0.65
# Lower bar for a bounded structural review shortlist.
SIMILAR_THRESHOLD = 0.35
# A very high formula-shape overlap also creates a shortlist candidate.
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
    """Component and overall similarity in [0, 1].

    ``skeleton``/``function`` read as *formula similarity*, ``header`` doubles
    as a proxy for *input/output similarity* (column labels cover both), and
    ``dependency`` (shared sheet names) is the closest signal this lighter,
    two-file-friendly fingerprint can offer for dependency similarity — a
    fuller four-signal breakdown (separate input/output/topology) is
    available from the folder-wide estate comparison (``--estate``).
    """
    sk = _jaccard(a.skeletons, b.skeletons)
    hd = _jaccard(a.headers, b.headers)
    fn = _jaccard(a.functions, b.functions)
    dep = _jaccard(a.sheet_names, b.sheet_names)
    if a.skeletons or b.skeletons:
        overall = 0.6 * sk + 0.3 * hd + 0.1 * fn
    else:  # value/data workbooks with no formulas: lean on structure
        overall = 0.7 * hd + 0.3 * fn
    return {
        "skeleton": round(sk, 3),
        "header": round(hd, 3),
        "function": round(fn, 3),
        "dependency": round(dep, 3),
        "overall": round(overall, 3),
    }


def _field_value(item):
    return item.value if hasattr(item, "value") else item


def rationalization_context(assessment, fp: Fingerprint | None = None,
                           comparison_id: str | None = None) -> dict:
    """Build compact business evidence for cross-EUC review without cell values."""
    file_assessment = assessment.file if hasattr(assessment, "file") else assessment.get("file", {})
    tabs = assessment.tabs if hasattr(assessment, "tabs") else assessment.get("tabs", [])

    def value(name):
        item = getattr(file_assessment, name, None) if not isinstance(file_assessment, dict) else file_assessment.get(name)
        return _field_value(item.get("value") if isinstance(item, dict) and "value" in item else item)

    file_id = value("file_id") or (fp.file_id if fp else "")
    file_name = shorten_external(value("file_name") or (fp.file_name if fp else ""))
    entity_id = str(comparison_id or file_id)
    evidence = []

    def add_evidence(field_name, obj, owner=file_id):
        if obj is None:
            return
        if isinstance(obj, str):
            text = obj.strip()
            if text:
                evidence.append({"comparison_id": entity_id, "file_id": owner,
                                 "field": field_name, "text": text})
        elif isinstance(obj, (list, tuple, set)):
            for child in obj:
                add_evidence(field_name, child, owner)
        elif isinstance(obj, dict):
            for child in obj.values():
                add_evidence(field_name, child, owner)
        elif isinstance(obj, (int, float)):
            evidence.append({"comparison_id": entity_id, "file_id": owner,
                             "field": field_name, "text": str(obj)})

    field_values = {
        "Purpose of File": value("purpose_of_file"),
        "Key Output / Outcome": value("key_output_outcome"),
        "Key Inputs": value("key_inputs"),
        "Key Outputs": value("key_outputs"),
        "Key calculations / logic": value("key_calculations_logic"),
    }
    for field_name, field_value in field_values.items():
        add_evidence(field_name, field_value)

    tab_context = []
    for tab in tabs:
        def tab_value(name):
            item = getattr(tab, name, None) if not isinstance(tab, dict) else tab.get(name)
            return _field_value(item.get("value") if isinstance(item, dict) and "value" in item else item)
        tab_name = tab_value("tab_name")
        summary = {
            "name": tab_name,
            "category": tab_value("tab_category"),
            "roles": tab_value("tab_roles"),
            "purpose": tab_value("tab_purpose_description"),
            "information_role": tab_value("tab_information_analysis"),
            "calculation_logic": tab_value("key_calculation_transformation_logic"),
            "upstream_dependencies": tab_value("upstream_dependencies"),
        }
        tab_context.append(summary)
        for field_name, field_value in summary.items():
            if field_name != "name":
                add_evidence(f"Worksheet {tab_name} {field_name}", field_value)
        add_evidence("Worksheet name", tab_name)

    if fp:
        add_evidence("Workbook headers", sorted(fp.headers))

    return {
        "comparison_id": entity_id,
        "file_id": str(file_id),
        "file_name": str(file_name),
        "process": value("process"),
        "sub_process": value("sub_process"),
        "business_area_purpose": value("business_area_process"),
        **field_values,
        "tabs": tab_context,
        "evidence": evidence,
    }


def _candidate_summary(scored) -> list[dict]:
    """Select structural shortlist only; these scores do not establish duplication."""
    return [
        {"file": shorten_external(peer.file_name), "file_id": peer.file_id,
         "similarity": score["overall"], "signals": score}
        for peer, score in scored
        if score["overall"] >= SIMILAR_THRESHOLD or score["skeleton"] >= SKELETON_DUP
    ][:20]


def _apply_corpus(assessment, fp: Fingerprint, others: list[Fingerprint]) -> list[dict]:
    """Create structural shortlists without treating them as business findings."""
    fa = assessment.file
    scored = sorted(((peer, similarity(fp, peer)) for peer in others),
                    key=lambda pair: pair[1]["overall"], reverse=True)
    candidates = _candidate_summary(scored)
    note = (f"{len(others)} other workbook(s) compared; structural scores shortlist candidates only"
            if others else "only one workbook in the corpus — no functional comparison possible")
    fa.similar_duplicate_files = Field(
        value="Not established — functional comparison required",
        basis="needs_corpus", confidence=None, evidence=[note],
    )
    fa.potential_duplication = Field(
        value={"verdict": "Not established — functional comparison required",
               "matches": candidates, "corpus_compared": len(others),
               "similarity_threshold": DUP_THRESHOLD},
        basis="needs_corpus", confidence=None,
        evidence=[note, "shared process labels or structural similarity alone do not establish duplicate activity"],
    )
    fa.potential_consolidation = Field(
        value={"verdict": "Not established — functional comparison required",
               "candidates": [], "structural_candidates": candidates},
        basis="needs_corpus", confidence=None,
        evidence=[note, "consolidation requires business-purpose fit and blocker review"],
    )
    return candidates


def _valid_citations(citations, contexts: dict[str, dict], allowed_comparison_ids: set[str]) -> list[dict]:
    accepted = []
    for cite in citations or []:
        if not isinstance(cite, dict):
            continue
        comparison_id = str(cite.get("comparison_id") or "")
        file_id = str(cite.get("file_id") or "")
        field_name = str(cite.get("field") or "")
        text = str(cite.get("text") or "")
        if comparison_id not in allowed_comparison_ids or not text or comparison_id not in contexts:
            continue
        if file_id != contexts[comparison_id].get("file_id"):
            continue
        known = contexts[comparison_id].get("evidence", [])
        if any(e["field"].casefold() == field_name.casefold()
               and e["text"].casefold() == text.casefold() for e in known):
            accepted.append({"comparison_id": comparison_id, "file_id": file_id,
                             "field": field_name, "text": text})
    return accepted


def _normalise_rationalization(assessment, current: dict, peers: list[dict], data: dict) -> bool:
    """Apply only recommendations whose evidence cites supplied workbook facts."""
    if not isinstance(data, dict):
        return False
    comparison_id = current["comparison_id"]
    results = data.get("results")
    if not isinstance(results, list):
        results = [data]
    result = next((row for row in results if isinstance(row, dict)
                   and str(row.get("comparison_id", row.get("file_id"))) == comparison_id), None)
    if result is None:
        return False
    peer_by_id = {str(p["comparison_id"]): p for p in peers}
    contexts = {str(c["comparison_id"]): c for c in [current, *peers]}
    business_fields = {"purpose of file", "key output / outcome", "key inputs",
                       "key outputs", "key calculations / logic"}

    def cited(items, allowed_ids):
        return _valid_citations(items, contexts, allowed_ids)

    def object_value(value):
        return value if isinstance(value, dict) else {}

    def list_value(value):
        return value if isinstance(value, list) else []

    duplicate_matches = []
    for match in list_value(object_value(result.get("potential_duplication")).get("matches")):
        if not isinstance(match, dict):
            continue
        peer_id = str(match.get("comparison_id") or match.get("file_id") or "")
        if peer_id not in peer_by_id:
            continue
        citations = cited(match.get("evidence"), {comparison_id, peer_id})
        citation_files = {c["comparison_id"] for c in citations}
        cited_business = {c["field"].casefold() for c in citations} & business_fields
        business_sources = {c["comparison_id"] for c in citations
                            if c["field"].casefold() in business_fields}
        similarities = [str(x) for x in match.get("functional_similarities", []) if str(x).strip()]
        differences = [str(x) for x in match.get("material_differences", []) if str(x).strip()]
        if (citation_files == {comparison_id, peer_id}
                and business_sources == {comparison_id, peer_id}
                and len(cited_business) >= 2 and similarities and differences):
            duplicate_matches.append({
                "file": peer_by_id[peer_id]["file_name"],
                "file_id": peer_by_id[peer_id]["file_id"],
                "comparison_id": peer_id,
                "functional_similarities": similarities,
                "material_differences": differences,
                "evidence": citations,
            })

    similar_files = []
    for match in list_value(result.get("similar_duplicate_files")):
        if not isinstance(match, dict):
            continue
        peer_id = str(match.get("comparison_id") or match.get("file_id") or "")
        if peer_id not in peer_by_id:
            continue
        citations = cited(match.get("evidence"), {comparison_id, peer_id})
        business_sources = {c["comparison_id"] for c in citations
                            if c["field"].casefold() in business_fields}
        cited_business = {c["field"].casefold() for c in citations} & business_fields
        if ({c["comparison_id"] for c in citations} == {comparison_id, peer_id}
                and business_sources == {comparison_id, peer_id} and len(cited_business) >= 2
                and match.get("functional_similarities") and match.get("material_differences")):
            similar_files.append({
                "file": peer_by_id[peer_id]["file_name"],
                "file_id": peer_by_id[peer_id]["file_id"],
                "comparison_id": peer_id,
                "functional_similarities": match["functional_similarities"],
                "material_differences": match["material_differences"], "evidence": citations,
            })

    def local_candidates(field_name, *, require_manual_evidence=False):
        raw = object_value(result.get(field_name)).get("candidates", [])
        out = []
        for candidate in list_value(raw):
            if not isinstance(candidate, dict):
                continue
            citations = cited(candidate.get("evidence"), {comparison_id})
            action = next((candidate.get(key) for key in
                           ("recommendation", "proposed_change", "candidate_action")
                           if str(candidate.get(key) or "").strip()), None)
            manual_terms = ("manual", "paste", "override", "keyed", "hardcode", "hard-coded")
            manual_evidence = any(
                any(term in citation["text"].casefold() for term in manual_terms)
                for citation in citations
            )
            if not citations or not action or (require_manual_evidence and not manual_evidence):
                continue
            out.append({**candidate, "evidence": citations})
        return out

    simplification = local_candidates("potential_simplification")
    automation = local_candidates("potential_automation", require_manual_evidence=True)

    def retain_existing(field_name, candidates):
        existing = object_value(getattr(assessment.file, field_name).value).get("candidates", [])
        combined = []
        seen = set()
        for candidate in [*list_value(existing), *candidates]:
            if not isinstance(candidate, dict):
                continue
            key = (candidate.get("worksheet"),
                   candidate.get("proposed_change") or candidate.get("recommendation")
                   or candidate.get("candidate_action"),
                   candidate.get("current_complexity"))
            key = key if any(part is not None for part in key) else repr(candidate)
            if key not in seen:
                combined.append(candidate)
                seen.add(key)
        return combined

    simplification = retain_existing("potential_simplification", simplification)
    automation = retain_existing("potential_automation", automation)

    consolidation = []
    for candidate in list_value(object_value(
            result.get("potential_consolidation")).get("candidates")):
        if not isinstance(candidate, dict):
            continue
        peer_id = str(candidate.get("comparison_id") or candidate.get("file_id") or "")
        if peer_id not in peer_by_id:
            continue
        citations = cited(candidate.get("evidence"), {comparison_id, peer_id})
        business_sources = {c["comparison_id"] for c in citations
                            if c["field"].casefold() in business_fields}
        cited_business = {c["field"].casefold() for c in citations} & business_fields
        if ({c["comparison_id"] for c in citations} == {comparison_id, peer_id}
                and business_sources == {comparison_id, peer_id}
                and len(cited_business) >= 2 and candidate.get("shared_solution")):
            consolidation.append({
                **candidate, "file": peer_by_id[peer_id]["file_name"],
                "file_id": peer_by_id[peer_id]["file_id"], "comparison_id": peer_id,
                "evidence": citations,
                "blockers": candidate.get("blockers") or [],
            })

    retirement_raw = object_value(result.get("potential_retirement"))
    replacements = [str(x) for x in list_value(retirement_raw.get("replacement_comparison_ids"))
                    if str(x) in peer_by_id]
    retirement_basis = str(retirement_raw.get("retirement_basis") or "").casefold()
    retirement_citations = cited(retirement_raw.get("evidence"),
                                 {comparison_id, *replacements})
    output_fields = {"key output / outcome", "key outputs"}
    output_sources = {c["comparison_id"] for c in retirement_citations
                      if c["field"].casefold() in output_fields}
    cited_ids = {c["comparison_id"] for c in retirement_citations}
    explicit_retired_terms = ("no longer required", "discontinued", "obsolete", "retired")
    explicit_not_required = (
        retirement_basis == "no longer required" and comparison_id in cited_ids
        and any(any(term in c["text"].casefold() for term in explicit_retired_terms)
                for c in retirement_citations if c["comparison_id"] == comparison_id)
    )
    duplicate_peers = {m["comparison_id"] for m in duplicate_matches}
    duplicate_purpose = (
        retirement_basis == "duplicated purpose" and bool(replacements)
        and set(replacements) <= duplicate_peers
        and cited_ids >= {comparison_id, *replacements}
    )
    replacement_coverage = (
        retirement_basis == "fully replaced" and bool(replacements)
        and cited_ids >= {comparison_id, *replacements}
        and output_sources >= {comparison_id, *replacements}
        and bool(retirement_raw.get("replacement_coverage"))
    )
    retirement_valid = explicit_not_required or duplicate_purpose or replacement_coverage

    basis = "inferred"
    confidence = 0.65
    fa = assessment.file
    if peers:
        fa.potential_duplication = Field(
            value={"verdict": "Potential functional duplication — confirm with process owners"
                   if duplicate_matches else
                   "No evidence-backed functional duplicate identified among reviewed candidates",
                   "matches": duplicate_matches,
                   "corpus_compared": current.get("corpus_compared", len(peers))},
            basis=basis, confidence=confidence,
            evidence=[f"Functional review checked {len(peers)} shortlisted candidate(s) from "
                      f"{current.get('corpus_compared', len(peers))} other EUC(s); "
                      "matching business evidence and material differences were required"],
        )
        fa.similar_duplicate_files = Field(
            value=similar_files or
            "No evidence-backed similar business use case identified among reviewed candidates",
            basis=basis, confidence=confidence,
            evidence=["file names and shared taxonomy labels are not used as the sole match signal"],
        )
    fa.potential_simplification = Field(
        value={"verdict": ("Potential simplification — owner review required" if simplification
                            else "No evidence-backed simplification opportunity identified"),
               "candidates": simplification},
        basis=basis, confidence=confidence,
        evidence=[f"{len(simplification)} candidate(s) retained after source-evidence validation"],
    )
    if peers:
        fa.potential_consolidation = Field(
            value={"verdict": ("Potential consolidation — owner review required" if consolidation
                                else "No evidence-backed consolidation opportunity identified among reviewed candidates"),
                   "candidates": consolidation},
            basis=basis, confidence=confidence,
            evidence=[f"{len(consolidation)} candidate(s) cite both EUCs; blockers retained"],
        )
    fa.potential_automation = Field(
        value={"verdict": ("Potential automation — owner review required" if automation
                            else "No evidence-backed automation opportunity identified"),
               "candidates": automation},
        basis=basis, confidence=confidence,
        evidence=[f"{len(automation)} repeatable/manual workflow candidate(s) retained after evidence validation"],
    )
    review = getattr(assessment, "review", None)
    if review is not None:
        review.simplification_candidates = simplification
        review.automation_candidates = automation
    fa.potential_retirement = Field(
        value=({"verdict": str(retirement_raw.get("verdict") or "Potential retirement — owner decision required"),
                "replacement_file_ids": [peer_by_id[x]["file_id"] for x in replacements],
                "replacement_coverage": retirement_raw.get("replacement_coverage"),
                "evidence": retirement_citations}
               if retirement_valid else
               {"verdict": "Not established — no evidence that this EUC is redundant or fully replaced",
                "replacement_file_ids": [], "replacement_coverage": None, "evidence": []}),
        basis=basis, confidence=confidence,
        evidence=["retirement is suppressed unless a supplied replacement EUC and coverage evidence are cited"],
    )
    return True


def _run_rationalization(assessment, current: dict, peers: list[dict], assessor) -> None:
    method = getattr(assessor, "rationalize", None)
    if not callable(method):
        return
    try:
        data = method({"current": current, "candidates": peers})
    except Exception:
        # Keep the structural shortlist, but never upgrade it to a business finding.
        assessment.file.potential_duplication.evidence.append(
            "LLM functional comparison was unavailable; structural matches are not duplicate findings")
        return
    _normalise_rationalization(assessment, current, peers, data)


def assess_corpus(workbooks: list[WorkbookXray], assessor=None) -> list:
    """Assess every workbook, then fill each one's corpus findings by comparing
    its fingerprint against all the others."""
    results = [assess(wx, assessor) for wx in workbooks]
    fps = [fingerprint(wx, a) for wx, a in zip(workbooks, results)]
    for i, (a, fp) in enumerate(zip(results, fps)):
        scored = sorted(
            ((j, fps[j], similarity(fp, fps[j])) for j in range(len(fps)) if j != i),
            key=lambda row: row[2]["overall"], reverse=True,
        )
        _apply_corpus(a, fp, [row[1] for row in scored])
        peers = []
        for j, peer_fp, score in scored:
            if score["overall"] < SIMILAR_THRESHOLD and score["skeleton"] < SKELETON_DUP:
                continue
            peer = rationalization_context(
                results[j], peer_fp, comparison_id=f"{j}:{peer_fp.file_id}")
            peer["structural_similarity"] = score
            peers.append(peer)
            if len(peers) == 20:
                break
        current = rationalization_context(a, fp, comparison_id=f"{i}:{fp.file_id}")
        current["corpus_compared"] = len(fps) - 1
        _run_rationalization(a, current, peers, assessor)
    return results
