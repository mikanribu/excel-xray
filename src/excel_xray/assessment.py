"""Assessment layer: map extracted evidence onto the EUC review schema.

The scanner (:mod:`excel_xray.scan`) is the *evidence* layer. This module turns
that evidence into the reviewer-facing fields of the EUC (End-User Computing)
assessment — file-level and tab-level — matching the target schema.

Every field is a :class:`Field` carrying not just a value but its ``basis``, so
the output is honest about what is grounded in the file versus what still needs a
model, a human, or the wider corpus to answer:

* ``extracted``    — read directly from the workbook.
* ``derived``      — a heuristic over extracted metrics (with evidence).
* ``needs_llm``    — a narrative/judgement call for the model step.
* ``needs_human``  — not knowable from the file (e.g. usage frequency).
* ``needs_corpus`` — needs the whole folder of workbooks (e.g. duplication).

The heavier structured rules — error grouping, hidden-sheet purpose, dependency
grouping, business-calculation descriptions, targeted validation questions,
automation/simplification/retirement/reconciliation — live in
:mod:`excel_xray.review_rules` and are wired in here.
"""

from __future__ import annotations

import datetime as _dt
import re
from collections import Counter
from dataclasses import asdict, dataclass, field

from . import review_rules as rr
from .review_rules import TabFacts, visibility_of
from .scan import WorkbookXray
from .util import safe_scan_message

# --------------------------------------------------------------------- Field

Basis = str  # one of: extracted | derived | needs_llm | needs_human | needs_corpus


@dataclass
class Field:
    """One assessment field: its value, where it came from, and why."""

    value: object = None
    basis: Basis = "needs_human"
    confidence: float | None = None
    evidence: list[str] = field(default_factory=list)

    @classmethod
    def extracted(cls, value, evidence=None) -> "Field":
        return cls(value=value, basis="extracted", confidence=1.0,
                   evidence=evidence or [])

    @classmethod
    def derived(cls, value, confidence, evidence) -> "Field":
        return cls(value=value, basis="derived", confidence=round(confidence, 2),
                   evidence=list(evidence))

    @classmethod
    def pending(cls, basis: Basis, note: str = "") -> "Field":
        return cls(value=None, basis=basis, evidence=[note] if note else [])

    @classmethod
    def needs_human(cls, question: str, evidence=None) -> "Field":
        """Human-only field: no value invented, plus a specific reviewer question."""
        return cls(value=None, basis="needs_human",
                   evidence=(evidence or []) + [f"reviewer question: {question}"])


# ------------------------------------------------------------- schema shapes


@dataclass
class FileAssessment:
    """File-level summary — the top block of the target schema."""

    # Fact Assessment
    file_id: Field = field(default_factory=Field)
    file_name: Field = field(default_factory=Field)
    business_area_process: Field = field(default_factory=Field)
    process: Field = field(default_factory=Field)
    sub_process: Field = field(default_factory=Field)
    purpose_of_file: Field = field(default_factory=Field)
    key_output_outcome: Field = field(default_factory=Field)
    complexity: Field = field(default_factory=Field)
    key_inputs: Field = field(default_factory=Field)
    source_system: Field = field(default_factory=Field)
    key_outputs: Field = field(default_factory=Field)
    usage_frequency: Field = field(default_factory=Field)
    completion_timeline: Field = field(default_factory=Field)
    euc_preparer: Field = field(default_factory=Field)
    output_recipient: Field = field(default_factory=Field)
    # Key AI Finding / Observation
    potential_duplication: Field = field(default_factory=Field)
    similar_duplicate_files: Field = field(default_factory=Field)
    potential_simplification: Field = field(default_factory=Field)
    potential_consolidation: Field = field(default_factory=Field)
    potential_automation: Field = field(default_factory=Field)
    potential_retirement: Field = field(default_factory=Field)
    # Workbook logic / Automation
    logic_type: Field = field(default_factory=Field)
    logic_types: Field = field(default_factory=Field)
    key_calculations_logic: Field = field(default_factory=Field)
    reconciliation_logic: Field = field(default_factory=Field)
    manual_intervention: Field = field(default_factory=Field)
    macros_vba_external_links: Field = field(default_factory=Field)


@dataclass
class TabAssessment:
    """Tab-level detail — the bottom block of the target schema (Step 2)."""

    tab_name: Field = field(default_factory=Field)
    tab_visibility: Field = field(default_factory=Field)
    tab_category: Field = field(default_factory=Field)
    tab_roles: Field = field(default_factory=Field)
    tab_purpose_description: Field = field(default_factory=Field)
    tab_information_analysis: Field = field(default_factory=Field)
    key_calculation_transformation_logic: Field = field(default_factory=Field)
    upstream_dependencies: Field = field(default_factory=Field)
    downstream_dependencies: Field = field(default_factory=Field)
    human_validation_required: Field = field(default_factory=Field)
    validation_reason: Field = field(default_factory=Field)


@dataclass
class ReviewFindings:
    """Structured findings outside the fixed EUC schema table — the report
    header (error/hidden summaries) and the rollups the Excel/HTML exporters
    render as their own sheets/sections."""

    error_groups: list = field(default_factory=list)
    error_details: list = field(default_factory=list)
    hidden_summary: dict = field(default_factory=dict)
    macro_details: dict = field(default_factory=dict)
    simplification_candidates: list = field(default_factory=list)
    automation_candidates: list = field(default_factory=list)
    key_inputs_grouped: dict = field(default_factory=dict)


@dataclass
class Assessment:
    file: FileAssessment
    tabs: list[TabAssessment] = field(default_factory=list)
    review: ReviewFindings = field(default_factory=ReviewFindings)
    scan: dict = field(default_factory=dict)


# ---------------------------------------------------------- workbook metrics


def _totals(wx: WorkbookXray) -> dict:
    """Roll per-sheet numbers up to the workbook, for the file-level fields."""
    cells = sum(s.populated_cells for s in wx.sheets)
    formulas = sum(s.formula_profile.get("total", 0) for s in wx.sheets)
    distinct = sum(s.formula_profile.get("distinct_skeletons", 0) for s in wx.sheets)
    functions: Counter = Counter()
    for s in wx.sheets:
        for fn, n in s.formula_profile.get("top_functions", []):
            functions[fn] += n
    cross_sheet = sum(s.formula_profile.get("cross_sheet_count", 0) for s in wx.sheets)
    hardcoded = sum(s.formula_profile.get("hardcoded_literal_count", 0) for s in wx.sheets)
    volatile = sum(s.formula_profile.get("volatile_count", 0) for s in wx.sheets)
    return {
        "cells": cells,
        "formulas": formulas,
        "distinct": distinct,
        "functions": functions,
        "cross_sheet": cross_sheet,
        "hardcoded": hardcoded,
        "volatile": volatile,
        "non_formula_cells": cells - formulas,
        "linked_workbooks": len(wx.external_links),
        "connections": len(wx.connections),
    }


# ------------------------------------------------------- deterministic fields


def _complexity(wx: WorkbookXray, t: dict) -> Field:
    """High / Medium / Low from data volume, calculation logic and linkage."""
    ev: list[str] = []
    score = 0

    if t["cells"] > 20_000:
        score += 2; ev.append(f"{t['cells']:,} populated cells (large)")
    elif t["cells"] > 2_000:
        score += 1; ev.append(f"{t['cells']:,} populated cells (moderate)")
    else:
        ev.append(f"{t['cells']:,} populated cells (small)")

    if t["formulas"] > 5_000 or t["distinct"] > 60:
        score += 2; ev.append(f"{t['formulas']:,} formulas / {t['distinct']} distinct calcs (heavy logic)")
    elif t["formulas"] > 200 or t["distinct"] > 15:
        score += 1; ev.append(f"{t['formulas']:,} formulas / {t['distinct']} distinct calcs (moderate logic)")
    elif t["formulas"]:
        ev.append(f"{t['formulas']:,} formulas / {t['distinct']} distinct calcs (light logic)")

    links = t["linked_workbooks"] + t["connections"]
    if links > 3:
        score += 2; ev.append(f"{links} external link(s)/connection(s)")
    elif links:
        score += 1; ev.append(f"{links} external link(s)/connection(s)")

    if len(wx.sheets) > 12:
        score += 1; ev.append(f"{len(wx.sheets)} sheets")
    if wx.has_vba:
        score += 2; ev.append("contains VBA macros")
    if wx.has_power_query:
        score += 1; ev.append("uses Power Query")

    verdict = "High" if score >= 5 else "Medium" if score >= 2 else "Low"
    # confidence rises as the score sits clearly inside a band, not on a boundary.
    conf = 0.6 + 0.1 * min(3, abs(score - 3.5))
    return Field.derived(verdict, min(conf, 0.9), ev)


def _logic_type(wx: WorkbookXray, t: dict) -> tuple[Field, Counter]:
    """Score the controlled workbook activity taxonomy from structural evidence.

    Returns the primary logic-type Field plus the raw score Counter, so
    :func:`_logic_types` can list every material activity without recomputing.
    """
    fns = t["functions"]
    ev: list[str] = []
    scores: Counter = Counter()

    lookup = sum(fns[f] for f in ("VLOOKUP", "HLOOKUP", "XLOOKUP", "MATCH", "INDEX"))
    agg = sum(fns[f] for f in ("SUM", "SUMIF", "SUMIFS", "AVERAGE", "COUNT", "COUNTIF", "SUBTOTAL"))
    textfn = sum(fns[f] for f in ("TEXT", "CONCATENATE", "CONCAT", "LEFT", "RIGHT", "MID", "TRIM", "SUBSTITUTE"))

    if lookup:
        scores["Data Transformation"] += lookup
        ev.append(f"{lookup} lookup/match call(s) — mapping or data preparation")
    if agg:
        scores["Calculation"] += agg
        ev.append(f"{agg} aggregation call(s)")
    if textfn:
        scores["Data Transformation"] += textfn
        ev.append(f"{textfn} text-manipulation call(s)")
    if wx.has_power_query:
        scores["Data Transformation"] += 5
        ev.append("Power Query present")
    if wx.pivot_cache_sources:
        scores["Reporting"] += 3
        ev.append(f"{len(wx.pivot_cache_sources)} pivot source(s) — reporting")

    if t["formulas"] and not scores:
        scores["Calculation"] += 1
        ev.append("formula calculations detected")

    if not scores:
        return Field.derived("Other", 0.4, ev or ["no dominant logic signal"]), scores
    primary = scores.most_common(1)[0][0]
    total = sum(scores.values())
    conf = 0.5 + 0.4 * (scores[primary] / total)
    ranked = [k for k, _ in scores.most_common()]
    return Field.derived(primary, min(conf, 0.9), ev + [f"ranked: {', '.join(ranked)}"]), scores


def _logic_types(scores: Counter, wx: WorkbookXray, tab_index: dict[str, TabFacts]) -> Field:
    """Every material activity the workbook supports, not just the top one —
    a workbook can be Calculation *and* Reconciliation *and* Reporting."""
    types: list[str] = []
    ev: list[str] = []
    if scores:
        top = max(scores.values())
        types = [k for k, v in scores.most_common() if v >= 0.25 * top]
        ev.append("scored from function mix: " + ", ".join(f"{k}={v:.1f}" for k, v in scores.most_common()))
    if any("Mapping" in (tf.roles or []) or tf.category == "Mapping" for tf in tab_index.values()):
        if "Data Transformation" not in types:
            types.append("Data Transformation")
        ev.append("one or more tabs carry the Mapping role")
    if any("Output" in (tf.roles or []) or tf.category == "Output" for tf in tab_index.values()):
        if "Reporting" not in types:
            types.append("Reporting")
        ev.append("one or more tabs carry the Output role")
    if any("Control Check" in tf.roles or "Validation" in tf.roles
           or tf.category in {"Control Check", "Validation"} for tf in tab_index.values()):
        types.append("Reconciliation / Control")
        ev.append("one or more tabs carry a Control Check or Validation role")
    if any("Input" in tf.roles or tf.category == "Input" for tf in tab_index.values()):
        types.append("Manual Input")
        ev.append("one or more tabs carry an Input role")
    allowed = {"Calculation", "Reporting", "Data Transformation",
               "Reconciliation / Control", "Manual Input", "Other"}
    types = list(dict.fromkeys(t for t in types if t in allowed))
    if not types:
        types = ["Other"]
    return Field.derived(types, 0.6, ev or ["no dominant activity signal"])


_BUSINESS_AREA = {
    "Calculation": "Calculation / modelling",
    "Reconciliation / Control": "Reconciliation / control",
    "Data Transformation": "Data preparation / transformation",
    "Reporting": "Reporting / MI",
    "Manual Input": "Manual data capture",
    "Other": "General-purpose / other",
}


def _business_area(logic_types: Field) -> Field:
    """Best-fit process classification — every detected activity joined,
    since a workbook can support more than one business process."""
    types = logic_types.value or ["Other"]
    areas = list(dict.fromkeys(_BUSINESS_AREA.get(t, t) for t in types))
    return Field.derived("; ".join(areas), logic_types.confidence or 0.5,
                         [f"mapped from logic types: {', '.join(types)}"])


def _key_inputs(wx: WorkbookXray, tab_index: dict[str, TabFacts],
                read_by: dict[str, set]) -> Field:
    grouped = rr.group_key_inputs(wx, tab_index, read_by)
    return Field.extracted(
        grouped,
        ["grouped by in-workbook sheets / lookup-mapping tables / data connections / "
         "external workbooks / other unresolved sources"],
    )


def _source_system(wx: WorkbookXray) -> Field:
    systems = sorted({str(c.get("name")).strip() for c in wx.connections
                      if c.get("name") and not re.search(
                          r"[/\\:@]|password=|server=|uid=|provider=|database=|token=", str(c.get("name")), re.I)})
    if systems:
        return Field.extracted(systems, ["connection display names only; connection strings are excluded"])
    return Field(value="Not established — confirm the source system with the process owner",
                 basis="needs_human",
                 evidence=["No safe source-system name was identified from formal connection metadata",
                           "reviewer question: Which system provides the source data?"])


def _euc_preparer(wx: WorkbookXray) -> Field:
    creator = (wx.core_props or {}).get("creator")
    last = (wx.core_props or {}).get("last_modified_by")
    names = [n for n in {creator, last} if n]
    if names:
        return Field(value=names, basis="derived", confidence=0.5,
                     evidence=["from document metadata (author/last-modified-by); "
                               "may reflect a template author, not the preparer"])
    return Field.pending("needs_human", "no author recorded in document metadata")


def _key_calculations(wx: WorkbookXray, t: dict, tab_index: dict[str, TabFacts]) -> Field:
    if not t["formulas"]:
        return Field.extracted([], ["workbook has no formulas"])
    descriptions = []
    for s in wx.sheets:
        d = rr.business_calculation_description(s, tab_index.get(s.name))
        if d:
            descriptions.append(d)
    return Field.extracted(
        {"business_descriptions": descriptions or [
            "Formula calculations are present; the business calculation purpose is not established from structure."
        ]},
        ["plain-language business descriptions only; technical formula patterns are in the technical appendix"],
    )


def _manual_intervention(wx: WorkbookXray) -> Field:
    d = rr.manual_intervention_file(wx)
    return Field(value=d["value"], basis="needs_human",
                 evidence=d["evidence"] + [f"reviewer question: {d['question']}"])


def _macros_links(wx: WorkbookXray) -> Field:
    d = rr.macro_details(wx)
    return Field.extracted(
        d,
        ["VBA / Power Query / data connections kept separate; external workbook "
         "dependencies are analysed under Key Inputs and Upstream Dependencies"],
    )


# ------------------------------------------------- AI findings (Step 3)


def _months_since(iso: str | None) -> int | None:
    """Whole months between an ISO timestamp and now; None if unparseable."""
    if not iso:
        return None
    try:
        s = str(iso).replace("Z", "+00:00")
        when = _dt.datetime.fromisoformat(s)
        if when.tzinfo is not None:
            when = when.replace(tzinfo=None)
    except (ValueError, TypeError):
        return None
    days = (_dt.datetime.now() - when).days
    return max(0, days // 30)


def _simplification(wx: WorkbookXray, tab_index: dict[str, TabFacts]) -> Field:
    candidates = rr.simplification_candidates(wx, tab_index)
    verdict = ("Potential simplification lead — confirm with process owner"
               if candidates else "No structural simplification lead detected")
    conf = 0.6
    return Field.derived(
        {"verdict": verdict, "candidates": candidates},
        min(conf, 0.85),
        [f"{len(candidates)} simplification candidate(s) across tabs, "
         "each identifying worksheet / activity / current complexity / proposed "
         "change / expected benefit"],
    )


def _automation(wx: WorkbookXray, tab_index: dict[str, TabFacts]) -> Field:
    candidates = rr.automation_candidates(wx, tab_index)
    verdict = ("Potential workflow lead — confirm manual steps with process owner"
               if candidates else
               "Not established — confirm the observed manual process with its owner")
    return Field.derived(
        {"verdict": verdict, "candidates": candidates},
        0.6,
        [f"{len(candidates)} explicitly labelled input/adjustment workflow lead(s); owner confirmation required"],
    )


def _retirement() -> Field:
    d = rr.retirement_assessment()
    return Field(value=d, basis="needs_human", evidence=[d["note"]])


def _reconciliation_logic(wx: WorkbookXray) -> Field:
    candidates = []
    questions = []
    for s in wx.sheets:
        d = rr.reconciliation_detail(s)
        if d.get("status") == "not_detected":
            continue
        if "reviewer_question" in d:
            questions.append(d["reviewer_question"])
            candidates.append({"tab": s.name, **d})
    if not candidates:
        return Field.derived(
            "No reconciliation pattern detected", 0.6,
            ["no lookup/variance logic or reconciliation-labelled headers on any tab"],
        )
    return Field.derived(
        {"candidates": candidates, "reviewer_questions": questions}, 0.6,
        ["structural candidates only; agreement, tolerance and exception handling require owner confirmation"],
    )


def _key_outputs(tabs: list["TabAssessment"]) -> Field:
    """List final business deliverables only when report-like evidence exists."""
    deliverables: list[str] = []
    output_terms = ("report", "summary", "dashboard", "statement", "pack", "submission",
                    "result", "claims paid", "claim performance", "reserve movement",
                    "management account", "p&l", "profit and loss", "balance sheet",
                    "income statement", "cash flow")
    for ta in tabs:
        roles = set(ta.tab_roles.value or []) | {ta.tab_category.value}
        name = ta.tab_name.value
        name_is_report = any(term in name.lower() for term in output_terms)
        if "Output" in roles and name_is_report:
            deliverables.append(name)
    ev = [f"{len(deliverables)} report-like final business deliverable(s) identified by worksheet name"]
    if not deliverables:
        ev.append("confirm final business deliverables with the process owner")
    value = {"final_business_deliverables": deliverables}
    if not deliverables:
        value["status"] = "Not established — confirm the final business deliverable with the process owner"
    return Field.derived(value, 0.6, ev)


# ------------------------------------------------------------------- entry


def assess_file(wx: WorkbookXray, tabs: list["TabAssessment"],
                 tab_index: dict[str, TabFacts],
                 read_by: dict[str, set] | None = None) -> FileAssessment:
    """Populate the file-level fields, grounded in the evidence plus the
    already-computed tab-level roles/categories."""
    t = _totals(wx)
    fa = FileAssessment()

    # Fact Assessment — extracted / derived
    fa.file_id = Field.extracted(wx.sha256[:12], ["content hash (stable across renames)"])
    fa.file_name = Field.extracted(wx.filename)
    fa.complexity = _complexity(wx, t)
    fa.key_inputs = _key_inputs(wx, tab_index, read_by or {})
    fa.source_system = _source_system(wx)
    fa.euc_preparer = _euc_preparer(wx)
    fa.key_outputs = _key_outputs(tabs)

    # Fact Assessment — deferred (narrative fields filled by the assessor step)
    fa.purpose_of_file = Field.pending("needs_llm", "narrative from structure + headers")
    fa.key_output_outcome = Field.pending("needs_llm", "business outcome the model supports")
    fa.process = Field.pending("needs_llm", "identify only when supported by workbook evidence")
    fa.sub_process = Field.pending("needs_llm", "identify only when supported by workbook evidence")
    fa.usage_frequency = Field.needs_human(
        "How often is this workbook used (e.g. daily/monthly/quarterly/ad hoc)?",
        ["not knowable from workbook structure"])
    fa.completion_timeline = Field.needs_human(
        "What is the preparation timeline for this workbook (e.g. month-end + N days)?",
        ["not knowable from workbook structure"])
    fa.output_recipient = Field.needs_human(
        "Who receives or consumes the output of this workbook?",
        ["downstream consumer is business context, not structural"])

    # Key AI Finding / Observation — heuristic (Step 3); corpus ones deferred
    fa.potential_duplication = Field.pending("needs_corpus", "needs the folder of workbooks")
    fa.similar_duplicate_files = Field.pending("needs_corpus", "needs the folder of workbooks")
    fa.potential_consolidation = Field.pending("needs_corpus", "needs the folder of workbooks")
    fa.potential_simplification = _simplification(wx, tab_index)
    fa.potential_automation = _automation(wx, tab_index)
    fa.potential_retirement = _retirement()

    # Workbook logic / Automation — extracted / derived
    fa.logic_type, logic_scores = _logic_type(wx, t)
    fa.logic_types = _logic_types(logic_scores, wx, tab_index)
    fa.business_area_process = _business_area(fa.logic_types)
    fa.key_calculations_logic = _key_calculations(wx, t, tab_index)
    fa.manual_intervention = _manual_intervention(wx)
    fa.macros_vba_external_links = _macros_links(wx)
    fa.reconciliation_logic = _reconciliation_logic(wx)

    return fa


# ------------------------------------------------------- tab-level (Step 2)

# Tabs whose best category scores below this are left "Uncertain" for the model,
# rather than forced into a bucket.
_CATEGORY_MIN_CONF = 0.55

_LOOKUP_FNS = ("VLOOKUP", "HLOOKUP", "XLOOKUP", "MATCH", "INDEX", "LOOKUP")
_COND_FNS = ("IF", "IFS", "IFERROR", "IFNA")


def _sheet_ratio(s) -> float:
    return s.formula_profile.get("total", 0) / max(1, s.populated_cells)


def _dependency_maps(wx: WorkbookXray) -> tuple[dict, dict]:
    """Return (reads, read_by) over in-workbook cross-sheet references.

    ``reads[name]``   = set of sheets whose cells this tab's formulas reference.
    ``read_by[name]`` = set of sheets that reference this tab (its downstream).
    """
    names = {s.name for s in wx.sheets}
    reads: dict[str, set] = {}
    read_by: dict[str, set] = {n: set() for n in names}
    for s in wx.sheets:
        refs = {
            r for r in s.formula_profile.get("referenced_sheets", {})
            if r in names and r != s.name
        }
        reads[s.name] = refs
        for r in refs:
            read_by.setdefault(r, set()).add(s.name)
    return reads, read_by


def _tab_scores(s, wx, has_downstream: bool) -> tuple[Counter, list[str]]:
    """Raw category scores for one tab — shared by the category Field and by
    role detection, so a tab can carry more than one role."""
    name = s.name.lower()
    fp = s.formula_profile
    total_f = fp.get("total", 0)
    ratio = _sheet_ratio(s)
    fns = dict(fp.get("top_functions", []))
    lookup = sum(fns.get(f, 0) for f in _LOOKUP_FNS)
    cond = sum(fns.get(f, 0) for f in _COND_FNS)
    kinds = [r.kind for r in s.regions]
    scores: Counter = Counter()
    ev: list[str] = []

    def hint(*words) -> bool:
        return any(w in name for w in words)

    # Output
    if hint("output", "report", "summary", "mi ", "mi_", "pack", "dashboard", "result"):
        scores["Output"] += 2; ev.append("name suggests output/reporting")
    if total_f and not has_downstream:
        scores["Output"] += 1; ev.append("terminal tab (not referenced by other tabs)")
    if wx.pivot_cache_sources:
        scores["Output"] += 0.5

    # Input
    if hint("input", "data", "raw", "extract", "source", "feed", "import"):
        scores["Input"] += 2; ev.append("name suggests input/data")
    if ratio < 0.1 and s.populated_cells > 10:
        scores["Input"] += 1.5; ev.append(f"low formula ratio {ratio:.0%} — mostly stored data")
        if has_downstream:
            scores["Input"] += 1; ev.append("read by other tabs — feeds the model")

    # Calculation
    if hint("calc", "working", "model", "engine", "compute"):
        scores["Calculation"] += 2; ev.append("name suggests calculation/working")
    if "calculation" in kinds:
        scores["Calculation"] += 2; ev.append("calculation region detected")
    if ratio > 0.4:
        scores["Calculation"] += 1.5; ev.append(f"high formula ratio {ratio:.0%}")

    # Mapping
    if hint("map", "lookup", "reference", "xref", "lob", "code"):
        scores["Mapping"] += 2; ev.append("name suggests mapping/lookup table")
    if total_f and lookup >= total_f * 0.3:
        scores["Mapping"] += 1.5; ev.append(f"{lookup} lookup/match call(s)")

    # Control Check (automated tie-out) vs Validation (manual review)
    if hint("check", "control", "tie", "recon", "reconcil", "proof", "agree"):
        scores["Control Check"] += 2.5; ev.append("name suggests a control/reconciliation check")
    if total_f and cond >= total_f * 0.4:
        scores["Control Check"] += 1; ev.append("comparison/conditional-heavy")
    if s.error_cells:
        scores["Control Check"] += 0.5; ev.append(f"{len(s.error_cells)} cached error cell(s)")
    if hint("valid", "review", "qa", "signoff", "sign-off", "approv"):
        scores["Validation"] += 2.5; ev.append("name suggests validation/review")

    return scores, ev


def _tab_category_field(scores: Counter, ev: list[str]) -> Field:
    """Auto-map to Input / Calculation / Mapping / Control Check / Validation /
    Output, flagging low-confidence tabs as Uncertain (needs_llm)."""
    if not scores:
        return Field(value="Uncertain", basis="needs_llm", confidence=0.3,
                     evidence=ev + ["no strong category signal — defer to model"])
    top, sc = scores.most_common(1)[0]
    conf = 0.4 + 0.5 * (sc / sum(scores.values()))
    ranked = ", ".join(f"{k}={v:.1f}" for k, v in scores.most_common())
    ev = ev + [f"scores: {ranked}"]
    if conf < _CATEGORY_MIN_CONF:
        return Field(value="Uncertain", basis="needs_llm", confidence=round(conf, 2),
                     evidence=ev + [f"top category below {_CATEGORY_MIN_CONF} confidence"])
    return Field.derived(top, conf, ev)


def _tab_roles_field(scores: Counter, category: Field, has_downstream: bool,
                     total_formulas: int) -> Field:
    """Every detected role, not just the primary one — a formula-heavy tab
    can be both a Calculation tab and the workbook's final Output."""
    cat_val = category.value
    if not scores:
        roles = [] if cat_val == "Uncertain" else [cat_val]
        return Field.derived(roles, category.confidence or 0.4,
                             ["single role from category (no score breakdown)"])
    top_score = max(scores.values())
    roles = [k for k, v in scores.items() if v > 0 and v >= 0.5 * top_score]
    ev = [f"roles scoring within 50% of the top score {top_score:.1f}: "
         + ", ".join(f"{k}={v:.1f}" for k, v in scores.most_common())]
    # A terminal, formula-bearing tab is a final output even when its primary
    # category is Calculation — do not force a Calculation/Output conflict.
    if cat_val == "Calculation" and not has_downstream and total_formulas and "Output" not in roles:
        roles.append("Output")
        ev.append("terminal formula-bearing tab added as Output — a calculation "
                 "can still be the final deliverable")
    if cat_val != "Uncertain" and cat_val not in roles:
        roles.insert(0, cat_val)
    elif cat_val in roles:
        roles = [cat_val] + [r for r in roles if r != cat_val]
    return Field.derived(roles, category.confidence or 0.5, ev)


def _tab_information(s, roles: list[str], has_upstream: bool, has_downstream: bool) -> Field:
    """data input / intermediate working / supporting schedule / mapping /
    control/check / final output — explained whenever a role combination
    (e.g. Calculation + Output) might otherwise look contradictory."""
    ratio = _sheet_ratio(s)
    kinds = {r.kind for r in s.regions}
    role_set = set(roles)
    ev = [f"formula ratio {ratio:.0%}",
          f"upstream={'yes' if has_upstream else 'no'}, "
          f"downstream={'yes' if has_downstream else 'no'}"]

    if "Mapping" in role_set:
        val = "mapping"
    elif "Control Check" in role_set or "Validation" in role_set:
        val = "control/check"
    elif "Output" in role_set:
        val = "final output"
        if ratio > 0.1:
            ev.append("terminal, formula-bearing tab classified as a final output rather "
                     "than an intermediate working — a calculation can still be the "
                     "deliverable")
    elif s.populated_cells and ratio < 0.1:
        val = "data input" if has_downstream else "supporting schedule"
    elif has_downstream and (has_upstream or ratio > 0.2):
        val = "intermediate working"
    elif ratio > 0.1 and not has_downstream:
        val = "final output"
    elif kinds <= {"notes", "title", "key_value", "unknown"}:
        val = "supporting schedule"
    else:
        val = "intermediate working"
    return Field.derived(val, 0.6, ev)


def _tab_key_calc(s, tf: TabFacts | None) -> Field:
    fp = s.formula_profile
    if not fp.get("total"):
        return Field.extracted([], ["tab has no formulas"])
    desc = rr.business_calculation_description(s, tf)
    value = {"business_description": desc or
             "Formula calculations are present; the business purpose is not established from structure."}
    ev = ["plain-language business description only; formula patterns remain in the technical appendix"]
    return Field.extracted(value, ev)


def _tab_upstream(s, wx, reads_set: set, tab_index: dict[str, TabFacts]) -> Field:
    """Separate what this tab depends on into in-workbook sheets, lookup/mapping
    tables, data connections and external workbooks, without exposing local
    paths or claiming that links can be assigned to a specific source."""
    in_workbook = sorted(reads_set)
    lookup_tables = [r for r in in_workbook
                     if "Mapping" in (set(tab_index[r].roles) if r in tab_index else set())]
    other_sheets = [r for r in in_workbook if r not in lookup_tables]

    external: list[dict] = []
    if s.formula_profile.get("external_count", 0):
        external = [{"file_name": "External workbook reference",
                     "source_status": "Observed on this worksheet; exact workbook not determinable"}]

    uses_connection = bool(wx.connections) and (
        s.formula_profile.get("total", 0) > 0
        or any(p.split("!", 1)[0] == s.name for p in wx.pivot_cache_sources)
    )
    data_conn = ([f"Formal data connection {i + 1}"
                  for i, _ in enumerate(wx.connections)] if uses_connection else [])

    value = {
        "in_workbook_sheets": other_sheets or ["none"],
        "lookup_mapping_tables": lookup_tables or ["none"],
        "data_connections": data_conn or ["none"],
        "external_workbooks": external or ["none"],
    }
    if not (in_workbook or external or data_conn):
        return Field.extracted(value, ["self-contained tab — no in-workbook, lookup, "
                                       "connection or external dependency detected"])
    return Field.extracted(
        value,
        ["dependencies grouped by kind; private source paths and connection strings are excluded"],
    )


def _tab_downstream(read_by_set: set, tab_index: dict[str, TabFacts]) -> Field:
    in_wb = [{"sheet": r, "role": rr.classify_consumer_role(r, tab_index)}
            for r in sorted(read_by_set)]
    return Field(
        value={"in_workbook": in_wb or ["none within this workbook"],
               "other_files": None},
        basis="derived" if in_wb else "extracted",
        confidence=1.0,
        evidence=["each in-workbook consumer is classified as intermediate working / "
                 "supporting schedule / control-check / final output; in-workbook edges "
                 "are exact, cross-file consumers need the corpus (needs_corpus)"],
    )


def _human_validation(s, tab_index: dict[str, TabFacts]) -> tuple[Field, Field]:
    """Only material issues trigger review — cached errors, unresolved
    external sources, material hardcodes, volatile calcs affecting outputs,
    uncertain final-output classification, or unresolved reconciliation
    logic. Region-detection uncertainty alone is not a trigger."""
    reasons: list[str] = []
    tech: list[str] = []
    fp = s.formula_profile

    if s.error_cells:
        err_types = sorted({rr._parse_error_entry(e)[2] for e in s.error_cells})
        reasons.append(
            f"After refreshing the sources and recalculating, do these errors remain, "
            f"and which reporting output changes? ({len(s.error_cells)} cached error "
            f"cell(s) on '{s.name}': {', '.join(err_types)}.)"
        )
        tech.append(f"{len(s.error_cells)} cached error cell(s): "
                    + ", ".join(s.error_cells[:3]))

    if fp.get("external_count"):
        reasons.append(
            f"Can the external source(s) referenced by '{s.name}' be verified and "
            f"confirmed current — are they still available and refreshed?"
        )
        tech.append("references external workbook(s) — values not verifiable here")

    if fp.get("hardcoded_literal_count", 0) > 3:
        reasons.append(
            f"Are the fixed values in the formulas on '{s.name}' approved assumptions, "
            f"and who maintains them?"
        )
        tech.append(f"{fp['hardcoded_literal_count']} formula(s) with a hardcoded number")

    if fp.get("volatile_count"):
        reasons.append(
            f"'{s.name}' uses volatile function(s) whose result depends on recalculation "
            f"state — confirm the output is current before relying on it."
        )
        tech.append(f"{fp['volatile_count']} volatile function(s)")

    recon = rr.reconciliation_detail(s)
    if "reviewer_question" in recon:
        reasons.append(recon["reviewer_question"])
        tech.append("lookup/variance logic present but source/target/key not "
                    "resolvable from headers")

    tf = tab_index.get(s.name)
    if tf and tf.category == "Uncertain":
        reasons.append(
            f"Confirm whether '{s.name}' is a final business output or an intermediate "
            f"working schedule."
        )
        tech.append("tab category confidence below threshold")

    manual = rr.manual_intervention_tab(s)
    manual_label_evidence = manual.get("evidence") or []
    if (manual["basis"] == "needs_human" and fp.get("total", 0) == 0
            and s.populated_cells > 20 and manual_label_evidence):
        reasons.append(manual["question"])
        tech.append("manual-entry label observed; actual workflow requires owner confirmation")

    required = bool(reasons)
    yn = Field.derived(
        "Y" if required else "N", 0.7 if required else 0.6,
        [f"{len(reasons)} material issue(s)" if required else
         "no material issue (cached error, unverifiable external source, material "
         "hardcode, volatile calculation, uncertain output classification, or "
         "unresolved reconciliation logic) detected"],
    )
    reason = Field.derived(
        reasons or ["none — no material trigger from automated checks"],
        0.7, tech or ["no material red flags"],
    )
    return yn, reason


def assess_tab_core(s, wx, has_downstream: bool) -> tuple[Field, Field]:
    """Category + roles only — the part other tabs' dependency grouping needs,
    computed before the full tab index exists."""
    scores, ev = _tab_scores(s, wx, has_downstream)
    category = _tab_category_field(scores, ev)
    roles = _tab_roles_field(scores, category, has_downstream,
                             s.formula_profile.get("total", 0))
    return category, roles


def assess_tab(s, wx, reads_set, read_by_set, category: Field, roles: Field,
               tab_index: dict[str, TabFacts]) -> TabAssessment:
    ta = TabAssessment()
    has_up = bool(reads_set)
    has_down = bool(read_by_set)
    ta.tab_name = Field.extracted(s.name)
    ta.tab_visibility = Field.extracted(
        visibility_of(s.state),
        [f"original workbook position {s.position + 1} of {len(wx.sheets)} (preserved "
         f"as evidence even though visible tabs are presented first)"],
    )
    ta.tab_category = category
    ta.tab_roles = roles
    ta.tab_information_analysis = _tab_information(s, roles.value or [], has_up, has_down)
    ta.key_calculation_transformation_logic = _tab_key_calc(s, tab_index.get(s.name))
    ta.upstream_dependencies = _tab_upstream(s, wx, reads_set, tab_index)
    ta.downstream_dependencies = _tab_downstream(read_by_set, tab_index)
    ta.human_validation_required, ta.validation_reason = _human_validation(s, tab_index)
    ta.tab_purpose_description = Field.pending(
        "needs_llm", "narrative purpose from headers + category + calculations")
    return ta


def _apply_narrative(a: Assessment, narr, basis: str, label: str) -> None:
    """Write an assessor's :class:`~excel_xray.narrative.Narrative` onto the
    narrative fields, tagging each with the assessor's basis. ``key_outputs``
    is deterministic (Step 12) and is not overwritten by the narrative step."""
    ev = [f"{basis} by {label}"]

    def put(fld, value):
        if value is not None:
            fld.value = value
            fld.basis = basis
            fld.confidence = 0.7 if basis == "inferred" else 0.4
            fld.evidence = ev + (
                [] if basis == "inferred"
                else ["offline template — enable the LLM assessor for a considered answer"]
            )

    put(a.file.purpose_of_file, narr.purpose_of_file)
    put(a.file.key_output_outcome, narr.key_output_outcome)
    # Accept process labels only when the assessor supplies exact source text
    # that occurs in the structural evidence bundle.
    business_evidence = {
        str(item).casefold()
        for item in getattr(narr, "_evidence_bundle", {}).get("business_evidence", [])
    }
    generic_terms = {
        "input", "output", "calculation", "model", "data", "summary", "report",
        "dashboard", "working", "worksheet", "table", "amount", "date", "total",
        "formula", "sheet", "unknown", "general", "euc",
    }
    def supported_evidence(items):
        if not items:
            return False
        for item in items:
            exact = str(item).strip().casefold()
            terms = {token.casefold() for token in re.findall(r"[A-Za-z][A-Za-z0-9&/-]*", exact)}
            if exact not in business_evidence or not (terms - generic_terms):
                return False
        return True
    for fld, value, evidence in (
        (a.file.process, getattr(narr, "process", None), getattr(narr, "process_evidence", [])),
        (a.file.sub_process, getattr(narr, "sub_process", None), getattr(narr, "sub_process_evidence", [])),
    ):
        supported = bool(value and supported_evidence(evidence))
        if supported:
            put(fld, value)
            fld.evidence = ev + [f"source evidence: {item}" for item in evidence]
        else:
            fld.value = "Not established — confirm with the process owner"
            fld.basis = "needs_human"
            fld.confidence = None
            fld.evidence = ["No direct supporting workbook evidence was supplied",
                            "reviewer question: What business process and sub-process does this workbook support?"]
    for ta in a.tabs:
        put(ta.tab_purpose_description, narr.tabs.get(ta.tab_name.value))


def _order_visible_first(tabs: list[TabAssessment]) -> list[TabAssessment]:
    """Present and analyse visible worksheets first; original position stays
    on the tab as evidence (tab_visibility), never lost, just not the sort key."""
    def key(ta):
        visible = ta.tab_visibility.value == "Visible"
        pos_ev = ta.tab_visibility.evidence[0] if ta.tab_visibility.evidence else ""
        # position was recorded as "...position N of M..."; fall back to name order.
        try:
            pos = int(pos_ev.split("position ", 1)[1].split(" of", 1)[0])
        except (IndexError, ValueError):
            pos = 0
        return (0 if visible else 1, pos)
    return sorted(tabs, key=key)


def assess(wx: WorkbookXray, assessor=None) -> Assessment:
    """Full assessment: deterministic fields (Steps 1-3) plus the narrative
    fields (Step 4) filled by ``assessor`` — the offline stub by default, so the
    pipeline stays network-free unless a model-backed assessor is passed."""
    from .narrative import OfflineAssessor, build_bundle

    reads, read_by = _dependency_maps(wx)

    # Pass 1: category + roles only, needed to build the cross-tab index.
    core = {s.name: assess_tab_core(s, wx, bool(read_by.get(s.name))) for s in wx.sheets}
    tab_index: dict[str, TabFacts] = {}
    for s in wx.sheets:
        category, roles = core[s.name]
        has_down = bool(read_by.get(s.name))
        info = _tab_information(s, roles.value or [], bool(reads.get(s.name)), has_down)
        tab_index[s.name] = TabFacts(
            name=s.name, position=s.position, visibility=visibility_of(s.state),
            category=category.value, roles=list(roles.value or []),
            information=info.value,
        )

    # Pass 2: full tab assessment, with the index available for dependency
    # grouping, downstream-role classification and validation questions.
    tabs = [
        assess_tab(s, wx, reads.get(s.name, set()), read_by.get(s.name, set()),
                  *core[s.name], tab_index)
        for s in wx.sheets
    ]
    tabs = _order_visible_first(tabs)

    fa = assess_file(wx, tabs, tab_index, read_by)
    review = ReviewFindings(
        error_groups=rr.build_error_groups(wx, tab_index, read_by),
        macro_details=fa.macros_vba_external_links.value,
        simplification_candidates=fa.potential_simplification.value.get("candidates", []),
        automation_candidates=fa.potential_automation.value.get("candidates", []),
        key_inputs_grouped=fa.key_inputs.value,
    )
    review.error_details = rr.error_detail_rows(review.error_groups)
    review.hidden_summary = rr.build_hidden_summary(wx, tab_index, read_by)

    a = Assessment(file=fa, tabs=tabs, review=review, scan={
        "status": wx.parse_status,
        "error": (safe_scan_message("; ".join(wx.warnings[:3]) or
                                    "Workbook scan was partial; review diagnostics.", wx.path)
                   if wx.parse_status == "partial" else None),
        "total_sheets": len(wx.sheets),
        "hidden_sheets": sum(s.state != "visible" for s in wx.sheets),
    })

    assessor = assessor or OfflineAssessor()
    evidence_bundle = build_bundle(a, wx)
    narr = assessor.narrate(evidence_bundle)
    narr._evidence_bundle = evidence_bundle
    _apply_narrative(a, narr, assessor.basis, assessor.label)
    for candidate in a.review.simplification_candidates + a.review.automation_candidates:
        candidate["process"] = fa.process.value
        candidate["sub_process"] = fa.sub_process.value
    return a


def to_dict(a: Assessment) -> dict:
    file_data = asdict(a.file)
    # Keep one controlled Logic Types output; the internal primary score is
    # not exported as a duplicate singular "Logic Type" field.
    file_data.pop("logic_type", None)
    return {"scan": a.scan, "file": file_data,
            "tabs": [asdict(t) for t in a.tabs], "review": asdict(a.review)}
