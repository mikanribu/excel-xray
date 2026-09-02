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

Step 1 implements the file-level ``extracted`` and ``derived`` fields; the rest
are returned with the correct basis and a ``null`` value, ready for later steps.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from collections import Counter

from .scan import WorkbookXray

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


# ------------------------------------------------------------- schema shapes


@dataclass
class FileAssessment:
    """File-level summary — the top block of the target schema."""

    # Fact Assessment
    file_id: Field = field(default_factory=Field)
    file_name: Field = field(default_factory=Field)
    business_area_process: Field = field(default_factory=Field)
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
    key_calculations_logic: Field = field(default_factory=Field)
    reconciliation_logic: Field = field(default_factory=Field)
    manual_intervention: Field = field(default_factory=Field)
    macros_vba_external_links: Field = field(default_factory=Field)


@dataclass
class TabAssessment:
    """Tab-level detail — the bottom block of the target schema (Step 2)."""

    tab_name: Field = field(default_factory=Field)
    tab_category: Field = field(default_factory=Field)
    tab_purpose_description: Field = field(default_factory=Field)
    tab_information_analysis: Field = field(default_factory=Field)
    key_calculation_transformation_logic: Field = field(default_factory=Field)
    upstream_dependencies: Field = field(default_factory=Field)
    downstream_dependencies: Field = field(default_factory=Field)
    human_validation_required: Field = field(default_factory=Field)
    validation_reason: Field = field(default_factory=Field)


@dataclass
class Assessment:
    file: FileAssessment
    tabs: list[TabAssessment] = field(default_factory=list)


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


def _logic_type(wx: WorkbookXray, t: dict) -> Field:
    """Reconciliation / Calculation / Data Transformation / Manual Input / Reporting."""
    fns = t["functions"]
    ev: list[str] = []
    scores: Counter = Counter()

    lookup = sum(fns[f] for f in ("VLOOKUP", "HLOOKUP", "XLOOKUP", "MATCH", "INDEX"))
    agg = sum(fns[f] for f in ("SUM", "SUMIF", "SUMIFS", "AVERAGE", "COUNT", "COUNTIF", "SUBTOTAL"))
    textfn = sum(fns[f] for f in ("TEXT", "CONCATENATE", "CONCAT", "LEFT", "RIGHT", "MID", "TRIM", "SUBSTITUTE"))
    cond = sum(fns[f] for f in ("IF", "IFS", "IFERROR"))

    if lookup:
        scores["Reconciliation"] += lookup + cond * 0.5
        ev.append(f"{lookup} lookup/match call(s) — matching across sources")
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

    formula_ratio = t["formulas"] / max(1, t["cells"])
    if formula_ratio < 0.05:
        scores["Manual Input"] += 4
        ev.append(f"only {formula_ratio:.0%} of cells are formulas — largely manual")

    if not scores:
        return Field.derived("Other", 0.4, ev or ["no dominant logic signal"])
    primary = scores.most_common(1)[0][0]
    total = sum(scores.values())
    conf = 0.5 + 0.4 * (scores[primary] / total)
    ranked = [k for k, _ in scores.most_common()]
    return Field.derived(primary, min(conf, 0.9), ev + [f"ranked: {', '.join(ranked)}"])


def _key_inputs(wx: WorkbookXray) -> Field:
    inputs: list[str] = []
    inputs += [f"linked workbook: {x}" for x in wx.external_links]
    inputs += [f"connection: {c.get('name') or c.get('type') or 'unnamed'}"
               for c in wx.connections]
    inputs += [f"pivot source: {p}" for p in wx.pivot_cache_sources]
    if inputs:
        return Field.extracted(inputs)
    return Field(value=None, basis="needs_llm",
                 evidence=["no external links/connections — inputs are manual or "
                           "embedded; naming the business datasets needs the model"])


def _source_system(wx: WorkbookXray) -> Field:
    systems: list[str] = []
    for c in wx.connections:
        s = c.get("connection_string") or c.get("command") or c.get("description")
        if s:
            systems.append(str(s))
    systems += [p for p in wx.pivot_cache_sources if ("!" not in p and p)]
    if systems:
        return Field.extracted(sorted(set(systems)))
    return Field.pending("needs_human",
                         "no data connections declare a source system")


def _euc_preparer(wx: WorkbookXray) -> Field:
    creator = (wx.core_props or {}).get("creator")
    last = (wx.core_props or {}).get("last_modified_by")
    names = [n for n in {creator, last} if n]
    if names:
        return Field(value=names, basis="derived", confidence=0.5,
                     evidence=["from document metadata (author/last-modified-by); "
                               "may reflect a template author, not the preparer"])
    return Field.pending("needs_human", "no author recorded in document metadata")


def _key_calculations(wx: WorkbookXray, t: dict) -> Field:
    if not t["formulas"]:
        return Field.extracted([], ["workbook has no formulas"])
    skeletons: Counter = Counter()
    for s in wx.sheets:
        for sk, n in s.formula_profile.get("top_skeletons", []):
            skeletons[sk] += n
    top = [f"{sk}  (×{n})" for sk, n in skeletons.most_common(8)]
    top_fns = [f"{fn}×{n}" for fn, n in t["functions"].most_common(8)]
    return Field.extracted(
        {"top_functions": top_fns, "top_formula_shapes": top},
        [f"{t['formulas']:,} formulas reduce to {t['distinct']} distinct shapes"],
    )


def _manual_intervention(wx: WorkbookXray, t: dict) -> Field:
    ev = [f"{t['non_formula_cells']:,} non-formula populated cells (potential manual entry)"]
    if t["hardcoded"]:
        ev.append(f"{t['hardcoded']} formula(s) contain a hardcoded number (buried input)")
    manual_tabs = [
        s.name for s in wx.sheets
        if s.populated_cells and s.formula_profile.get("total", 0) / s.populated_cells < 0.05
    ]
    if manual_tabs:
        ev.append(f"largely-manual tab(s): {', '.join(manual_tabs)}")
    ratio = t["non_formula_cells"] / max(1, t["cells"])
    return Field.derived(
        "High" if ratio > 0.7 else "Medium" if ratio > 0.3 else "Low",
        0.6, ev,
    )


def _macros_links(wx: WorkbookXray) -> Field:
    parts: list[str] = []
    if wx.has_vba:
        parts.append("VBA macros present")
    if wx.has_power_query:
        parts.append("Power Query / DataMashup present")
    if wx.external_links:
        parts.append(f"{len(wx.external_links)} external workbook link(s)")
    if wx.connections:
        parts.append(f"{len(wx.connections)} data connection(s)")
    return Field.extracted(parts or ["none detected"])


# ------------------------------------------------------------------- entry


def assess_file(wx: WorkbookXray) -> FileAssessment:
    """Step 1: populate the file-level fields we can ground in the evidence."""
    t = _totals(wx)
    fa = FileAssessment()

    # Fact Assessment — extracted / derived
    fa.file_id = Field.extracted(wx.sha256[:12], ["content hash (stable across renames)"])
    fa.file_name = Field.extracted(wx.filename)
    fa.complexity = _complexity(wx, t)
    fa.key_inputs = _key_inputs(wx)
    fa.source_system = _source_system(wx)
    fa.euc_preparer = _euc_preparer(wx)

    # Fact Assessment — deferred
    fa.business_area_process = Field.pending("needs_llm", "classify from purpose + logic")
    fa.purpose_of_file = Field.pending("needs_llm", "narrative from structure + headers")
    fa.key_output_outcome = Field.pending("needs_llm", "business outcome the model supports")
    fa.key_outputs = Field.pending("needs_llm", "name outputs from terminal/reporting tabs")
    fa.usage_frequency = Field.pending("needs_human", "not knowable from the file")
    fa.completion_timeline = Field.pending("needs_human", "not knowable from the file")
    fa.output_recipient = Field.pending("needs_human", "downstream consumer is external context")

    # Key AI Finding / Observation — later steps
    fa.potential_duplication = Field.pending("needs_corpus", "needs the folder of workbooks")
    fa.similar_duplicate_files = Field.pending("needs_corpus", "needs the folder of workbooks")
    fa.potential_consolidation = Field.pending("needs_corpus", "needs the folder of workbooks")
    fa.potential_simplification = Field.pending("needs_llm", "heuristic + judgement (Step 3)")
    fa.potential_automation = Field.pending("needs_llm", "heuristic + judgement (Step 3)")
    fa.potential_retirement = Field.pending("needs_llm", "heuristic + judgement (Step 3)")

    # Workbook logic / Automation — extracted / derived
    fa.logic_type = _logic_type(wx, t)
    fa.key_calculations_logic = _key_calculations(wx, t)
    fa.manual_intervention = _manual_intervention(wx, t)
    fa.macros_vba_external_links = _macros_links(wx)
    fa.reconciliation_logic = Field.pending(
        "needs_llm", "matching criteria/tolerances/ageing (Step 3 heuristic)")

    return fa


def assess(wx: WorkbookXray) -> Assessment:
    """Full assessment. Step 1 fills the file level; tabs come in Step 2."""
    return Assessment(file=assess_file(wx), tabs=[])


def to_dict(a: Assessment) -> dict:
    return {"file": asdict(a.file), "tabs": [asdict(t) for t in a.tabs]}
