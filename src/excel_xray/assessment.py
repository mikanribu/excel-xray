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


def _tab_category(s, wx, has_downstream: bool) -> Field:
    """Auto-map to Input / Calculation / Mapping / Control Check / Validation /
    Output, flagging low-confidence tabs as Uncertain (needs_llm)."""
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

    if not scores:
        return Field(value="Uncertain", basis="needs_llm", confidence=0.3,
                     evidence=ev + ["no strong category signal — defer to model"])
    top, sc = scores.most_common(1)[0]
    conf = 0.4 + 0.5 * (sc / sum(scores.values()))
    ranked = ", ".join(f"{k}={v:.1f}" for k, v in scores.most_common())
    ev.append(f"scores: {ranked}")
    if conf < _CATEGORY_MIN_CONF:
        return Field(value="Uncertain", basis="needs_llm", confidence=round(conf, 2),
                     evidence=ev + [f"top category below {_CATEGORY_MIN_CONF} confidence"])
    return Field.derived(top, conf, ev)


def _tab_information(s, has_upstream: bool, has_downstream: bool) -> Field:
    """data input / intermediate workings / output / supporting."""
    ratio = _sheet_ratio(s)
    kinds = {r.kind for r in s.regions}
    if s.populated_cells and ratio < 0.1:
        val = "data input" if has_downstream else "supporting"
    elif has_downstream and (has_upstream or ratio > 0.2):
        val = "intermediate workings"
    elif ratio > 0.1 and not has_downstream:
        val = "output"
    elif kinds <= {"notes", "title", "key_value", "unknown"}:
        val = "supporting"
    else:
        val = "intermediate workings"
    ev = [f"formula ratio {ratio:.0%}",
          f"upstream={'yes' if has_upstream else 'no'}, "
          f"downstream={'yes' if has_downstream else 'no'}"]
    return Field.derived(val, 0.6, ev)


def _tab_key_calc(s) -> Field:
    fp = s.formula_profile
    if not fp.get("total"):
        return Field.extracted([], ["tab has no formulas"])
    shapes = [f"{sk}  (×{n})" for sk, n in fp.get("top_skeletons", [])[:6]]
    fns = [f"{fn}×{n}" for fn, n in fp.get("top_functions", [])[:8]]
    return Field.extracted(
        {"top_functions": fns, "top_formula_shapes": shapes},
        [f"{fp['total']} formulas, {fp.get('distinct_skeletons', 0)} distinct shapes"],
    )


def _tab_upstream(s, wx, reads_set: set) -> Field:
    deps: list[str] = []
    deps += [f"sheet: {r}" for r in sorted(reads_set)]
    if s.formula_profile.get("external_count", 0):
        deps += [f"external workbook: {x}" for x in wx.external_links] or \
                ["external workbook link(s)"]
    fns = dict(s.formula_profile.get("top_functions", []))
    if any(fns.get(f) for f in _LOOKUP_FNS):
        deps.append("lookup tables (VLOOKUP/INDEX/MATCH targets)")
    return Field.extracted(deps or ["none — self-contained tab"])


def _tab_downstream(s, read_by_set: set) -> Field:
    in_wb = [f"sheet: {r}" for r in sorted(read_by_set)]
    return Field(
        value={"in_workbook": in_wb or ["none within this workbook"],
               "other_files": None},
        basis="derived" if in_wb else "extracted",
        confidence=1.0,
        evidence=["in-workbook edges are exact; cross-file consumers need the "
                  "corpus (needs_corpus)"],
    )


def _human_validation(s, wx) -> tuple[Field, Field]:
    reasons: list[str] = []
    fp = s.formula_profile
    if s.error_cells:
        reasons.append(f"{len(s.error_cells)} cached error cell(s): "
                       + ", ".join(s.error_cells[:3]))
    if fp.get("hardcoded_literal_count"):
        reasons.append(f"{fp['hardcoded_literal_count']} formula(s) with a hardcoded "
                       "number (buried assumption)")
    if fp.get("volatile_count"):
        reasons.append(f"{fp['volatile_count']} volatile function(s) — result depends "
                       "on recalculation state")
    if fp.get("external_count"):
        reasons.append("references external workbook(s) — values not verifiable here")
    low_conf = [r.ref for r in s.regions if r.detect_confidence < 0.70]
    if low_conf:
        reasons.append(f"ambiguous region layout at {', '.join(low_conf)}")
    if wx.has_vba:
        reasons.append("workbook contains VBA — logic may be hidden from the grid")

    required = bool(reasons)
    yn = Field.derived("Y" if required else "N", 0.7 if required else 0.55,
                       [f"{len(reasons)} red flag(s)" if required
                        else "no automated red flags on this tab"])
    reason = Field.derived(
        reasons or ["none from automated checks; still subject to normal sampling"],
        0.7, ["triggers are heuristic; a clean tab is not a guarantee"],
    )
    return yn, reason


def assess_tab(s, wx, reads_set, read_by_set) -> TabAssessment:
    ta = TabAssessment()
    has_up = bool(reads_set)
    has_down = bool(read_by_set)
    ta.tab_name = Field.extracted(s.name,
                                  ["hidden sheet" for _ in [1] if s.state != "visible"])
    ta.tab_category = _tab_category(s, wx, has_down)
    ta.tab_information_analysis = _tab_information(s, has_up, has_down)
    ta.key_calculation_transformation_logic = _tab_key_calc(s)
    ta.upstream_dependencies = _tab_upstream(s, wx, reads_set)
    ta.downstream_dependencies = _tab_downstream(s, read_by_set)
    ta.human_validation_required, ta.validation_reason = _human_validation(s, wx)
    ta.tab_purpose_description = Field.pending(
        "needs_llm", "narrative purpose from headers + category + calculations")
    return ta


def assess(wx: WorkbookXray) -> Assessment:
    """Full assessment: file level (Step 1) and every tab (Step 2)."""
    reads, read_by = _dependency_maps(wx)
    tabs = [
        assess_tab(s, wx, reads.get(s.name, set()), read_by.get(s.name, set()))
        for s in wx.sheets
    ]
    return Assessment(file=assess_file(wx), tabs=tabs)


def to_dict(a: Assessment) -> dict:
    return {"file": asdict(a.file), "tabs": [asdict(t) for t in a.tabs]}
