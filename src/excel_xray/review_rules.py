"""Review-rules layer: turns raw evidence into the structured, reviewer-facing
findings the EUC review schema asks for.

Kept separate from :mod:`excel_xray.assessment` because these rules are meaty —
several are table-shaped and reused by both the file- and tab-level fields and
by the HTML/Excel report exporters: error grouping, hidden-sheet purpose,
dependency grouping (in-workbook / mapping / connections / external
workbooks), business-calculation descriptions, targeted validation questions,
and the automation/simplification/retirement/reconciliation rules that must
not be inferred from a single raw percentage.

Nothing here reads cell values — only structure already captured on
:class:`~excel_xray.scan.WorkbookXray` (headers, formula shapes, function
counts, names, dependency edges).
"""

from __future__ import annotations

import io
import posixpath
import re
import zipfile
from collections import defaultdict
from dataclasses import dataclass, field

from .util import range_boundaries

# --------------------------------------------------------------- tab index


@dataclass
class TabFacts:
    """Cheap, cross-referenceable facts about one sheet, built in an
    assessment's first pass so the second pass can classify dependencies,
    errors and hidden-sheet purpose without re-deriving category/role."""

    name: str
    position: int
    visibility: str          # Visible | Hidden | Very Hidden
    category: str            # primary tab_category
    roles: list[str] = field(default_factory=list)
    information: str = ""    # tab_information_analysis value


def visibility_of(state: str) -> str:
    return {"visible": "Visible", "hidden": "Hidden",
            "veryHidden": "Very Hidden"}.get(state, "Visible")


def _roles_of(tf: TabFacts | None) -> set:
    if tf is None:
        return set()
    return set(tf.roles) | {tf.category}


# --------------------------------------------------------------- error groups

_ERR_RE = re.compile(r"^([A-Za-z]{1,3}\d{1,7}) (#[A-Z0-9/!?]+)$")


def _parse_error_entry(entry: str) -> tuple[str, str, str]:
    """``'SheetName!C5 #DIV/0!'`` -> ``('SheetName', 'C5', '#DIV/0!')``."""
    sheet, _, rest = entry.partition("!")
    m = _ERR_RE.match(rest)
    if not m:
        return sheet, rest, "#UNKNOWN"
    return sheet, m.group(1), m.group(2)


def _cell_region(sheet_obj, col: int, row: int):
    if sheet_obj is None:
        return None
    for r in sheet_obj.regions:
        if r.top <= row <= r.bottom and r.left <= col <= r.right:
            return r
    return None


def area_type_for(tf: TabFacts | None) -> str:
    """Input / calculation-working / supporting schedule / final output /
    control-check / mapping / unknown, from a tab's detected roles."""
    if tf is None:
        return "unknown"
    roles = _roles_of(tf)
    if "Output" in roles:
        return "final output"
    if "Control Check" in roles or "Validation" in roles:
        return "control / check"
    if "Mapping" in roles:
        return "mapping / lookup table"
    if "Input" in roles:
        return "input"
    if "Calculation" in roles:
        return "calculation / working area"
    if tf.information == "supporting schedule":
        return "supporting schedule"
    return "unknown"


@dataclass
class ErrorGroup:
    error_type: str
    count: int
    sheet: str
    region_ref: str
    area_type: str
    downstream_outputs: list[str]
    recalc_required: bool
    reviewer_action: str
    sample_cells: list[str]


def build_error_groups(wx, tab_index: dict[str, TabFacts],
                        read_by: dict[str, set]) -> list[ErrorGroup]:
    """Group raw cached-error cells by (sheet, error type, region) so the
    report header can summarise instead of listing every cell reference."""
    sheets_by_name = {s.name: s for s in wx.sheets}
    buckets: dict[tuple, list[tuple[int, int, str]]] = defaultdict(list)
    for s in wx.sheets:
        for entry in s.error_cells:
            sheet_name, cellref, err = _parse_error_entry(entry)
            try:
                col, row, _mc, _mr = range_boundaries(cellref)
            except Exception:
                continue
            region = _cell_region(sheets_by_name.get(sheet_name), col, row)
            region_ref = region.ref if region else cellref
            buckets[(sheet_name, err, region_ref)].append((row, col, cellref))

    groups: list[ErrorGroup] = []
    for (sheet_name, err, region_ref), cells in buckets.items():
        tf = tab_index.get(sheet_name)
        area = area_type_for(tf)
        downstream_names = sorted(read_by.get(sheet_name, set()))
        downstream_outputs = [
            n for n in downstream_names if "Output" in _roles_of(tab_index.get(n))
        ]
        if not downstream_outputs and area == "final output":
            downstream_outputs = [sheet_name]

        broken_ref = err == "#REF!"
        recalc_required = not broken_ref
        if broken_ref:
            action = (
                f"{len(cells)} {err} error(s) on {sheet_name}!{region_ref} indicate a "
                f"broken reference — repair the formula (recalculation will not fix this)"
                + (f"; check impact on {', '.join(downstream_outputs)}."
                   if downstream_outputs else
                   "; no downstream output currently identified for this area.")
            )
        else:
            action = (
                f"After refreshing linked sources and recalculating, confirm whether "
                f"these {len(cells)} {err} error(s) on {sheet_name}!{region_ref} persist"
                + (f", and check impact on {', '.join(downstream_outputs)}."
                   if downstream_outputs else
                   "; no downstream output currently identified for this area.")
            )
        groups.append(ErrorGroup(
            error_type=err, count=len(cells), sheet=sheet_name, region_ref=region_ref,
            area_type=area, downstream_outputs=downstream_outputs,
            recalc_required=recalc_required, reviewer_action=action,
            sample_cells=[c for _, _, c in sorted(cells)[:25]],
        ))
    groups.sort(key=lambda g: (-g.count, g.sheet, g.error_type))
    return groups


def error_detail_rows(groups: list[ErrorGroup]) -> list[dict]:
    """Flat per-cell rows for a detailed table — individual references live
    here, never in the report header."""
    out = []
    for g in groups:
        for cell in g.sample_cells:
            out.append({
                "sheet": g.sheet, "cell": cell, "error_type": g.error_type,
                "region": g.region_ref, "area_type": g.area_type,
            })
    return out


# -------------------------------------------------------------- hidden sheets

_PURPOSE_HINTS = [
    ("Input", ("input", "raw", "source", "feed", "extract", "import", "data")),
    ("Calculation", ("calc", "working", "model", "engine", "compute", "workings")),
    ("Mapping", ("map", "lookup", "reference", "xref", "code", "lob")),
    ("Control/validation", ("check", "control", "valid", "tie", "recon", "qa",
                            "signoff", "sign-off", "approv")),
    ("Output", ("output", "report", "summary", "pack", "dashboard", "result", "mi")),
]


def classify_hidden_purpose(sheet_obj, tf: TabFacts | None) -> tuple[str, str]:
    """(purpose_group, explanation) for one hidden sheet."""
    name = sheet_obj.name.lower()
    if tf is not None:
        roles = _roles_of(tf)
        if "Output" in roles:
            return "Output", "formula/role evidence classifies this tab as an output"
        if "Mapping" in roles:
            return "Mapping", "lookup/mapping formula signature"
        if "Control Check" in roles or "Validation" in roles:
            return "Control/validation", "reconciliation/validation formula signature"
        if "Input" in roles:
            return "Input", "low formula ratio and/or read by other tabs"
        if "Calculation" in roles:
            return "Calculation", "high formula ratio — calculation role"
    for group, hints in _PURPOSE_HINTS:
        if any(h in name for h in hints):
            return group, f"sheet name suggests {group.lower()}"
    if sheet_obj.populated_cells and not sheet_obj.formula_profile.get("total"):
        return "Supporting", "populated but has no formulas — likely a supporting/reference sheet"
    return "Unknown", "no reliable structural or naming signal — needs a reviewer look"


@dataclass
class HiddenGroup:
    purpose: str
    sheets: list[str]
    count: int
    depended_on_by: list[str]
    explanation: str


def build_hidden_summary(wx, tab_index: dict[str, TabFacts], read_by: dict[str, set]) -> dict:
    hidden = [s for s in wx.sheets if s.state != "visible"]
    total = len(wx.sheets)
    header = f"{len(hidden)} of {total} worksheets are hidden."
    by_purpose: dict[str, list[str]] = defaultdict(list)
    explanations: dict[str, list[str]] = defaultdict(list)
    for s in hidden:
        purpose, expl = classify_hidden_purpose(s, tab_index.get(s.name))
        by_purpose[purpose].append(s.name)
        explanations[purpose].append(expl)

    groups = []
    for purpose, names in by_purpose.items():
        depends = sorted({d for n in names for d in read_by.get(n, set())})
        groups.append(HiddenGroup(
            purpose=purpose, sheets=sorted(names), count=len(names),
            depended_on_by=depends or ["none detected in-workbook — may still be read "
                                       "manually or by an external link"],
            explanation="; ".join(sorted(set(explanations[purpose])))[:400],
        ))
    groups.sort(key=lambda g: -g.count)
    return {"header": header, "hidden_count": len(hidden), "total": total, "groups": groups}


# -------------------------------------------------------------- key inputs

_INPUT_PURPOSE_HINTS = [
    ("Trial balance", ("trial balance", " tb ", " tb.", "tb_", "gl balance")),
    ("Claims data", ("claim",)),
    ("Account mapping", ("mapping", "chart of accounts", "coa")),
    ("Exchange rates", ("exchange", "fx ", "fx_", "fx.", "rate")),
    ("Previous-period workbook", ("prior", "previous", "py ", "py_", "last month", "last year")),
    ("Manual adjustments", ("adjustment", "manual", "override")),
]


def classify_input_purpose(name: str) -> str:
    low = name.lower()
    for label, hints in _INPUT_PURPOSE_HINTS:
        if any(h in low for h in hints):
            return label
    return "Unclassified — requires review"


def shorten_external(path: str) -> str:
    """Display filename only; the caller keeps the full path as evidence."""
    p = str(path).replace("\\", "/")
    return posixpath.basename(p) or str(path)


def group_key_inputs(wx, tab_index: dict[str, TabFacts]) -> dict:
    """Separate Key Inputs into in-workbook / lookup tables / connections /
    external workbooks (grouped by purpose, filename only) / unresolved."""
    input_sheets = sorted(n for n, tf in tab_index.items() if "Input" in _roles_of(tf))
    lookup_tables = sorted(n for n, tf in tab_index.items() if "Mapping" in _roles_of(tf))
    data_connections = [
        {"name": c.get("name") or c.get("type") or "unnamed connection",
         "type": c.get("type")}
        for c in wx.connections
    ]
    external_consumers = sorted(
        s.name for s in wx.sheets if s.formula_profile.get("external_count")
    )
    external_groups: dict[str, dict] = {}
    for link in wx.external_links:
        fname = shorten_external(link)
        purpose = classify_input_purpose(fname)
        g = external_groups.setdefault(purpose, {
            "purpose": purpose, "files": [], "consuming_worksheets": external_consumers,
            "essentiality": "requires review — not automatically confirmed",
        })
        g["files"].append({"file_name": fname, "full_path_evidence": link})
    if external_groups:
        note = ("consuming-worksheet mapping is approximate: OOXML external references "
                "are not individually attributed to a source link, so every sheet with an "
                "external reference is listed against every external workbook")
        for g in external_groups.values():
            g["evidence_note"] = note

    other_unresolved = sorted({p for p in wx.pivot_cache_sources if p and "!" not in p})

    return {
        "in_workbook_sheets": input_sheets or ["none identified as a dedicated input tab"],
        "lookup_mapping_tables": lookup_tables or ["none identified"],
        "data_connections": data_connections or ["none"],
        "external_workbooks": list(external_groups.values()) or ["none"],
        "other_unresolved_sources": other_unresolved or ["none"],
    }


# --------------------------------------------------------- downstream roles


def classify_consumer_role(name: str, tab_index: dict[str, TabFacts]) -> str:
    tf = tab_index.get(name)
    if tf is None:
        return "unknown"
    roles = _roles_of(tf)
    if "Output" in roles:
        return "final output"
    if "Control Check" in roles or "Validation" in roles:
        return "control / check"
    if tf.information == "supporting schedule":
        return "supporting schedule"
    return "intermediate working"


# --------------------------------------------------- business calc descriptions

_ROLL_FORWARD_HINTS = ("opening", "movement", "closing", "roll forward", "roll-forward",
                       "brought forward", "carried forward")
_FX_HINTS = ("fx", "exchange rate", "exchange-rate", "ccy", "currency")
_MAPPING_FNS = ("VLOOKUP", "XLOOKUP", "INDEX", "MATCH", "HLOOKUP", "LOOKUP")
_AGG_FNS = ("SUM", "SUMIF", "SUMIFS", "SUBTOTAL", "AVERAGE", "COUNT", "COUNTIF")


def business_calculation_description(s, tf: TabFacts | None) -> str | None:
    """Plain-language description of a tab's calculation, tied to its name and
    likely output — kept alongside (not instead of) the technical evidence."""
    fp = s.formula_profile
    if not fp.get("total"):
        return None
    headers = [h.lower() for r in s.regions for h in r.headers if h]
    fns = dict(fp.get("top_functions", []))
    total = fp["total"]
    agg = sum(fns.get(f, 0) for f in _AGG_FNS)
    lookup = sum(fns.get(f, 0) for f in _MAPPING_FNS)
    roles = _roles_of(tf)

    if any(any(h in hd for h in _ROLL_FORWARD_HINTS) for hd in headers):
        verb = "Rolls opening balances through movements to closing balances"
    elif any(any(h in hd for h in _FX_HINTS) for hd in headers):
        verb = "Converts balances using exchange-rate assumptions"
    elif "Mapping" in roles or lookup >= total * 0.3:
        verb = "Maps codes or values through a lookup/reference table"
    elif "Control Check" in roles or "Validation" in roles:
        verb = "Matches records between sources and reports exceptions"
    elif "Output" in roles:
        verb = "Produces a reporting or management-information schedule"
    elif agg >= total * 0.3:
        verb = "Aggregates detailed data into summary totals"
    else:
        return None

    output = "this tab is itself the deliverable" if "Output" in roles else \
        "feeds a downstream calculation or output tab"
    return f"{verb} (tab: {s.name}; {output})."


# ------------------------------------------------------------- reconciliation

_RECON_SRC_HINTS = ("source", "from", "gl", "tb", "ledger")
_RECON_TGT_HINTS = ("target", "to", "report", "control", "reserve")
_RECON_KEY_HINTS = ("id", "code", "ref", "key", "account", "policy")


def reconciliation_detail(s) -> dict:
    """Source / target / key / tolerance / exception / output when it can be
    read off headers and formula shapes — otherwise a targeted question. A
    count of VLOOKUP/MATCH/INDEX alone is never treated as sufficient."""
    fns = dict(s.formula_profile.get("top_functions", []))
    has_lookup = any(fns.get(f) for f in _MAPPING_FNS)
    has_abs = bool(fns.get("ABS"))
    if not (has_lookup or has_abs):
        return {"status": "not_detected"}

    headers = [h for r in s.regions for h in r.headers if h]
    source = next((h for h in headers if any(k in h.lower() for k in _RECON_SRC_HINTS)), None)
    target = next((h for h in headers if any(k in h.lower() for k in _RECON_TGT_HINTS)), None)
    key = next((h for h in headers if any(k in h.lower() for k in _RECON_KEY_HINTS)), None)
    if source and target and key:
        return {
            "source": source, "comparison_target": target, "matching_key": key,
            "tolerance": "ABS()/threshold comparison present" if has_abs else "not evidenced",
            "exception_rule": "not evidenced from headers — confirm with preparer",
            "result_output": s.name,
        }
    return {"question": (
        f"What source is being reconciled, what is the comparison target, and where "
        f"are exceptions reviewed on tab '{s.name}'? (lookup/variance formulas were "
        f"detected, but the source, target and matching key could not be resolved "
        f"from headers alone.)"
    )}


# ------------------------------------------------------------ manual entry

_MANUAL_HINTS = ("input", "override", "adjust", "manual", "paste", "hardcode")


def manual_intervention_tab(s) -> dict:
    """Explicit input/paste/adjustment/override evidence, or needs_human with
    a targeted question — never a High/Medium/Low from non-formula-cell %."""
    name = s.name.lower()
    headers = [h.lower() for r in s.regions for h in r.headers if h]
    hits = [h for h in [name] + headers if any(k in h for k in _MANUAL_HINTS)]
    fp = s.formula_profile
    if hits:
        return {"basis": "derived",
                "value": f"Explicit manual-entry/override area evidenced on '{s.name}'",
                "evidence": sorted(set(hits))[:5]}
    if fp.get("hardcoded_literal_count"):
        return {"basis": "derived",
                "value": (f"{fp['hardcoded_literal_count']} formula(s) on '{s.name}' "
                          f"contain a hardcoded value — a possible manual override "
                          f"buried in a formula"),
                "evidence": [f"{fp['hardcoded_literal_count']} hardcoded literal(s) in formulas"]}
    return {"basis": "needs_human", "value": None, "evidence": [],
            "question": f"What do users enter, paste, adjust or override on '{s.name}'?"}


def manual_intervention_file(wx) -> dict:
    per_tab = {s.name: manual_intervention_tab(s) for s in wx.sheets}
    concrete = {n: v["value"] for n, v in per_tab.items() if v["basis"] == "derived"}
    if concrete:
        return {"basis": "derived", "value": concrete,
                "evidence": [f"{len(concrete)}/{len(wx.sheets)} tab(s) show explicit "
                            f"manual-entry/override evidence"]}
    return {"basis": "needs_human", "value": None,
            "evidence": ["no explicit manual-entry/override area evidenced structurally "
                        "on any tab"],
            "question": "What do users enter, paste, adjust or override in this "
                        "workbook, and on which tabs?"}


# -------------------------------------------------------------- automation

def automation_candidates(wx, tab_index: dict[str, TabFacts]) -> list[dict]:
    """Per-tab automation candidates. Formula reuse/lookups/connections are
    supporting signals only — never the sole basis for a verdict."""
    out = []
    for s in wx.sheets:
        tf = tab_index.get(s.name)
        if tf is None:
            continue
        roles = _roles_of(tf)
        fp = s.formula_profile
        distinct = fp.get("distinct_skeletons", 0)
        comp = fp.get("total", 0) / distinct if distinct else 0
        signals = []
        if comp >= 5:
            signals.append(f"formula reuse {comp:.0f}x — a regular, codifiable rule")
        if wx.connections or wx.external_links:
            signals.append("existing data connection/link — source is already systemised")
        if "Control Check" in roles or "Validation" in roles:
            signals.append("reconciliation/matching pattern — a classic automation target")
        manual = manual_intervention_tab(s)
        if not signals:
            out.append({
                "worksheet": s.name, "verdict": "Candidate — workflow confirmation required",
                "suspected_manual_step": manual.get("value") or "not established",
                "required_inputs": "not established", "repeated_business_rule": "not established",
                "frequency": "unknown", "exceptions": "unknown",
                "review_or_approval_controls": "unknown",
                "output_produced": tf.information, "evidence": [],
            })
            continue
        out.append({
            "worksheet": s.name, "verdict": "Candidate",
            "suspected_manual_step": manual.get("value") or "not evidenced structurally",
            "required_inputs": "needs_human",
            "repeated_business_rule": (f"{distinct} distinct formula shape(s)"
                                       if fp.get("total") else "n/a"),
            "frequency": "unknown — needs_human", "exceptions": "unknown — needs_human",
            "review_or_approval_controls": "unknown — needs_human",
            "output_produced": tf.information, "evidence": signals,
        })
    return out


# ----------------------------------------------------------- simplification

def simplification_candidates(wx, tab_index: dict[str, TabFacts]) -> list[dict]:
    """One row per (worksheet, complexity issue) rather than one mixed list of
    hidden sheets, errors, volatile functions and hardcoded values."""
    out = []
    for s in wx.sheets:
        tf = tab_index.get(s.name)
        activity = tf.information if tf else "unknown"
        fp = s.formula_profile
        issues = []
        if fp.get("hardcoded_literal_count", 0) > 5:
            issues.append((
                f"{fp['hardcoded_literal_count']} hardcoded numbers buried in formulas",
                "Lift the hardcoded values to a named assumptions block",
                "Assumptions become visible and auditable, not buried in formula text",
            ))
        if fp.get("volatile_count"):
            issues.append((
                f"{fp['volatile_count']} volatile function(s) (OFFSET/INDIRECT/…)",
                "Replace with structured references or Excel Tables",
                "Faster, more predictable recalculation",
            ))
        distinct = fp.get("distinct_skeletons", 0)
        if fp.get("total", 0) > 50 and distinct and fp["total"] / distinct < 3:
            issues.append((
                f"{fp['total']} formulas reduce to only {distinct} distinct shapes "
                f"({fp['total']/distinct:.1f}x reuse)",
                "Standardise the formula pattern down the column/row",
                "Fewer one-off formulas to maintain and review",
            ))
        for current, change, benefit in issues:
            out.append({
                "worksheet": s.name, "business_activity": activity,
                "current_complexity": current, "proposed_change": change,
                "expected_benefit": benefit, "evidence": [current],
                "requires_human_confirmation": True,
            })
    return out


# ------------------------------------------------------------- retirement

def retirement_assessment() -> dict:
    """Retirement is never inferred from hidden sheets, cached errors, file
    age or a suggestive filename — those are not evidence of disuse."""
    return {
        "verdict": "Not established — owner decision required.",
        "confirmation_needed": ["active usage", "owner", "recipients",
                                "replacement coverage",
                                "regulatory or retention requirements"],
        "note": ("Hidden sheets, cached errors, file age, or a filename containing "
                 "'old'/'copy'/'backup' are not evidence of retirement on their own."),
    }


# ---------------------------------------------------------------- duplication

def duplication_corpus_note(n_others: int, threshold: int = 5) -> str | None:
    """Caveat when the comparison population is too small to be definitive."""
    if n_others == 0:
        return "no other workbooks were supplied in this run — duplication cannot be assessed"
    if n_others < threshold:
        return (f"conclusion is limited to the {n_others} other workbook(s) supplied in "
                f"this run, not a definitive answer across the full estate")
    return None


# ------------------------------------------------------------- macros / VBA

def read_vba_module_names(path: str) -> list[str] | None:
    """Best-effort VBA module names via olefile (optional dependency). Only
    the module *names* are read — never macro source code."""
    try:
        import olefile
    except ImportError:
        return None
    try:
        with zipfile.ZipFile(path) as zf:
            if "xl/vbaProject.bin" not in zf.namelist():
                return None
            data = zf.read("xl/vbaProject.bin")
    except Exception:
        return None
    try:
        ole = olefile.OleFileIO(io.BytesIO(data))
        streams = ole.listdir()
        ole.close()
    except Exception:
        return None
    skip = {"PROJECT", "PROJECTwm", "VBA", "__SRP_0", "__SRP_1"}
    names = sorted({s[-1] for s in streams if s and s[-1] not in skip and "VBA" not in s})
    return names or None


def macro_details(wx) -> dict:
    """VBA / Power Query / data connections / external links, kept separate
    so external workbook dependencies surface under Key Inputs instead."""
    modules = read_vba_module_names(wx.path) if wx.has_vba else None
    return {
        "vba": {
            "present": wx.has_vba,
            "modules": modules or (
                ["present — module names not extractable in this environment"]
                if wx.has_vba else []
            ),
        },
        "power_query": wx.has_power_query,
        "data_connections": [c.get("name") or c.get("type") or "unnamed"
                             for c in wx.connections] or ["none"],
        "external_workbook_links": len(wx.external_links),
    }
