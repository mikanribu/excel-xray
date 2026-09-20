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
import urllib.parse
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
    ("Trial balance", ("trial balance", "tb", "gl balance", "general ledger")),
    ("Claims data", ("claim",)),
    ("Account mapping", ("mapping", "chart of accounts", "coa")),
    ("Exchange rates", ("exchange", "fx ", "fx_", "fx.", "rate")),
    ("Previous-period workbook", ("prior", "previous", "py ", "py_", "last month", "last year")),
    ("Manual adjustments", ("adjustment", "manual", "override")),
]


def classify_input_purpose(name: str) -> str:
    low = name.lower()
    if re.search(r"(?:^|[._\-\s])tb(?:$|[._\-\s])", low):
        return "Trial balance"
    for label, hints in _INPUT_PURPOSE_HINTS:
        if any(h != "tb" and h in low for h in hints):
            return label
    return "Unclassified — requires review"


def shorten_external(path: str) -> str:
    """Return a normalized workbook name without URL/path details."""
    raw = urllib.parse.unquote(str(path))
    parsed = urllib.parse.urlsplit(raw)
    candidate = parsed.path if parsed.scheme else raw.split("?", 1)[0].split("#", 1)[0]
    candidate = candidate.replace("\\", "/")
    return posixpath.basename(candidate.rstrip("/")) or "unresolved external workbook"


def _connection_label(connection: dict) -> str:
    """Return a friendly connection label without server or credential text."""
    name = str(connection.get("name") or "").strip()
    if name and not re.search(r"[/\\:@]|password=|server=|uid=|provider=|database=|token=|api[_-]?key=", name, re.I):
        return name[:80]
    kind = str(connection.get("type") or "").strip()
    if kind and not re.search(r"[/\\:@]|password=|server=|uid=|provider=|database=|token=|api[_-]?key=", kind, re.I):
        return f"{kind} connection"
    return "Unnamed formal data connection"


def group_key_inputs(wx, tab_index: dict[str, TabFacts],
                     read_by: dict[str, set] | None = None) -> dict:
    """Separate Key Inputs into in-workbook / lookup tables / connections /
    external workbooks (grouped by purpose, filename only) / unresolved."""
    lookup_tables = sorted(n for n, tf in tab_index.items() if "Mapping" in _roles_of(tf))
    input_sheets = sorted(n for n, tf in tab_index.items()
                          if "Input" in _roles_of(tf) and n not in lookup_tables)
    data_connections = [
        {"name": _connection_label(c),
         "type": c.get("type")}
        for c in wx.connections
    ]
    source_details = []
    for name in input_sheets:
        consumers = sorted((read_by or {}).get(name, set()))
        reference_count = sum(
            int(s.formula_profile.get("referenced_sheets", {}).get(name, 0) or 0)
            for s in wx.sheets
        )
        source_details.append({
            "business_purpose": classify_input_purpose(name),
            "source_type": "In-workbook input worksheet", "source_name": name,
            "reference_count": reference_count,
            "consuming_worksheets": consumers or "No in-workbook consumer detected",
            "source_status": "Observed worksheet; input role is structurally inferred",
            "essentiality": "Not established — confirm with process owner",
        })
    for name in lookup_tables:
        consumers = sorted((read_by or {}).get(name, set()))
        reference_count = sum(
            int(s.formula_profile.get("referenced_sheets", {}).get(name, 0) or 0)
            for s in wx.sheets
        )
        source_details.append({
            "business_purpose": classify_input_purpose(name),
            "source_type": "Lookup or mapping worksheet", "source_name": name,
            "reference_count": reference_count,
            "consuming_worksheets": consumers or "No in-workbook consumer detected",
            "source_status": "Observed worksheet; mapping role is structurally inferred",
            "essentiality": "Not established — confirm with process owner",
        })
    for connection in data_connections:
        source_details.append({
            "business_purpose": "Not established — confirm with process owner",
            "source_type": "Formal data connection", "source_name": connection["name"],
            "reference_count": 1,
            "consuming_worksheets": "Not determinable from workbook structure",
            "source_status": "Connection metadata observed; source contents not verified",
            "essentiality": "Not established — confirm with process owner",
        })
    external_groups: dict[str, dict] = {}
    normalized_links: dict[str, list] = {}
    for link in wx.external_links:
        fname = shorten_external(link)
        entry = normalized_links.setdefault(fname.casefold(), [fname, 0])
        entry[1] += 1
    for fname, reference_count in sorted(normalized_links.values(), key=lambda item: item[0].casefold()):
        purpose = classify_input_purpose(fname)
        g = external_groups.setdefault(purpose, {
            "business_purpose": purpose, "source_type": "External workbook",
            "files": [],
            "consuming_worksheets": "Not determinable from workbook structure",
            "essentiality": "Not established — confirm with the process owner",
        })
        existing = next((f for f in g["files"] if f["file_name"] == fname), None)
        if existing:
            existing["reference_count"] += reference_count
        else:
            g["files"].append({"file_name": fname, "reference_count": reference_count,
                               "source_status": "Observed reference; availability and contents not verified"})
        source_details.append({
            "business_purpose": purpose, "source_type": "External workbook",
            "source_name": fname, "reference_count": reference_count,
            "consuming_worksheets": "Not determinable from workbook structure",
            "source_status": "Observed reference; availability and contents not verified",
            "essentiality": "Not established — confirm with the process owner",
        })

    other_unresolved = sorted({
        (shorten_external(p) if re.search(r"(?i)\.xls[xmb]?(?:$|[?#])", str(p))
         else "Unresolved pivot or external data source")
        for p in wx.pivot_cache_sources if p and "!" not in p
    })
    for name in other_unresolved:
        if name != "none":
            source_details.append({
                "business_purpose": "Not established — confirm with process owner",
                "source_type": "Unresolved pivot or external data source",
                "source_name": name, "reference_count": 1,
                "consuming_worksheets": "Not determinable from workbook structure",
                "source_status": "Source reference observed; source identity not established",
                "essentiality": "Not established — confirm with process owner",
            })

    return {
        "in_workbook_sheets": input_sheets or ["none identified as a dedicated input tab"],
        "lookup_mapping_tables": lookup_tables or ["none identified"],
        "data_connections": data_connections or ["none"],
        "external_workbooks": list(external_groups.values()) or ["none"],
        "other_unresolved_sources": other_unresolved or ["none"],
        "source_details": source_details,
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
        verb = "Supports a control or validation check; confirm the matching rule and exception handling"
    elif "Output" in roles:
        verb = "Supports a reporting-oriented worksheet; confirm the final business deliverable"
    elif agg >= total * 0.3:
        verb = "Aggregates detailed data into summary totals"
    else:
        return None

    output = "tab has a structurally inferred output role" if "Output" in roles else \
        "feeds a downstream calculation or output tab"
    return f"{verb} (tab: {s.name}; {output})."


# ------------------------------------------------------------- reconciliation

_RECON_SRC_HINTS = ("source", "from", "gl", "tb", "ledger")
_RECON_TGT_HINTS = ("target", "to", "report", "control", "reserve")
_RECON_KEY_HINTS = ("id", "code", "ref", "key", "account", "policy")


def reconciliation_detail(s) -> dict:
    """Return a review candidate, never a confirmed tie-out from headers alone.

    The structural scan does not inspect calculated values or prove that two
    sources agree, so confirmation always remains with the process owner.
    """
    fns = dict(s.formula_profile.get("top_functions", []))
    has_abs = bool(fns.get("ABS"))
    headers = [h for r in s.regions for h in r.headers if h]
    header_text = " ".join(headers).lower()
    labeled_recon = any(word in header_text or word in s.name.lower()
                        for word in ("recon", "variance", "difference", "tie out", "tie-out", "matched"))
    source = next((h for h in headers if any(k in h.lower() for k in _RECON_SRC_HINTS)), None)
    target = next((h for h in headers if any(k in h.lower() for k in _RECON_TGT_HINTS)), None)
    key = next((h for h in headers if any(k in h.lower() for k in _RECON_KEY_HINTS)), None)
    if not (labeled_recon or (has_abs and source and target)):
        return {"status": "not_detected"}
    evidence = []
    if has_abs:
        evidence.append("absolute-difference formula observed")
    if labeled_recon:
        evidence.append("reconciliation/variance label observed")
    return {
        "status": "candidate — owner confirmation required",
        "source": source or "not established",
        "comparison_target": target or "not established",
        "matching_key": key or "not established",
        "agreement_evidence": "not established from workbook structure",
        "variance_rule": "not established from workbook structure",
        "exception_output": "not established from workbook structure",
        "observed_evidence": evidence,
        "reviewer_question": (f"Confirm the two sources, matching key, agreement or tolerance rule, "
                              f"and where exceptions are reviewed for '{s.name}'."),
    }


# ------------------------------------------------------------ manual entry

_MANUAL_HINTS = ("input", "override", "adjust", "manual", "paste", "hardcode")


def manual_intervention_tab(s) -> dict:
    """Explicit input/paste/adjustment/override evidence, or needs_human with
    a targeted question — never a High/Medium/Low from non-formula-cell %."""
    name = s.name.lower()
    headers = [h.lower() for r in s.regions for h in r.headers if h]
    hits = [h for h in [name] + headers if any(k in h for k in _MANUAL_HINTS)]
    return {"basis": "needs_human",
            "value": ("Not established — confirm which inputs are keyed, pasted, "
                      "adjusted or overridden, by whom and how often."),
            "evidence": sorted(set(hits))[:5],
            "question": f"What do users enter, paste, adjust or override on '{s.name}', and how often?"}


def manual_intervention_file(wx) -> dict:
    return {"basis": "needs_human",
            "value": ("Not established — confirm which inputs are keyed, pasted, "
                      "adjusted or overridden, by whom and how often."),
            "evidence": ["manual actions and ownership are not established by workbook structure"],
            "question": "Which inputs are keyed, pasted, adjusted or overridden, by whom and how often?"}


# -------------------------------------------------------------- automation

def automation_candidates(wx, tab_index: dict[str, TabFacts]) -> list[dict]:
    """Show only evidence-backed workflow leads that need owner validation."""
    out = []
    for s in wx.sheets:
        tf = tab_index.get(s.name)
        if tf is None:
            continue
        fp = s.formula_profile
        distinct = fp.get("distinct_skeletons", 0)
        comp = fp.get("total", 0) / distinct if distinct else 0
        label_text = " ".join([s.name] + [h for r in s.regions for h in r.headers if h]).lower()
        manual_terms = ("manual", "paste", "override", "adjustment", "input", "keyed")
        if not any(term in label_text for term in manual_terms):
            continue
        signals = ["worksheet labels suggest a manual input or adjustment step"]
        if comp >= 5:
            signals.append(f"repeated formula pattern observed ({comp:.0f} uses per shape)")
        out.append({
            "process": "Not established — confirm with process owner",
            "sub_process": "Not established — confirm with process owner",
            "worksheet": s.name,
            "verdict": "Potential workflow lead — owner confirmation required",
            "observed_manual_step": "Not established — confirm with process owner",
            "required_inputs": "Confirm with process owner",
            "repeated_business_rule": f"{distinct} distinct formula shape(s)" if fp.get("total") else "not observed",
            "frequency": "Confirm with process owner",
            "exceptions_and_controls": "Confirm with process owner",
            "candidate_action": "Map the manual steps, inputs, approval controls and exceptions before assessing automation.",
            "evidence": signals,
            "requires_human_confirmation": True,
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
                "process": "Not established — confirm with process owner",
                "sub_process": activity if activity != "unknown" else "Not established — confirm with process owner",
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
        "verdict": "Not established — confirm with the process owner.",
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
        "formal_data_connections": {
            "count": len(wx.connections),
            "summary": (f"{len(wx.connections)} formal data connection(s) detected"
                        if wx.connections else "No formal data connections detected"),
        },
        "external_workbook_links": len(wx.external_links),
    }
