"""Presentation schema (Step 6): lay the assessment out as the review table.

One ordered list of ``(section, attribute, label)`` per level mirrors the target
sheet exactly (File level summary, Tab level details), and drives both the HTML
render and the CSV export so the two never drift.
"""

from __future__ import annotations

import csv
import json

# (Type / section, dataclass attribute, human label) — order and wording match
# the target schema.
FILE_FIELDS = [
    ("Fact Assessment", "file_id", "File ID"),
    ("Fact Assessment", "file_name", "File Name"),
    ("Fact Assessment", "business_area_process", "Business Area / Process"),
    ("Fact Assessment", "purpose_of_file", "Purpose of File"),
    ("Fact Assessment", "key_output_outcome", "Key Output / Outcome"),
    ("Fact Assessment", "complexity", "Complexity"),
    ("Fact Assessment", "key_inputs", "Key Inputs"),
    ("Fact Assessment", "source_system", "Source System"),
    ("Fact Assessment", "key_outputs", "Key Outputs"),
    ("Fact Assessment", "usage_frequency", "Usage Frequency"),
    ("Fact Assessment", "completion_timeline", "Completion Timeline"),
    ("Fact Assessment", "euc_preparer", "EUC Preparer"),
    ("Fact Assessment", "output_recipient", "Output Recipient"),
    ("Key AI Finding / Observation", "potential_duplication", "Potential Duplication"),
    ("Key AI Finding / Observation", "similar_duplicate_files", "Similar / Duplicate File(s)"),
    ("Key AI Finding / Observation", "potential_simplification", "Potential Simplification"),
    ("Key AI Finding / Observation", "potential_consolidation", "Potential Consolidation"),
    ("Key AI Finding / Observation", "potential_automation", "Potential Automation"),
    ("Key AI Finding / Observation", "potential_retirement", "Potential Retirement"),
    ("Workbook logic / Automation", "logic_type", "Logic Type"),
    ("Workbook logic / Automation", "key_calculations_logic", "Key calculations / logic"),
    ("Workbook logic / Automation", "reconciliation_logic", "Reconciliation logic"),
    ("Workbook logic / Automation", "manual_intervention", "Manual intervention"),
    ("Workbook logic / Automation", "macros_vba_external_links", "Macros / VBA / external links"),
]

TAB_FIELDS = [
    ("Fact Assessment", "tab_name", "Tab Name"),
    ("Fact Assessment", "tab_category", "Tab Category"),
    ("Fact Assessment", "tab_purpose_description", "Tab Purpose / Description"),
    ("Fact Assessment", "tab_information_analysis", "Tab Information Analysis"),
    ("Fact Assessment", "key_calculation_transformation_logic",
     "Key Calculation / Transformation Logic Analysis"),
    ("Fact Assessment", "upstream_dependencies", "Upstream Dependencies"),
    ("Fact Assessment", "downstream_dependencies", "Downstream Dependencies"),
    ("Key AI Finding / Observation", "human_validation_required", "Human Validation Required?"),
    ("Key AI Finding / Observation", "validation_reason", "Validation Reason"),
]


def fmt_value(v) -> str:
    """Collapse any field value to a compact, human-readable string."""
    if v is None:
        return "—"
    if isinstance(v, bool):
        return "Yes" if v else "No"
    if isinstance(v, (int, float, str)):
        return str(v)
    if isinstance(v, list):
        return "; ".join(fmt_value(x) for x in v) if v else "—"
    if isinstance(v, dict):
        if "verdict" in v:
            extras = []
            for k in ("matches", "candidates", "opportunities", "drivers", "signals"):
                items = v.get(k)
                if items:
                    vals = [x.get("file") if isinstance(x, dict) else str(x) for x in items]
                    extras.append(f"{k}: " + ", ".join(vals))
            s = str(v["verdict"])
            return s + (" — " + "; ".join(extras) if extras else "")
        if "top_functions" in v:
            fns = ", ".join(v.get("top_functions", []))
            shapes = "; ".join(v.get("top_formula_shapes", [])[:3])
            return " | ".join(p for p in (fns, shapes) if p) or "—"
        if "in_workbook" in v:  # downstream deps
            return "; ".join(v.get("in_workbook") or []) or "—"
        return json.dumps(v, default=str)
    return str(v)


def _field(obj, attr):
    return getattr(obj, attr)


def file_rows(assessment) -> list[dict]:
    """Flat rows for the File level summary block."""
    out = []
    for section, attr, label in FILE_FIELDS:
        fld = _field(assessment.file, attr)
        out.append({
            "type": section, "field": label,
            "value": fmt_value(fld.value), "basis": fld.basis,
            "confidence": fld.confidence,
            "evidence": "; ".join(fld.evidence),
        })
    return out


def tab_rows(assessment) -> list[dict]:
    """Flat rows for the Tab level details block (one group per tab)."""
    out = []
    for ta in assessment.tabs:
        name = ta.tab_name.value
        for section, attr, label in TAB_FIELDS:
            fld = _field(ta, attr)
            out.append({
                "tab": name, "type": section, "field": label,
                "value": fmt_value(fld.value), "basis": fld.basis,
                "confidence": fld.confidence,
                "evidence": "; ".join(fld.evidence),
            })
    return out


def to_csv(named_assessments, path: str) -> str:
    """Write one long-format CSV for one or more (file_name, assessment) pairs.

    Columns: File, Level, Type, Field, Value, Basis, Confidence, Evidence.
    """
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["File", "Level", "Type", "Field", "Value", "Basis",
                    "Confidence", "Evidence"])
        for fname, a in named_assessments:
            for r in file_rows(a):
                w.writerow([fname, "File", r["type"], r["field"], r["value"],
                            r["basis"], r["confidence"], r["evidence"]])
            for r in tab_rows(a):
                w.writerow([fname, f"Tab: {r['tab']}", r["type"], r["field"],
                            r["value"], r["basis"], r["confidence"], r["evidence"]])
    return path
