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
    ("Fact Assessment", "business_area_process", "Business Area Purpose"),
    ("Fact Assessment", "process", "Process"),
    ("Fact Assessment", "sub_process", "Sub-Process"),
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
    ("Workbook logic / Automation", "logic_types", "Logic Types"),
    ("Workbook logic / Automation", "key_calculations_logic", "Key calculations / logic"),
    ("Workbook logic / Automation", "reconciliation_logic", "Reconciliation logic"),
    ("Workbook logic / Automation", "manual_intervention", "Manual intervention"),
    ("Workbook logic / Automation", "macros_vba_external_links", "Macros / VBA / external links"),
]

TAB_FIELDS = [
    ("Fact Assessment", "tab_name", "Tab Name"),
    ("Fact Assessment", "tab_visibility", "Tab Visibility"),
    ("Fact Assessment", "tab_category", "Tab Category"),
    ("Fact Assessment", "tab_roles", "Tab Roles"),
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
                    vals = [_candidate_label(x) for x in items]
                    extras.append(f"{k}: " + ", ".join(vals))
            if v.get("replacement_coverage"):
                extras.append(f"replacement coverage: {v['replacement_coverage']}")
            s = str(v["verdict"])
            return s + (" — " + "; ".join(extras) if extras else "")
        if "functional_similarities" in v or "shared_solution" in v:
            return _candidate_label(v)
        if "business_descriptions" in v or "business_description" in v:
            descs = v.get("business_descriptions") or [v.get("business_description")]
            return "; ".join(str(d) for d in descs if d) or "Not established — confirm with process owner"
        if "in_workbook" in v:  # tab downstream_dependencies
            items = v.get("in_workbook") or []
            return "; ".join(_candidate_label(x) for x in items) or "—"
        if "final_business_deliverables" in v:  # file key_outputs
            vals = v.get("final_business_deliverables") or []
            return "; ".join(vals) if vals else v.get(
                "status", "Not established — confirm the final business deliverable with the process owner")
        if "vba" in v:  # macros_vba_external_links
            vba = v.get("vba", {})
            formal = v.get("formal_data_connections", {})
            count = formal.get("count", 0)
            parts = [f"VBA: {'present' if vba.get('present') else 'none detected'}",
                    f"Power Query: {'present' if v.get('power_query') else 'none detected'}",
                    f"Formal data connections: {count} detected" if count else "Formal data connections: none detected",
                    f"External workbook links: {v.get('external_workbook_links', 0)} reference(s)"]
            return "; ".join(parts)
        if "reconciliation_count" in v or "resolved" in v or "reviewer_questions" in v:
            reconciliations = v.get("reconciliations") or v.get("candidates", []) + v.get("resolved", [])
            count = v.get("reconciliation_count", len(reconciliations))
            parts = [f"{count} reconciliation(s) identified"]
            parts.extend(
                f"{d.get('worksheet', d.get('tab', 'Worksheet'))}: {d.get('overview', 'comparison')}; "
                f"Source A={d.get('source_a', d.get('source', 'not established'))}; "
                f"Source B={d.get('source_b', d.get('comparison_target', 'not established'))}; "
                f"matching criteria={d.get('matching_criteria', d.get('matching_key', 'not established'))}; "
                f"tolerance={d.get('tolerance', 'not established')}; "
                f"exception logic={d.get('exception_logic', 'not established')}"
                for d in reconciliations
            )
            parts += v.get("reviewer_questions", [])
            if v.get("status"):
                parts.append(str(v["status"]))
            return "; ".join(parts)
        if "in_workbook_sheets" in v or "other_tabs_within_this_euc" in v:
            parts = []
            tabs = v.get("other_tabs_within_this_euc")
            if tabs is None:
                tabs = [x for x in (v.get("in_workbook_sheets") or [])
                        + (v.get("lookup_mapping_tables") or [])
                        if x not in {"none", "none identified", "none identified as a dedicated input tab"}]
            other_eucs = v.get("other_eucs")
            if other_eucs is None:
                other_eucs = [
                    {"file_name": file.get("file_name"),
                     "reference_count": file.get("reference_count", 1)}
                    for group in v.get("external_workbooks", [])
                    if isinstance(group, dict)
                    for file in group.get("files", [])
                ]
            if tabs:
                parts.append("Other tabs within this EUC: " + ", ".join(map(str, tabs)))
            elif "other_tabs_within_this_euc" in v:
                parts.append("Other tabs within this EUC: none identified")
            if other_eucs:
                parts.append("Other EUCs: " + ", ".join(
                    f"{item.get('file_name')} ({item.get('reference_count', 1)} reference(s))"
                    for item in other_eucs if isinstance(item, dict)
                ))
            elif "other_eucs" in v or "external_workbooks" in v:
                parts.append("Other EUCs: none identified")
            for k, label in (("data_connections", "Formal data connections"),
                             ("other_unresolved_sources", "Other unresolved sources")):
                items = v.get(k)
                if items and items not in (["none"], []):
                    parts.append(f"{label}: " + ", ".join(_candidate_label(x) for x in items))
            return "; ".join(parts) or "—"
        return json.dumps(v, default=str)
    return str(v)


def _candidate_label(x) -> str:
    """One short label for a structured list item (candidate/match/consumer)."""
    if isinstance(x, dict):
        name = None
        for k in ("worksheet", "file", "sheet", "file_name"):
            if k in x:
                name = str(x[k])
                break
        details = []
        for key, label in (
            ("functional_similarities", "similarities"),
            ("material_differences", "differences"),
            ("shared_solution", "shared solution"),
            ("proposed_change", "proposed change"),
            ("recommendation", "recommendation"),
            ("candidate_action", "next step"),
            ("suspected_manual_step", "manual step"),
            ("observed_manual_step", "manual step"),
        ):
            value = x.get(key)
            if value:
                rendered = "; ".join(map(str, value)) if isinstance(value, list) else str(value)
                details.append(f"{label}: {rendered}")
        blockers = x.get("blockers")
        if blockers:
            details.append("blockers: " + "; ".join(map(str, blockers)))
        if not details:
            extra = x.get("role") or x.get("verdict") or x.get("current_complexity")
            if extra:
                details.append(str(extra))
        if name is not None:
            return name + (" — " + "; ".join(details) if details else "")
        return str(x)
    return str(x)


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
