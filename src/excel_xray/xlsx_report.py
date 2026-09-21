"""Excel (.xlsx) report exporter — the default report format.

Builds one workbook per assessed file with a fixed set of sheets: File
assessment, Tab assessments, Sheet inventory, Regions, Formula patterns,
Warnings, Report information, Error summary, Hidden sheet groups, Input
sources, Calculation steps, Review opportunities, and — for the folder-level
``estate.xlsx`` — an Estate comparison sheet.

Every assessment worksheet carries blank, editable **Reviewer Value** and
**Reviewer Notes** columns: the human-only anchor the review schema calls
for. Structure only — no cell values from the scanned workbook are embedded,
matching the HTML report's privacy stance.
"""

from __future__ import annotations

import datetime as _dt

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as _col_letter
from openpyxl.worksheet.worksheet import Worksheet

from .tabular import FILE_FIELDS, TAB_FIELDS, fmt_value

HEADER_FILL = PatternFill("solid", fgColor="10192B")
HEADER_FONT = Font(color="FFFFFF", bold=True)
TITLE_FONT = Font(bold=True, size=13)
SUBTITLE_FONT = Font(bold=True, size=11, color="5B6577")
WRAP = Alignment(wrap_text=True, vertical="top")
TOP = Alignment(vertical="top")


def _new_sheet(wb: Workbook, name: str) -> Worksheet:
    return wb.create_sheet(title=name[:31])


def _header_row(ws: Worksheet, headers: list[str], row: int = 1) -> None:
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.freeze_panes = ws.cell(row=row + 1, column=1).coordinate
    ws.auto_filter.ref = f"A{row}:{_col_letter(len(headers))}{row}"
    ws.row_dimensions[row].height = 30


def _widths(ws: Worksheet, widths: list[int]) -> None:
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[_col_letter(i)].width = w


def _write_rows(ws: Worksheet, rows: list[list], start_row: int = 2,
                wrap_cols: set[int] | None = None) -> int:
    """Write ``rows`` from ``start_row``; returns the next free row."""
    wrap_cols = wrap_cols or set()
    r = start_row
    for row in rows:
        for c, val in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.alignment = WRAP if c in wrap_cols else TOP
        r += 1
    return r


def _join(items) -> str:
    if not items:
        return "—"
    return "; ".join(str(x) for x in items)


# --------------------------------------------------------------- sheets


def _file_sheet(wb: Workbook, assessment) -> None:
    ws = _new_sheet(wb, "File assessment")
    headers = ["Type", "Field", "Value", "Basis", "Confidence", "Evidence",
              "Reviewer Value", "Reviewer Notes"]
    _header_row(ws, headers)
    rows = []
    for section, attr, label in FILE_FIELDS:
        fld = getattr(assessment.file, attr)
        rows.append([section, label, fmt_value(fld.value), fld.basis, fld.confidence,
                    _join(fld.evidence), None, None])
    _write_rows(ws, rows, wrap_cols={2, 3, 6, 7, 8})
    _widths(ws, [24, 30, 60, 12, 11, 60, 26, 32])


def _tab_sheet(wb: Workbook, assessment) -> None:
    ws = _new_sheet(wb, "Tab assessments")
    headers = ["Tab", "Type", "Field", "Value", "Basis", "Confidence", "Evidence",
              "Reviewer Value", "Reviewer Notes"]
    _header_row(ws, headers)
    rows = []
    for ta in assessment.tabs:
        name = ta.tab_name.value
        for section, attr, label in TAB_FIELDS:
            fld = getattr(ta, attr)
            rows.append([name, section, label, fmt_value(fld.value), fld.basis,
                        fld.confidence, _join(fld.evidence), None, None])
    _write_rows(ws, rows, wrap_cols={3, 4, 7, 8, 9})
    _widths(ws, [18, 24, 30, 50, 12, 11, 60, 26, 32])


def _sheet_inventory(wb: Workbook, wx) -> None:
    ws = _new_sheet(wb, "Sheet inventory")
    headers = ["Sheet", "Original position", "Visibility", "Populated cells",
              "Formulas", "Distinct formula shapes", "Density", "Dimension",
              "Merged ranges", "Cached errors"]
    _header_row(ws, headers)
    rows = []
    for s in sorted(wx.sheets, key=lambda s: (s.state != "visible", s.position)):
        fp = s.formula_profile
        rows.append([s.name, s.position + 1,
                    {"visible": "Visible", "hidden": "Hidden",
                     "veryHidden": "Very Hidden"}.get(s.state, s.state),
                    s.populated_cells, fp.get("total", 0), fp.get("distinct_skeletons", 0),
                    s.density, s.dimension, len(s.merges), len(s.error_cells)])
    _write_rows(ws, rows, wrap_cols=set())
    _widths(ws, [24, 16, 12, 14, 10, 20, 10, 14, 12, 12])


def _regions_sheet(wb: Workbook, wx) -> None:
    ws = _new_sheet(wb, "Regions")
    headers = ["Sheet", "Range", "Kind", "Origin", "Confidence", "Data rows",
              "Formula cells", "Totals row", "Headers"]
    _header_row(ws, headers)
    rows = []
    for s in wx.sheets:
        for r in s.regions:
            rows.append([s.name, r.ref, r.kind, r.origin, r.detect_confidence,
                        r.n_data_rows, r.formula_cells, r.totals_row,
                        _join([h for h in r.headers if h])])
    _write_rows(ws, rows, wrap_cols={9})
    _widths(ws, [22, 14, 14, 12, 11, 10, 12, 11, 50])


def _formula_patterns_sheet(wb: Workbook, wx) -> None:
    ws = _new_sheet(wb, "Formula patterns")
    headers = ["Sheet", "Total formulas", "Distinct shapes", "Compression",
              "Cross-sheet", "External", "Hardcoded literals", "Volatile",
              "Top functions", "Top formula shapes"]
    _header_row(ws, headers)
    rows = []
    for s in wx.sheets:
        fp = s.formula_profile
        if not fp.get("total"):
            continue
        top_fns = ", ".join(f"{fn}×{n}" for fn, n in fp.get("top_functions", [])[:8])
        top_sk = "; ".join(f"{sk} (×{n})" for sk, n in fp.get("top_skeletons", [])[:6])
        rows.append([s.name, fp.get("total", 0), fp.get("distinct_skeletons", 0),
                    fp.get("compression", 0), fp.get("cross_sheet_count", 0),
                    fp.get("external_count", 0), fp.get("hardcoded_literal_count", 0),
                    fp.get("volatile_count", 0), top_fns, top_sk])
    _write_rows(ws, rows, wrap_cols={9, 10})
    _widths(ws, [22, 14, 14, 12, 12, 10, 16, 10, 45, 55])


def _warnings_sheet(wb: Workbook, wx) -> None:
    ws = _new_sheet(wb, "Warnings")
    _header_row(ws, ["Warning"])
    from .util import safe_scan_message
    _write_rows(ws, [[safe_scan_message(w, wx.path)] for w in wx.warnings] or [["none"]], wrap_cols={1})
    _widths(ws, [110])


def _report_info_sheet(wb: Workbook, wx) -> None:
    ws = _new_sheet(wb, "Report information")
    _header_row(ws, ["Field", "Value"])
    from .util import safe_scan_message
    scan_error = (safe_scan_message("; ".join(wx.warnings[:3]) or
                                    "Workbook scan was partial; review diagnostics.", wx.path)
                  if wx.parse_status == "partial" else "N/A")
    rows = [
        ["File name", wx.filename],
        ["Scan status", wx.parse_status],
        ["Scan error", scan_error],
        ["No. of Sheets - Total", len(wx.sheets)],
        ["No. of Sheets - Hidden", sum(s.state != "visible" for s in wx.sheets)],
        ["Filesystem modified", wx.fs_modified],
        ["Document last modified", (wx.core_props or {}).get("modified")],
        ["Last saved by application", (wx.app_props or {}).get("application")],
        ["Report generated", _dt.datetime.now().isoformat(timespec="seconds")],
    ]
    _write_rows(ws, rows, wrap_cols={2})
    _widths(ws, [32, 60])


def _error_summary_sheet(wb: Workbook, review) -> None:
    ws = _new_sheet(wb, "Error summary")
    headers = ["Error type", "Count", "Sheet", "Region", "Area type",
              "Downstream outputs", "Recalc required", "Reviewer action"]
    _header_row(ws, headers)
    rows = [[g.error_type, g.count, g.sheet, g.region_ref, g.area_type,
            _join(g.downstream_outputs), "Yes" if g.recalc_required else "No — formula repair",
            g.reviewer_action] for g in review.error_groups]
    next_row = _write_rows(ws, rows, wrap_cols={6, 8}) if rows else 2
    _widths(ws, [12, 8, 20, 14, 20, 26, 20, 60])
    if not rows:
        ws.cell(row=2, column=1, value="No cached error cells detected.")
        return
    # Detail block: individual cell references, kept separate from the summary.
    title_row = next_row + 1
    t = ws.cell(row=title_row, column=1, value="Detail: individual cell references")
    t.font = SUBTITLE_FONT
    detail_headers = ["Sheet", "Cell", "Error type", "Region", "Area type"]
    for i, h in enumerate(detail_headers, start=1):
        c = ws.cell(row=title_row + 1, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
    for r, d in enumerate(review.error_details, start=title_row + 2):
        ws.cell(row=r, column=1, value=d["sheet"])
        ws.cell(row=r, column=2, value=d["cell"])
        ws.cell(row=r, column=3, value=d["error_type"])
        ws.cell(row=r, column=4, value=d["region"])
        ws.cell(row=r, column=5, value=d["area_type"])


def _hidden_groups_sheet(wb: Workbook, review) -> None:
    ws = _new_sheet(wb, "Hidden sheet groups")
    hs = review.hidden_summary
    t = ws.cell(row=1, column=1, value=hs.get("header", "0 worksheets are hidden."))
    t.font = TITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    ws.cell(row=2, column=1, value="Hidden status is never treated as evidence a "
                                   "worksheet or workbook is obsolete.").font = SUBTITLE_FONT
    headers = ["Purpose", "Worksheets", "Count", "Depended on by", "Likely purpose"]
    _header_row(ws, headers, row=4)
    rows = [[g.purpose, _join(g.sheets), g.count, _join(g.depended_on_by), g.explanation]
           for g in hs.get("groups", [])]
    _write_rows(ws, rows, start_row=5, wrap_cols={2, 4, 5})
    _widths(ws, [18, 30, 8, 40, 55])


def _input_sources_sheet(wb: Workbook, grouped: dict) -> None:
    ws = _new_sheet(wb, "Input sources")
    headers = ["Business purpose", "Source type", "Source name", "Reference count",
              "Consuming worksheets", "Source status", "Essentiality"]
    _header_row(ws, headers)
    rows = [[item.get("business_purpose"), item.get("source_type"),
             item.get("source_name"), item.get("reference_count"),
             _join(item.get("consuming_worksheets")),
             item.get("source_status"), item.get("essentiality")]
            for item in grouped.get("source_details", [])]
    if not rows:
        rows = [["—", "—", "no inputs identified", "", "", "", ""]]
    _write_rows(ws, rows, wrap_cols=set(range(1, len(headers) + 1)))
    _widths(ws, [35, 28, 40, 16, 42, 48, 42])


def _calculation_steps_sheet(wb: Workbook, assessment) -> None:
    ws = _new_sheet(wb, "Calculation steps")
    headers = ["Worksheet", "Business calculation or transformation"]
    _header_row(ws, headers)
    rows = []
    for ta in assessment.tabs:
        kc = ta.key_calculation_transformation_logic.value
        if not isinstance(kc, dict) or not (kc.get("business_description") or kc.get("business_descriptions")):
            continue
        description = kc.get("business_description") or "; ".join(kc.get("business_descriptions", []))
        rows.append([ta.tab_name.value, description])
    if not rows:
        rows = [["—", "No business calculation summary established from workbook structure."]]
    _write_rows(ws, rows, wrap_cols={2})
    _widths(ws, [28, 90])


def _review_opportunities_sheet(wb: Workbook, assessment) -> None:
    ws = _new_sheet(wb, "Review opportunities")
    headers = ["Kind", "Process", "Sub-Process", "Worksheet", "Detail", "Proposed action / question",
              "Requires human confirmation", "Reviewer Value", "Reviewer Notes"]
    _header_row(ws, headers)
    rows = []
    for c in assessment.review.simplification_candidates:
        rows.append(["Simplification", c.get("process"), c.get("sub_process"), c["worksheet"],
                    c["current_complexity"], c["proposed_change"], "Yes", None, None])
    for c in assessment.review.automation_candidates:
        detail = c.get("observed_manual_step", "Not established — confirm with process owner")
        rows.append(["Automation", c.get("process"), c.get("sub_process"), c["worksheet"],
                    detail, c.get("candidate_action", c["verdict"]), "Yes", None, None])
    for ta in assessment.tabs:
        if ta.human_validation_required.value == "Y":
            for reason in ta.validation_reason.value:
                rows.append(["Validation", "Not established", "Not established",
                             ta.tab_name.value, "", reason, "Yes", None, None])
    if not rows:
        rows = [["—", "Not established", "Not established", "—",
                 "no material review opportunities identified", "", "", None, None]]
    _write_rows(ws, rows, wrap_cols={2, 3, 5, 6, 8, 9})
    _widths(ws, [16, 38, 38, 24, 50, 65, 27, 20, 35])


def _estate_sheet(wb: Workbook, estate, insight=None) -> None:
    ws = _new_sheet(wb, "Estate comparison")
    fps = estate.fingerprints
    t = ws.cell(row=1, column=1,
               value=f"{len(fps)} workbooks · {len(estate.clusters)} families · "
                     f"{len(estate.pairs)} linked pairs · {len(estate.singletons)} unrelated")
    t.font = TITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    if insight is not None and insight.estate_summary:
        ws.cell(row=2, column=1, value=insight.estate_summary).font = SUBTITLE_FONT
    headers = ["Workbook A", "Workbook B", "Relationship", "Overall",
              "Formula shapes", "Inputs", "Outputs", "Topology"]
    _header_row(ws, headers, row=4)
    rows = [[fps[i].file_name, fps[j].file_name, c["relationship"], c["overall"],
            c["skeleton"], c["input"], c["output"], c["topology"]]
           for i, j, c in estate.pairs]
    _write_rows(ws, rows, start_row=5)
    _widths(ws, [30, 30, 26, 10, 12, 10, 10, 10])


# ------------------------------------------------------------------ entry


def build_xlsx_report(wx, assessment, estate=None, estate_insight=None) -> Workbook:
    """One workbook: File/Tab assessments, sheet inventory, regions, formula
    patterns, warnings, report info, error summary, hidden groups, input
    sources, calculation steps, review opportunities, and — when ``estate``
    is supplied — an Estate comparison sheet."""
    wb = Workbook()
    wb.remove(wb.active)  # drop the default blank sheet

    _file_sheet(wb, assessment)
    _tab_sheet(wb, assessment)
    _sheet_inventory(wb, wx)
    _regions_sheet(wb, wx)
    _formula_patterns_sheet(wb, wx)
    _warnings_sheet(wb, wx)
    _report_info_sheet(wb, wx)
    _error_summary_sheet(wb, assessment.review)
    _hidden_groups_sheet(wb, assessment.review)
    _input_sources_sheet(wb, assessment.review.key_inputs_grouped)
    _calculation_steps_sheet(wb, assessment)
    _review_opportunities_sheet(wb, assessment)
    if estate is not None:
        _estate_sheet(wb, estate, estate_insight)
    return wb


def write_xlsx_report(wx, assessment, path: str, estate=None, estate_insight=None) -> str:
    wb = build_xlsx_report(wx, assessment, estate=estate, estate_insight=estate_insight)
    wb.save(path)
    return path


def build_estate_xlsx(estate, insight=None) -> Workbook:
    """A standalone estate.xlsx for a folder-level comparison — same content
    as the per-file report's Estate comparison sheet, as its own workbook."""
    wb = Workbook()
    wb.remove(wb.active)
    _estate_sheet(wb, estate, insight)
    return wb


def write_estate_xlsx(estate, path: str, insight=None) -> str:
    wb = build_estate_xlsx(estate, insight)
    wb.save(path)
    return path
