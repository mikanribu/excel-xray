"""Step 6 — presentation: table schema, CSV export, report rendering."""

from __future__ import annotations

import csv

from excel_xray import assess
from excel_xray.report import build_report
from excel_xray.tabular import (
    FILE_FIELDS,
    TAB_FIELDS,
    file_rows,
    fmt_value,
    tab_rows,
    to_csv,
)


def test_fmt_value_handles_shapes():
    assert fmt_value(None) == "—"
    assert fmt_value(True) == "Yes"
    assert fmt_value(["a", "b"]) == "a; b"
    assert fmt_value({"verdict": "Yes", "matches": [{"file": "x.xlsx"}]}) == \
        "Yes — matches: x.xlsx"
    assert "SUM" in fmt_value({"top_functions": ["SUM×3"], "top_formula_shapes": []})


def test_file_rows_cover_the_whole_schema(xray):
    rows = file_rows(assess(xray))
    assert len(rows) == len(FILE_FIELDS)
    labels = {r["field"] for r in rows}
    assert "Complexity" in labels and "Potential Duplication" in labels


def test_tab_rows_one_group_per_tab(xray):
    a = assess(xray)
    rows = tab_rows(a)
    assert len(rows) == len(a.tabs) * len(TAB_FIELDS)


def test_csv_export_roundtrips(xray, tmp_path):
    a = assess(xray)
    path = tmp_path / "euc.csv"
    to_csv([("wb.xlsx", a)], str(path))
    with open(path, newline="") as fh:
        rows = list(csv.DictReader(fh))
    assert rows[0].keys() >= {"File", "Level", "Type", "Field", "Value", "Basis"}
    # File-level plus every tab's fields are present.
    assert any(r["Field"] == "Logic Type" and r["Level"] == "File" for r in rows)
    assert any(r["Level"].startswith("Tab:") for r in rows)


def test_report_embeds_assessment(xray):
    html = build_report(xray, assess(xray))
    assert "EUC assessment" in html
    assert "File level summary" in html
    assert "Tab level details" in html


def test_report_without_assessment_still_renders(xray):
    html = build_report(xray)  # assessment optional
    assert "EUC assessment" not in html
    assert "Workbook X-ray" in html
