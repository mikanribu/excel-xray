"""HTML report layout: high-level header vs. detailed Diagnostics tab."""

from __future__ import annotations

import re

from excel_xray import assess
from excel_xray.report import build_report


def _panel_ids(html: str) -> list[str]:
    return re.findall(r"<div class='tabpanel' id='([^']+)'", html)


def test_assessment_tab_is_first_and_default():
    from excel_xray import xray_workbook
    wx = xray_workbook("tests/fixtures/messy_reserving_model.xlsx")
    html = build_report(wx, assess(wx))
    ids = _panel_ids(html)
    assert ids[0] == "tab-assessment"
    # JS falls back to 'tab-assessment' explicitly, not just DOM order.
    assert "fallback = ids.includes('tab-assessment') ? 'tab-assessment' : ids[0]" in html


def test_diagnostics_tab_holds_the_detail_not_the_header():
    from excel_xray import xray_workbook
    wx = xray_workbook("tests/fixtures/messy_reserving_model.xlsx")
    html = build_report(wx, assess(wx))
    ids = _panel_ids(html)
    assert "tab-diagnostics" in ids
    assert ids.index("tab-diagnostics") == 1  # right after Assessment

    header = html[:html.find("<div class='tabbar'>")]
    diagnostics = html[html.find("id='tab-diagnostics'"):html.find("id='tab-sheet-0'")]

    # The always-visible header gets a one-line warning, not the grouped table.
    assert "cached error cell(s)" in header
    assert "<table" not in header.split("<div class='vitals'>")[1].split("tabbar")[0]
    # The grouped table and per-cell detail live in the Diagnostics tab.
    assert "Error summary" in diagnostics
    assert "Reviewer action" in diagnostics
    assert "Detailed error cell references" in diagnostics


def test_hidden_detail_lives_in_diagnostics_tab():
    from excel_xray import xray_workbook
    wx = xray_workbook("tests/fixtures/complex_euc_model.xlsx")
    html = build_report(wx, assess(wx))
    header = html[:html.find("<div class='tabbar'>")]
    diagnostics = html[html.find("id='tab-diagnostics'"):html.find("id='tab-sheet-0'")]

    assert "worksheets are hidden" in header
    assert "Likely purpose" not in header  # the grouped table's column header
    assert "Hidden sheet groups" in diagnostics
    assert "Likely purpose" in diagnostics


def test_report_without_assessment_has_no_diagnostics_tab():
    from excel_xray import xray_workbook
    wx = xray_workbook("tests/fixtures/messy_reserving_model.xlsx")
    html = build_report(wx)  # no assessment
    assert "tab-diagnostics" not in html
