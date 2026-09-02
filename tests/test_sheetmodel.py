"""The single-pass lxml reader must recover formulas, cached values and styles.

These are the signals openpyxl gave for free and that region detection depends
on; if the OOXML reader gets them wrong, everything downstream silently drifts.
"""

from __future__ import annotations

import zipfile

import pytest

from excel_xray.ooxml import read_structure
from excel_xray.sheetmodel import read_shared_strings, read_sheet_cells, read_styles


@pytest.fixture(scope="module")
def calc_cells(fixture_path):
    struct = read_structure(fixture_path)
    part = next(s.part for s in struct.sheets if s.name == "Calc")
    with zipfile.ZipFile(fixture_path) as zf:
        shared = read_shared_strings(zf)
        styles = read_styles(zf)
        cells, profile, errors, *_ = read_sheet_cells(zf, part, shared, styles)
    return cells, profile, errors


def test_formula_and_cached_value_in_one_pass(calc_cells):
    cells, _, _ = calc_cells
    # Calc!C5 = B5 * Assumptions!$B$6, cached to 59937.255 -- both in one <c>.
    c5 = cells[(5, 3)]
    assert c5.formula == "=B5*Assumptions!$B$6"
    assert c5.value == pytest.approx(59937.255)
    assert not c5.is_text


def test_shared_string_resolved(calc_cells):
    cells, _, _ = calc_cells
    # A5 is a text label pulled from the shared-string table.
    a5 = cells[(5, 1)]
    assert a5.is_text
    assert isinstance(a5.value, str) and a5.value


def test_cached_error_surfaced(calc_cells):
    cells, _, errors = calc_cells
    i10 = cells[(10, 9)]  # =I6/I9 -> #DIV/0!
    assert i10.formula == "=I6/I9"
    assert i10.is_error
    assert any("#DIV/0!" in e for e in errors)


def test_bold_style_read(calc_cells):
    cells, _, _ = calc_cells
    # The totals row (row 13) is styled; at least one of its cells is bold.
    assert any(cells[(13, c)].bold for c in range(1, 8) if (13, c) in cells)


def test_formula_compression(calc_cells):
    _, profile, _ = calc_cells
    # Many near-identical roll-forward formulas collapse to few skeletons.
    assert profile.total > profile.distinct_skeletons
    assert profile.compression > 1.0
