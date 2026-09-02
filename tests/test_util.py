"""Unit tests for the A1-notation helpers that replaced openpyxl.utils."""

from __future__ import annotations

import pytest

from excel_xray.util import (
    column_index_from_string,
    get_column_letter,
    range_boundaries,
)


@pytest.mark.parametrize(
    "idx,letter", [(1, "A"), (26, "Z"), (27, "AA"), (52, "AZ"), (702, "ZZ"), (703, "AAA")]
)
def test_column_letter_roundtrip(idx, letter):
    assert get_column_letter(idx) == letter
    assert column_index_from_string(letter) == idx


def test_range_boundaries_single_cell():
    assert range_boundaries("A1") == (1, 1, 1, 1)


def test_range_boundaries_rectangle():
    # (min_col, min_row, max_col, max_row) -- columns first, like openpyxl.
    assert range_boundaries("A1:E41") == (1, 1, 5, 41)


def test_range_boundaries_ignores_absolute_markers():
    assert range_boundaries("$B$3:$C$3") == (2, 3, 3, 3)


def test_range_boundaries_rejects_garbage():
    with pytest.raises(ValueError):
        range_boundaries("not-a-range")
