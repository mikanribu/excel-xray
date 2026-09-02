"""Shared fixtures."""

from __future__ import annotations

from pathlib import Path

import pytest

FIXTURE = Path(__file__).parent / "fixtures" / "messy_reserving_model.xlsx"


@pytest.fixture(scope="session")
def fixture_path() -> str:
    assert FIXTURE.exists(), f"missing fixture {FIXTURE}"
    return str(FIXTURE)


@pytest.fixture(scope="session")
def xray(fixture_path):
    from excel_xray import xray_workbook

    return xray_workbook(fixture_path)


@pytest.fixture(scope="session")
def regions_by_ref(xray):
    """(sheet, 'A1:E41') -> Region."""
    return {(s.name, r.ref): r for s in xray.sheets for r in s.regions}
