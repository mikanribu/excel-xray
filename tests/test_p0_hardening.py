"""P0 regression tests: conservative formula references and auditable assessments."""
from excel_xray.formulas import analyse
from excel_xray import assess


def test_adjacent_ranges_do_not_swallow_following_reference():
    f = analyse("=SUM(A1:B2)+C3", 5, 5)
    assert set(f.precedent_cells) == {(1, 1), (1, 2), (2, 1), (2, 2), (3, 3)}


def test_multiple_ranges_are_independent():
    f = analyse("=SUM(A1:A2)+SUM(C4:C5)", 6, 6)
    assert set(f.precedent_cells) == {(1, 1), (2, 1), (4, 3), (5, 3)}


def test_invalid_excel_coordinates_are_not_precedents():
    f = analyse("=XFE1+XFD1048577+A1", 2, 2)
    assert f.precedent_cells == [(1, 1)]


def test_assessment_has_traceable_category_evidence(xray):
    a = assess(xray)
    for tab in a.tabs:
        category = tab.tab_category
        assert category.confidence_kind == "heuristic"
        for ref in category.evidence_refs:
            assert ref.startswith(tab.tab_name.value + "!")
    assert a.file.complexity.confidence_kind == "heuristic"
