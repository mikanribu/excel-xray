"""Region detection scored against hand-marked ground truth.

Ground truth was marked by reading the fixture, not by running the detector. It
was written to exercise known failure modes (a merged multi-row header, a blank
spacer row inside a table, two blocks side by side, a hidden working sheet), so
it is a regression guard, not a corpus-accuracy claim.
"""

from __future__ import annotations

import pytest

# (sheet, ref, kind) as marked by hand.
GROUND_TRUTH = [
    ("Assumptions", "A1:B7", "key_value"),
    ("Assumptions", "A10:A10", "notes"),
    ("Data", "A1:E41", "data_table"),
    ("Data", "G1:I6", "data_table"),
    ("Calc", "A1:A1", "title"),
    ("Calc", "A3:G13", "calculation"),
    ("Calc", "I5:I10", "calculation"),
    ("Calc", "A16:A17", "notes"),
    ("MI Pack", "A1:F7", "data_table"),
    ("Old_Working", "A1:B19", "data_table"),
]

EXPECTED_HEADERS = {
    ("Data", "A1:E41"): [
        "Policy ID", "Line of Business", "Written Premium",
        "Earned Premium", "Region",
    ],
    ("Calc", "A3:G13"): [
        "Accident Year", "Gross::Opening", "Gross::Movement",
        "Reinsurance::Opening", "Reinsurance::Movement",
        "Net::Opening", "Net::Movement",
    ],
}


def test_all_region_boundaries_detected(regions_by_ref):
    """Every hand-marked region boundary is found (100% on this fixture)."""
    missing = [(s, ref) for s, ref, _ in GROUND_TRUTH if (s, ref) not in regions_by_ref]
    assert not missing, f"missed region boundaries: {missing}"


def test_region_kind_accuracy(regions_by_ref):
    """Kind accuracy stays at the measured 8/10; regressions below fail."""
    hits = sum(
        1 for s, ref, kind in GROUND_TRUTH
        if (s, ref) in regions_by_ref and regions_by_ref[(s, ref)].kind == kind
    )
    assert hits >= 8, f"kind accuracy regressed to {hits}/{len(GROUND_TRUTH)}"


def test_no_spurious_regions(regions_by_ref):
    """The detector invents no regions the ground truth does not list."""
    expected = {(s, ref) for s, ref, _ in GROUND_TRUTH}
    spurious = [k for k in regions_by_ref if k not in expected]
    assert not spurious, f"spurious regions: {spurious}"


@pytest.mark.parametrize("key", list(EXPECTED_HEADERS))
def test_header_flattening(regions_by_ref, key):
    """Merged multi-row headers flatten to parent::child labels exactly."""
    assert regions_by_ref[key].headers == EXPECTED_HEADERS[key]


def test_low_confidence_is_where_it_is_wrong(xray, regions_by_ref):
    """The two kind errors must be among regions flagged below 0.70.

    'Wrong only where it says it is unsure' is the property that makes the
    confidence score trustworthy.
    """
    wrong = {
        (s, ref) for s, ref, kind in GROUND_TRUTH
        if (s, ref) in regions_by_ref and regions_by_ref[(s, ref)].kind != kind
    }
    for key in wrong:
        assert regions_by_ref[key].detect_confidence < 0.70, (
            f"{key} is misclassified but presented at high confidence "
            f"{regions_by_ref[key].detect_confidence}"
        )


def test_hidden_sheet_flagged_low(xray):
    """The hidden Old_Working sheet is genuinely ambiguous -> low confidence."""
    old = next(s for s in xray.sheets if s.name == "Old_Working")
    assert old.state != "visible"
    assert all(r.detect_confidence < 0.5 for r in old.regions)
