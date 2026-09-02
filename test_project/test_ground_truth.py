"""Score the detector against hand-labelled regions.

The point of this file is to let the demo state a measured number instead of an
assertion. Ground truth was marked by reading the fixture, not by running the
detector -- if you regenerate expectations from output, the number is worthless.
"""
import sys

sys.path.insert(0, "/home/claude/xray_project")

from xray.scan import xray_workbook  # noqa: E402

FIXTURE = "/home/claude/xray_project/samples/messy_reserving_model.xlsx"

# (sheet, ref, kind) as marked by hand.
GROUND_TRUTH = [
    ("Assumptions", "A1:B7",  "key_value"),
    ("Assumptions", "A10:A10", "notes"),
    ("Data",        "A1:E41", "data_table"),
    ("Data",        "G1:I6",  "data_table"),
    ("Calc",        "A1:A1",  "title"),
    ("Calc",        "A3:G13", "calculation"),
    ("Calc",        "I5:I10", "calculation"),
    ("Calc",        "A16:A17", "notes"),
    ("MI Pack",     "A1:F7",  "data_table"),
    ("Old_Working", "A1:B19", "data_table"),
]

EXPECTED_HEADERS = {
    ("Calc", "A3:G13"): [
        "Accident Year", "Gross::Opening", "Gross::Movement",
        "Reinsurance::Opening", "Reinsurance::Movement",
        "Net::Opening", "Net::Movement",
    ],
    ("Data", "A1:E41"): [
        "Policy ID", "Line of Business", "Written Premium",
        "Earned Premium", "Region",
    ],
}


def main() -> int:
    wx = xray_workbook(FIXTURE)
    found = {(s.name, r.ref): r for s in wx.sheets for r in s.regions}

    boundary_hits = kind_hits = 0
    print(f"{'sheet':<13}{'expected ref':<11}{'kind':<13}{'boundary':<10}kind")
    print("-" * 62)
    for sheet, ref, kind in GROUND_TRUTH:
        reg = found.get((sheet, ref))
        b = "OK" if reg else "MISS"
        k = "-"
        if reg:
            boundary_hits += 1
            k = "OK" if reg.kind == kind else f"got {reg.kind}"
            kind_hits += reg.kind == kind
        print(f"{sheet:<13}{ref:<11}{kind:<13}{b:<10}{k}")

    spurious = [ref for ref in found
                if ref not in {(s, r) for s, r, _ in GROUND_TRUTH}]

    n = len(GROUND_TRUTH)
    print("-" * 62)
    print(f"boundaries : {boundary_hits}/{n} ({boundary_hits/n:.0%})")
    print(f"kinds      : {kind_hits}/{n} ({kind_hits/n:.0%})")
    print(f"spurious   : {len(spurious)} {spurious if spurious else ''}")

    print("\nheader flattening:")
    for key, exp in EXPECTED_HEADERS.items():
        reg = found.get(key)
        got = reg.headers if reg else None
        print(f"  {key[0]}!{key[1]}: {'OK' if got == exp else 'MISMATCH'}")
        if got != exp:
            print(f"    expected {exp}\n    got      {got}")

    low = [(s, r.ref, r.detect_confidence) for s in wx.sheets for r in s.regions
           if r.detect_confidence < 0.7]
    print(f"\nflagged for review (<0.70): {len(low)}")
    for s, ref, c in low:
        print(f"  {s.name}!{ref} {c:.2f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
