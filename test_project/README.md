# Excel X-ray — Stage 1

Structural scanner for financial workbooks. Takes a hand-built model and produces
a verifiable map of what is in it: regions, headers, calculation profile,
dependencies, and quality flags.

Read-only. Never writes to, moves or renames a source file.

```bash
python xray_cli.py samples/messy_reserving_model.xlsx -o out/
python xray_cli.py /path/to/corpus -o out/          # folder
python xray_cli.py file.xlsx --json                 # machine-readable
python tests/test_ground_truth.py                   # accuracy score
```

## Architecture

Three layers, deliberately separated.

| Module | Does | Why not the obvious thing |
|---|---|---|
| `ooxml.py` | Structure straight from the zip via `lxml` | openpyxl in read-only mode **silently drops merged ranges, ListObjects, connections, external links and pivot caches** — exactly the parts that say what a workbook is for. Reading the XML is complete and ~10x faster |
| `regions.py` | Occupancy runs → connected components → header inference → classification | Nothing off the shelf detects hand-built table regions |
| `formulas.py` | A1 → R1C1 → literal abstraction | Collapses a filled-down column to one skeleton |
| `scan.py` | Two-pass cell load, triage, orchestration | `data_only=False` gives formulas with no results; `data_only=True` gives results with no formulas. One pass cannot give both |
| `report.py` | Self-contained HTML | No CDN, no network, structure only |

### The asymmetry that matters

```
row_gap = 1   a blank row inside a table is common (spacer, visual break)
col_gap = 0   a blank column between blocks is the standard separator
```

Getting this backwards merges an entire sheet into one region. It is the single
most consequential constant in the detector.

### Privacy

The report and the JSON contain cell **coordinates, types and normalised formula
shapes**. No cell values are embedded. A workbook X-ray can be circulated to
people who are not cleared for the underlying figures.

## Measured accuracy

Against hand-marked ground truth on the fixture (`tests/test_ground_truth.py`).
Ground truth was marked by reading the file, not by running the detector.

```
boundaries : 10/10 (100%)
kinds      :  8/10 (80%)
spurious   :  0
```

Both kind errors are among the two regions the detector flagged below 0.70. It
is wrong where it says it is unsure, which is the property that matters.

This is one workbook. **Do not quote these numbers as corpus accuracy** — the
fixture was written to exercise known failure modes, so it is neither a random
sample nor an adversarial one. A defensible number needs ~40 hand-labelled real
workbooks.

## What it catches on the fixture

- ListObject read exactly, with declared column names
- Two blocks side by side on one sheet, correctly separated
- A blank spacer row *inside* a table, correctly bridged (`A3:G13` spans the gap)
- Multi-row merged header flattened: `Gross::Opening`, `Reinsurance::Movement`
- Totals row identified and excluded from the data row count
- Key/value assumptions block distinguished from a table
- Cached `#DIV/0!` surfaced
- Hidden `Old_Working` sheet flagged, and scored 0.29 because it genuinely is
  ambiguous
- 100,000 formulas → 1 distinct calculation

## Known gaps

**Performance.** ~10k cells/sec — 60s for a 600k-cell sheet. Acceptable for a
demo, too slow for a corpus. The fix is to stop using openpyxl for cell
streaming and `iterparse` `sheetData` directly, reading `<c>` and `<f>` in one
pass instead of two. Expect 5–10x. Not done yet.

**Formats.** `.xlsb` is values-only via `pyxlsb` (no formulas — a real blind
spot). `.xls` needs LibreOffice conversion first. Both are triaged and reported,
not silently skipped.

**Confidence calibration.** Weights and the 0.70 threshold are set by hand. They
are plausible, not fitted. A clean machine-generated extract initially scored
0.65 because its headers are not styled — corrected with a column type-consistency
signal, but this is exactly the kind of thing only a golden set settles.

**Region detection on presentation-formatted MI packs** is the weakest area, and
it is where a financial corpus is thickest. Expect 60–75% boundary accuracy there
versus 100% on this fixture.

## Not built yet (Stage 1 step 4)

Intra-workbook dependency graph. `formulas.analyse()` already returns
`precedent_cells`, so the edges are available; collapsing cell→cell edges to
region→region and deriving input / calculation / output roles is the next step.
That is the part that demonstrates comprehension rather than extraction.
