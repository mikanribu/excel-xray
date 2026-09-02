# Excel X-ray

Structural scanner and explainer for complicated financial workbooks. Takes a
hand-built model and produces a verifiable map of what is in it: regions,
headers, calculation profile, dependencies and quality flags — as a
self-contained HTML report or JSON.

Read-only. Never writes to, moves or renames a source file.

```bash
uv run excel-xray tests/fixtures/messy_reserving_model.xlsx -o out/
uv run excel-xray /path/to/corpus -o out/          # a whole folder
uv run excel-xray file.xlsx --json                 # machine-readable to stdout
uv run pytest                                      # accuracy + reader tests
```

## Why not openpyxl

openpyxl is the obvious library, and it is the wrong tool for both halves of this
job:

1. **Structure.** In read-only mode openpyxl *silently drops* merged ranges,
   ListObjects, external links, connections and pivot caches — exactly the parts
   that tell you what a workbook is *for*. [ooxml.py](src/excel_xray/ooxml.py)
   reads them straight from the zip with `lxml`, and about 10× faster.

2. **Cells.** openpyxl exposes *either* formulas (`data_only=False`) *or* cached
   values (`data_only=True`) per load, so people load the file twice. But the
   file keeps them together:

   ```xml
   <c r="C5" t="n"><f>B5*Assumptions!$B$6</f><v>59937.255</v></c>
   ```

   The "you need two passes" limitation is openpyxl's, not the format's.
   [sheetmodel.py](src/excel_xray/sheetmodel.py) streams `<sheetData>` once with
   `lxml.iterparse`, reading the formula, its cached value **and** the style
   index (for the bold/fill/border/date signals) in a single pass — ~40k
   cells/sec versus openpyxl's ~10k, with the whole workbook's formulas and
   values available at once.

The package has **no openpyxl dependency**. Its cell reader is verified
cell-for-cell against openpyxl on the fixture (values, formulas, bold, fill,
border — zero mismatches); openpyxl is a test-only tool.

## Architecture

Layers, deliberately separated.

| Module | Does |
|---|---|
| [ooxml.py](src/excel_xray/ooxml.py) | Structure straight from the zip via `lxml`: merges, ListObjects, connections, external links, pivot caches, VBA, Power Query |
| [sheetmodel.py](src/excel_xray/sheetmodel.py) | Single `lxml` pass over `<sheetData>` → cells (value + formula + style) and a formula profile. Resolves `sharedStrings` and `styles` |
| [regions.py](src/excel_xray/regions.py) | Occupancy runs → connected components → header inference → classification. Detects hand-built table regions nothing off-the-shelf finds |
| [formulas.py](src/excel_xray/formulas.py) | A1 → R1C1 → literal abstraction. Collapses a filled-down column to one skeleton |
| [scan.py](src/excel_xray/scan.py) | Triage, orchestration, occupancy plate |
| [report.py](src/excel_xray/report.py) | Self-contained HTML — no CDN, no network, structure only |
| [util.py](src/excel_xray/util.py) | A1-notation helpers (replaces `openpyxl.utils`) |

### The asymmetry that matters

```
row_gap = 1   a blank row inside a table is common (spacer, visual break)
col_gap = 0   a blank column between blocks is the standard separator
```

Getting this backwards merges an entire sheet into one region. It is the single
most consequential constant in the detector.

### Privacy

The report and the JSON contain cell **coordinates, types and normalised formula
shapes**. No cell values are embedded, so an X-ray can be circulated to people
who are not cleared for the underlying figures.

## Measured accuracy

Against hand-marked ground truth on the fixture
([tests/test_ground_truth.py](tests/test_ground_truth.py)), marked by reading
the file, not by running the detector:

```
boundaries : 10/10 (100%)
kinds      :  8/10 (80%)
spurious   :  0
```

Both kind errors are among the regions the detector flagged **below 0.70** — it
is wrong only where it says it is unsure, which the test suite enforces. This is
one workbook written to exercise known failure modes; do not quote it as corpus
accuracy.

## Setup

```bash
uv sync                      # runtime deps: lxml only
uv sync --extra encrypted    # + olefile, to name encrypted files in triage
```

## Not built yet

Intra-workbook dependency graph. `formulas.analyse()` already returns
`precedent_cells`, so the cell→cell edges exist; collapsing them to region→region
and deriving input / calculation / output roles is the next step — the part that
demonstrates comprehension rather than extraction.
