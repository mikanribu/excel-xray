# Excel X-ray

Scans a complicated financial workbook and produces two things:

1. a verifiable **structural map** — regions, headers, calculation profile,
   dependencies and quality flags, and
2. an **EUC assessment** — the reviewer-facing fields a controls team needs
   (purpose, complexity, logic type, inputs, key findings, per-tab detail,
   duplication across a folder) — as a self-contained HTML report, JSON, or CSV.

Read-only. Never writes to, moves or renames a source file.

```bash
uv run excel-xray file.xlsx -o out/            # HTML report (structure + assessment)
uv run excel-xray /path/to/folder -o out/      # one report per workbook + corpus findings
uv run excel-xray file.xlsx --assess           # EUC assessment as JSON to stdout
uv run excel-xray /path/to/folder --csv euc.csv   # assessment as a flat CSV table
uv run excel-xray file.xlsx --json             # raw structural scan as JSON
uv run pytest                                  # accuracy + reader + assessment tests
```

## Two layers

**Evidence** — the scanner reads the workbook and computes hard facts: cell
values/formulas, detected table regions, normalised formula shapes, the
dependency graph, macros/links/connections.

**Assessment** — a layer maps that evidence onto the EUC review schema. Every
field declares its **basis**, so the output is honest about what is grounded in
the file versus what still needs a model, a human, or the wider corpus:

| Basis | Meaning |
|---|---|
| `extracted` | read directly from the file |
| `derived` | a heuristic over extracted metrics (with evidence) |
| `drafted` | offline narrative template from the evidence |
| `inferred` | written by the LLM assessor (`--llm`) |
| `needs_human` | not knowable from the file (e.g. usage frequency) |
| `needs_corpus` | needs the whole folder (only when scanning one file) |

Nothing is fabricated: a `drafted` value is never presented as a considered one,
and fields that genuinely aren't in the file stay `needs_human`.

## Why not openpyxl

openpyxl is the obvious library, and it is the wrong tool for both halves of the
scan:

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
| [assessment.py](src/excel_xray/assessment.py) | Evidence → EUC schema: complexity, logic type, tab categories, dependencies, human-validation, heuristic findings |
| [narrative.py](src/excel_xray/narrative.py) | Narrative fields behind an `Assessor` interface: offline template (default) or Claude (`--llm`) |
| [corpus.py](src/excel_xray/corpus.py) | Fingerprint + pairwise compare across a folder → duplication / consolidation |
| [tabular.py](src/excel_xray/tabular.py) | The review-table schema; drives the HTML tables and the CSV export |
| [report.py](src/excel_xray/report.py) | Self-contained HTML — no CDN, no network |
| [util.py](src/excel_xray/util.py) | A1-notation helpers (replaces `openpyxl.utils`) |

### The asymmetry that matters

```
row_gap = 1   a blank row inside a table is common (spacer, visual break)
col_gap = 0   a blank column between blocks is the standard separator
```

Getting this backwards merges an entire sheet into one region. It is the single
most consequential constant in the detector.

## The EUC assessment

Produced at two levels, matching the review template:

**File level summary** — File ID/Name, Business Area, Purpose, Key Output /
Outcome, Complexity, Key Inputs, Source System, Key Outputs, Usage Frequency,
Completion Timeline, EUC Preparer, Output Recipient; the AI findings (Potential
Duplication, Similar/Duplicate Files, Simplification, Consolidation, Automation,
Retirement); and workbook logic (Logic Type, Key calculations, Reconciliation
logic, Manual intervention, Macros/VBA/links).

**Tab level details** — per sheet: Tab Name, Category (Input / Calculation /
Mapping / Control Check / Validation / Output, or Uncertain), Purpose,
Information Analysis, Key Calculation/Transformation Logic, Upstream and
Downstream dependencies, and a Human-Validation-Required flag with reasons.

```python
from excel_xray import xray_workbook, assess, assess_corpus
a = assess(xray_workbook("workbook.xlsx"))
print(a.file.complexity.value, a.file.logic_type.value)
for t in a.tabs:
    print(t.tab_name.value, "→", t.tab_category.value)
```

### Narrative fields and the LLM

Purpose, Key Output/Outcome, Key Outputs and each tab's Purpose are written by an
`Assessor`. The default is offline (network-free, `drafted`). Pass `--llm` to use
Claude instead (`inferred`); this needs the optional `anthropic` package and a
credential:

```bash
uv sync --extra llm
ANTHROPIC_API_KEY=sk-ant-... uv run excel-xray file.xlsx --assess --llm
```

Only a **value-free structural bundle** (headers + normalised formula shapes,
never cell values) is sent to the model.

### Privacy

The report, JSON and CSV contain cell **coordinates, types, normalised formula
shapes, headers and derived assessments** — **no cell values**. An X-ray can be
circulated to people who are not cleared for the underlying figures, and the
`--llm` bundle carries no values either.

## Measured accuracy

Region detection against hand-marked ground truth on the fixture
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
uv sync --extra llm          # + anthropic, for the --llm narrative assessor
```

## Known limits

- Reads `.xlsx` / `.xlsm`. Legacy `.xls` is triaged and reported, not parsed.
- Three fields are genuinely not in the file and stay `needs_human`: Usage
  Frequency, Completion Timeline, Output Recipient. A reviewer supplies these.
- The heuristic findings and tab categories carry a confidence and evidence;
  treat anything the tool flags `Uncertain` or below `0.70` as a prompt to look.
