# Excel X-ray

Scans a complicated financial workbook and produces two things:

1. a verifiable **structural map** — regions, headers, calculation profile,
   dependencies and quality flags, and
2. an **EUC assessment** — the reviewer-facing fields a controls team needs
   (purpose, complexity, logic type, inputs, key findings, per-tab detail,
   duplication across a folder) — as an Excel report or local portfolio dashboard.

The scanner reads source workbooks without changing them. It writes all results
to a separate output folder.

## Quick start: review a folder of EUCs

Use Python 3.12 or newer and [uv](https://docs.astral.sh/uv/getting-started/installation/) on macOS, Linux,
or Windows. These commands are for a macOS/Linux terminal; on Windows, run the
equivalent commands in PowerShell and use the printed run-folder path in step 4.
No API key or model service is needed for the default offline assessment.

1. Get **this branch** and install its dependencies:

   ```bash
   git clone --branch codex/portfolio-dashboard-exports https://github.com/mikanribu/excel-xray.git
   cd excel-xray
   uv sync
   ```

2. Put the two (or more) `.xlsx`/`.xlsm` workbooks you want to review in a
   folder such as `inputs/`. This folder is Git-ignored. For confidential
   originals, you can instead pass an absolute path to a folder outside the
   repository.

   ```bash
   mkdir -p inputs out
   cp "/path/to/first.xlsx" "/path/to/second.xlsx" inputs/
   ```

   Replace the two example paths with the paths to your actual workbooks.
   `inputs/` and `out/` are Git-ignored so local workbook contents and reports
   are not included in a code commit.

3. Scan the folder. The folder command produces **one portfolio run**, not one
   JSON/report file per workbook:

   ```bash
   uv run excel-xray inputs/ -o out/
   ```

   The command prints `Portfolio ready: out/xray_<timestamp>_<suffix>`.
   The suffix makes repeated runs unique. A run containing two workbooks has
   two data rows in `file_summary.csv`, plus one header row.

4. Open the dashboard. This selects the newest run under `out/`:

   ```bash
   RUN_DIR="$(ls -dt out/xray_* | head -n 1)"
   uv run excel-xray serve "$RUN_DIR"
   ```

   Open the printed URL, normally <http://127.0.0.1:8765/>. Keep this terminal
   running while using the dashboard; press `Ctrl+C` to stop it. If port 8765
   is occupied, add `--port 8877` and open the URL printed for that port.

5. In the dashboard, click a file name to inspect its assessment, worksheets,
   cached-error diagnostics and hidden-sheet groups. Tick multiple files and
   click **Compare selected** for a side-by-side view and structural matches.
   Export buttons use the selected files; with none selected they export the
   entire portfolio. **Review Excel** is the easiest file to send for review.
   **Download original EUC** appears in a file's detail view. **Originals +
   analysis ZIP** bundles source workbooks with their analysis.

### Email the Excel report

Email delivery is opt-in. A folder scan sends the single consolidated
`portfolio_review.xlsx`; a one-workbook scan sends its generated Excel report.
It never sends one message per EUC. The default recipient is
`jacobweglarz@gmail.com`, and `--email-to` can override it.

Set `GMAIL_ADDRESS` to the Gmail account that will send the report and
`GMAIL_APP_PASSWORD` to an [App Password](https://support.google.com/accounts/answer/2461835)
created for that account. Google requires 2-Step Verification for App Passwords.
Use an App Password instead of your normal Google password, and keep it in your
shell environment or a local secret manager; do not commit it to the repository.
The sender uses Gmail's encrypted SMTP service at `smtp.gmail.com`.

```bash
export GMAIL_ADDRESS="your-sender@gmail.com"
export GMAIL_APP_PASSWORD="your-16-character-app-password"

# One message with the consolidated workbook for the whole folder
uv run excel-xray inputs/ -o out/ --email-report

# Send to a different reviewer
uv run excel-xray inputs/ -o out/ --email-report --email-to reviewer@example.com

# One workbook: send its individual Excel report
uv run excel-xray "inputs/one-euc.xlsx" -o out/ --email-report
```

The report is generated locally even if email delivery fails; the command then
prints the delivery error and exits nonzero so the send can be retried. Reports
larger than 17 MiB are kept locally and not attached, to leave room for email
attachment encoding. The workbook contains analysis for every EUC in the folder,
so check that the recipient is cleared to review it before enabling delivery.

### What the run creates

| File in `RUN_DIR` | Use |
|---|---|
| `portfolio_review.xlsx` | Shareable Excel review workbook: File summary, Worksheet details, Diagnostics, Portfolio findings |
| `file_summary.csv` | One row per submitted EUC, including a failed-scan row where applicable |
| `worksheet_details.csv` | One row per worksheet with visibility, role, dependencies and validation reason |
| `diagnostics.csv` | Grouped cached errors and hidden-sheet purpose details |
| `portfolio_findings.csv` | Findings consolidated across the EUCs in the run |
| `portfolio.sqlite` | Indexed dashboard data; keep this file and the source workbooks available to use the dashboard later |

The consolidated file summary includes scan status, concise scan errors,
total/hidden worksheet counts, Business Area / Process, Process and
Sub-Process, plus the file-level review fields. A process label is left as
“Not established” when the workbook does not contain direct supporting
evidence. Input sources are grouped by purpose and type, with workbook names,
reference counts, known worksheet consumers and an explicit owner check for
source essentiality.

Reviewer-facing CSV and Excel exports do not include local source paths, file
sizes or raw SHA-256 values. Technical formula patterns remain in the
individual report's formula appendix; the Key calculations field contains
plain-language descriptions. Original workbooks can still be downloaded
individually or included in the selected-files ZIP.

The default folder run creates no per-file JSON or HTML reports. Individual
HTML/Excel reports are generated on demand in the dashboard. To scan one
workbook directly, use `uv run excel-xray file.xlsx -o out/`; it writes the
existing individual Excel report. Add `--format html` for individual HTML.

### Other commands

```bash
# Resume an interrupted portfolio run; unchanged files are reused and failures retried.
uv run excel-xray inputs/ --resume "$RUN_DIR"

# Also copy the consolidated, one-row-per-EUC CSV to a chosen location.
uv run excel-xray inputs/ -o out/ --csv euc_summary.csv

# Older one-report-per-workbook folder output, if specifically needed.
uv run excel-xray inputs/ -o out/ --individual-reports

# Raw structure or assessment JSON for a single workbook, printed to the terminal.
uv run excel-xray file.xlsx --json
uv run excel-xray file.xlsx --assess
```

`--resume` must point to the **run folder** containing `portfolio.sqlite`, not
the parent `out/` folder. The dashboard reads originals from their recorded
source paths. If an original is moved or modified after scanning, resume the
run from its current input folder before downloading that original or opening
its on-demand report.

The analysis exports contain structural evidence and assessment text, not raw
cell values. The **original EUC download and ZIP do contain the source workbook
bytes and cell values**; share them only with recipients cleared for those
files. Generated output stays local and is not committed to GitHub.

At 2,000 EUCs, the portfolio scan holds one workbook at a time and keeps its
assessments in one SQLite database. Cross-EUC comparison uses compact
fingerprints and is quadratic in file count; only the 20 strongest review
matches per EUC are stored. The dashboard recomputes selected-file similarity
so selected pairs remain comparable even when outside that shortlist.

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

The scanner does not depend on openpyxl, but the Excel report writer does. The
scanner's cell reader is verified cell-for-cell against openpyxl on the fixture
(values, formulas, bold, fill, border — zero mismatches).

## Architecture

Layers, deliberately separated.

| Module | Does |
|---|---|
| [ooxml.py](src/excel_xray/ooxml.py) | Structure straight from the zip via `lxml`: merges, ListObjects, connections, external links, pivot caches, VBA, Power Query |
| [sheetmodel.py](src/excel_xray/sheetmodel.py) | Single `lxml` pass over `<sheetData>` → cells (value + formula + style) and a formula profile. Resolves `sharedStrings` and `styles` |
| [regions.py](src/excel_xray/regions.py) | Occupancy runs → connected components → header inference → classification. Detects hand-built table regions nothing off-the-shelf finds |
| [formulas.py](src/excel_xray/formulas.py) | A1 → R1C1 → literal abstraction. Collapses a filled-down column to one skeleton |
| [scan.py](src/excel_xray/scan.py) | Triage, orchestration, occupancy plate |
| [assessment.py](src/excel_xray/assessment.py) | Evidence → EUC schema: complexity, controlled logic types, tab categories, dependencies, human-validation, heuristic findings |
| [narrative.py](src/excel_xray/narrative.py) | Narrative fields behind an `Assessor` interface: offline template (default) or Claude (`--llm`) |
| [corpus.py](src/excel_xray/corpus.py) | Per-file duplication/consolidation fields (formula shapes + headers) |
| [portfolio.py](src/excel_xray/portfolio.py) | One-at-a-time folder scan, SQLite store, cross-EUC findings, consolidated CSV/Excel and selected bundles |
| [portfolio_web.py](src/excel_xray/portfolio_web.py) | Local dashboard, drill-down, comparison and download routes |
| [xlsx_report.py](src/excel_xray/xlsx_report.py) | Existing individual and estate Excel reports |
| [estate.py](src/excel_xray/estate.py) | Estate comparison: four-signal fingerprints, relationship typing, clustering |
| [estate_insight.py](src/excel_xray/estate_insight.py) | Interprets families → recommendations (offline template or Claude) |
| [estate_report.py](src/excel_xray/estate_report.py) | Standalone estate HTML + pairs CSV |
| [tabular.py](src/excel_xray/tabular.py) | The review-table schema; drives the HTML tables and the CSV export |
| [report.py](src/excel_xray/report.py) | Self-contained HTML — no CDN, no network |
| [xlsx_report.py](src/excel_xray/xlsx_report.py) | The default `.xlsx` report: one workbook per assessed file, plus the estate comparison sheet. Blank Reviewer Value/Notes columns for the human sign-off |
| [portfolio.py](src/excel_xray/portfolio.py) | Bounded-memory folder scan for the consolidated portfolio run: `portfolio.sqlite` store, CSV/xlsx exports, `--resume` |
| [portfolio_web.py](src/excel_xray/portfolio_web.py) | The `excel-xray serve` dashboard: a local-only HTTP server reading `portfolio.sqlite` for paged drill-down, search, comparison and on-demand exports |
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

**File level summary** — File ID/Name, Business Area / Process, Process,
Sub-Process, Purpose, Key Output / Outcome, Complexity, Key Inputs, Source System, Key Outputs, Usage Frequency,
Completion Timeline, EUC Preparer, Output Recipient; the AI findings (Potential
Duplication, Similar/Duplicate Files, Simplification, Consolidation, Automation,
Retirement); and workbook logic (Logic Types, Key calculations, Reconciliation
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

Purpose, Key Output/Outcome, Process, Sub-Process and each tab's Purpose are written by an
`Assessor`. The default is offline (network-free, `drafted`); unknown business
labels stay “Not established” for owner confirmation. Pass `--llm` to use a
model instead (`inferred`); process labels must be supported by exact workbook
text. This needs the optional `anthropic` and/or `openai`
package and a credential:

```bash
uv sync --extra llm
ANTHROPIC_API_KEY=sk-ant-... uv run excel-xray file.xlsx --assess --llm
```

`--provider` chooses the backend: `claude` (Anthropic, the default) or `openai`
(the public OpenAI API, or an Azure OpenAI deployment once `--azure-endpoint` /
`$AZURE_OPENAI_ENDPOINT` is set — which also implies `--provider openai`).
`--model` overrides the per-provider default model (Claude: `$ANTHROPIC_MODEL`
or `claude-opus-5`; OpenAI: `$OPENAI_MODEL` or `gpt-4o`; Azure: the deployment
name, from `$AZURE_OPENAI_DEPLOYMENT` or `--model`).

```bash
OPENAI_API_KEY=sk-... uv run excel-xray file.xlsx --assess --llm --provider openai
AZURE_OPENAI_API_KEY=... uv run excel-xray file.xlsx --assess --llm \
  --azure-endpoint https://<resource>.openai.azure.com --model <deployment-name>
```

Or put credentials in a `.env` file (copy [.env.example](.env.example) to
`.env`) — with `--llm` the CLI loads it automatically. `.env` is git-ignored, so
keys are never committed. The Claude path also accepts an `ant auth login`
profile if you have one. Only a **value-free structural bundle** (headers +
normalised formula shapes, never cell values) is sent to the model — the same
holds for the estate insight layer's `ClaudeEstateAssessor` /
`OpenAIEstateAssessor` below.

For a folder of EUCs, enable OpenAI narrative analysis on the consolidated
portfolio run with:

```bash
uv sync --extra llm
uv run excel-xray inputs/ -o out/ --llm --provider openai
```

Set `OPENAI_API_KEY` in the shell running the command or in the project-root
`.env` file. If using `.env.example`, uncomment and fill its OpenAI key line.
The run still produces one consolidated `portfolio_review.xlsx` for the folder.

## Estate comparison

`--estate` is the older, separate estate-analysis mode. Over a folder it
compares every workbook against every other on four independent signals and
writes `estate.xlsx` + `estate_pairs.csv` by default (use `--format html` for
`estate.html`):

| Signal | Captures | A match means |
|---|---|---|
| Formula shapes (R1C1 skeletons) | the calculation method | *same method* |
| Input signature (input headers, sources, connections, links) | what it consumes | *same source / lineage* |
| Output signature (output-tab headers) | the deliverable | *same output* |
| Dependency topology (tab-role counts + cross-sheet edges) | the pipeline shape | *same process shape* |

Instead of one blended score, each pair gets a **relationship**: **Duplicate**
(same method, inputs and outputs), **Same output, different method** (→
consolidate), **Overlapping logic** (→ extract a reusable component), or **Shared
source** (common data lineage). Linked pairs are clustered (connected components)
into **families**, and the report shows the families, a similarity matrix, and
the per-signal breakdown for each pair.

**Insight layer.** The scores, relationships and clusters stay deterministic
(reproducible and auditable). On top of them an *insight* layer interprets each
family — what it is, and a recommended action (consolidate / keep one / extract
shared logic / align source) — plus a ranked estate-level opportunities list. It
runs offline by default (basis `drafted`) and upgrades to a model with `--llm`
(basis `inferred`, same `--provider`/`--model` choice as above); only
fingerprint metadata is sent, never cell values.

```python
from excel_xray import xray_workbook, assess, build_estate
pairs = [(wx, assess(wx)) for wx in map(xray_workbook, paths)]
estate = build_estate(pairs)
for i, j, c in estate.pairs:
    print(estate.fingerprints[i].file_name, c["relationship"], c["overall"])
```

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
uv sync                      # runtime deps: lxml and openpyxl
uv sync --extra encrypted    # + olefile, to name encrypted files in triage
uv sync --extra llm          # + anthropic, openai, python-dotenv, for --llm
```

## Known limits

- Reads `.xlsx` / `.xlsm`. Legacy `.xls` is triaged and reported, not parsed.
- Three fields are genuinely not in the file and stay `needs_human`: Usage
  Frequency, Completion Timeline, Output Recipient. A reviewer supplies these.
- The heuristic findings and tab categories carry a confidence and evidence;
  treat anything the tool flags `Uncertain` or below `0.70` as a prompt to look.
