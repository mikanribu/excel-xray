#!/usr/bin/env python3
"""Excel X-ray command line.

  excel-xray FILE.xlsx                 xlsx report in a timestamped folder
                                        next to the file
  excel-xray FOLDER -o out/            every workbook in a folder, reports in
                                        out/xray_<timestamp>/
  excel-xray FILE.xlsx --format html   HTML report instead of xlsx
  excel-xray FILE.xlsx --json          machine-readable JSON to stdout

Every run creates a fresh xray_<YYYYMMDD_HHMMSS>/ subfolder under the given
(or default) output path, so repeated runs never overwrite an earlier report.

Reads only. Never writes to, moves or renames a source file.
"""

from __future__ import annotations

import argparse
import os
import sys
import time
import traceback

import json
import shutil
from pathlib import Path

from .assessment import assess
from .assessment import to_dict as assessment_to_dict
from .report import write_report
from .scan import UnreadableWorkbook, to_json, xray_workbook

EXTS = {".xlsx", ".xlsm", ".xltx", ".xltm"}
SKIP_PREFIX = ("~$", ".")


def _narrative_assessor(args):
    """Build the --llm narrative assessor for the chosen --provider, or None."""
    if not args.llm:
        return None
    max_tokens = int(os.environ.get("LLM_MAX_TOKENS") or 2000)
    if args.provider == "openai":
        from .narrative import OpenAIAssessor
        return OpenAIAssessor(model=args.model, max_tokens=max_tokens,
                               azure_endpoint=args.azure_endpoint,
                               api_version=args.api_version)
    from .narrative import ClaudeAssessor
    return ClaudeAssessor(model=args.model, max_tokens=max_tokens)


def _estate_assessor(args):
    """Build the --llm estate-insight assessor for the chosen --provider, or None."""
    if not args.llm:
        return None
    if args.provider == "openai":
        from .estate_insight import OpenAIEstateAssessor
        return OpenAIEstateAssessor(model=args.model,
                                     azure_endpoint=args.azure_endpoint,
                                     api_version=args.api_version)
    from .estate_insight import ClaudeEstateAssessor
    return ClaudeEstateAssessor(model=args.model)


def collect(target: str) -> list[str]:
    if os.path.isfile(target):
        return [target]
    out = []
    for root, dirs, files in os.walk(target):
        # A repeat run under the same folder must not ingest a previous run's
        # generated portfolio_review.xlsx as though it were a source EUC.
        dirs[:] = [d for d in dirs if not d.startswith("xray_") and d not in
                   {".git", ".venv", "__pycache__"}]
        for f in files:
            if f.startswith(SKIP_PREFIX):
                continue
            if os.path.splitext(f)[1].lower() in EXTS:
                out.append(os.path.join(root, f))
    return sorted(out)


def main() -> int:
    if len(sys.argv) > 1 and sys.argv[1] == "serve":
        serve_ap = argparse.ArgumentParser(prog="excel-xray serve",
                                           description="Serve a portfolio run on localhost")
        serve_ap.add_argument("run_dir", help="folder containing portfolio.sqlite")
        serve_ap.add_argument("--port", type=int, default=8765)
        serve_args = serve_ap.parse_args(sys.argv[2:])
        from .portfolio_web import serve
        serve(serve_args.run_dir, port=serve_args.port)
        return 0
    ap = argparse.ArgumentParser(
        prog="excel-xray", description="Excel X-ray structural scanner"
    )
    ap.add_argument("target", help="workbook or folder")
    ap.add_argument("-o", "--out", default=None,
                    help="base output directory (default: next to the file, or "
                         "the target folder itself); a timestamped "
                         "xray_<YYYYMMDD_HHMMSS>/ subfolder is created under it "
                         "on every run")
    ap.add_argument("--json", action="store_true", help="emit scan JSON to stdout")
    ap.add_argument("--assess", action="store_true",
                    help="emit the EUC assessment JSON to stdout")
    ap.add_argument("--llm", action="store_true",
                    help="use a model assessor for narrative fields "
                         "(needs a credential for the chosen --provider)")
    ap.add_argument("--provider", choices=["claude", "openai"], default=None,
                    help="LLM backend for --llm: 'claude' (Anthropic, needs the "
                         "'anthropic' package) or 'openai' (GPT via the public "
                         "OpenAI API, or an Azure OpenAI deployment if "
                         "--azure-endpoint/$AZURE_OPENAI_ENDPOINT is set; needs "
                         "the 'openai' package). Default: claude, unless an "
                         "Azure endpoint is given, which implies openai.")
    ap.add_argument("--model", default=None,
                    help="model id for --llm. Claude default: $ANTHROPIC_MODEL "
                         "or claude-opus-5. OpenAI default: $OPENAI_MODEL or "
                         "gpt-4o; on Azure this is the deployment name "
                         "(else $AZURE_OPENAI_DEPLOYMENT).")
    ap.add_argument("--azure-endpoint", default=None,
                    help="Azure OpenAI resource endpoint, e.g. "
                         "https://<resource>.openai.azure.com "
                         "(else $AZURE_OPENAI_ENDPOINT). Implies --provider openai.")
    ap.add_argument("--api-version", default=None,
                    help="Azure OpenAI api-version (else $AZURE_OPENAI_API_VERSION, "
                         "default 2024-10-21)")
    ap.add_argument("--csv", default=None, metavar="PATH",
                    help="also write the EUC assessment as a CSV table")
    ap.add_argument("--estate", action="store_true",
                    help="compare workbooks across the folder: write an estate "
                         "report (estate.<format> + estate_pairs.csv) to the out dir, "
                         "and embed an Estate comparison sheet in each xlsx report")
    ap.add_argument("--format", choices=["xlsx", "html"], default="xlsx",
                    help="report format for the per-workbook and estate report "
                         "(default: xlsx)")
    ap.add_argument("--max-rows", type=int, default=200_000)
    ap.add_argument("--portfolio", action="store_true",
                    help="create a consolidated portfolio run (default for folders)")
    ap.add_argument("--individual-reports", action="store_true",
                    help="for a folder, retain the legacy one-report-per-EUC output")
    ap.add_argument("--resume", metavar="RUN_DIR", default=None,
                    help="resume a portfolio run, reusing unchanged scanned EUCs")
    args = ap.parse_args()

    if args.llm:
        # Load a local .env so ANTHROPIC_API_KEY / OPENAI_API_KEY /
        # AZURE_OPENAI_* can live there. Best-effort: python-dotenv ships
        # with the [llm] extra, so this is a no-op otherwise.
        try:
            from dotenv import load_dotenv
            load_dotenv()
        except ImportError:
            pass

    args.azure_endpoint = args.azure_endpoint or os.environ.get("AZURE_OPENAI_ENDPOINT")
    args.api_version = args.api_version or os.environ.get("AZURE_OPENAI_API_VERSION")
    # An Azure endpoint only makes sense for the openai provider.
    args.provider = args.provider or ("openai" if args.azure_endpoint else "claude")

    # Resolve model / token budget: explicit flag wins, then the environment
    # (loaded from .env above), then a per-provider default.
    if args.provider == "openai":
        args.model = (args.model or os.environ.get("AZURE_OPENAI_DEPLOYMENT")
                      or os.environ.get("OPENAI_MODEL") or "gpt-4o")
    else:
        args.model = args.model or os.environ.get("ANTHROPIC_MODEL") or "claude-opus-5"

    paths = collect(args.target)
    if not paths:
        print(f"no workbooks found under {args.target}", file=sys.stderr)
        return 2

    folder_portfolio = os.path.isdir(args.target) and not (
        args.json or args.assess or args.estate or args.individual_reports)
    if args.portfolio or args.resume or folder_portfolio:
        from .portfolio import DB_NAME, SUMMARY_NAME, scan_portfolio
        base = args.out or (args.target if os.path.isdir(args.target)
                            else os.path.dirname(os.path.abspath(args.target)))
        if args.resume and not (Path(args.resume) / DB_NAME).is_file():
            ap.error(f"--resume requires an existing portfolio run containing {DB_NAME}")
        run_dir = args.resume or os.path.join(
            base, f"xray_{time.strftime('%Y%m%d_%H%M%S')}_{time.time_ns() % 1_000_000_000:09d}")
        result = scan_portfolio(paths, run_dir, max_rows=args.max_rows,
                                assessor=_narrative_assessor(args), resume=bool(args.resume))
        if args.csv:
            shutil.copyfile(os.path.join(run_dir, SUMMARY_NAME), args.csv)
        print(f"Portfolio ready: {run_dir} ({result['stored']} EUCs; "
              f"{result['failed']} failed; {result['reused']} reused)", file=sys.stderr)
        print(f"Open with: excel-xray serve '{run_dir}'", file=sys.stderr)
        return 0

    base_outdir = args.out or (args.target if os.path.isdir(args.target)
                               else os.path.dirname(os.path.abspath(args.target)))
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    outdir = os.path.join(base_outdir, f"xray_{timestamp}")
    # --json/--assess print to stdout and --csv writes to its own given path —
    # none of those touch outdir. Only the report loop and --estate do, so
    # skip creating an empty timestamped folder when neither applies.
    writes_to_outdir = args.estate or (not args.json and not args.assess)
    if writes_to_outdir:
        os.makedirs(outdir, exist_ok=True)
        print(f"writing reports to {outdir}", file=sys.stderr)

    ok = partial = failed = 0
    reasons: dict[str, list[str]] = {}
    batch: list = []  # (path, wx) for assessment-aware output modes
    timings: dict[str, float] = {}  # path -> extraction seconds, for the report line below
    for p in paths:
        t0 = time.perf_counter()
        try:
            wx = xray_workbook(p, max_rows=args.max_rows)
        except UnreadableWorkbook as e:
            elapsed = time.perf_counter() - t0
            failed += 1
            reasons.setdefault(e.category, []).append(os.path.basename(p))
            print(f"SKIPPED  {os.path.basename(p):46} {e}  ({elapsed:.2f}s)", file=sys.stderr)
            continue
        except Exception as e:  # noqa: BLE001 - report and keep going over a corpus
            elapsed = time.perf_counter() - t0
            failed += 1
            reasons.setdefault("unexpected", []).append(os.path.basename(p))
            print(f"FAILED   {os.path.basename(p):46} {type(e).__name__}: {e}  "
                  f"({elapsed:.2f}s)", file=sys.stderr)
            if os.environ.get("XRAY_DEBUG"):
                traceback.print_exc()
            continue
        elapsed = time.perf_counter() - t0
        timings[p] = elapsed

        ok += wx.parse_status == "full"
        partial += wx.parse_status == "partial"
        print(f"EXTRACT  {os.path.basename(p):46} {elapsed:6.2f}s", file=sys.stderr)
        if args.json:  # raw scan, no assessment
            print(to_json(wx))
        else:
            batch.append((p, wx))

    # Assessment-aware modes (HTML report, --assess, --csv). One corpus pass so
    # duplication/consolidation see the whole set.
    assessments = None
    if batch and (args.assess or args.csv or args.estate or not args.json):
        assessor = _narrative_assessor(args)
        wxs = [wx for _, wx in batch]
        if len(wxs) > 1:
            from .corpus import assess_corpus
            assessments = assess_corpus(wxs, assessor)
        else:
            assessments = [assess(wxs[0], assessor)]

    if args.assess and assessments is not None:
        payload = [assessment_to_dict(a) for a in assessments]
        print(json.dumps(payload if len(payload) != 1 else payload[0],
                         indent=2, default=str))

    # Estate comparison is computed before the per-file reports so an xlsx
    # report can embed it as its own sheet, per the required Excel content.
    estate = insight = None
    if args.estate and assessments is not None:
        if len(assessments) < 2:
            print("--estate needs more than one workbook to compare", file=sys.stderr)
        else:
            from .estate import build_estate
            from .estate_insight import generate_estate_insight
            pairs_in = [(wx, a) for (_, wx), a in zip(batch, assessments)]
            estate = build_estate(pairs_in)
            insight = generate_estate_insight(estate, _estate_assessor(args))

    if not args.json and not args.assess and assessments is not None:
        ext = "xlsx" if args.format == "xlsx" else "html"
        for (p, wx), a in zip(batch, assessments):
            name = os.path.splitext(os.path.basename(p))[0]
            dest = os.path.join(outdir, f"xray_{name}.{ext}")
            if args.format == "xlsx":
                from .xlsx_report import write_xlsx_report
                write_xlsx_report(wx, a, dest, estate=estate, estate_insight=insight)
            else:
                write_report(wx, dest, a)
            regions = sum(len(s.regions) for s in wx.sheets)
            need_review = sum(1 for t in a.tabs if t.human_validation_required.value == "Y")
            print(f"{wx.parse_status:8} {os.path.basename(p):46} "
                  f"{len(wx.sheets):3} sheets  {regions:3} regions  "
                  f"{need_review:2} need review  {timings.get(p, 0):5.2f}s  -> {dest}")

    if args.csv and assessments is not None:
        from .tabular import to_csv
        named = [(wx.filename, a) for (_, wx), a in zip(batch, assessments)]
        to_csv(named, args.csv)
        print(f"wrote {args.csv} ({len(named)} workbook(s))", file=sys.stderr)

    if estate is not None:
        ext = "xlsx" if args.format == "xlsx" else "html"
        dest = os.path.join(outdir, f"estate.{ext}")
        csv_path = os.path.join(outdir, "estate_pairs.csv")
        if args.format == "xlsx":
            from .xlsx_report import write_estate_xlsx
            write_estate_xlsx(estate, dest, insight)
        else:
            from .estate_report import write_estate_report
            write_estate_report(estate, dest, insight)
        from .estate_report import write_estate_csv
        write_estate_csv(estate, csv_path)
        print(f"{len(estate.fingerprints)} workbooks  "
              f"{len(estate.clusters)} families  {len(estate.pairs)} linked pairs"
              f"  -> {dest}", file=sys.stderr)

    total = len(paths)
    print(f"\ncoverage: {ok}/{total} full, {partial} partial, {failed} unreadable",
          file=sys.stderr)
    for cat, files in sorted(reasons.items()):
        print(f"  {cat:12} {len(files):3}  {', '.join(files[:4])}"
              + (" ..." if len(files) > 4 else ""), file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
