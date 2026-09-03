#!/usr/bin/env python3
"""Excel X-ray command line.

  excel-xray FILE.xlsx                 HTML report next to the file
  excel-xray FOLDER -o out/            every workbook in a folder
  excel-xray FILE.xlsx --json          machine-readable JSON to stdout

Reads only. Never writes to, moves or renames a source file.
"""

from __future__ import annotations

import argparse
import os
import sys
import time
import traceback

import json

from .assessment import assess
from .assessment import to_dict as assessment_to_dict
from .report import write_report
from .scan import UnreadableWorkbook, to_json, xray_workbook

EXTS = {".xlsx", ".xlsm", ".xltx", ".xltm"}
SKIP_PREFIX = ("~$", ".")


def collect(target: str) -> list[str]:
    if os.path.isfile(target):
        return [target]
    out = []
    for root, _dirs, files in os.walk(target):
        for f in files:
            if f.startswith(SKIP_PREFIX):
                continue
            if os.path.splitext(f)[1].lower() in EXTS:
                out.append(os.path.join(root, f))
    return sorted(out)


def main() -> int:
    ap = argparse.ArgumentParser(
        prog="excel-xray", description="Excel X-ray structural scanner"
    )
    ap.add_argument("target", help="workbook or folder")
    ap.add_argument("-o", "--out", default=None, help="output directory")
    ap.add_argument("--json", action="store_true", help="emit scan JSON to stdout")
    ap.add_argument("--assess", action="store_true",
                    help="emit the EUC assessment JSON to stdout")
    ap.add_argument("--llm", action="store_true",
                    help="use the Claude assessor for narrative fields "
                         "(needs the 'anthropic' package + a credential)")
    ap.add_argument("--model", default="claude-opus-5",
                    help="model id for --llm (default: claude-opus-5)")
    ap.add_argument("--csv", default=None, metavar="PATH",
                    help="also write the EUC assessment as a CSV table")
    ap.add_argument("--max-rows", type=int, default=200_000)
    args = ap.parse_args()

    if args.llm:
        # Load a local .env so ANTHROPIC_API_KEY can live there. Best-effort:
        # python-dotenv ships with the [llm] extra, so this is a no-op otherwise.
        try:
            from dotenv import load_dotenv
            load_dotenv()
        except ImportError:
            pass

    paths = collect(args.target)
    if not paths:
        print(f"no workbooks found under {args.target}", file=sys.stderr)
        return 2

    outdir = args.out or (args.target if os.path.isdir(args.target)
                          else os.path.dirname(os.path.abspath(args.target)))
    os.makedirs(outdir, exist_ok=True)

    ok = partial = failed = 0
    reasons: dict[str, list[str]] = {}
    batch: list = []  # (path, wx) for assessment-aware output modes
    for p in paths:
        try:
            wx = xray_workbook(p, max_rows=args.max_rows)
        except UnreadableWorkbook as e:
            failed += 1
            reasons.setdefault(e.category, []).append(os.path.basename(p))
            print(f"SKIPPED  {os.path.basename(p):46} {e}", file=sys.stderr)
            continue
        except Exception as e:  # noqa: BLE001 - report and keep going over a corpus
            failed += 1
            reasons.setdefault("unexpected", []).append(os.path.basename(p))
            print(f"FAILED   {os.path.basename(p):46} {type(e).__name__}: {e}",
                  file=sys.stderr)
            if os.environ.get("XRAY_DEBUG"):
                traceback.print_exc()
            continue

        ok += wx.parse_status == "full"
        partial += wx.parse_status == "partial"
        if args.json:  # raw scan, no assessment
            print(to_json(wx))
        else:
            batch.append((p, wx))

    # Assessment-aware modes (HTML report, --assess, --csv). One corpus pass so
    # duplication/consolidation see the whole set.
    assessments = None
    if batch and (args.assess or args.csv or not args.json):
        assessor = None
        if args.llm:
            from .narrative import ClaudeAssessor
            assessor = ClaudeAssessor(model=args.model)
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
    elif not args.json and assessments is not None:
        for (p, wx), a in zip(batch, assessments):
            name = os.path.splitext(os.path.basename(p))[0]
            dest = os.path.join(outdir, f"xray_{name}.html")
            write_report(wx, dest, a)
            regions = sum(len(s.regions) for s in wx.sheets)
            low = sum(1 for s in wx.sheets for r in s.regions
                      if r.detect_confidence < 0.70)
            print(f"{wx.parse_status:8} {os.path.basename(p):46} "
                  f"{len(wx.sheets):3} sheets  {regions:3} regions  "
                  f"{low:2} low-conf  -> {dest}")

    if args.csv and assessments is not None:
        from .tabular import to_csv
        named = [(wx.filename, a) for (_, wx), a in zip(batch, assessments)]
        to_csv(named, args.csv)
        print(f"wrote {args.csv} ({len(named)} workbook(s))", file=sys.stderr)

    total = len(paths)
    print(f"\ncoverage: {ok}/{total} full, {partial} partial, {failed} unreadable",
          file=sys.stderr)
    for cat, files in sorted(reasons.items()):
        print(f"  {cat:12} {len(files):3}  {', '.join(files[:4])}"
              + (" ..." if len(files) > 4 else ""), file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
