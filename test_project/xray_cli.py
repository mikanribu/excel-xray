#!/usr/bin/env python3
"""Excel X-ray -- stage 1 scanner.

  python xray_cli.py FILE.xlsx                 report next to the file
  python xray_cli.py FOLDER -o out/            every workbook in a folder
  python xray_cli.py FILE.xlsx --json          machine-readable to stdout

Reads only. Never writes to, moves or renames a source file.
"""
from __future__ import annotations

import argparse
import os
import sys
import time
import traceback

from xray.report import write_report
from xray.scan import UnreadableWorkbook, to_json, xray_workbook

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
    ap = argparse.ArgumentParser(description="Excel X-ray structural scanner")
    ap.add_argument("target", help="workbook or folder")
    ap.add_argument("-o", "--out", default=None, help="output directory")
    ap.add_argument("--json", action="store_true", help="emit JSON to stdout")
    ap.add_argument("--max-rows", type=int, default=200_000)
    args = ap.parse_args()

    paths = collect(args.target)
    if not paths:
        print(f"no workbooks found under {args.target}", file=sys.stderr)
        return 2

    outdir = args.out or (args.target if os.path.isdir(args.target)
                          else os.path.dirname(os.path.abspath(args.target)))
    os.makedirs(outdir, exist_ok=True)

    ok = partial = failed = 0
    reasons: dict[str, list[str]] = {}
    for p in paths:
        t0 = time.time()
        try:
            wx = xray_workbook(p, max_rows=args.max_rows)
        except UnreadableWorkbook as e:
            failed += 1
            reasons.setdefault(e.category, []).append(os.path.basename(p))
            print(f"SKIPPED  {os.path.basename(p):46} {e}", file=sys.stderr)
            continue
        except Exception as e:
            failed += 1
            reasons.setdefault("unexpected", []).append(os.path.basename(p))
            print(f"FAILED   {os.path.basename(p):46} {type(e).__name__}: {e}",
                  file=sys.stderr)
            if os.environ.get("XRAY_DEBUG"):
                traceback.print_exc()
            continue

        if args.json:
            print(to_json(wx))
        else:
            name = os.path.splitext(os.path.basename(p))[0]
            dest = os.path.join(outdir, f"xray_{name}.html")
            write_report(wx, dest)
            regions = sum(len(s.regions) for s in wx.sheets)
            low = sum(1 for s in wx.sheets for r in s.regions
                      if r.detect_confidence < 0.70)
            print(f"{wx.parse_status:8} {os.path.basename(p):46} "
                  f"{len(wx.sheets):3} sheets  {regions:3} regions  "
                  f"{low:2} low-conf  {time.time()-t0:5.1f}s  -> {dest}")
        ok += wx.parse_status == "full"
        partial += wx.parse_status == "partial"

    total = len(paths)
    print(f"\ncoverage: {ok}/{total} full, {partial} partial, {failed} unreadable",
          file=sys.stderr)
    for cat, files in sorted(reasons.items()):
        print(f"  {cat:12} {len(files):3}  {', '.join(files[:4])}"
              + (" ..." if len(files) > 4 else ""), file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
