"""Excel X-ray — structural scanner and explainer for financial workbooks.

Reads structure straight from the OOXML package (openpyxl silently drops merged
ranges, tables, connections, external links and pivot caches) and cells in a
single ``lxml`` pass (formula and cached value together). Detects hand-built
table regions, normalises formulas to counted skeletons, and explains the
result as a self-contained HTML report or JSON.
"""

from __future__ import annotations

from .assessment import Assessment, FileAssessment, TabAssessment, assess
from .ooxml import WorkbookStructure, read_structure
from .report import build_report, write_report
from .scan import (
    SheetXray,
    UnreadableWorkbook,
    WorkbookXray,
    to_json,
    triage,
    xray_workbook,
)

__all__ = [
    "xray_workbook",
    "to_json",
    "triage",
    "WorkbookXray",
    "SheetXray",
    "UnreadableWorkbook",
    "read_structure",
    "WorkbookStructure",
    "build_report",
    "write_report",
    "assess",
    "Assessment",
    "FileAssessment",
    "TabAssessment",
    "main",
]

__version__ = "0.1.0"


def main() -> None:
    """Console-script entry point."""
    from .cli import main as _main

    raise SystemExit(_main())
