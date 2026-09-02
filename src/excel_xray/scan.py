"""Workbook X-ray: orchestrates structure, cells, regions and formulas.

Structure (merges, tables, connections, external links, pivot caches, VBA,
Power Query) comes from :mod:`excel_xray.ooxml` — openpyxl silently drops all of
it in read-only mode. Cell values, formulas and styles come from
:mod:`excel_xray.sheetmodel`, a single ``lxml`` pass over ``<sheetData>`` that
reads the formula and its cached value together. Nothing here uses openpyxl.
"""

from __future__ import annotations

import datetime as _dt
import hashlib
import json
import os
import zipfile
from dataclasses import asdict, dataclass, field

from .ooxml import WorkbookStructure, read_structure
from .regions import Region, detect_regions
from .sheetmodel import read_shared_strings, read_sheet_cells, read_styles
from .util import get_column_letter


@dataclass
class SheetXray:
    name: str
    position: int
    state: str
    dimension: str | None
    max_row: int
    max_col: int
    populated_cells: int
    density: float
    regions: list[Region] = field(default_factory=list)
    formula_profile: dict = field(default_factory=dict)
    merges: list[str] = field(default_factory=list)
    error_cells: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    # Compact occupancy for the report plate: [row, col, flag] where flag is
    # 1 value, 2 formula, 3 error, 4 text. Structure only -- no cell values, so
    # the report can be shared without carrying confidential figures.
    occupancy: list[list[int]] = field(default_factory=list)
    occupancy_truncated: bool = False


@dataclass
class WorkbookXray:
    path: str
    filename: str
    size_bytes: int
    sha256: str
    fs_modified: str
    parse_status: str
    core_props: dict = field(default_factory=dict)
    app_props: dict = field(default_factory=dict)
    has_vba: bool = False
    has_power_query: bool = False
    external_links: list[str] = field(default_factory=list)
    connections: list[dict] = field(default_factory=list)
    pivot_cache_sources: list[str] = field(default_factory=list)
    defined_names: list = field(default_factory=list)
    sheets: list[SheetXray] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def _sha256(path: str, chunk: int = 1 << 20) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for block in iter(lambda: fh.read(chunk), b""):
            h.update(block)
    return h.hexdigest()


OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
ZIP_MAGIC = b"PK\x03\x04"


class UnreadableWorkbook(Exception):
    """Raised with a human reason, so the coverage report is actionable."""

    def __init__(self, reason: str, category: str):
        super().__init__(reason)
        self.category = category


def triage(path: str) -> None:
    """Classify a file before parsing. Raises UnreadableWorkbook with a reason.

    'BadZipFile' tells an analyst nothing. 'Password-protected' tells them to go
    and ask for the password, which is the whole point of a coverage report.
    """
    size = os.path.getsize(path)
    if size == 0:
        raise UnreadableWorkbook("zero-byte file", "empty")
    with open(path, "rb") as fh:
        head = fh.read(8)

    if head.startswith(OLE_MAGIC):
        # OLE container: either a legacy .xls or an encrypted OOXML package.
        try:
            import olefile  # optional
            if olefile.isOleFile(path):
                ole = olefile.OleFileIO(path)
                streams = {"/".join(s) for s in ole.listdir()}
                ole.close()
                if "EncryptedPackage" in streams:
                    raise UnreadableWorkbook(
                        "password-protected (encrypted OOXML package)", "encrypted")
        except ImportError:
            pass
        raise UnreadableWorkbook(
            "legacy OLE format (.xls or encrypted) -- needs conversion first",
            "legacy_ole")

    if not head.startswith(ZIP_MAGIC):
        raise UnreadableWorkbook(
            f"not a zip container (magic {head[:4]!r}) -- extension may be wrong",
            "not_ooxml")

    try:
        with zipfile.ZipFile(path) as zf:
            if "xl/workbook.xml" not in zf.namelist():
                raise UnreadableWorkbook(
                    "zip without xl/workbook.xml -- not a spreadsheet package",
                    "not_workbook")
    except zipfile.BadZipFile as e:
        raise UnreadableWorkbook(f"corrupt or truncated zip: {e}", "corrupt") from e


def xray_workbook(path: str, max_rows: int = 200_000) -> WorkbookXray:
    triage(path)
    st = os.stat(path)
    wx = WorkbookXray(
        path=os.path.abspath(path),
        filename=os.path.basename(path),
        size_bytes=st.st_size,
        sha256=_sha256(path),
        fs_modified=_dt.datetime.fromtimestamp(st.st_mtime).isoformat(timespec="seconds"),
        parse_status="full",
    )

    struct: WorkbookStructure = read_structure(path)
    wx.core_props = struct.core_props
    wx.app_props = struct.app_props
    wx.has_vba = struct.has_vba
    wx.has_power_query = struct.has_power_query
    wx.external_links = struct.external_links
    wx.connections = struct.connections
    wx.pivot_cache_sources = struct.pivot_cache_sources
    wx.defined_names = [{"name": n, "refers_to": r} for n, r in struct.defined_names]
    wx.warnings.extend(struct.parse_notes)

    # A file last written by something other than Excel usually has no cached
    # formula values -- worth flagging, though we no longer need a second load.
    app = (struct.app_props.get("application") or "")
    if app and "Excel" not in app:
        wx.warnings.append(
            f"last written by {app!r}, not Excel -- cached formula values may be "
            "absent, so value-derived signals are unreliable"
        )

    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        shared = read_shared_strings(zf)
        styles = read_styles(zf)
        for sref in struct.sheets:
            if not sref.part or sref.part not in names:
                wx.warnings.append(f"sheet {sref.name!r} part missing")
                wx.parse_status = "partial"
                continue
            sstruct = struct.structures.get(sref.name)
            wx.sheets.append(
                _xray_sheet(zf, sref, sstruct, shared, styles, max_rows)
            )

    if all(s.populated_cells == 0 for s in wx.sheets) and wx.sheets:
        wx.parse_status = "partial"
        wx.warnings.append("no populated cells found in any sheet")
    return wx


def _xray_sheet(zf, sref, sstruct, shared, styles, max_rows: int) -> SheetXray:
    merges = sstruct.merges if sstruct else []
    tables = sstruct.tables if sstruct else []

    cells, profile, raw_errors, truncated, max_r, max_c = read_sheet_cells(
        zf, sref.part, shared, styles, max_rows=max_rows
    )

    notes: list[str] = []
    if truncated:
        notes.append(f"row scan capped at {max_rows}")
    errors = [f"{sref.name}!{e}" for e in raw_errors]

    regions = detect_regions(cells, merges, tables)

    # Plate render caps: beyond this a miniature is unreadable anyway.
    PLATE_ROWS, PLATE_COLS = 120, 60
    occ: list[list[int]] = []
    occ_trunc = False
    for (r, c), f in cells.items():
        if r > PLATE_ROWS or c > PLATE_COLS:
            occ_trunc = True
            continue
        flag = 3 if f.is_error else 2 if f.formula else 4 if f.is_text else 1
        occ.append([r, c, flag])
    occ.sort()

    area = max_r * max_c
    return SheetXray(
        name=sref.name,
        position=sref.position,
        state=sref.state,
        dimension=sstruct.dimension if sstruct else None,
        max_row=max_r,
        max_col=max_c,
        populated_cells=len(cells),
        density=round(len(cells) / area, 4) if area else 0.0,
        regions=regions,
        merges=merges,
        error_cells=errors,
        notes=notes,
        occupancy=occ,
        occupancy_truncated=occ_trunc,
        formula_profile={
            "total": profile.total,
            "distinct_skeletons": profile.distinct_skeletons,
            "compression": round(profile.compression, 1),
            "volatile_count": profile.volatile_count,
            "cross_sheet_count": profile.cross_sheet_count,
            "external_count": profile.external_count,
            "hardcoded_literal_count": profile.hardcoded_literal_count,
            "referenced_sheets": dict(profile.referenced_sheets),
            "top_functions": profile.functions.most_common(12),
            "top_skeletons": profile.top(15),
        },
    )


def to_json(wx: WorkbookXray) -> str:
    def enc(o):
        if isinstance(o, (_dt.datetime, _dt.date)):
            return o.isoformat()
        if isinstance(o, set):
            return sorted(o)
        return str(o)
    d = asdict(wx)
    for s in d["sheets"]:
        for r in s["regions"]:
            r["ref"] = f"{r['top']},{r['left']}-{r['bottom']},{r['right']}"
    return json.dumps(d, indent=2, default=enc)


if __name__ == "__main__":
    import sys
    print(to_json(xray_workbook(sys.argv[1])))
