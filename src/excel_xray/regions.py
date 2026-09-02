"""Region detection.

ListObjects cover maybe 10-20% of real financial sheets. The rest are hand-built:
merged multi-row headers, blank spacer rows inside one logical table, two blocks
side by side, totals rows that look like data, notes in column A.

Approach: run-length encode occupied cells per row, then connected-component
label the runs. Deliberately asymmetric gap tolerance --

    row_gap = 1   a single blank row inside a table is common
    col_gap = 0   a blank column between blocks is the standard separator

Getting that asymmetry wrong is what makes naive detectors merge the whole sheet
into one region.

Every region carries a detect_confidence. Do not hide it: a detector that says
"0.42, please check" reads as engineering; one that presents everything at equal
certainty invites the reader to hunt for the error.
"""
from __future__ import annotations

import re
import statistics
from dataclasses import dataclass, field

from .util import get_column_letter, range_boundaries

TOTALS_RE = re.compile(
    r"^\s*(grand\s+)?(total|totals|subtotal|sub-total|sum|net|balance)\b",
    re.I,
)


@dataclass
class CellFact:
    """One populated cell, flattened to the signals detection needs."""
    value: object = None
    formula: str | None = None
    is_text: bool = False
    is_number: bool = False
    is_date: bool = False
    is_error: bool = False
    bold: bool = False
    filled: bool = False
    bordered: bool = False
    text: str = ""


@dataclass
class Region:
    kind: str                      # data_table | matrix | key_value | calculation
    #                                | notes | unknown
    origin: str                    # listobject | detected
    top: int
    left: int
    bottom: int
    right: int
    header_rows: list[int] = field(default_factory=list)
    headers: list[str] = field(default_factory=list)
    n_data_rows: int = 0
    populated_cells: int = 0
    density: float = 0.0
    formula_cells: int = 0
    totals_row: int | None = None
    detect_confidence: float = 0.0
    evidence: list[str] = field(default_factory=list)
    table_name: str | None = None
    columns: list[dict] = field(default_factory=list)

    @property
    def ref(self) -> str:
        return (f"{get_column_letter(self.left)}{self.top}:"
                f"{get_column_letter(self.right)}{self.bottom}")

    @property
    def n_cols(self) -> int:
        return self.right - self.left + 1


# ------------------------------------------------------------------ components

def _row_runs(cols: list[int], col_gap: int) -> list[tuple[int, int]]:
    """Merge sorted column indices into runs, bridging gaps <= col_gap."""
    runs: list[tuple[int, int]] = []
    start = prev = cols[0]
    for c in cols[1:]:
        if c - prev - 1 > col_gap:
            runs.append((start, prev))
            start = c
        prev = c
    runs.append((start, prev))
    return runs


def find_components(
    occupied: dict[int, list[int]], row_gap: int = 1, col_gap: int = 0
) -> list[tuple[int, int, int, int, int]]:
    """Connected-component label over run-length encoded occupancy.

    occupied: row -> sorted list of populated column indices.
    Returns (top, left, bottom, right, cell_count) per component.

    Run-based rather than per-cell so a 100k-row sheet stays linear in rows
    rather than quadratic in cells.
    """
    parent: dict[int, int] = {}

    def find(x: int) -> int:
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a: int, b: int) -> None:
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[rb] = ra

    # (row, c1, c2, count, id)
    runs: list[tuple[int, int, int, int, int]] = []
    for r in sorted(occupied):
        cols = sorted(occupied[r])
        for c1, c2 in _row_runs(cols, col_gap):
            n = sum(1 for c in cols if c1 <= c <= c2)
            rid = len(runs)
            parent[rid] = rid
            runs.append((r, c1, c2, n, rid))

    # Union runs on nearby rows whose column spans touch within col_gap.
    by_row: dict[int, list[int]] = {}
    for idx, (r, *_rest) in enumerate(runs):
        by_row.setdefault(r, []).append(idx)

    rows_sorted = sorted(by_row)
    for i, r in enumerate(rows_sorted):
        for r2 in rows_sorted[i + 1:]:
            if r2 - r - 1 > row_gap:
                break
            for a in by_row[r]:
                _, a1, a2, _, _ = runs[a]
                for b in by_row[r2]:
                    _, b1, b2, _, _ = runs[b]
                    if a1 - col_gap <= b2 + col_gap and b1 - col_gap <= a2 + col_gap:
                        union(a, b)

    agg: dict[int, list[int]] = {}
    for r, c1, c2, n, rid in runs:
        root = find(rid)
        box = agg.get(root)
        if box is None:
            agg[root] = [r, c1, r, c2, n]
        else:
            box[0] = min(box[0], r)
            box[1] = min(box[1], c1)
            box[2] = max(box[2], r)
            box[3] = max(box[3], c2)
            box[4] += n
    return [tuple(v) for v in agg.values()]


# ------------------------------------------------------------------- headers

def _row_profile(
    cells: dict[tuple[int, int], CellFact], row: int, left: int, right: int
) -> dict:
    present = [cells[(row, c)] for c in range(left, right + 1)
               if (row, c) in cells]
    if not present:
        return {"n": 0, "text_ratio": 0.0, "styled": 0.0, "distinct": 0.0,
                "density": 0.0, "formula_ratio": 0.0}
    n = len(present)
    texts = [f.text for f in present if f.text]
    return {
        "n": n,
        "text_ratio": sum(f.is_text for f in present) / n,
        "styled": sum(f.bold or f.filled or f.bordered for f in present) / n,
        "distinct": (len(set(texts)) / len(texts)) if texts else 0.0,
        "density": n / max(1, right - left + 1),
        "formula_ratio": sum(bool(f.formula) for f in present) / n,
    }


def infer_header(
    cells: dict[tuple[int, int], CellFact],
    top: int, left: int, bottom: int, right: int,
    merges: list[tuple[int, int, int, int]],
) -> tuple[list[int], list[str], float, list[str]]:
    """Choose the header row(s) and flatten merged parents into 'parent::child'.

    Scores candidate rows on four signals rather than assuming row 1: text ratio
    versus the body below, style delta, value distinctness, and density.
    """
    evidence: list[str] = []
    height = bottom - top + 1
    max_candidate = min(4, height)

    body_rows = list(range(top + max_candidate, min(bottom, top + 12) + 1))
    body = [_row_profile(cells, r, left, right) for r in body_rows]
    body_text = statistics.mean([b["text_ratio"] for b in body]) if body else 0.0
    body_style = statistics.mean([b["styled"] for b in body]) if body else 0.0

    # A lone text cell on the top row of a wide block is a sheet title, not a
    # header. Absorbing it produces headers like ['Reserve Roll-Forward', '', ''].
    width = right - left + 1
    title: str | None = None
    scan_from = top
    if width >= 2:
        for r in range(top, top + max_candidate):
            p = _row_profile(cells, r, left, right)
            if p["n"] == 0:
                continue
            if p["n"] == 1 and width >= 2 and (r, left) in cells \
                    and cells[(r, left)].is_text:
                title = cells[(r, left)].text
                scan_from = r + 1
                evidence.append(f"row {r} treated as title, not header: {title!r}")
            break

    max_candidate = min(4, bottom - scan_from + 1)
    if max_candidate <= 0:
        return [], [], 0.3, evidence + ["no rows left after title"]

    best_row, best_score = None, 0.0
    for r in range(scan_from, scan_from + max_candidate):
        p = _row_profile(cells, r, left, right)
        if p["n"] == 0:
            continue
        score = (
            0.40 * max(0.0, p["text_ratio"] - body_text)
            + 0.25 * max(0.0, p["styled"] - body_style)
            + 0.20 * p["distinct"]
            + 0.15 * p["density"]
        )
        # A row full of formulas is a calculation row, not a header.
        score *= (1.0 - 0.8 * p["formula_ratio"])
        if score > best_score:
            best_row, best_score = r, score

    if best_row is None or best_score < 0.15:
        return [], [], min(0.45, 0.2 + best_score), evidence + [
            f"no header row scored above threshold (best {best_score:.2f})"]

    # Multi-row header. If the winning row is itself a merged parent band, the
    # real column labels are the row below it -- descend, then flatten.
    spanning = [m for m in merges
                if m[0] == best_row and m[2] == best_row and m[3] > m[1]]
    parent_row = None
    child_row = best_row
    if spanning and best_row + 1 <= bottom:
        below = _row_profile(cells, best_row + 1, left, right)
        if below["n"] >= _row_profile(cells, best_row, left, right)["n"]:
            parent_row, child_row = best_row, best_row + 1

    headers = [
        (cells[(child_row, c)].text if (child_row, c) in cells else "")
        for c in range(left, right + 1)
    ]
    header_rows = [child_row]
    evidence.append(
        f"header row {child_row} scored {best_score:.2f} "
        f"(text {_row_profile(cells,child_row,left,right)['text_ratio']:.0%} vs "
        f"body {body_text:.0%})"
    )

    if parent_row is not None:
        parents: dict[int, str] = {}
        for r1, c1, r2, c2 in spanning:
            label = cells[(r1, c1)].text if (r1, c1) in cells else ""
            for c in range(c1, c2 + 1):
                parents[c] = label
        headers = [
            (f"{parents[c]}::{headers[c-left]}"
             if c in parents and parents[c] and headers[c - left]
             else headers[c - left])
            for c in range(left, right + 1)
        ]
        header_rows.insert(0, parent_row)
        evidence.append(
            f"multi-row header: {len(spanning)} merged parent band(s) on row "
            f"{parent_row} flattened to parent::child"
        )

    # Calibrate. Saturating at 1.0 makes the number decoration; spread it, then
    # penalise the things that genuinely make a header doubtful.
    conf = 0.35 + 0.55 * min(1.0, best_score)
    blanks = sum(1 for h in headers if not h)
    if blanks:
        conf *= (1 - 0.5 * blanks / len(headers))
        evidence.append(f"{blanks}/{len(headers)} header cells blank")
    if title:
        conf = min(1.0, conf + 0.05)
    return header_rows, headers, round(conf, 2), evidence


# ------------------------------------------------------------- classification

def _profile_columns(
    cells: dict[tuple[int, int], CellFact],
    top: int, left: int, bottom: int, right: int,
    headers: list[str],
) -> list[dict]:
    out = []
    n_rows = max(1, bottom - top + 1)
    for c in range(left, right + 1):
        col = [cells[(r, c)] for r in range(top, bottom + 1) if (r, c) in cells]
        n = len(col)
        if n == 0:
            kind = "empty"
        else:
            num = sum(f.is_number for f in col)
            txt = sum(f.is_text for f in col)
            dt = sum(f.is_date for f in col)
            kind = ("date" if dt > n * 0.6
                    else "number" if num > n * 0.6
                    else "text" if txt > n * 0.6
                    else "mixed")
        vals = [f.text for f in col if f.text]
        out.append({
            "ordinal": c - left,
            "column": get_column_letter(c),
            "header": headers[c - left] if c - left < len(headers) else "",
            "inferred_type": kind,
            "null_rate": round(1 - n / n_rows, 3),
            "distinct_ratio": round(len(set(vals)) / len(vals), 3) if vals else 0.0,
            "formula_cells": sum(bool(f.formula) for f in col),
            "error_cells": sum(f.is_error for f in col),
        })
    return out


def classify_region(
    r: Region, cells: dict[tuple[int, int], CellFact]
) -> tuple[str, list[str]]:
    ev: list[str] = []
    n_cols = r.n_cols
    height = r.bottom - r.top + 1
    formula_ratio = r.formula_cells / max(1, r.populated_cells)

    if r.origin == "listobject":
        ev.append("declared ListObject")
        return "data_table", ev

    # Long free text in a single column is a notes block, not a table.
    texts = [cells[(rr, cc)].text for rr in range(r.top, r.bottom + 1)
             for cc in range(r.left, r.right + 1) if (rr, cc) in cells]
    if texts:
        avg_len = statistics.mean(len(t) for t in texts)
        if n_cols <= 2 and avg_len > 40:
            ev.append(f"single narrow column, mean text length {avg_len:.0f}")
            return "notes", ev

    # Label/value pairs: narrow, text on the left, no repeating structure.
    if n_cols <= 3 and height <= 25 and not r.headers:
        left_col = [cells[(rr, r.left)] for rr in range(r.top, r.bottom + 1)
                    if (rr, r.left) in cells]
        if left_col and sum(f.is_text for f in left_col) / len(left_col) > 0.7:
            ev.append("narrow block, left column all text, no header row detected")
            return "key_value", ev

    if formula_ratio > 0.55:
        ev.append(f"{formula_ratio:.0%} of populated cells are formulas")
        return "calculation", ev

    if len(r.header_rows) > 1:
        ev.append("multi-row header with merged parents")
        return "matrix", ev

    if r.headers and r.n_data_rows >= 2:
        ev.append(f"header row plus {r.n_data_rows} data rows")
        return "data_table", ev

    return "unknown", ev


# ----------------------------------------------------------------- entrypoint

def detect_regions(
    cells: dict[tuple[int, int], CellFact],
    merges: list[str],
    tables: list,
    row_gap: int = 1,
    col_gap: int = 0,
) -> list[Region]:
    """Detect every region on one sheet. ListObjects first, then the rest."""
    if not cells:
        return []

    merge_boxes: list[tuple[int, int, int, int]] = []
    for m in merges:
        try:
            c1, r1, c2, r2 = range_boundaries(m)
            merge_boxes.append((r1, c1, r2, c2))
        except Exception:
            continue

    regions: list[Region] = []
    claimed: set[tuple[int, int]] = set()

    # 1. Declared tables are exact. Take them and remove from the search space.
    for t in tables:
        try:
            c1, r1, c2, r2 = range_boundaries(t.ref)
        except Exception:
            continue
        reg = Region(kind="data_table", origin="listobject",
                     top=r1, left=c1, bottom=r2, right=c2,
                     header_rows=[r1] if t.header_row_count else [],
                     headers=list(t.columns),
                     table_name=t.display_name or t.name,
                     detect_confidence=1.0,
                     evidence=[f"declared ListObject {t.display_name!r} at {t.ref}"])
        _finalise(reg, cells)
        regions.append(reg)
        for rr in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                claimed.add((rr, cc))

    # 2. Connected components over whatever is left.
    occupied: dict[int, list[int]] = {}
    for (rr, cc) in cells:
        if (rr, cc) in claimed:
            continue
        occupied.setdefault(rr, []).append(cc)

    for top, left, bottom, right, count in find_components(
        occupied, row_gap=row_gap, col_gap=col_gap
    ):
        reg = Region(kind="unknown", origin="detected",
                     top=top, left=left, bottom=bottom, right=right)

        shape, shape_ev = _pre_shape(cells, top, left, bottom, right)
        if shape:
            # Titles and label/value blocks have no column headers. Running
            # header inference on them just promotes the first label to a header.
            reg.evidence.extend(shape_ev)
            _finalise(reg, cells)
            reg.kind = shape
            reg.detect_confidence = 0.85 if shape == "title" else 0.80
            regions.append(reg)
            continue

        hdr_rows, headers, conf, ev = infer_header(
            cells, top, left, bottom, right, merge_boxes)
        reg.header_rows, reg.headers = hdr_rows, headers
        reg.detect_confidence = conf
        reg.evidence.extend(ev)
        _finalise(reg, cells)
        reg.kind, kev = classify_region(reg, cells)
        reg.evidence.extend(kev)
        regions.append(reg)

    regions.sort(key=lambda r: (r.top, r.left))
    return regions


def _pre_shape(
    cells: dict[tuple[int, int], CellFact],
    top: int, left: int, bottom: int, right: int,
) -> tuple[str | None, list[str]]:
    """Recognise shapes that have no column headers, before header inference.

    A key/value assumptions block and a standalone title both get mangled by
    header scoring, which promotes the first label to a column header.
    """
    width = right - left + 1
    height = bottom - top + 1
    present = [(r, c) for r in range(top, bottom + 1)
               for c in range(left, right + 1) if (r, c) in cells]

    if len(present) == 1:
        f = cells[present[0]]
        if f.is_text and not f.formula:
            if len(f.text) > 50:
                return "notes", [f"single long text cell ({len(f.text)} chars)"]
            return "title", [f"single populated text cell: {f.text[:60]!r}"]
        return None, []

    if width == 2 and height <= 30:
        labels = [cells[(r, left)] for r in range(top, bottom + 1)
                  if (r, left) in cells]
        values = [cells[(r, left + 1)] for r in range(top, bottom + 1)
                  if (r, left + 1) in cells]
        if not labels or not values:
            return None, []
        label_text = sum(f.is_text for f in labels) / len(labels)
        label_vals = [f.text for f in labels if f.text]
        distinct = len(set(label_vals)) / len(label_vals) if label_vals else 0
        paired = len(values) / len(labels)
        if label_text > 0.8 and distinct > 0.9 and 0.6 <= paired <= 1.0:
            return "key_value", [
                f"two columns, left {label_text:.0%} text and {distinct:.0%} "
                f"distinct, right column paired -- label/value block"]
    return None, []


def _finalise(reg: Region, cells: dict[tuple[int, int], CellFact]) -> None:
    """Fill counts, totals row, column profiles, and confidence adjustments."""
    pop = 0
    formulas = 0
    for rr in range(reg.top, reg.bottom + 1):
        for cc in range(reg.left, reg.right + 1):
            f = cells.get((rr, cc))
            if f is None:
                continue
            pop += 1
            formulas += bool(f.formula)
    reg.populated_cells = pop
    reg.formula_cells = formulas
    area = (reg.bottom - reg.top + 1) * (reg.right - reg.left + 1)
    reg.density = round(pop / area, 3) if area else 0.0

    first_data = (max(reg.header_rows) + 1) if reg.header_rows else reg.top
    reg.n_data_rows = max(0, reg.bottom - first_data + 1)

    # Totals row: label match, or a bold last row of SUM formulas.
    last = reg.bottom
    label = cells.get((last, reg.left))
    row_cells = [cells[(last, c)] for c in range(reg.left, reg.right + 1)
                 if (last, c) in cells]
    sumish = sum(1 for f in row_cells
                 if f.formula and "SUM" in f.formula.upper())
    if (label and TOTALS_RE.match(label.text or "")) or (
        row_cells and sumish >= max(1, len(row_cells) // 2)
    ):
        reg.totals_row = last
        reg.n_data_rows = max(0, reg.n_data_rows - 1)
        reg.evidence.append(f"row {last} identified as totals row")

    if reg.origin == "detected":
        # Ragged blocks are less likely to be a single clean table.
        if reg.density < 0.5:
            reg.detect_confidence = round(reg.detect_confidence * 0.75, 2)
            reg.evidence.append(f"low density {reg.density:.0%} reduced confidence")
        reg.detect_confidence = round(min(1.0, reg.detect_confidence), 2)

    reg.columns = _profile_columns(
        cells, first_data, reg.left, reg.bottom, reg.right, reg.headers)

    # A header is corroborated by what sits under it. Columns that are densely
    # populated and each hold one consistent type are strong evidence the header
    # row was read correctly -- independent of how the header itself was styled,
    # which matters because machine-generated extracts rarely bold their headers.
    if reg.origin == "detected" and reg.headers and reg.n_data_rows >= 3:
        typed = [c for c in reg.columns if c["inferred_type"] != "empty"]
        if typed:
            clean = sum(1 for c in typed
                        if c["inferred_type"] != "mixed" and c["null_rate"] < 0.1)
            share = clean / len(typed)
            if share >= 0.8:
                reg.detect_confidence = round(
                    min(1.0, reg.detect_confidence + 0.15 * share), 2)
                reg.evidence.append(
                    f"{clean}/{len(typed)} columns single-typed and near-complete "
                    f"-- corroborates the header row")
