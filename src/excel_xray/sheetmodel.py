"""Single-pass cell reader, straight from the OOXML — no openpyxl.

The two-openpyxl-pass dance (``data_only=False`` for formulas, then
``data_only=True`` for cached values) exists only because openpyxl exposes one
or the other per load. The file itself keeps them together::

    <c r="C5" t="n"><f>B5*Assumptions!$B$6</f><v>59937.255</v></c>

The formula and its last cached result sit in the same cell element. Streaming
``<sheetData>`` with ``lxml.iterparse`` therefore yields both in one pass — and
also the style index, from which we resolve the bold / fill / border / date
signals the region detector needs — while never materialising a 100k-row sheet.

Produces the exact ``CellFact`` objects :mod:`excel_xray.regions` consumes and a
populated :class:`~excel_xray.formulas.FormulaProfile`, so the rest of the
pipeline is unchanged.
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass, field

from lxml import etree

from .formulas import ERROR_VALUES, FormulaProfile, analyse
from .regions import CellFact
from .util import column_index_from_string, get_column_letter

_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_STRICT = "http://purl.oclc.org/ooxml/spreadsheetml/main"

# Builtin number-format ids that Excel renders as a date or time.
_DATE_BUILTINS = set(range(14, 23)) | set(range(45, 48)) | {27, 30, 36, 50, 57}


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


# ----------------------------------------------------------- workbook globals


def read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    """The ``sharedStrings.xml`` table. ``t="s"`` cells index into this."""
    try:
        data = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    out: list[str] = []
    root = etree.fromstring(data)
    for si in root:
        if _local(si.tag) != "si":
            continue
        # <si> is either a single <t> or a run of <r><t>… pieces.
        out.append("".join(t.text or "" for t in si.iter() if _local(t.tag) == "t"))
    return out


@dataclass
class StyleTable:
    """Just the per-style signals the detector cares about, keyed by ``s`` index."""

    bold: list[bool] = field(default_factory=list)
    filled: list[bool] = field(default_factory=list)
    bordered: list[bool] = field(default_factory=list)
    is_date: list[bool] = field(default_factory=list)

    def _get(self, seq: list[bool], idx: int) -> bool:
        return seq[idx] if 0 <= idx < len(seq) else False


def read_styles(zf: zipfile.ZipFile) -> StyleTable:
    """Resolve ``cellXfs`` into flat bold/fill/border/date lookups by style index."""
    st = StyleTable()
    try:
        root = etree.fromstring(zf.read("xl/styles.xml"))
    except KeyError:
        return st

    fonts_bold: list[bool] = []
    fills_solid: list[bool] = []
    borders_edged: list[bool] = []
    custom_date: dict[int, bool] = {}

    for section in root:
        tag = _local(section.tag)
        if tag == "fonts":
            for font in section:
                fonts_bold.append(
                    any(_local(c.tag) == "b" for c in font)
                )
        elif tag == "fills":
            for fill in section:
                pat = next((c for c in fill if _local(c.tag) == "patternFill"), None)
                ptype = pat.get("patternType") if pat is not None else None
                fills_solid.append(bool(ptype and ptype != "none"))
        elif tag == "borders":
            for border in section:
                edged = False
                for edge in border:
                    if _local(edge.tag) in ("top", "bottom") and edge.get("style"):
                        edged = True
                borders_edged.append(edged)
        elif tag == "numFmts":
            for nf in section:
                fid = int(nf.get("numFmtId", "0"))
                code = (nf.get("formatCode") or "").lower()
                custom_date[fid] = _looks_like_date(code)

    for section in root:
        if _local(section.tag) != "cellXfs":
            continue
        for xf in section:
            fid = int(xf.get("fontId", "0"))
            fill_id = int(xf.get("fillId", "0"))
            bid = int(xf.get("borderId", "0"))
            num_id = int(xf.get("numFmtId", "0"))
            st.bold.append(fonts_bold[fid] if fid < len(fonts_bold) else False)
            st.filled.append(fills_solid[fill_id] if fill_id < len(fills_solid) else False)
            st.bordered.append(borders_edged[bid] if bid < len(borders_edged) else False)
            st.is_date.append(num_id in _DATE_BUILTINS or custom_date.get(num_id, False))
    return st


def _looks_like_date(code: str) -> bool:
    """True if a custom number-format code renders a date/time.

    Strips quoted literals and bracket sections first, so a currency format like
    ``[$-409]#,##0`` is not mistaken for a date because of its 'd'-free brackets.
    """
    stripped = []
    quote = False
    bracket = 0
    for ch in code:
        if ch == '"':
            quote = not quote
            continue
        if ch == "[":
            bracket += 1
            continue
        if ch == "]":
            bracket = max(0, bracket - 1)
            continue
        if quote or bracket:
            continue
        stripped.append(ch)
    body = "".join(stripped)
    return any(tok in body for tok in ("y", "m", "d", "h", "s"))


# ----------------------------------------------------------------- cell stream


def _coord_col(ref: str) -> int:
    """Column index from a cell ref like ``'C5'`` -> 3."""
    letters = "".join(ch for ch in ref if ch.isalpha())
    return column_index_from_string(letters) if letters else 0


def _num(text: str):
    """Parse a numeric ``<v>`` to int when integral, else float."""
    try:
        f = float(text)
    except (TypeError, ValueError):
        return text
    return int(f) if f.is_integer() else f


def _make_fact(
    ref: str,
    t: str | None,
    style_idx: int,
    formula_text: str | None,
    v_text: str | None,
    shared: list[str],
    inline_text: str | None,
    styles: StyleTable,
) -> CellFact | None:
    """Build one :class:`CellFact` from the parsed pieces of a ``<c>`` element."""
    formula = ("=" + formula_text) if formula_text is not None else None

    # Resolve the shown value by cell type.
    shown: object
    if t == "s":
        idx = int(v_text) if v_text not in (None, "") else -1
        shown = shared[idx] if 0 <= idx < len(shared) else ""
    elif t == "inlineStr":
        shown = inline_text or ""
    elif t == "str":
        shown = v_text
    elif t == "b":
        shown = bool(int(v_text)) if v_text not in (None, "") else False
    elif t == "e":
        shown = v_text
    elif t == "d":
        shown = v_text  # ISO 8601 date string
    else:  # "n" or unspecified
        shown = _num(v_text) if v_text not in (None, "") else None

    if shown is None and formula is None:
        return None

    fact = CellFact(value=shown, formula=formula)
    is_dated = styles._get(styles.is_date, style_idx)

    # Flag semantics mirror the original openpyxl-fed path exactly: a string-
    # valued cell is is_text, and is_error is an *additional* flag on top (an
    # error like "#DIV/0!" is both). The region detector's hand-tuned thresholds
    # depend on that overlap -- separating them shifts header/body text ratios.
    date_by_format = (
        is_dated and isinstance(shown, (int, float)) and not isinstance(shown, bool)
    )
    if t == "d" or date_by_format:
        fact.is_date = True
        fact.text = ""
    elif isinstance(shown, str):
        fact.is_text = not shown.startswith("=")
        fact.is_error = shown in ERROR_VALUES
        fact.text = shown
    elif isinstance(shown, bool):
        fact.text = ""
    elif isinstance(shown, (int, float)):
        fact.is_number = True
        fact.text = ""
    else:  # formula with no cached value
        fact.text = ""

    fact.bold = styles._get(styles.bold, style_idx)
    fact.filled = styles._get(styles.filled, style_idx)
    fact.bordered = styles._get(styles.bordered, style_idx)
    return fact


def read_sheet_cells(
    zf: zipfile.ZipFile,
    part: str,
    shared: list[str],
    styles: StyleTable,
    max_rows: int = 200_000,
) -> tuple[dict[tuple[int, int], CellFact], FormulaProfile, list[str], bool, int, int]:
    """Stream one worksheet's ``<sheetData>`` into cells + a formula profile.

    Returns ``(cells, profile, error_cells, truncated, max_row, max_col)``.
    """
    cells: dict[tuple[int, int], CellFact] = {}
    profile = FormulaProfile()
    errors: list[str] = []
    truncated = False
    max_r = max_c = 0

    row_idx = 0
    col_cursor = 0  # for cells that omit their r="" attribute
    with zf.open(part) as fh:
        for _event, el in etree.iterparse(fh, events=("end",)):
            tag = _local(el.tag)
            if tag == "row":
                el.clear()
                continue
            if tag != "c":
                continue

            ref = el.get("r")
            if ref:
                r = int("".join(ch for ch in ref if ch.isdigit()))
                c = _coord_col(ref)
                row_idx, col_cursor = r, c
            else:
                r, col_cursor = row_idx or 1, col_cursor + 1
                c = col_cursor
            if r > max_rows:
                truncated = True
                el.clear()
                continue

            t = el.get("t")
            style_idx = int(el.get("s", "0") or 0)
            f_text = None
            v_text = None
            inline_text = None
            for child in el:
                ct = _local(child.tag)
                if ct == "f":
                    f_text = child.text or ""
                elif ct == "v":
                    v_text = child.text
                elif ct == "is":
                    inline_text = "".join(
                        t.text or "" for t in child.iter() if _local(t.tag) == "t"
                    )

            fact = _make_fact(
                ref or f"{get_column_letter(c)}{r}", t, style_idx,
                f_text, v_text, shared, inline_text, styles,
            )
            el.clear()
            if fact is None:
                continue

            cells[(r, c)] = fact
            max_r, max_c = max(max_r, r), max(max_c, c)
            if fact.formula:
                profile.add(analyse(fact.formula, r, c))
            if fact.is_error:
                errors.append(f"{get_column_letter(c)}{r} {fact.text}")

    return cells, profile, errors, truncated, max_r, max_c
