"""Formula normalisation.

Collapses a column of near-identical formulas to a single skeleton, so a
40,000-formula model reduces to a countable set of distinct calculations.

Pipeline per formula:
  1. mask string literals so refs inside them are not rewritten
  2. rewrite every A1 reference to R1C1 relative to the owning cell
  3. abstract numeric and string literals
  4. tag the reference scope (local / cross-sheet / external)

`=B5*$C$2` at C5 and `=B6*$C$2` at C6 both become  RC[-1]*R2C3 .
"""
from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field

# Functions whose presence forces recalculation on every edit. High counts of
# these are a performance smell and a strong "hand-built model" signal.
VOLATILE = {
    "NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT",
    "CELL", "INFO", "RANDARRAY",
}

# Cached error values Excel bakes into the file.
ERROR_VALUES = {
    "#REF!", "#N/A", "#DIV/0!", "#VALUE!", "#NAME?", "#NULL!", "#NUM!",
    "#SPILL!", "#CALC!",
}

_STRING = re.compile(r'"(?:[^"]|"")*"')
_FUNC = re.compile(r"\b([A-Z][A-Z0-9_.]*)\s*\(")

# sheet prefix (optional) + A1 reference. Handles 'quoted sheet'!, [1]External!,
# and bare Sheet1!. Column letters capped at 3 (XFD is the Excel maximum).
_REF = re.compile(
    r"(?P<ext>\[\d+\])?"
    r"(?P<sheet>'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)?"
    r"(?P<bang>!)?"
    r"(?P<cabs>\$)?(?P<col>[A-Za-z]{1,3})(?P<rabs>\$)?(?P<row>\d{1,7})"
    r"(?![A-Za-z0-9_(])"
)

_NUMBER = re.compile(r"(?<![A-Za-z0-9_.$])\d+(?:\.\d+)?(?:[eE][+-]?\d+)?")


def col_to_num(col: str) -> int:
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


@dataclass
class FormulaFacts:
    """What a single formula tells us, before aggregation."""
    skeleton: str
    functions: list[str] = field(default_factory=list)
    is_volatile: bool = False
    local_refs: int = 0
    cross_sheet_refs: int = 0
    external_refs: int = 0
    referenced_sheets: set[str] = field(default_factory=set)
    # (row, col) cells this formula reads, local sheet only -- feeds the
    # dependency graph in stage 4.
    precedent_cells: list[tuple[int, int]] = field(default_factory=list)
    has_literal_number: bool = False


# Sentinels deliberately contain no digits: the numeric-literal pass runs after
# both substitutions and would otherwise rewrite an index into "N".
_STR_SENTINEL = "\x00"
_REF_SENTINEL = "\x01"


def _mask_strings(f: str) -> tuple[str, int]:
    """Replace string literals with a digit-free sentinel.

    Every literal collapses to "S" in the final skeleton, so we only need the
    count, not the text.
    """
    n = 0

    def repl(m: re.Match) -> str:
        nonlocal n
        n += 1
        return _STR_SENTINEL

    return _STRING.sub(repl, f), n


def _expand_range_cells(
    r1: int, c1: int, r2: int, c2: int, cap: int = 4096
) -> list[tuple[int, int]]:
    """Enumerate a rectangular range, capped so SUM(A:A) cannot explode."""
    if (r2 - r1 + 1) * (c2 - c1 + 1) > cap:
        return [(r1, c1), (r2, c2)]  # endpoints only
    return [(r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]


def analyse(formula: str, row: int, col: int) -> FormulaFacts:
    """Normalise one formula owned by cell (row, col), both 1-based."""
    if not formula:
        return FormulaFacts(skeleton="")
    f = formula.lstrip("=").strip()
    masked, _nstr = _mask_strings(f)

    facts = FormulaFacts(skeleton="")
    facts.functions = sorted(set(_FUNC.findall(masked.upper())))
    facts.is_volatile = any(fn in VOLATILE for fn in facts.functions)

    # Track pending range starts so B5:B8 records as one span, not two cells.
    pending: dict[int, tuple[int, int]] = {}
    rewritten: list[str] = []

    def ref_repl(m: re.Match) -> str:
        sheet = m.group("sheet")
        ext = m.group("ext")
        bang = m.group("bang")
        # A bare "sheet" with no "!" is not a sheet -- it is part of something
        # else the regex over-reached on. Guard it.
        if sheet and not bang:
            sheet = None
        c = col_to_num(m.group("col"))
        r = int(m.group("row"))
        cabs, rabs = bool(m.group("cabs")), bool(m.group("rabs"))

        if ext:
            facts.external_refs += 1
            scope = "EXT!"
        elif sheet:
            facts.cross_sheet_refs += 1
            facts.referenced_sheets.add(sheet.strip("'").replace("''", "'"))
            scope = "SHEET!"
        else:
            facts.local_refs += 1
            scope = ""
            pending[m.start()] = (r, c)

        rpart = f"R{r}" if rabs else (f"R[{r-row}]" if r != row else "R")
        cpart = f"C{c}" if cabs else (f"C[{c-col}]" if c != col else "C")
        rewritten.append(f"{scope}{rpart}{cpart}")
        return _REF_SENTINEL

    skel = _REF.sub(ref_repl, masked)

    # Resolve local precedents, merging A:B into spans where a colon joins them.
    positions = sorted(pending)
    used: set[int] = set()
    for i, p in enumerate(positions):
        if p in used:
            continue
        if i + 1 < len(positions):
            q = positions[i + 1]
            between = masked[p:q]
            if ":" in between and between.count(":") == 1:
                r1, c1 = pending[p]
                r2, c2 = pending[q]
                facts.precedent_cells.extend(
                    _expand_range_cells(min(r1, r2), min(c1, c2),
                                        max(r1, r2), max(c1, c2))
                )
                used.add(p)
                used.add(q)
                continue
        facts.precedent_cells.append(pending[p])
        used.add(p)

    # Order matters: refs are sentinels at this point, so the only digits left
    # are genuine literals. A hardcoded number inside a formula is exactly what
    # we want to surface -- it is the classic buried-assumption smell.
    facts.has_literal_number = bool(_NUMBER.search(skel))
    skel = _NUMBER.sub("N", skel)

    out, it = [], iter(rewritten)
    for ch in skel:
        if ch == _REF_SENTINEL:
            out.append(next(it, "?"))
        elif ch == _STR_SENTINEL:
            out.append("S")
        else:
            out.append(ch)
    facts.skeleton = re.sub(r"\s+", "", "".join(out))
    return facts


@dataclass
class FormulaProfile:
    """Aggregate over all formulas on one sheet."""
    total: int = 0
    skeletons: Counter = field(default_factory=Counter)
    functions: Counter = field(default_factory=Counter)
    volatile_count: int = 0
    cross_sheet_count: int = 0
    external_count: int = 0
    referenced_sheets: Counter = field(default_factory=Counter)
    hardcoded_literal_count: int = 0

    @property
    def distinct_skeletons(self) -> int:
        return len(self.skeletons)

    @property
    def compression(self) -> float:
        """Formulas per distinct skeleton. High = disciplined fill-down."""
        return self.total / len(self.skeletons) if self.skeletons else 0.0

    def add(self, facts: FormulaFacts) -> None:
        self.total += 1
        self.skeletons[facts.skeleton] += 1
        for fn in facts.functions:
            self.functions[fn] += 1
        self.volatile_count += int(facts.is_volatile)
        self.cross_sheet_count += int(facts.cross_sheet_refs > 0)
        self.external_count += int(facts.external_refs > 0)
        for s in facts.referenced_sheets:
            self.referenced_sheets[s] += 1
        self.hardcoded_literal_count += int(facts.has_literal_number)

    def top(self, n: int = 15) -> list[tuple[str, int]]:
        return self.skeletons.most_common(n)
