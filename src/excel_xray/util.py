"""Small A1-notation helpers.

These replace the three ``openpyxl.utils`` functions the rest of the package
used, so nothing here depends on openpyxl any more. Column indices are 1-based,
matching Excel and openpyxl's own convention.
"""

from __future__ import annotations

import re

_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d{1,7})$")


def get_column_letter(idx: int) -> str:
    """1 -> 'A', 26 -> 'Z', 27 -> 'AA'."""
    if idx < 1:
        raise ValueError(f"column index must be >= 1, got {idx}")
    letters = ""
    while idx:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def column_index_from_string(col: str) -> int:
    """'A' -> 1, 'AA' -> 27. Case-insensitive."""
    n = 0
    for ch in col.upper():
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"invalid column {col!r}")
        n = n * 26 + (ord(ch) - 64)
    return n


def range_boundaries(ref: str) -> tuple[int, int, int, int]:
    """Return ``(min_col, min_row, max_col, max_row)`` for an A1 range.

    Mirrors ``openpyxl.utils.range_boundaries``: columns first, 1-based, ``$``
    absolute markers ignored. A single cell like ``"A1"`` collapses to a 1x1
    range. Raises ``ValueError`` on anything it cannot parse (open-ended ranges
    such as ``A:A`` are not supported — callers here never pass them).
    """
    ref = ref.strip()
    if ":" in ref:
        start, end = ref.split(":", 1)
    else:
        start = end = ref
    ms, me = _CELL_RE.match(start), _CELL_RE.match(end)
    if not ms or not me:
        raise ValueError(f"cannot parse range {ref!r}")
    c1 = column_index_from_string(ms.group(1))
    r1 = int(ms.group(2))
    c2 = column_index_from_string(me.group(1))
    r2 = int(me.group(2))
    return (min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2))
