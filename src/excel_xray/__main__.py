"""``python -m excel_xray`` — same entry point as the ``excel-xray`` console script."""

from .cli import main as _main

if __name__ == "__main__":
    raise SystemExit(_main())
