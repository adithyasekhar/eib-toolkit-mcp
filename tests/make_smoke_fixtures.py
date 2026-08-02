"""Build synthetic files for the CI smoke run:  python tests/make_smoke_fixtures.py DIR

Writes template.xlsx, filled.xlsx (a clean load), spec.yaml, and the CSVs the
spec maps — everything the CLI end-to-end smoke needs. All data is synthetic
(fake names, fake IDs; no real PII or tenant data). Slice 5 replaces this
with bundled example fixtures.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from helpers import make_filled, make_spec_and_csvs, make_template


def main() -> int:
    if len(sys.argv) != 2:
        print(__doc__, file=sys.stderr)
        return 2
    directory = Path(sys.argv[1])
    directory.mkdir(parents=True, exist_ok=True)
    make_template(directory / "template.xlsx")
    make_filled(directory / "filled.xlsx")
    make_spec_and_csvs(directory)
    print(f"Smoke fixtures written to {directory}/")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
