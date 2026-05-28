#!/usr/bin/env python3
"""Merge historical shipping data from year-based folders into a single CSV.

Usage:
    python scripts/merge_historical.py

Reads:  data/{year}/*.xlsx (Actual sheet), for year != current calendar year
Writes: data/Over the Years/historical.csv (UTF-8-BOM, Excel-compatible)
"""

import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = REPO_ROOT / "data"
CURRENT_YEAR = datetime.now().year

REQUIRED_COLS = [
    "Customer Name", "Ship Date", "QTY",
    "SALES Total AMT", "final GP(NTD,data from Financial Report)",
    "Part Number", "Category",
]
OPTIONAL_COLS = ["DES", "SALE_Person", "Currency", "UP", "TP(USD)"]


def _find_latest_xlsx(year_dir: Path) -> Path | None:
    files = list(year_dir.glob("*.xlsx"))
    return max(files, key=lambda f: f.stat().st_mtime) if files else None


def main() -> None:
    if not DATA_DIR.exists():
        print(f"Data directory not found: {DATA_DIR}")
        sys.exit(1)

    year_dirs = sorted(
        (
            entry for entry in DATA_DIR.iterdir()
            if entry.is_dir()
            and entry.name.isdigit()
            and 2019 <= int(entry.name) <= 2099
            and int(entry.name) != CURRENT_YEAR
        ),
        key=lambda p: int(p.name),
    )

    if not year_dirs:
        print(f"No historical year folders found under {DATA_DIR} (looking for years != {CURRENT_YEAR}).")
        sys.exit(0)

    frames = []
    for d in year_dirs:
        xlsx = _find_latest_xlsx(d)
        if xlsx is None:
            print(f"  SKIP {d.name}/: no .xlsx file found")
            continue
        try:
            try:
                xl = pd.ExcelFile(xlsx, engine="calamine")
            except ImportError:
                xl = pd.ExcelFile(xlsx)
        except Exception as e:
            print(f"  SKIP {d.name}/{xlsx.name}: cannot open ({e})")
            continue
        if "Actual" not in xl.sheet_names:
            print(f"  SKIP {d.name}/{xlsx.name}: 'Actual' sheet not found (got {xl.sheet_names})")
            continue
        raw = xl.parse("Actual")
        missing = [c for c in REQUIRED_COLS if c not in raw.columns]
        if missing:
            print(f"  SKIP {d.name}/{xlsx.name}: missing required columns {missing}")
            continue
        use_cols = REQUIRED_COLS + [c for c in OPTIONAL_COLS if c in raw.columns]
        frames.append(raw[use_cols].copy())
        print(f"  OK   {d.name}/{xlsx.name}: {len(raw):,} rows")

    if not frames:
        print("No data to merge — output file not written.")
        sys.exit(0)

    merged = pd.concat(frames, ignore_index=True)
    out_dir = DATA_DIR / "Over the Years"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / "historical.csv"
    merged.to_csv(out_path, index=False, encoding="utf-8-sig")
    print(f"\nWrote {len(merged):,} rows → {out_path}")


if __name__ == "__main__":
    main()
