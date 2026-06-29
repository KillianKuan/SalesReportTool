#!/usr/bin/env python3
"""Merge historical shipping data from year-based folders into a single CSV.

Usage:
    python scripts/merge_historical.py

Reads:  data/{year}/*.xlsx (Actual sheet), for year != current calendar year
Writes: data/Over the Years/historical.csv (UTF-8-BOM, Excel-compatible)
        data/Over the Years/historical.parquet (fast cold-start loading)
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


def _get_all_xlsx(year_dir: Path) -> list[Path]:
    return sorted(year_dir.glob("*.xlsx"), key=lambda p: p.name)


def main() -> None:
    if not DATA_DIR.exists():
        print(f"Data directory not found: {DATA_DIR}")
        sys.exit(1)

    year_dirs = sorted(
        (
            entry for entry in DATA_DIR.iterdir()
            if entry.is_dir()
            and entry.name.isdigit()
            and int(entry.name) != CURRENT_YEAR
        ),
        key=lambda p: int(p.name),
    )

    if not year_dirs:
        print(f"No historical year folders found under {DATA_DIR} (looking for years != {CURRENT_YEAR}).")
        sys.exit(0)

    frames = []
    seen_optional = set()
    for d in year_dirs:
        xlsx_files = _get_all_xlsx(d)
        if not xlsx_files:
            print(f"  SKIP {d.name}/: no .xlsx file found")
            continue

        for xlsx in xlsx_files:
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
            frame = raw[use_cols].copy()
            frames.append(frame)
            seen_optional.update(col for col in OPTIONAL_COLS if col in frame.columns)
            print(f"  OK   {d.name}/{xlsx.name}: {len(frame):,} rows")

    if not frames:
        print("No data to merge — output file not written.")
        sys.exit(0)

    merged = pd.concat(frames, ignore_index=True)
    output_cols = REQUIRED_COLS + [col for col in OPTIONAL_COLS if col in seen_optional]
    for col in OPTIONAL_COLS:
        if col not in merged.columns:
            merged[col] = pd.NA
    merged = merged.loc[:, output_cols]

    out_dir = DATA_DIR / "Over the Years"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / "historical.csv"
    merged.to_csv(out_path, index=False, encoding="utf-8-sig")
    print(f"\nWrote {len(merged):,} rows → {out_path}")

    # Also write a Parquet copy for fast cold-start loading in the app.
    # load_historical_csv() prefers this file when present, falling back to CSV.
    parquet_path = out_dir / "historical.parquet"
    try:
        merged.to_parquet(parquet_path, index=False)
        print(f"Wrote {len(merged):,} rows → {parquet_path}")
    except Exception as e:  # pragma: no cover - depends on optional parquet engine
        print(f"  WARN: could not write Parquet ({e}); CSV is still available.")


if __name__ == "__main__":
    main()
