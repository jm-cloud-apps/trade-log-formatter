"""Reconcile master-trades.xlsx open positions to a known cash state.

When the FIFO matcher misses closes (e.g. the closing day's IB PDF never
came through email, or the position was closed outside of IB's email
reports), `master-trades.xlsx` ends up with phantom open rows that
distort the Open Positions report forever after.

This script finds every row where Exit Qty / Exit Price is blank and
fills them in so the position is treated as closed:

    --mode manual    Exit Price = Entry Price (P&L zeroed). Best for
                     "I genuinely don't remember the close price and I
                     just want the row out of the open queue."
    --mode flat      Exit Price = 0 + Notes='reconcile-zeroed'. Use only
                     if you want P&L to reflect the full open notional
                     as a loss (rarely what you want).
    --mode price     Exit Price = --price for every row.

A timestamped backup is always taken before writing.

Usage:
    python reconcile_open_positions.py --dry-run
    python reconcile_open_positions.py --mode manual --exit-date 2026-05-19
    python reconcile_open_positions.py --mode price --price 100 --symbol AEHR
"""

from __future__ import annotations

import argparse
import shutil
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

BASE_PATH_TRADES = Path("/Users/michaeljacinto/Library/CloudStorage/OneDrive-Personal/Desktop - onedrive/trades")
MASTER_FILE = BASE_PATH_TRADES / "master-trades.xlsx"
SHEET = "Trades"


def is_open(row) -> bool:
    eq = row.get("Exit Qty")
    ep = row.get("Exit Price")
    return pd.isna(eq) or pd.isna(ep)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--mode", choices=("manual", "flat", "price"), default="manual")
    ap.add_argument("--price", type=float, help="Exit price for --mode price")
    ap.add_argument("--exit-date", default=datetime.now().strftime("%Y-%m-%d"),
                    help="Exit Date to stamp (YYYY-MM-DD). Defaults to today.")
    ap.add_argument("--exit-time", default=datetime.now().strftime("%H:%M:%S"),
                    help="Exit Time to stamp. Defaults to now.")
    ap.add_argument("--symbol", help="Limit to a single symbol (default: all)")
    ap.add_argument("--notes-tag", default="reconcile",
                    help="Tag appended to Notes so reconciled rows are traceable.")
    ap.add_argument("--dry-run", action="store_true",
                    help="Show what would change without writing.")
    args = ap.parse_args()

    if args.mode == "price" and args.price is None:
        print("error: --mode price requires --price", file=sys.stderr)
        return 2

    if not MASTER_FILE.exists():
        print(f"error: {MASTER_FILE} not found", file=sys.stderr)
        return 1

    # Read all sheets so we can write the workbook back intact.
    all_sheets = pd.read_excel(MASTER_FILE, sheet_name=None)
    df = all_sheets[SHEET]
    print(f"Loaded {SHEET}: {len(df)} rows total")

    open_mask = df.apply(is_open, axis=1)
    if args.symbol:
        open_mask &= df["Symbol"] == args.symbol.upper()
    candidates = df[open_mask]
    if candidates.empty:
        print("Nothing to reconcile — no rows match the filter.")
        return 0

    print(f"\nWill reconcile {len(candidates)} row(s):")
    cols = ["Symbol", "Qty", "Side", "Entry Price", "Entry Date"]
    print(candidates[cols].to_string(index=False))

    # Compute new exit values per row
    def new_exit(row):
        qty = abs(int(row["Qty"]))
        side = row["Side"]
        # For LONG (BUY) Qty=+, Exit Qty negative.
        # For SHORT (SELL) Qty=-, Exit Qty positive.
        sign = -1 if str(side).upper() in ("BUY", "LONG") else 1
        if args.mode == "manual":
            ep = float(row["Entry Price"])
        elif args.mode == "flat":
            ep = 0.0
        else:  # price
            ep = float(args.price)
        return sign * qty, ep

    if args.dry_run:
        print(f"\n--dry-run set — no changes written. mode={args.mode}")
        return 0

    # Backup before mutation
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = MASTER_FILE.with_name(f"master-trades_reconcile_backup_{ts}.xlsx")
    shutil.copy2(MASTER_FILE, backup)
    print(f"\n📑 Backup: {backup.name}")

    notes_tag = f"[{args.notes_tag}-{args.mode}-{args.exit_date}]"
    for idx in candidates.index:
        eq, ep = new_exit(df.loc[idx])
        df.at[idx, "Exit Qty"] = eq
        df.at[idx, "Exit Price"] = ep
        df.at[idx, "Exit Date"] = args.exit_date
        df.at[idx, "Exit Time"] = args.exit_time
        existing_notes = df.at[idx, "Notes"] if "Notes" in df.columns else ""
        existing_notes = "" if pd.isna(existing_notes) else str(existing_notes)
        df.at[idx, "Notes"] = (existing_notes + " " + notes_tag).strip()

    all_sheets[SHEET] = df
    with pd.ExcelWriter(MASTER_FILE, engine="openpyxl") as w:
        for name, frame in all_sheets.items():
            frame.to_excel(w, sheet_name=name, index=False)

    print(f"✅ Wrote {MASTER_FILE.name}: {len(candidates)} row(s) reconciled "
          f"with Exit Date {args.exit_date}, mode={args.mode}.")
    print(f"   Backup retained at {backup}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
