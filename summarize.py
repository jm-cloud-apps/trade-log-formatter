"""Append a one-line daily summary to daily_summary.md after formatter runs.

For each new report date we processed, reads master-trades.xlsx and computes:
  - count of closed lots (Exit Date == report date)
  - realized P&L: sum over closed lots of (Exit Price - Entry Price) * Qty
    (sign-adjusted for SELL/short entries)
"""

import logging
import re
from datetime import datetime
from pathlib import Path

import pandas as pd

logger = logging.getLogger(__name__)

SCRIPT_DIR = Path(__file__).resolve().parent
SUMMARY_FILE = SCRIPT_DIR / "daily_summary.md"

# Match the date inside filenames like DailyTradeReport.20260511.pdf
FILENAME_DATE_RE = re.compile(r"(\d{8})")


def report_dates_from_paths(paths: list[Path]) -> list[datetime]:
    dates: list[datetime] = []
    for p in paths:
        m = FILENAME_DATE_RE.search(p.name)
        if not m:
            continue
        try:
            dates.append(datetime.strptime(m.group(1), "%Y%m%d"))
        except ValueError:
            continue
    return sorted(set(dates))


def _to_date(val) -> datetime | None:
    if pd.isna(val):
        return None
    try:
        return pd.to_datetime(val).to_pydatetime()
    except Exception:
        return None


def compute_day_summary(master_xlsx: Path, report_date: datetime) -> dict:
    """Return {trade_count, realized_pnl} for closed lots whose Exit Date matches."""
    if not master_xlsx.exists():
        return {"trade_count": 0, "realized_pnl": 0.0, "note": "master file missing"}
    try:
        df = pd.read_excel(master_xlsx)
    except Exception as e:
        return {"trade_count": 0, "realized_pnl": 0.0, "note": f"read error: {e}"}

    needed = {"Exit Date", "Exit Price", "Entry Price", "Side"}
    if not needed.issubset(df.columns):
        return {"trade_count": 0, "realized_pnl": 0.0, "note": "schema mismatch"}

    target = report_date.date()
    df = df.copy()
    df["_exit"] = df["Exit Date"].apply(lambda v: _to_date(v).date() if _to_date(v) else None)
    closed = df[df["_exit"] == target].copy()

    if closed.empty:
        return {"trade_count": 0, "realized_pnl": 0.0}

    qty_col = "Exit Qty" if "Exit Qty" in closed.columns and closed["Exit Qty"].notna().any() else "Qty"
    closed["_qty"] = pd.to_numeric(closed[qty_col], errors="coerce").fillna(0)
    closed["_entry"] = pd.to_numeric(closed["Entry Price"], errors="coerce").fillna(0)
    closed["_exit_p"] = pd.to_numeric(closed["Exit Price"], errors="coerce").fillna(0)

    def pnl_row(r):
        sign = 1 if str(r["Side"]).strip().upper() in ("BUY", "LONG") else -1
        return sign * (r["_exit_p"] - r["_entry"]) * r["_qty"]

    closed["_pnl"] = closed.apply(pnl_row, axis=1)
    return {
        "trade_count": int(len(closed)),
        "realized_pnl": float(closed["_pnl"].sum()),
    }


def append_summary(master_xlsx: Path, report_dates: list[datetime]) -> list[str]:
    """Append one line per report date to daily_summary.md. Returns the lines written."""
    if not report_dates:
        return []

    lines = []
    for d in report_dates:
        s = compute_day_summary(master_xlsx, d)
        date_str = d.strftime("%Y-%m-%d")
        if "note" in s:
            line = f"- {date_str}: ({s['note']})"
        else:
            pnl = s["realized_pnl"]
            sign = "+" if pnl >= 0 else "−"
            line = (
                f"- {date_str}: {s['trade_count']} closed trade(s), "
                f"realized {sign}${abs(pnl):,.2f}"
            )
        lines.append(line)

    header_needed = not SUMMARY_FILE.exists()
    with open(SUMMARY_FILE, "a") as f:
        if header_needed:
            f.write("# Daily Trade Summary\n\n")
        for line in lines:
            f.write(line + "\n")

    return lines
