"""Orchestrator: fetch new IB reports from Gmail, run formatter, summarize.

Usage:
    python run_daily.py            # fetch + format + summarize (default)
    python run_daily.py --fetch    # fetch only
    python run_daily.py --format   # format only
"""

import argparse
import logging
import os
import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path

from email_fetcher import fetch_new_reports, load_config, load_state
from notifier import report_failure
from summarize import append_summary, report_dates_from_paths

SCRIPT_DIR = Path(__file__).resolve().parent
FORMATTER = SCRIPT_DIR / "trade-log-formatter.py"

# The formatter writes the master workbook into the trades base folder (see
# trade-log-formatter.py: os.path.join(BASE_PATH_TRADES, "master-trades.xlsx")),
# NOT into this repo dir. summarize.py must read it from the same place — the
# previous SCRIPT_DIR path never existed, which is why every day showed
# "(master file missing)".
TRADES_BASE = Path(os.getenv(
    "TRADES_BASE_PATH",
    "/Users/michaeljacinto/Library/CloudStorage/OneDrive-Personal/Desktop - onedrive/trades",
))
MASTER_XLSX = TRADES_BASE / "master-trades.xlsx"

_MONTH_RE = re.compile(r"^\d{2}\.\d{4}$")


def months_from_paths(paths: list[Path]) -> list[str]:
    """Distinct MM.YYYY folder names the fetched reports landed in, in order."""
    out: list[str] = []
    for p in paths:
        folder = p.parent.name
        if _MONTH_RE.match(folder) and folder not in out:
            out.append(folder)
    return out


def run_formatter(month: str, apply: bool = True) -> int:
    """Run the formatter, piping the target MM.YYYY month + y/N confirmation.

    The formatter prompts twice via input():
        1. the month folder (we always feed `month`)
        2. 'Apply these changes? (y/N):' — feed 'y' to write, 'n' to dry-run.

    Pre-feeding both lines before stdin closes prevents the EOFError that
    used to bubble up as 'Unexpected error processing folder' whenever this
    script ran from automation.
    """
    answer = "y" if apply else "n"
    logging.info("Running formatter for %s (apply=%s): %s", month, apply, FORMATTER)
    result = subprocess.run(
        [sys.executable, str(FORMATTER)],
        cwd=SCRIPT_DIR,
        input=f"{month}\n{answer}\n",
        text=True,
    )
    return result.returncode


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--fetch", action="store_true", help="Only fetch emails")
    parser.add_argument("--format", dest="do_format", action="store_true", help="Only run formatter")
    parser.add_argument(
        "--month",
        default=datetime.now().strftime("%m.%Y"),
        help="Month folder to format (MM.YYYY). Defaults to current month.",
    )
    parser.add_argument(
        "--no-apply",
        dest="apply",
        action="store_false",
        help="Run formatter in preview mode (won't write to Trades.xlsx).",
    )
    parser.set_defaults(apply=True)
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

    do_fetch = args.fetch or not args.do_format
    do_format = args.do_format or not args.fetch

    config = load_config()
    saved: list[Path] = []

    if do_fetch:
        try:
            state = load_state()
            saved = fetch_new_reports(config, state)
            logging.info("Fetched %d new report(s).", len(saved))
        except Exception as e:
            report_failure(config, "email_fetch", e)
            logging.error("Email fetch failed: %s", e)
            return 1

    if do_format:
        # Format the month(s) the new reports actually landed in (parsed from
        # the saved file paths), plus the explicitly-requested month. Without
        # this, a month rollover meant June reports were fetched into 06.YYYY
        # while the formatter only ran on the selected 05.YYYY — so the new
        # trades were never written to the master and the summary came up empty.
        months_to_format = months_from_paths(saved)
        if args.month not in months_to_format:
            months_to_format.append(args.month)
        logging.info("Formatting month(s): %s", ", ".join(months_to_format))
        try:
            for mth in months_to_format:
                rc = run_formatter(mth, apply=args.apply)
                if rc != 0:
                    raise RuntimeError(f"formatter exited with code {rc} for {mth}")
        except Exception as e:
            report_failure(config, "formatter", e)
            logging.error("Formatter failed: %s", e)
            return 1

        # Summarize newly fetched report dates after formatter has updated master.
        try:
            dates = report_dates_from_paths(saved)
            lines = append_summary(MASTER_XLSX, dates)
            if lines:
                print()
                print("Daily summary appended to daily_summary.md:")
                for line in lines:
                    print(f"  {line}")
        except Exception as e:
            # Summary failure is non-fatal — log it but don't error out.
            report_failure(config, "summarize", e)
            logging.warning("Summary generation failed (non-fatal): %s", e)

    return 0


if __name__ == "__main__":
    sys.exit(main())
