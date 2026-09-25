# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A pipeline that turns Interactive Brokers "Daily Trade Report" PDF emails into a consolidated Excel workbook (`master-trades.xlsx`) used for trade tracking. The QuantForge app (sibling repo `../quantforge`) shells out to this repo's scripts via `backend/formatter/` and reads the workbook for all of its trade analytics — changes to the output format here ripple into QuantForge.

## Commands

```bash
source venv/bin/activate
pip install -r requirements.txt

python run_daily.py             # full pipeline: fetch from Gmail + format + summarize
python run_daily.py --fetch     # fetch PDFs only
python run_daily.py --format    # format only (process already-downloaded PDFs)

python trade-log-formatter.py   # interactive: enter MM.YYYY to process a month, or RESET

python reconcile_open_positions.py --dry-run          # list phantom open rows
python reconcile_open_positions.py --mode manual --exit-date YYYY-MM-DD
```

Setup requires `config.json` (copy from `config.example.json`): Gmail IMAP credentials (app password), the IB sender address, and `download_base` (the trades folder).

`test_trade_formatter.py` exists but is not runnable as-is: it does `from trade_log_formatter.py import ...`, which fails because the main script's filename is hyphenated (`trade-log-formatter.py`) and can't be imported as a module.

## Pipeline architecture

1. **`email_fetcher.py`** — IMAP-searches Gmail for IB report emails, saves each PDF attachment as `DailyTradeReport.YYYYMMDD.pdf` into the month folder `<download_base>/MM.YYYY/`. Dedupes via `email_state.json` in this repo dir.
2. **`trade-log-formatter.py`** — the core (~1700 lines, all module-level functions). For a given `MM.YYYY` folder: extracts trade lines from each PDF with PyMuPDF + a regex over masked account rows (`U***1234 SYMBOL date,time ... BUY/SELL qty price`), consolidates same-symbol/date/side fills, FIFO-matches sells against buys, then rewrites the master workbook.
3. **`summarize.py`** — appends one line per report date (closed lots + realized P&L) to `daily_summary.md`.
4. **`notifier.py`** — on any failure, writes `last_error.log` and emails a self-alert via the same Gmail account.
5. **`run_daily.py`** — orchestrates 1→3 and routes exceptions to 4.

### File locations and state

- Input PDFs + per-month `processed_files.json` (dedupe — delete an entry to allow reprocessing): `<trades base>/MM.YYYY/`
- Output: `<trades base>/master-trades.xlsx` (sheets: `Trades` = FIFO-matched positions, `Raw Trades`, `Consolidated Trades`), plus `master-copy-backup.xlsx` written before each update
- Trades base path is hardcoded as `BASE_PATH_TRADES` in `trade-log-formatter.py` and `reconcile_open_positions.py`; `run_daily.py` honors the `TRADES_BASE_PATH` env var
- `DEBUG` and `TEST_MODE` toggles are module-level constants at the top of `trade-log-formatter.py` (TEST_MODE switches to `master-copy-test.xlsx` / `processed_files_test.json`)

### Gotchas that have caused real data corruption

- **Missing PDF ⇒ phantom SHORTs.** FIFO matching processes days in order. If one day's PDF never arrived (email missed, fetch failure), the BUYs from that day don't exist, so later SELLs can't match and get recorded as new SHORT entries in the `Trades` sheet. Fix: get the missing PDF into the month folder, remove the affected rows/processed-files entries, and reprocess — or use `reconcile_open_positions.py` for positions closed outside IB email reports.
- **Scale-out exits are written as Excel formulas.** Multi-fill exits store the weighted-average exit price as a literal arithmetic formula (e.g. `=((100*49.401)+(125*54.61))/225`). Plain `pd.read_excel` returns NaN for those cells — QuantForge compensates in its `read_trades_excel()`; any new reader here must too.
- **Options trades**: symbols like `UNH 16JAN26 550 C` are detected by pattern and their prices are multiplied by 100.
- **`RESET`** (interactive prompt) clears all three sheets and every month's `processed_files.json` after taking a timestamped backup — full reprocess from PDFs is the recovery path.
- `config.json` contains a real Gmail app password — never commit it (only `config.example.json` belongs in git).
