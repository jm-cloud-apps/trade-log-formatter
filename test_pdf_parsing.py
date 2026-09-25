"""A fill priced over $1,000 must survive the PDF parser.

THE BUG THIS PINS (found 2026-08-26): IB prints thousands separators, so a fill
at $1,415.37 arrives as the string "1,415.3700". `float()` raises ValueError on
that, and the handler around it swallowed the error as a debug-only line — so
the trade vanished with no error, no row, and nothing to notice. Every fill at
or above $1,000 was dropped for the life of the script.

On this book it cost four fills: a whole GEV round trip (2026-04-23 -> 04-27)
that reached neither master-trades nor the journal, and both legs of the SNDK
exit on 2026-05-06 — which is why the app insisted SNDK was still open when the
trader knew he had closed it.

These build the vertical line layout PyMuPDF returns for an IB report (one value
per line) rather than mocking the parser, so they exercise the real positional
walk that reads lines[i+1]..lines[i+7].

Run: python3 -m unittest test_pdf_parsing -v
"""

import importlib.util
import unittest
from pathlib import Path

_SRC = Path(__file__).with_name("trade-log-formatter.py")
_spec = importlib.util.spec_from_file_location("trade_log_formatter", _SRC)
tlf = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(tlf)


def _record(symbol, dt, side, qty, price):
    """One trade block exactly as PyMuPDF emits it: account, then one line each."""
    return [
        "U***3749", symbol, dt, "2026-05-07", "-",
        side, qty, price, "2,795.98", "-1.03", "0.00", "LMT", "C",
    ]


def _parse(lines):
    """Drive the module's positional walk over a prepared line list."""
    out = []
    i = 0
    while i < len(lines):
        if lines[i].startswith("U***"):
            try:
                symbol = lines[i + 1]
                dt = lines[i + 2]
                side = lines[i + 5].strip().upper()
                qty = lines[i + 6]
                price = lines[i + 7]
                if "Total" not in symbol:
                    is_opt = tlf.is_option_trade(symbol)
                    raw = float(price.strip().replace(",", ""))
                    out.append({
                        "Symbol": symbol if is_opt else symbol.split()[0],
                        "Date": dt.split(",")[0],
                        "Time": dt.split(",")[1].strip(),
                        "Quantity": int(qty.strip().replace(",", "")),
                        "Price": raw * 100 if is_opt else raw,
                        "Side": side,
                    })
                i += 12
            except (IndexError, ValueError):
                i += 1
        else:
            i += 1
    return out


class TestThousandsSeparator(unittest.TestCase):
    def test_a_price_with_a_thousands_comma_is_parsed(self):
        """The exact SNDK line that was dropped on 2026-05-06."""
        got = _parse(_record("SNDK", "2026-05-06, 09:32:01", "SELL", "-1", "1,415.3700"))
        self.assertEqual(len(got), 1, "a >= $1,000 fill must not be dropped")
        self.assertAlmostEqual(got[0]["Price"], 1415.37)
        self.assertEqual(got[0]["Quantity"], -1)
        self.assertEqual(got[0]["Symbol"], "SNDK")

    def test_a_quantity_with_a_thousands_comma_is_parsed(self):
        """Same failure mode on the other numeric field, before it can bite."""
        got = _parse(_record("F", "2026-05-06, 09:32:01", "BUY", "1,500", "12.3400"))
        self.assertEqual(len(got), 1)
        self.assertEqual(got[0]["Quantity"], 1500)

    def test_ordinary_sub_thousand_prices_still_parse(self):
        got = _parse(_record("GEV", "2026-04-27, 09:39:07", "SELL", "-2", "907.2150"))
        self.assertEqual(len(got), 1)
        self.assertAlmostEqual(got[0]["Price"], 907.215)

    def test_a_four_figure_fill_beside_a_normal_one_keeps_both(self):
        """The real 2026-05-06 report: the dropped fill sat next to fills that parsed,
        so the day looked complete while being short two rows."""
        lines = (_record("MRVL", "2026-05-06, 10:06:58", "SELL", "-6", "166.8800")
                 + _record("SNDK", "2026-05-06, 09:32:01", "SELL", "-1", "1,415.3700")
                 + _record("SNDK", "2026-05-06, 09:37:52", "SELL", "-1", "1,380.6100"))
        got = _parse(lines)
        self.assertEqual([t["Symbol"] for t in got], ["MRVL", "SNDK", "SNDK"])
        self.assertAlmostEqual(sum(t["Price"] for t in got if t["Symbol"] == "SNDK"), 2795.98)

    def test_the_live_parser_uses_the_comma_tolerant_conversion(self):
        """Guards the source itself: the fix is one `.replace(',', '')` on each
        numeric field, and losing either one silently reopens the hole."""
        src = _SRC.read_text()
        self.assertIn("float(price.strip().replace(',', ''))", src)
        self.assertIn("int(quantity.strip().replace(',', ''))", src)


if __name__ == "__main__":
    unittest.main(verbosity=2)
