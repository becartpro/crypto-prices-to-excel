"""Fetch market data from the public Binance API and export a formatted Excel report.

Usage:
    python crypto_report.py --symbols BTCUSDT ETHUSDT SOLUSDT --days 30 --out report.xlsx
"""
from __future__ import annotations

import argparse
import logging
import time
from dataclasses import dataclass
from datetime import datetime, timezone

import requests
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE_URL = "https://api.binance.com/api/v3/klines"
log = logging.getLogger("crypto_report")


@dataclass
class Candle:
    date: datetime
    open: float
    high: float
    low: float
    close: float
    volume: float


class ApiError(RuntimeError):
    pass


def fetch_daily_candles(symbol: str, days: int, session: requests.Session | None = None,
                        retries: int = 3, backoff: float = 1.0) -> list[Candle]:
    """Return the last `days` daily candles for `symbol`, retrying on network errors and rate limits."""
    session = session or requests.Session()
    params = {"symbol": symbol.upper(), "interval": "1d", "limit": days}
    for attempt in range(1, retries + 1):
        try:
            resp = session.get(BASE_URL, params=params, timeout=10)
        except requests.RequestException as exc:
            log.warning("%s: network error (%s), attempt %d/%d", symbol, exc, attempt, retries)
        else:
            if resp.status_code == 200:
                return [_parse(row) for row in resp.json()]
            if resp.status_code in (429, 418) or resp.status_code >= 500:
                log.warning("%s: HTTP %d, attempt %d/%d", symbol, resp.status_code, attempt, retries)
            else:
                raise ApiError(f"{symbol}: HTTP {resp.status_code} - {resp.text[:200]}")
        time.sleep(backoff * attempt)
    raise ApiError(f"{symbol}: failed after {retries} attempts")


def _parse(row: list) -> Candle:
    return Candle(
        date=datetime.fromtimestamp(row[0] / 1000, tz=timezone.utc).replace(tzinfo=None),
        open=float(row[1]), high=float(row[2]), low=float(row[3]),
        close=float(row[4]), volume=float(row[5]),
    )


HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def build_workbook(data: dict[str, list[Candle]]) -> Workbook:
    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    _header(summary, ["Symbol", "First close", "Last close", "Change %", "Max high", "Min low"])

    for row_idx, (symbol, candles) in enumerate(data.items(), start=2):
        ws = wb.create_sheet(symbol)
        _header(ws, ["Date", "Open", "High", "Low", "Close", "Volume"])
        for i, c in enumerate(candles, start=2):
            ws.append([c.date, c.open, c.high, c.low, c.close, c.volume])
            ws.cell(row=i, column=1).number_format = "DD/MM/YYYY"
        last = len(candles) + 1
        for col in range(2, 7):
            for r in range(2, last + 1):
                ws.cell(row=r, column=col).number_format = "#,##0.00"
        _autosize(ws)
        _chart(ws, symbol, last)

        # Summary rows use formulas so the file stays live if data is edited.
        summary.cell(row=row_idx, column=1, value=symbol)
        summary.cell(row=row_idx, column=2, value=f"='{symbol}'!E2")
        summary.cell(row=row_idx, column=3, value=f"='{symbol}'!E{last}")
        summary.cell(row=row_idx, column=4, value=f"=IFERROR(C{row_idx}/B{row_idx}-1,0)").number_format = "0.0%"
        summary.cell(row=row_idx, column=5, value=f"=MAX('{symbol}'!C2:C{last})")
        summary.cell(row=row_idx, column=6, value=f"=MIN('{symbol}'!D2:D{last})")
        for col in (2, 3, 5, 6):
            summary.cell(row=row_idx, column=col).number_format = "#,##0.00"
    _autosize(summary)
    return wb


def _header(ws, labels):
    ws.append(labels)
    for cell in ws[1]:
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"


def _autosize(ws):
    for i, column in enumerate(ws.columns, start=1):
        width = max(len(str(c.value)) if c.value is not None else 0 for c in column)
        ws.column_dimensions[get_column_letter(i)].width = min(max(12, width + 2), 40)


def _chart(ws, symbol, last_row):
    chart = LineChart()
    chart.title = f"{symbol} - daily close"
    chart.height, chart.width = 8, 18
    chart.add_data(Reference(ws, min_col=5, min_row=1, max_row=last_row), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=last_row))
    ws.add_chart(chart, "H2")


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--symbols", nargs="+", default=["BTCUSDT", "ETHUSDT"])
    parser.add_argument("--days", type=int, default=30)
    parser.add_argument("--out", default="crypto_report.xlsx")
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")

    with requests.Session() as session:
        data = {s.upper(): fetch_daily_candles(s, args.days, session) for s in args.symbols}
    build_workbook(data).save(args.out)
    log.info("Report written to %s (%d symbols)", args.out, len(data))


if __name__ == "__main__":
    main()
