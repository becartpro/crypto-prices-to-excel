# crypto-prices-to-excel

Small Python tool that pulls daily market data from the public **Binance REST API** and turns it into a ready-to-share **Excel report**: one sheet per symbol, a line chart, and a summary sheet driven by live formulas (change %, max, min).

It's a compact example of the kind of job I do for clients: connect to an API, handle its failures properly, and deliver the data in the format the business actually uses.

## Features
- REST API client with timeouts, retries and backoff on rate limits (HTTP 429/418) and server errors
- Client errors (bad symbol, bad params) fail fast with a clear message
- Formatted Excel output (header styling, number formats, frozen panes, auto column widths, charts)
- Summary sheet built with Excel formulas, so the numbers stay correct if the data is edited
- Unit tests with mocked HTTP calls (no network needed)

## Usage
```bash
pip install -r requirements.txt
python crypto_report.py --symbols BTCUSDT ETHUSDT SOLUSDT --days 30 --out report.xlsx
```

## Tests
```bash
pytest -q
```

## Adapting it
Swapping the data source for another REST API (a CRM, an accounting tool, an open-data API) only means rewriting `fetch_daily_candles` and `_parse`; the Excel layer stays the same.

---
Simon Becart — freelance software engineer (Python, API integration, automation) · [LinkedIn](https://www.linkedin.com/in/simonbecart/)
