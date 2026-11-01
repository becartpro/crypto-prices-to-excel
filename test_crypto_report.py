import sys
from pathlib import Path
from unittest.mock import MagicMock

import pytest
import requests
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent))
import crypto_report as cr  # noqa: E402

ROWS = [
    [1730419200000, "100", "110", "95", "105", "12.5", 0, 0, 0, 0, 0, 0],
    [1730505600000, "105", "120", "101", "118", "20.0", 0, 0, 0, 0, 0, 0],
]


def _response(status, payload=None, text=""):
    resp = MagicMock()
    resp.status_code = status
    resp.json.return_value = payload
    resp.text = text
    return resp


def test_parses_candles():
    session = MagicMock()
    session.get.return_value = _response(200, ROWS)
    candles = cr.fetch_daily_candles("btcusdt", 2, session)
    assert len(candles) == 2
    assert candles[1].close == 118.0
    assert session.get.call_args.kwargs["params"]["symbol"] == "BTCUSDT"


def test_retries_on_rate_limit(monkeypatch):
    monkeypatch.setattr(cr.time, "sleep", lambda s: None)
    session = MagicMock()
    session.get.side_effect = [_response(429), requests.ConnectionError("boom"), _response(200, ROWS)]
    assert len(cr.fetch_daily_candles("ETHUSDT", 2, session)) == 2
    assert session.get.call_count == 3


def test_client_error_is_not_retried():
    session = MagicMock()
    session.get.return_value = _response(400, text="Invalid symbol")
    with pytest.raises(cr.ApiError):
        cr.fetch_daily_candles("NOPE", 2, session)
    assert session.get.call_count == 1


def test_workbook_has_summary_and_sheets(tmp_path):
    candles = [cr._parse(r) for r in ROWS]
    out = tmp_path / "r.xlsx"
    cr.build_workbook({"BTCUSDT": candles}).save(out)
    wb = load_workbook(out)
    assert wb.sheetnames == ["Summary", "BTCUSDT"]
    assert wb["Summary"]["D2"].value.startswith("=IFERROR")
    assert wb["BTCUSDT"]["E3"].value == 118.0
