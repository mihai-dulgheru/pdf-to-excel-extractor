import importlib
import unittest
from datetime import datetime
from unittest.mock import patch

import requests

from functions import ExchangeRateUnavailable, get_bnr_exchange_rate

# the package re-exports the function under the submodule's name, so fetch the module itself
bnr = importlib.import_module("functions.get_bnr_exchange_rate")

PAGE_WITH_RATE = ('<table id="table-currencies"><tbody>'
                  '<tr><td>EUR</td><td>Euro</td><td>4,9771</td><td>-0,0002</td></tr>'
                  '</tbody></table>')


class _Response:
    def __init__(self, text, status_code=200):
        self.text, self.status_code, self.reason = text, status_code, "OK"


class TestExchangeRateOffline(unittest.TestCase):
    """No network: requests.get is replaced in the module under test."""

    def setUp(self):
        get_bnr_exchange_rate.cache_clear()
        self.date = datetime(2030, 1, 15)

    def test_offline_raises_and_a_later_retry_succeeds(self):
        """A failed lookup must raise (not return None) and must not be cached for the session."""
        with patch.object(bnr.requests, "get", side_effect=requests.ConnectionError("offline")):
            with self.assertRaises(ExchangeRateUnavailable) as ctx:
                get_bnr_exchange_rate(self.date)
        self.assertIn("15.01.2030", str(ctx.exception))

        with patch.object(bnr.requests, "get", return_value=_Response(PAGE_WITH_RATE)):
            self.assertEqual(get_bnr_exchange_rate(self.date), 4.9771)

    def test_request_uses_a_timeout(self):
        seen = {}

        def fake_get(url, **kwargs):
            seen.update(kwargs)
            return _Response(PAGE_WITH_RATE)

        with patch.object(bnr.requests, "get", side_effect=fake_get):
            get_bnr_exchange_rate(self.date)
        self.assertGreater(seen.get("timeout", 0), 0)


if __name__ == "__main__":
    unittest.main()
