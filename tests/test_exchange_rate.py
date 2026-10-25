import importlib
import os
import unittest
from datetime import datetime
from unittest.mock import patch

from functions import get_bnr_exchange_rate

# the package re-exports the function under the submodule's name, so fetch the module itself
bnr = importlib.import_module("functions.get_bnr_exchange_rate")

# cursbnr.ro archive pages recorded for February 2025 (EUR row only); 15-16 February is a weekend without a page
RATES = {"2025-02-10": "4,9765", "2025-02-11": "4,9771", "2025-02-12": "4,9770", "2025-02-13": "4,9773",
         "2025-02-14": "4,9771"}


class _Response:
    def __init__(self, text, status_code=200):
        self.text, self.status_code = text, status_code
        self.reason = "OK" if status_code == 200 else "Not Found"


def _archive_page(url, **kwargs):
    rate = RATES.get(url[-10:])
    if rate is None:
        return _Response("", 404)
    return _Response(f'<table id="table-currencies"><tbody><tr><td>EUR</td><td>Euro</td><td>{rate}</td>'
                     f'<td>0,0000</td></tr></tbody></table>')


class TestExchangeRate(unittest.TestCase):
    """The BNR archive is replayed from recorded pages: no network, the same answers on every run."""

    def setUp(self):
        get_bnr_exchange_rate.cache_clear()
        patcher = patch.object(bnr.requests, "get", side_effect=_archive_page)
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_expected_exchange_rates(self):
        for date_str, expected in RATES.items():
            with self.subTest(date=date_str):
                rate = get_bnr_exchange_rate(datetime.strptime(date_str, "%Y-%m-%d"))
                self.assertIsInstance(rate, float)
                self.assertAlmostEqual(rate, float(expected.replace(",", ".")), places=4)

    def test_weekend_falls_back_to_the_previous_published_day(self):
        self.assertAlmostEqual(get_bnr_exchange_rate(datetime(2025, 2, 16)), 4.9771, places=4)

    def test_error_handling(self):
        """An invalid date type raises TypeError before any request is made."""
        with self.assertRaises(TypeError):
            get_bnr_exchange_rate("invalid_date")


@unittest.skipUnless(os.environ.get("BNR_LIVE"), "set BNR_LIVE=1 to check the real cursbnr.ro page layout on demand")
class TestExchangeRateLive(unittest.TestCase):
    """The replay above cannot notice a layout change on cursbnr.ro; run this before a release."""

    def test_live_site_still_parses(self):
        get_bnr_exchange_rate.cache_clear()
        self.assertAlmostEqual(get_bnr_exchange_rate(datetime(2025, 2, 14)), 4.9771, places=4)


if __name__ == "__main__":
    unittest.main()
