import unittest
from datetime import datetime

from functions import get_bnr_exchange_rate


class TestExchangeRate(unittest.TestCase):

    def test_valid_date(self):
        """Test if a valid date returns a float exchange rate."""
        date = datetime.strptime("2024-02-14", "%Y-%m-%d")
        rate = get_bnr_exchange_rate(date)
        self.assertIsInstance(rate, float)

    def test_expected_exchange_rates(self):
        """Test specific exchange rates for given dates."""
        test_cases = {"2025-02-14": 4.9771, "2025-02-13": 4.9773, "2025-02-12": 4.9770, "2025-02-11": 4.9771,
                      "2025-02-10": 4.9765, }

        for date_str, expected_rate in test_cases.items():
            with self.subTest(date=date_str, expected=expected_rate):
                date = datetime.strptime(date_str, "%Y-%m-%d")
                rate = get_bnr_exchange_rate(date)
                self.assertAlmostEqual(rate, expected_rate, places=4, msg=f"Expected {expected_rate} for {date_str}")

    def test_error_handling(self):
        """Test if the function raises TypeError for an invalid date format."""
        date = "invalid_date"
        with self.assertRaises(TypeError):
            get_bnr_exchange_rate(date)


if __name__ == "__main__":
    unittest.main()
