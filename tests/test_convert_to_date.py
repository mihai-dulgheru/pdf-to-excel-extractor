import unittest
from datetime import datetime

import numpy as np
import pandas as pd

from functions import convert_to_date


class TestConvertToDate(unittest.TestCase):

    def test_none_values(self):
        """Test that None and NaN values return None."""
        self.assertIsNone(convert_to_date(None))
        self.assertIsNone(convert_to_date(np.nan))
        self.assertIsNone(convert_to_date(pd.NA))

    def test_datetime_objects(self):
        """Test that datetime objects are returned unchanged."""
        date = datetime(2025, 3, 15)
        self.assertEqual(convert_to_date(date), date)

    def test_string_format(self):
        """Test conversion of strings in the primary format."""
        date_str = "15.03.2025"
        expected = datetime(2025, 3, 15)
        self.assertEqual(convert_to_date(date_str), expected)

    def test_alternative_string_formats(self):
        """Test conversion of strings in alternative formats."""
        formats = {"2025-03-15": datetime(2025, 3, 15), "15/03/2025": datetime(2025, 3, 15),
                   "Mar 15, 2025": datetime(2025, 3, 15)}

        for date_str, expected in formats.items():
            with self.subTest(date=date_str):
                result = convert_to_date(date_str)
                self.assertEqual(result, expected)

    def test_excel_numeric_date(self):
        """Test conversion of Excel numeric dates."""
        # Excel date for March 15, 2025 (days since 1899-12-30)
        excel_date = 45731.0  # Adjusted to the correct date
        expected = datetime(2025, 3, 15)
        result = convert_to_date(excel_date)
        self.assertEqual(result.date(), expected.date())

    def test_invalid_inputs(self):
        """Test that invalid inputs return None."""
        invalid_inputs = ["invalid", "32.13.2025", "abc"]
        for value in invalid_inputs:
            with self.subTest(value=value):
                self.assertIsNone(convert_to_date(value))

    def test_large_integers(self):
        """Test that large integers are converted to dates."""
        # The function treats integers as Excel dates, so large integers become future dates
        large_int = 123456
        result = convert_to_date(large_int)
        self.assertIsNotNone(result)
        self.assertIsInstance(result, pd.Timestamp)


if __name__ == "__main__":
    unittest.main()
