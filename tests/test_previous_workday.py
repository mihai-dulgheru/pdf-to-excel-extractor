import unittest
from datetime import datetime

from functions import get_previous_workday


class TestPreviousWorkday(unittest.TestCase):

    def test_regular_weekday(self):
        """Test a regular weekday (should return the previous day)."""
        date = datetime.strptime("2025-01-09", "%Y-%m-%d")  # Joi
        expected = datetime.strptime("2025-01-08", "%Y-%m-%d")  # Miercuri
        self.assertEqual(get_previous_workday(date), expected)

    def test_monday_returns_friday(self):
        """Test if a Monday returns the previous Friday."""
        date = datetime.strptime("2025-01-06", "%Y-%m-%d")  # Luni
        expected = datetime.strptime("2025-01-03", "%Y-%m-%d")  # Vineri
        self.assertEqual(get_previous_workday(date), expected)

    def test_saturday_returns_friday(self):
        """Test if a Saturday returns the previous Friday."""
        date = datetime.strptime("2025-01-04", "%Y-%m-%d")  # Sâmbătă
        expected = datetime.strptime("2025-01-03", "%Y-%m-%d")  # Vineri
        self.assertEqual(get_previous_workday(date), expected)

    def test_sunday_returns_friday(self):
        """Test if a Sunday returns the previous Friday."""
        date = datetime.strptime("2025-01-05", "%Y-%m-%d")  # Duminică
        expected = datetime.strptime("2025-01-03", "%Y-%m-%d")  # Vineri
        self.assertEqual(get_previous_workday(date), expected)


if __name__ == "__main__":
    unittest.main()
