import unittest

from functions import round_to_n_decimals


class TestRoundToNDecimals(unittest.TestCase):

    def test_default_rounding(self):
        """Test rounding with default precision (2 decimals)."""
        self.assertEqual(round_to_n_decimals(3.14159), 3.14)
        self.assertEqual(round_to_n_decimals(2.999), 3.0)
        self.assertEqual(round_to_n_decimals(0.001), 0.0)

    def test_custom_precision(self):
        """Test rounding with custom precision."""
        self.assertEqual(round_to_n_decimals(3.14159, 3), 3.142)
        self.assertEqual(round_to_n_decimals(2.999, 1), 3.0)
        self.assertEqual(round_to_n_decimals(0.0001, 4), 0.0001)

    def test_negative_numbers(self):
        """Test rounding negative numbers."""
        self.assertEqual(round_to_n_decimals(-3.14159, 2), -3.14)
        self.assertEqual(round_to_n_decimals(-2.999, 1), -3.0)

    def test_zero_precision(self):
        """Test rounding to zero decimal places (integers)."""
        self.assertEqual(round_to_n_decimals(3.14159, 0), 3)
        self.assertEqual(round_to_n_decimals(2.999, 0), 3)
        self.assertEqual(round_to_n_decimals(-3.14159, 0), -3)


if __name__ == "__main__":
    unittest.main()
