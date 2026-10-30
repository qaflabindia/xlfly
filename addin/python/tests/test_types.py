import datetime as dt
import math
import unittest

from xlfly._types import XLError, from_excel_date, from_wire, to_excel_date, to_wire, unwrap_single


class WireTests(unittest.TestCase):
    def test_scalars(self):
        self.assertEqual(to_wire(3), 3)
        self.assertEqual(to_wire(2.5), 2.5)
        self.assertEqual(to_wire("x"), "x")
        self.assertIs(to_wire(True), True)
        self.assertIsNone(to_wire(None))
        self.assertEqual(to_wire(math.nan), {"$err": "#NUM!"})
        self.assertEqual(to_wire(XLError("#N/A")), {"$err": "#N/A"})

    def test_sequences(self):
        self.assertEqual(to_wire([1, 2]), [[1], [2]])
        self.assertEqual(to_wire([[1], [2, 3]]), [[1, None], [2, 3]])
        self.assertEqual(to_wire([]), [[None]])

    def test_errors_round_trip(self):
        self.assertEqual(from_wire([[{"$err": "#DIV/0!"}, 1]]), [[XLError("#DIV/0!"), 1]])
        with self.assertRaises(ValueError):
            XLError("#BOGUS")

    def test_dates(self):
        self.assertEqual(to_excel_date(dt.date(2024, 1, 1)), 45292)
        self.assertEqual(to_wire(dt.datetime(2024, 1, 1, 12)), 45292.5)
        self.assertEqual(from_excel_date(45292.5), dt.datetime(2024, 1, 1, 12))

    def test_unwrap(self):
        self.assertEqual(unwrap_single([[5]]), 5)
        self.assertEqual(unwrap_single([[5, 6]]), [[5, 6]])


if __name__ == "__main__":
    unittest.main()
