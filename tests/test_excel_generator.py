import os
import tempfile
import unittest

import openpyxl
import pandas as pd

from config import Constants
from modules.excel_generator import ExcelGenerator


def _row(inv, code, origin="RO", eur=100.0, weight=10, vat="SK1", date="03.03.2025"):
    return {"company": "X", "invoice_number": inv, "nc8_code": code, "origin": origin, "destination": "SK",
            "invoice_value_eur": eur, "net_weight": weight, "shipment_date": date, "exchange_rate": 5.0,
            "value_ron": eur * 5, "vat_number": vat, "delivery_location": 1759, "delivery_condition": "FCA"}


def _df(*rows):
    return pd.DataFrame(list(rows), columns=Constants.COLUMNS)


def _data_rows(path):
    ws = openpyxl.load_workbook(path).active
    header = [c.value for c in ws[1]]
    rows = [dict(zip(header, r)) for r in ws.iter_rows(min_row=2, values_only=True)]
    return [r for r in rows if r["Nr Factura Marfa"] is not None]


class TestNrCrt(unittest.TestCase):
    def test_rows_of_one_invoice_share_one_number(self):
        """Customer convention: Nr Crt counts invoices inside a VAT group, not rows."""
        with tempfile.TemporaryDirectory() as td:
            out = os.path.join(td, "fresh.xlsx")
            ExcelGenerator(_df(_row("1001", "94019920"), _row("1001", "87082990"), _row("1002", "94019920"))) \
                .generate_excel(out)
            self.assertEqual([r["Nr Crt"] for r in _data_rows(out)], [1, 1, 2])


class TestAppendPerOrigin(unittest.TestCase):
    def test_new_invoice_with_two_origins_is_appended_whole_in_extraction_order(self):
        with tempfile.TemporaryDirectory() as td:
            base, out = os.path.join(td, "base.xlsx"), os.path.join(td, "out.xlsx")
            ExcelGenerator(_df(_row("1001", "94019920", "RO"), _row("1002", "94019920", "RO"))).generate_excel(base)
            ExcelGenerator(_df(_row("1003", "94019920", "RO"),
                               _row("1003", "94019920", "CZ", eur=50.0, weight=5))).generate_excel(out, base)
            rows = _data_rows(out)
            self.assertEqual([(r["Nr Factura Marfa"], r["Origine"], r["Nr Crt"]) for r in rows],
                             [(1001, "RO", 1), (1002, "RO", 2), (1003, "RO", 3), (1003, "CZ", 3)])

    def test_invoice_already_in_the_workbook_is_skipped_whole_even_if_its_rows_changed(self):
        """Re-running a PDF that an older version already wrote must not add origin rows next to the old ones."""
        with tempfile.TemporaryDirectory() as td:
            base, out = os.path.join(td, "base.xlsx"), os.path.join(td, "out.xlsx")
            ExcelGenerator(_df(_row("1001", "94019920", "RO"), _row("1002", "94019920", "RO"))).generate_excel(base)
            ExcelGenerator(_df(_row("1001", "94019920", "RO", eur=60.0),
                               _row("1001", "94019920", "CZ", eur=40.0),
                               _row("1003", "87082990", "RO"))).generate_excel(out, base)
            rows = _data_rows(out)
            self.assertEqual([(r["Nr Factura Marfa"], r["Origine"], r["Val Fact Euro"]) for r in rows],
                             [(1001, "RO", 100.0), (1002, "RO", 100.0), (1003, "RO", 100.0)])


if __name__ == "__main__":
    unittest.main()
