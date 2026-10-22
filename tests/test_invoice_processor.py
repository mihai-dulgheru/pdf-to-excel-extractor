import unittest

from modules.invoice_processor import InvoiceProcessor


class _FakePage:
    """Minimal stand-in for a pdfplumber page: whole page text is also every bbox's text."""

    def __init__(self, text):
        self.text = text

    def extract_text(self):
        return self.text

    def within_bbox(self, _coords):
        return self


class _FakePdf:
    def __init__(self, *page_texts):
        self.pages = [_FakePage(t) for t in page_texts]


def _nc8(text, is_credit_note=False):
    pdf = _FakePdf(text)
    return InvoiceProcessor._extract_nc8_codes(pdf, pdf.pages[0], 595, 842, is_credit_note, "RO")


ITEM_WRAPPED_3_LINES = """\
D_FB_CV SSS RH VALC CL HE RS-LN SL
0190 Part number: 378142703SLT 10PC104140.00 / 1000 PC EUR
1,041.40
0.00
Customer Part. Nb.: 378142703SLT
Purch. order no. : 5500850400
Commodity Code : 87082990
Country of origin : RO
.
D_FB_CV SSS RH VALC CL HE RS-LN SL-ER
0191 Part number: 378142703THC 5PC104140.00 / 1000 PC EUR
520.70 0.00
Customer Part. Nb.: 378142703THC
Purch. order no. : 5500850398
Commodity Code : 87082990
Country of origin : RO
"""

CREDIT_NOTE_WITH_INTERNAL_ORDER = """\
Pos. Details/Description Delivery Price per Total Total VAT
Quantity Unit net value amount
INTERNAL ORDER 711594219
REFERENCE
Y_RC40_CV KO COUPE LLL KR GLATT
0010 Part number: 221442507B10 1PC 62263.24/1000PC EUR 62,26 0,00
Customer Part. Nb.: 221442507B10
Purch. order no. : 4505260714/4505260699/4505260
Commodity Code : 94019920
Country of origin : RO
.
Y_2S BEZ KOM GLAT COUPE L SL
0040 Part number: 292715203Y8R 1PC 84037.14/1000PC EUR 84,04 0,00
Customer Part. Nb.: 292715203Y8R
Purch. order no. : 4505260714/4505260699/4505260
Commodity Code : 39269097
Country of origin : RO
"""


class TestExtractNc8Codes(unittest.TestCase):
    def test_amount_wrapped_over_three_lines_is_not_lost(self):
        """pdfplumber may split 'EUR <net> <vat>' as 'EUR' / '<net>' / '<vat>'; the net value must still count."""
        result = _nc8(ITEM_WRAPPED_3_LINES)
        self.assertEqual([r[0] for r in result], ["87082990"])
        self.assertAlmostEqual(result[0][1], 1041.40 + 520.70, places=2)

    def test_credit_note_is_split_per_code_with_negative_values(self):
        """A credit note with item lines gives one row per NC8 code, values negated, even with an INTERNAL ORDER header."""
        result = _nc8(CREDIT_NOTE_WITH_INTERNAL_ORDER, is_credit_note=True)
        self.assertEqual([(r[0], round(r[1], 2), r[2]) for r in result],
                         [("94019920", -62.26, "RO"), ("39269097", -84.04, "RO")])

    def test_space_grouped_thousands_amount(self):
        """Some layouts print thousands with a space: 'EUR 1 062,26 0,00' must give 1062.26, not 1."""
        text = CREDIT_NOTE_WITH_INTERNAL_ORDER.replace("EUR 62,26 0,00", "EUR 1 062,26 0,00")
        result = _nc8(text, is_credit_note=True)
        self.assertEqual([(r[0], round(r[1], 2)) for r in result], [("94019920", -1062.26), ("39269097", -84.04)])

    def test_credit_note_without_items_falls_back_to_label(self):
        result = _nc8("Credit Note\nno item lines here", is_credit_note=True)
        self.assertEqual(result, [("Credit Note", 0, "RO")])

    def test_invoice_with_internal_order_header_keeps_single_row(self):
        """Regression guard: regular invoices keep the existing INTERNAL ORDER rule."""
        result = _nc8(CREDIT_NOTE_WITH_INTERNAL_ORDER, is_credit_note=False)
        self.assertEqual(result, [("INTERNAL ORDER", 0, "RO")])



class TestProcessInvoicesProblems(unittest.TestCase):
    def test_missing_exchange_rate_aborts_the_run_with_a_message(self):
        from unittest.mock import patch
        from functions import ExchangeRateUnavailable

        def boom(_path):
            raise ExchangeRateUnavailable("Cursul BNR pentru 03.03.2025 nu a putut fi obținut")

        with patch.object(InvoiceProcessor, "_process_single_invoice", staticmethod(boom)):
            processor = InvoiceProcessor(["a.pdf", "b.pdf"])
            processor.process_invoices()
        self.assertIn("Cursul BNR", processor.fatal_error)
        self.assertTrue(processor.df.empty)

    def test_unreadable_pdf_is_listed_as_a_problem(self):
        processor = InvoiceProcessor([r"C:\does\not\exist\missing.pdf"])
        processor.process_invoices()
        self.assertIsNone(processor.fatal_error)
        self.assertEqual(len(processor.problems), 1)
        self.assertIn("missing.pdf", processor.problems[0])
        self.assertTrue(processor.df.empty)

if __name__ == "__main__":
    unittest.main()
