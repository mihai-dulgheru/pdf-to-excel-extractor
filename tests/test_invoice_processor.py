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


def _nc8(text, is_credit_note=False, default_origin="RO"):
    pdf = _FakePdf(text)
    return InvoiceProcessor._extract_nc8_codes(pdf, pdf.pages[0], 595, 842, is_credit_note, default_origin)


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

MULTI_ORIGIN = """\
0030 Part number: 92V106011001 30M 3.04 / 1 M EUR 91.20 0.00
Customer Part. Nb.: 92V106011001
Purch. order no. : 1
Commodity Code : 94019920
Country of origin : CZ
.
0060 Part number: 92Z001011000 15M 15.33 / 1 M EUR 229.95 0.00
Purch. order no. : 1
Commodity Code : 94019920
Country of origin : RO
.
0090 Part number: 92R312011000 20M 0.15 / 1 M EUR 3.00 0.00
Purch. order no. : 1
Commodity Code : 94019920
Country of origin : CN
.
0040 Part number: 92Z072011000 30M 3.77 / 1 M EUR 113.10 0.00
Purch. order no. : 1
Commodity Code : 94019920
Country of origin : CZ
"""

ORIGIN_ONLY_ON_SECOND_ITEM = """\
0010 Part number: A 1PC 10.00 / 1 PC EUR 10.00 0.00
Purch. order no. : 1
Commodity Code : 94019920
Country of origin :
.
0020 Part number: B 1PC 20.00 / 1 PC EUR 20.00 0.00
Purch. order no. : 1
Commodity Code : 87082990
Country of origin : RO
"""


class TestOriginPerItem(unittest.TestCase):
    def test_same_code_is_split_per_origin_with_subtotals(self):
        """Nae 2025-06-03 / exemplu.xlsx: one row per (code, origin), values summed per row."""
        result = _nc8(MULTI_ORIGIN)
        self.assertEqual([(c, round(v, 2), o) for c, v, o in result],
                         [("94019920", 204.30, "CZ"), ("94019920", 229.95, "RO"), ("94019920", 3.00, "CN")])

    def test_item_without_origin_takes_the_first_origin_printed_in_the_invoice(self):
        result = _nc8(ORIGIN_ONLY_ON_SECOND_ITEM, default_origin="-")
        self.assertEqual([(c, o) for c, _, o in result], [("94019920", "RO"), ("87082990", "RO")])

    def test_invoice_without_any_origin_keeps_the_dash(self):
        text = ORIGIN_ONLY_ON_SECOND_ITEM.replace("Country of origin : RO", "Country of origin :")
        result = _nc8(text, default_origin="-")
        self.assertEqual([o for _, _, o in result], ["-", "-"])

    def test_blank_origin_on_page_one_does_not_steal_letters_from_the_next_line(self):
        section_4 = "Commodity Code : 87089220\nCountry of origin:\noutput tax 0% goods external RO\nSub total 99,21"
        self.assertEqual(InvoiceProcessor._extract_origin(section_4), "-")


PER_LAYOUT_TWO_ORIGINS = """\
DESCRIPTION QUANTITY UNIT PRICE NET PRICE
VW413 DRV PAD LHD CERAMIQUE BW9
2922137BW9 PER 1.000 PC 590PC 16.430,40 9.693,94
S008BU-110
Purch. order no. : 5500004139
Commodity Code : 94019920
Country of origin: RO
VW413 DRV PAD LHD TornadoROT TS9
2922137TS9 PER 1.000 PC 500PC 16.199,40 8.099,70
S008BS-110
Purch. order no. : 5500004141
Commodity Code : 87082990
Country of origin: CZ
VW413 PASS PAD LHD TornadoROT TS9
2922206TS9 PER 1.000 PC 432PC 12.908,20 5.576,34
S008BN-110
Purch. order no. : 5500004144
Commodity Code : 87082990
Country of origin: RO
"""

CODE_INTERLEAVED = """\
0010 Part number: A 1PC 1.00 / 1 PC EUR 10.00 0.00
Purch. order no. : 1
Commodity Code : 94019920
Country of origin : CZ
.
0020 Part number: B 1PC 1.00 / 1 PC EUR 20.00 0.00
Purch. order no. : 1
Commodity Code : 39199080
Country of origin : CZ
.
0030 Part number: C 1PC 1.00 / 1 PC EUR 30.00 0.00
Purch. order no. : 1
Commodity Code : 94019920
Country of origin : RO
"""


class TestOriginPerItemLayouts(unittest.TestCase):
    def test_per_layout_reads_each_items_own_origin(self):
        """BROSE/JLR 'PER' lines: the origin line comes after the Commodity Code line and belongs to that item."""
        result = _nc8(PER_LAYOUT_TWO_ORIGINS)
        self.assertEqual([(c, round(v, 2), o) for c, v, o in result],
                         [("94019920", 9693.94, "RO"), ("87082990", 8099.70, "CZ"), ("87082990", 5576.34, "RO")])

    def test_rows_of_one_code_stay_adjacent(self):
        """exemplu.xlsx lists the origins of one code one under the other."""
        result = _nc8(CODE_INTERLEAVED)
        self.assertEqual([(c, o) for c, _, o in result], [("94019920", "CZ"), ("94019920", "RO"), ("39199080", "CZ")])


if __name__ == "__main__":
    unittest.main()
