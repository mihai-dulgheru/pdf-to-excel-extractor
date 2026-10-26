import os
import tempfile
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


GERMAN_SECTION_1 = """\
FAURECIA UST-IDNR : RO15599111
Bitte angeben
Gutschrift
ZEICHEN DATUM Kunden Nr Lieferant
9610178330 10.03.25 C032000000 74773"""

GERMAN_SECTION_2 = """\
Bestellt von : C032000000
DR. ING. H. C. F. PORSCHE AG
FFF3 KREDITORENBUCHHALTUNG
PORSCHEPLATZ 1
D-70435 STUTTGART
Incoterm : FCA -"""

GERMAN_SECTION_3 = """\
Berechnet am : C032000000
Fakturiert zu : C032000000
KUNDEN UST-IDNR : DE147799625
DR. ING. H. C. F. PORSCHE AG
Bankverbindung
BRD - Groupe Société Générale BRDEROBU"""

CODE_MISSING_ON_ONE_ITEM = """\
0010 Part number: A 1PC 1.00 / 1 PC EUR 10.00 0.00
Purch. order no. : 1
Commodity Code : 94019920
Country of origin : RO
.
0020 Part number: B 1PC 1.00 / 1 PC EUR 20.00 0.00
Purch. order no. : 1
Commodity Code :
Country of origin : RO
.
0030 Part number: C 1PC 1.00 / 1 PC EUR 30.00 0.00
Purch. order no. : 1
Commodity Code : 94019920
Country of origin : RO
"""


class TestGermanLayout(unittest.TestCase):
    def test_gutschrift_is_a_credit_note(self):
        self.assertTrue(InvoiceProcessor._is_credit_note(GERMAN_SECTION_1))
        self.assertTrue(InvoiceProcessor._is_credit_note("Original\nCredit Note\n9610210908 06.10.2026"))
        self.assertFalse(InvoiceProcessor._is_credit_note("Original\nInvoice\n9610201631 23.02.2026"))

    def test_german_field_labels_are_read(self):
        import re
        self.assertEqual(InvoiceProcessor._extract_company(GERMAN_SECTION_2), "DR. ING. H. C. F. PORSCHE AG")
        self.assertEqual(InvoiceProcessor._extract_field(GERMAN_SECTION_3, InvoiceProcessor.VAT_PATTERN, "Unknown",
                                                         re.IGNORECASE), "DE147799625")
        self.assertEqual(
            InvoiceProcessor._extract_field(GERMAN_SECTION_2, InvoiceProcessor.INCOTERMS_PATTERN, "Unknown",
                                            re.IGNORECASE), "FCA")
        destination = InvoiceProcessor._extract_field(GERMAN_SECTION_3, InvoiceProcessor.DESTINATION_PATTERN, "Unknown",
                                                      re.DOTALL)
        self.assertIn("DE147799625", destination)


class TestNetWeightLocale(unittest.TestCase):
    def _weight(self, text):
        return InvoiceProcessor._extract_net_weight(_FakePage(text), False)

    def test_dot_is_a_thousands_separator_when_the_gross_weight_says_so(self):
        self.assertEqual(self._weight("Total Gross Weight 4.218,300 KG\nTotal Net weight 1.035 KG"), 1035)

    def test_english_and_small_romanian_weights_are_unchanged(self):
        self.assertEqual(self._weight("Gross Weight 2,822.622 KG\nNet weight 1,243.382 KG"), 1243)
        self.assertEqual(self._weight("Gross Weight 20,088 KG\nNet weight 9,208 KG"), 9)
        self.assertEqual(self._weight("Net weight 144.745 KG"), 145)


class TestItemsWithoutCode(unittest.TestCase):
    def test_with_one_real_code_the_codeless_items_join_it(self):
        """What the customer did by hand in all 11 March 2025 cases: one row with the invoice's only code."""
        result = _nc8(CODE_MISSING_ON_ONE_ITEM)
        self.assertEqual([(c, round(v, 2), o) for c, v, o in result], [("94019920", 60.0, "RO")])

    def test_with_several_codes_the_codeless_items_stay_visible_under_a_dash(self):
        text = CODE_MISSING_ON_ONE_ITEM.replace(
            "0030 Part number: C 1PC 1.00 / 1 PC EUR 30.00 0.00\nPurch. order no. : 1\nCommodity Code : 94019920",
            "0030 Part number: C 1PC 1.00 / 1 PC EUR 30.00 0.00\nPurch. order no. : 1\nCommodity Code : 87082990")
        result = _nc8(text)
        self.assertEqual([(c, round(v, 2)) for c, v, _ in result],
                         [("94019920", 10.0), ("-", 20.0), ("87082990", 30.0)])


class TestDuplicateInvoicesInOneRun(unittest.TestCase):
    def test_same_invoice_from_two_files_is_kept_once_and_reported(self):
        from unittest.mock import patch

        def rows(path):
            return [{"company": "X", "invoice_number": "9610182120", "nc8_code": "94019920", "origin": "RO",
                     "destination": "SK", "invoice_value_eur": 143771.13, "net_weight": 7819,
                     "shipment_date": "11.04.2025", "exchange_rate": 5.0, "value_ron": 718855.65,
                     "vat_number": "SK2120249945", "delivery_location": 1759, "delivery_condition": "FCA"}]

        with patch.object(InvoiceProcessor, "_process_single_invoice", staticmethod(rows)):
            processor = InvoiceProcessor([r"C:\a\inv.pdf", r"C:\b\inv_copy.pdf"])
            processor.process_invoices()
        self.assertEqual(len(processor.df), 1)
        self.assertEqual(len(processor.problems), 1)
        self.assertIn("9610182120", processor.problems[0])
        self.assertIn("inv.pdf", processor.problems[0])
        self.assertIn("inv_copy.pdf", processor.problems[0])


GARBLED_CURRENCY_TOKENS = """\
CAR 13 KO GLATT SCHWAR
0010 Part number: 99P019010106 48PC 62263.24 / 1000 PECUR 2.988,6 4 0,00
Customer Part. Nb.: 2214425YAAG06
Purch. order no. : 5500632800
Commodity Code : 94019920
Country of origin : RO
.
CAV ST SP CL GLATT SCHWARZ SCHWARZ ORANG
0020 Part number: 99ETSO379011012 1PC 90.35 / 1 PCEUR 90,3 5 0,00
Customer Part. Nb.: 2258446YQQG12
Purch. order no. : 4504190099
Commodity Code : 87082990
Country of origin : RO
"""


class TestGarbledAmounts(unittest.TestCase):
    def test_pc_eur_glued_into_pecur_or_pceur_with_detached_last_digit_is_read(self):
        """FAURECIA SLOVAKIA layout: '1000 PECUR 2.988,6 4 0,00' means EUR 2.988,64."""
        result = _nc8(GARBLED_CURRENCY_TOKENS)
        self.assertEqual([(c, round(v, 2)) for c, v, _ in result], [("94019920", 2988.64), ("87082990", 90.35)])

    def test_unparsed_items_collapse_to_one_row_instead_of_zero_rows(self):
        collapsed = InvoiceProcessor._collapse_unparsed([("94019920", 0.0, "RO"), ("-", 0.0, "RO")], 148107.23)
        self.assertEqual(collapsed, [("94019920; -", 0, "RO")])
        kept = [("94019920", 100.0, "RO"), ("-", 20.0, "RO")]
        self.assertEqual(InvoiceProcessor._collapse_unparsed(kept, 120.0), kept)
        zero_invoice = [("94019920", 0.0, "RO")]
        self.assertEqual(InvoiceProcessor._collapse_unparsed(zero_invoice, 0.0), zero_invoice)


class TestUnreadableFilesAreNotDuplicates(unittest.TestCase):
    def test_two_unreadable_files_both_stay_and_no_duplicate_is_reported(self):
        from unittest.mock import patch

        def unknown_row(path):
            return [{"company": "Unknown", "invoice_number": "Unknown", "nc8_code": "Unknown", "origin": "-",
                     "destination": "Unknown", "invoice_value_eur": 0, "net_weight": 0, "shipment_date": "01.01.2026",
                     "exchange_rate": 5.0, "value_ron": 0, "vat_number": "Unknown", "delivery_location": 2093,
                     "delivery_condition": "Unknown"}]

        with patch.object(InvoiceProcessor, "_process_single_invoice", staticmethod(unknown_row)):
            processor = InvoiceProcessor([r"C:\x\scan1.pdf", r"C:\x\scan2.pdf"])
            processor.process_invoices()
        self.assertEqual(len(processor.df), 2)
        self.assertEqual(processor.problems, [])


def _blank_pdf(path):
    """A valid one-page PDF with no text at all (what a scanned document looks like to the parser)."""
    objects = [b"<< /Type /Catalog /Pages 2 0 R >>", b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
               b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] >>"]
    out, offsets = bytearray(b"%PDF-1.4\n"), []
    for i, obj in enumerate(objects, 1):
        offsets.append(len(out))
        out += f"{i} 0 obj\n".encode() + obj + b"\nendobj\n"
    xref = len(out)
    out += f"xref\n0 {len(objects) + 1}\n0000000000 65535 f \n".encode()
    for off in offsets:
        out += f"{off:010d} 00000 n \n".encode()
    out += f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode()
    with open(path, "wb") as f:
        f.write(out)


class TestScannedPdf(unittest.TestCase):
    def test_pdf_without_text_is_reported_and_produces_no_row(self):
        with tempfile.TemporaryDirectory() as td:
            path = os.path.join(td, "scan.pdf")
            _blank_pdf(path)
            processor = InvoiceProcessor([path])
            processor.process_invoices()
        self.assertTrue(processor.df.empty)
        self.assertEqual(len(processor.problems), 1)
        self.assertIn("scanat", processor.problems[0])


class TestWeightRounding(unittest.TestCase):
    def test_rounded_weights_add_up_to_the_invoice_weight(self):
        """Largest remainder: whole kilograms first, the leftover ones to the biggest fractions (ties: first code)."""
        thirds = InvoiceProcessor._calculate_proportional_weights([("a", 1.0), ("b", 1.0), ("c", 1.0)], 100)
        self.assertEqual(sum(w for _, w in thirds), 100)
        self.assertEqual(thirds[0], ("a", 34))
        jlr = InvoiceProcessor._calculate_proportional_weights(
            [("39269097", 58461.02), ("87082990", 31134.42), ("94019920", 446.78)], 1244)
        self.assertEqual(sum(w for _, w in jlr), 1244)
        self.assertEqual([w for _, w in jlr][1:], [430, 6])

    def test_no_row_goes_negative_and_credit_note_signs_do_not_matter(self):
        five = InvoiceProcessor._calculate_proportional_weights([(c, 1.0) for c in "abcde"], 3)
        self.assertEqual(sum(w for _, w in five), 3)
        self.assertTrue(all(w >= 0 for _, w in five), five)
        negative = InvoiceProcessor._calculate_proportional_weights([("a", -2.0), ("b", -1.0)], 3)
        self.assertEqual(negative, [("a", 2), ("b", 1)])


class TestCompanyName(unittest.TestCase):
    def test_trailing_comma_is_dropped(self):
        text = "ORDERED BY : 1359000000\nSASAL - ASSENTOS PARA AUTOMOVEIS,\nZONA INDUSTRIAL MONTE CAVALO\nP-3670-273 VOUZELA"
        self.assertEqual(InvoiceProcessor._extract_company(text), "SASAL - ASSENTOS PARA AUTOMOVEIS")
        self.assertEqual(InvoiceProcessor._extract_company("ORDERED BY : 1\nACME GMBH ,\nstreet"), "ACME GMBH")


if __name__ == "__main__":
    unittest.main()
