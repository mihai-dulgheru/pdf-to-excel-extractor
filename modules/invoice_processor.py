import locale
import re
from datetime import datetime

import pandas as pd
import pdfplumber

from config import Constants
from functions.get_bnr_exchange_rate import get_bnr_exchange_rate
from functions.get_country_code_from_address import get_country_code_from_address
from functions.get_delivery_location import get_delivery_location


class InvoiceProcessor:
    """
    Extracts invoice data from PDF files and stores it in a pandas DataFrame.
    """
    __slots__ = ("input_paths", "df", "progress_callback")

    def __init__(self, paths, progress_callback=None):
        """
        Initialize the processor with PDF file paths and an optional progress callback.

        :param paths: List[str]
            A list of paths to the input PDF files.
        :param progress_callback: Callable or None
            A callable to report progress (processed, total) or None.
        """
        self.input_paths = paths
        self.df = pd.DataFrame(columns=Constants.COLUMNS)
        self.progress_callback = progress_callback

    def process_invoices(self):
        """
        Extract and aggregate data from each PDF into the DataFrame.
        Calls the progress callback after processing each file, if provided.
        """
        total_invoices = len(self.input_paths)
        for idx, pdf_path in enumerate(self.input_paths):
            with pdfplumber.open(pdf_path) as pdf:
                first_page = pdf.pages[0]
                page_width, page_height = first_page.width, first_page.height

                section_2_text = self._extract_section_text(first_page, "section_2", page_width, page_height)
                company = self._extract_company(section_2_text)
                print(f"[LOG] Company: {company}")

                section_1_text = self._extract_section_text(first_page, "section_1", page_width, page_height)
                invoice_number = self._extract_invoice_number(section_1_text)
                print(f"[LOG] Invoice Number: {invoice_number}")

                nc8_codes = self._extract_nc8_codes(pdf, first_page, page_width, page_height)
                print(f"[LOG] NC8 Codes: {nc8_codes}")

                section_4_text = self._extract_section_text(first_page, "section_4", page_width, page_height)
                origin = self._extract_origin(section_4_text)
                print(f"[LOG] Origin: {origin}")

                section_3_text = self._extract_section_text(first_page, "section_3", page_width, page_height)
                destination = get_country_code_from_address(
                    self._extract_field(section_3_text, r"Invoiced to\s*:\s*(.+?)\nCredit transfer", "Unknown",
                                        re.DOTALL))
                print(f"[LOG] Destination: {destination}")

                is_credit_note = "CREDIT NOTE" in section_1_text.upper()
                print(f"[LOG] Is Credit Note: {is_credit_note}")

                invoice_value_eur, invoice_value_ron, currency = self._extract_invoice_values(pdf.pages[-1], page_width,
                                                                                              page_height,
                                                                                              is_credit_note)
                print(f"[LOG] Invoice Value EUR: {invoice_value_eur}")
                print(f"[LOG] Invoice Value RON: {invoice_value_ron}")
                print(f"[LOG] Currency: {currency}")

                is_invoice = "INVOICE" in section_1_text.upper() and "ORIGINAL" not in section_1_text.upper()
                print(f"[LOG] Is Invoice: {is_invoice}")

                net_weight = self._extract_net_weight(pdf.pages[-1], is_invoice or is_credit_note, currency)
                print(f"[LOG] Net Weight: {net_weight}")

                shipment_date = self._extract_shipment_date(section_1_text)
                print(f"[LOG] Shipment Date: {shipment_date.strftime('%d.%m.%Y')}")

                exchange_rate = get_bnr_exchange_rate(shipment_date, "RON")
                print(f"[LOG] Exchange Rate: {exchange_rate}")

                delivery_location = get_delivery_location(section_1_text)
                print(f"[LOG] Delivery Location: {delivery_location}")

                vat_number = self._extract_field(section_3_text, r"Tax number\s*:\s*(\w+)", "Unknown", re.IGNORECASE)
                print(f"[LOG] VAT Number: {vat_number}")

                delivery_condition = self._extract_field(section_2_text, r"Incoterms\s*:\s*(\w+)", "Unknown",
                                                         re.IGNORECASE)
                print(f"[LOG] Delivery Condition: {delivery_condition}")

                self.df.loc[idx] = [company, invoice_number, ", ".join(nc8_codes), origin, destination,
                                    invoice_value_eur, net_weight, shipment_date.strftime("%d.%m.%Y"), exchange_rate,
                                    invoice_value_ron, vat_number, delivery_location, delivery_condition]

            if self.progress_callback:
                self.progress_callback(idx + 1, total_invoices)

    @staticmethod
    def _extract_section_text(page, section_name, page_width, page_height):
        """
        Extract text from a predefined bounding box within the PDF page.

        :param page: pdfplumber.page.Page
            The page object to extract text from.
        :param section_name: str
            The key for the section in Constants.PROPORTIONS.
        :param page_width: float
            The page width.
        :param page_height: float
            The page height.
        :return: str
            The extracted text for the specified section.
        """
        from functions.calculate_coordinates import calculate_coordinates
        section_coords = calculate_coordinates(page_width, page_height, Constants.PROPORTIONS[section_name])
        return page.within_bbox(section_coords).extract_text()

    @staticmethod
    def _extract_company(text):
        """
        Return the company name found after the 'ORDERED BY' line, else 'Unknown'.

        :param text: str
            The text to search for the company name.
        :return: str
            The extracted company name or 'Unknown'.
        """
        if not text:
            return "Unknown"

        lines = text.split("\n")
        for i, line in enumerate(lines):
            if "ORDERED BY" in line and i + 1 < len(lines):
                return lines[i + 1].strip()

        return "Unknown"

    @staticmethod
    def _extract_field(text, pattern, default, flags=0):
        """
        Extract a single value from text using a regex pattern.

        :param text: str
            The text to search.
        :param pattern: str
            The regex pattern to find the desired value.
        :param default: str
            The default value if no match is found.
        :param flags: int
            Optional regex flags.
        :return: str
            The extracted value or the default if not found.
        """
        if not text:
            return default
        match = re.search(pattern, text, flags)
        return match.group(1).strip() if match else default

    @staticmethod
    def _extract_invoice_number(section_1_text):
        """
        Extract the invoice number from the last line of the given section text.

        :param section_1_text: str
            The text containing invoice number info.
        :return: str
            The extracted invoice number or 'Unknown'.
        """
        if not section_1_text:
            return "Unknown"
        lines = section_1_text.split("\n")
        if lines:
            return lines[-1].split(" ")[0]
        return "Unknown"

    @staticmethod
    def _extract_nc8_codes(pdf, first_page, page_width, page_height):
        """
        Extract NC8 commodity codes from all pages of the PDF.
        If both "REFERENCE" and "INTERNAL ORDER" are found in section_4_text,
        return ["REFERENCE; INTERNAL ORDER"] instead.

        :param pdf: pdfplumber.pdf.PDF
            The opened PDF object.
        :param first_page: pdfplumber.page.Page
            The first page object (used to check section_4_text).
        :param page_width: float
            The page width.
        :param page_height: float
            The page height.
        :return: list[str]
            A list of NC8 codes or a single-item list with "REFERENCE; INTERNAL ORDER".
        """
        section_4_text = InvoiceProcessor._extract_section_text(first_page, "section_4", page_width, page_height)

        if section_4_text and ("REFERENCE" in section_4_text and "INTERNAL ORDER" in section_4_text):
            return ["REFERENCE; INTERNAL ORDER"]

        codes = []
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                codes += re.findall(r"Commodity Code\s*:\s*(\d+)", text)

        return codes if codes else ["Unknown"]

    @staticmethod
    def _extract_origin(section_4_text):
        """
        Extract the two-letter country code from the 'Country of origin' line,
        or return '-' if not found.

        :param section_4_text: str
            The text extracted from section 4.
        :return: str
            The extracted country code or '-'.
        """
        if not section_4_text:
            return "-"
        match = re.search(r"Country of origin\s*:\s*([A-Z]{2})", section_4_text, re.IGNORECASE)
        return match.group(1).upper() if match else "-"

    @staticmethod
    def _extract_invoice_values(last_page, page_width, page_height, is_credit_note):
        """
        Extract the invoice values (EUR, RON) and the currency from a bounding box on the last page.

        :param last_page: pdfplumber.page.Page
            The last page object.
        :param page_width: float
            The page width.
        :param page_height: float
            The page height.
        :param is_credit_note: bool
            True if the document is identified as a credit note, False otherwise.
        :return: tuple
            (invoice_value_eur, invoice_value_ron, currency) where currency is str or None.
        """
        from functions.calculate_coordinates import calculate_coordinates
        bbox_ratio = (0.56, 0.88, 0.94, 0.94)
        bbox_coords = calculate_coordinates(page_width, page_height, bbox_ratio)
        bbox_text = last_page.within_bbox(bbox_coords).extract_text()

        if not bbox_text:
            return 0, 0, None

        lines = bbox_text.splitlines()
        invoice_value_eur = 0
        invoice_value_ron = 0
        currency = None

        for line in lines:
            if "*" in line:
                parts = line.split("*")
                if len(parts) >= 2:
                    currency_detected = parts[0].strip().lower()
                    raw_value = parts[-1].strip()
                    cleaned_value = raw_value.replace(".", "").replace(",", "")

                    try:
                        if len(cleaned_value) > 2:
                            numeric_value = float(cleaned_value[:-2] + "." + cleaned_value[-2:])
                        else:
                            numeric_value = float("0." + cleaned_value.zfill(2))
                    except ValueError:
                        continue

                    if is_credit_note:
                        numeric_value = -numeric_value

                    if currency_detected == "eur":
                        invoice_value_eur = numeric_value
                        currency = "EUR"
                    elif currency_detected == "ron":
                        invoice_value_ron = numeric_value
                        currency = "RON"

        return invoice_value_eur, invoice_value_ron, currency

    @staticmethod
    def _extract_net_weight(last_page, is_invoice_or_credit_note, currency):
        """
        Extract the net weight from the last page if the document is neither
        an invoice nor a credit note. Otherwise, return 0.

        :param last_page: pdfplumber.page.Page
            The page from which net weight is extracted.
        :param is_invoice_or_credit_note: bool
            True if it is an invoice or credit note, False otherwise.
        :param currency: str or None
            The detected currency (e.g., 'EUR' or 'RON').
        :return: int
            The net weight in KG (rounded) or 0.
        """
        if is_invoice_or_credit_note:
            return 0

        text = last_page.extract_text()
        if not text:
            return 0

        locale_code = "ro_RO.UTF-8" if currency == "RON" else "en_US.UTF-8"

        try:
            locale.setlocale(locale.LC_NUMERIC, locale_code)
        except locale.Error:
            print(f"[LOG] Locale {locale_code} not available, using default.")
            locale_code = locale.getdefaultlocale()[0]
            locale.setlocale(locale.LC_NUMERIC, locale_code)

        match = re.search(r"Net weight\s+([\d.,]+)\s+KG", text, re.IGNORECASE)
        if match:
            raw_value = match.group(1)
            try:
                numeric_value = locale.atof(raw_value)
                return round(numeric_value)
            except ValueError:
                return 0

        return 0

    @staticmethod
    def _extract_shipment_date(section_1_text):
        """
        Extract the shipment date (DD.MM.YYYY or DD.MM.YY) from the last line of section_1_text.
        Assumes DD.MM.YY is in the 20YY format if year < 50, else 19YY.

        :param section_1_text: str
            The text from which to extract the shipment date.
        :return: datetime.datetime
            A datetime object representing the extracted shipment date,
            or the current date/time if not found or invalid.
        """
        if not section_1_text:
            return datetime.now()

        lines = section_1_text.split("\n")
        last_line = lines[-1] if lines else ""

        match = re.search(r"(\d{2}\.\d{2}\.\d{4}|\d{2}\.\d{2}\.\d{2})", last_line)
        if match:
            date_str = match.group(1)
            try:
                if len(date_str) == 10:
                    return datetime.strptime(date_str, "%d.%m.%Y")
                elif len(date_str) == 8:
                    year = int(date_str[-2:])
                    year += 2000 if year < 50 else 1900
                    return datetime.strptime(f"{date_str[:6]}{year}", "%d.%m.%Y")
            except ValueError:
                pass

        return datetime.now()
