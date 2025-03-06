import concurrent.futures
import locale
import re
from datetime import datetime

import pandas as pd
import pdfplumber

from config import Constants
from functions import calculate_coordinates, get_country_code_from_address, get_bnr_exchange_rate, \
    get_delivery_location, get_previous_workday


class InvoiceProcessor:
    """
    Extracts invoice data from PDF files into a pandas DataFrame.
    """

    __slots__ = ("input_paths", "df", "progress_callback")

    def __init__(self, paths, progress_callback=None):
        """
        :param paths: List of PDF file paths.
        :param progress_callback: Callable or None to report progress.
        """
        self.input_paths = paths
        self.df = pd.DataFrame(columns=Constants.COLUMNS)
        self.progress_callback = progress_callback

    def process_invoices(self):
        """
        Parse each PDF file in parallel, then aggregate results in self.df.
        """
        try:
            total = len(self.input_paths)
            results = []
            processed = 0

            with concurrent.futures.ThreadPoolExecutor() as executor:
                future_map = {executor.submit(self._process_single_invoice, path): path for path in self.input_paths}

                for future in concurrent.futures.as_completed(future_map):
                    path = future_map[future]
                    try:
                        result = future.result()
                        results.append(result)
                    except Exception as e:
                        print(f"[LOG] Error processing {path}: {e}")
                    processed += 1
                    if self.progress_callback:
                        self.progress_callback(processed, total)

            self.df = pd.DataFrame(results, columns=Constants.COLUMNS)
        except Exception as e:
            print(f"[LOG] Error in process_invoices: {e}")

    @staticmethod
    def _process_single_invoice(pdf_path):
        """
        Process a single PDF file and return a dictionary of extracted fields.
        """
        try:
            with pdfplumber.open(pdf_path) as pdf:
                if len(pdf.pages) == 0:
                    raise ValueError("PDF has no pages")

                first_page = pdf.pages[0]
                last_page = pdf.pages[-1] if len(pdf.pages) > 0 else first_page
                page_width, page_height = first_page.width, first_page.height

                s1 = InvoiceProcessor._extract_section_text(first_page, "section_1", page_width, page_height)
                s2 = InvoiceProcessor._extract_section_text(first_page, "section_2", page_width, page_height)
                s3 = InvoiceProcessor._extract_section_text(first_page, "section_3", page_width, page_height)
                s4 = InvoiceProcessor._extract_section_text(first_page, "section_4", page_width, page_height)

                is_invoice = "INVOICE" in s1.upper() and "ORIGINAL" not in s1.upper() if s1 else False
                is_credit_note = "CREDIT NOTE" in s1.upper() if s1 else False

                company = InvoiceProcessor._extract_company(s2)
                invoice_number = InvoiceProcessor._extract_invoice_number(s1)
                nc8_codes = InvoiceProcessor._extract_nc8_codes(pdf, first_page, page_width, page_height,
                                                                is_credit_note)
                origin = InvoiceProcessor._extract_origin(s4)

                destination_field = InvoiceProcessor._extract_field(s3, r"Invoiced to\s*:\s*(.+?)\nCredit transfer",
                                                                    "Unknown", re.DOTALL)
                destination = get_country_code_from_address(destination_field)

                invoice_value_eur, invoice_value_ron, currency = InvoiceProcessor._extract_invoice_values(last_page,
                                                                                                          page_width,
                                                                                                          page_height,
                                                                                                          is_credit_note)
                net_weight = InvoiceProcessor._extract_net_weight(last_page, is_invoice or is_credit_note, currency)
                shipment_date = InvoiceProcessor._extract_shipment_date(s1)
                exchange_rate = get_bnr_exchange_rate(get_previous_workday(shipment_date))
                vat_number = InvoiceProcessor._extract_field(s3, r"Tax number\s*:\s*(\w+)", "Unknown", re.IGNORECASE)
                delivery_location = get_delivery_location(s1)
                delivery_condition = InvoiceProcessor._extract_field(s2, r"Incoterms\s*:\s*(\w+)", "Unknown",
                                                                     re.IGNORECASE)

                return {"company": company, "invoice_number": invoice_number, "nc8_code": ", ".join(nc8_codes),
                        "origin": origin, "destination": destination, "invoice_value_eur": invoice_value_eur,
                        "net_weight": net_weight, "shipment_date": shipment_date.strftime("%d.%m.%Y"),
                        "exchange_rate": exchange_rate, "value_ron": invoice_value_ron, "vat_number": vat_number,
                        "delivery_location": delivery_location, "delivery_condition": delivery_condition}
        except Exception as e:
            print(f"[LOG] Error processing PDF {pdf_path}: {e}")

    @staticmethod
    def _extract_section_text(page, section_name, page_width, page_height):
        """
        Extract text from a predefined bounding box on the page.
        """
        try:
            coords = calculate_coordinates(page_width, page_height, Constants.PROPORTIONS[section_name])
            text = page.within_bbox(coords).extract_text()
            return text if text else ""
        except Exception as e:
            print(f"[LOG] Error extracting section {section_name}: {e}")
            return ""

    @staticmethod
    def _extract_company(text):
        """
        Return the company name found after the 'ORDERED BY' line, otherwise 'Unknown'.
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
        """
        if not text:
            return default
        match = re.search(pattern, text, flags)
        return match.group(1).strip() if match else default

    @staticmethod
    def _extract_invoice_number(section_1_text):
        """
        Extract the invoice number from the last line of section_1_text.
        """
        if not section_1_text:
            return "Unknown"

        lines = section_1_text.split("\n")
        if not lines:
            return "Unknown"

        try:
            last_line = lines[-1]
            parts = last_line.split(" ")
            return parts[0] if parts else "Unknown"
        except (IndexError, Exception) as e:
            print(f"[LOG] Error extracting invoice number: {e}")
            return "Unknown"

    @staticmethod
    def _extract_nc8_codes(pdf, first_page, page_width, page_height, is_credit_note):
        """
        Extract NC8 codes from all pages.
        If it's a credit note, return ["Credit Note"].
        If 'REFERENCE' and 'INTERNAL ORDER' in section_4_text, return that as a single code.
        """
        if is_credit_note:
            return ["Credit Note"]

        s4 = InvoiceProcessor._extract_section_text(first_page, "section_4", page_width, page_height)
        if s4 and "REFERENCE" in s4 and "INTERNAL ORDER" in s4:
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
        Extract the country code from 'Country of origin', or '-' if not found.
        """
        if not section_4_text:
            return "-"
        match = re.search(r"Country of origin\s*:\s*([A-Z]{2})", section_4_text, re.IGNORECASE)
        return match.group(1).upper() if match else "-"

    @staticmethod
    def _extract_invoice_values(last_page, page_width, page_height, is_credit_note):
        """
        Extract invoice values (EUR, RON) and currency from the last page bounding box.
        """
        try:
            box = (0.52, 0.88, 0.94, 0.94)
            coords = calculate_coordinates(page_width, page_height, box)
            text = last_page.within_bbox(coords).extract_text()
            if not text:
                return 0, 0, None

            lines = [l.strip().replace('*', ' ') for l in text.splitlines() if any(cur in l for cur in ["EUR", "RON"])]
            if not lines:
                return 0, 0, None

            normalized_line = lines[0].replace(" ", "").replace(",", "").replace(".", "")
            currency_str = None
            amount_str = None

            for currency in ["EUR", "RON"]:
                if currency in normalized_line:
                    currency_str = currency.lower()
                    parts = normalized_line.split(currency, 1)
                    if len(parts) > 1:
                        amount_str = parts[1]
                    break

            if not currency_str or not amount_str:
                return 0, 0, None

            try:
                amount = float(amount_str[:-2] + "." + amount_str[-2:]) if len(amount_str) > 2 else float(
                    "0." + amount_str.zfill(2))
            except (ValueError, IndexError):
                return 0, 0, None

            if is_credit_note:
                amount = -amount

            if currency_str == "eur":
                return amount, 0, "EUR"
            elif currency_str == "ron":
                return 0, amount, "RON"

            return 0, 0, None
        except Exception as e:
            print(f"[LOG] Error extracting invoice values: {e}")
            return 0, 0, None

    @staticmethod
    def _extract_net_weight(last_page, is_invoice_or_credit_note, currency):
        """
        Extract net weight if not an invoice or credit note.
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
            try:
                default_loc = locale.getdefaultlocale()[0]
                locale.setlocale(locale.LC_NUMERIC, default_loc)
            except (locale.Error, TypeError):
                pass

        match = re.search(r"Net weight\s+([\d.,]+)\s+KG", text, re.IGNORECASE)
        if match:
            raw_val = match.group(1)
            try:
                num_val = locale.atof(raw_val)
                return round(num_val)
            except ValueError:
                return 0
        return 0

    @staticmethod
    def _extract_shipment_date(section_1_text):
        """
        Extract shipment date (DD.MM.YYYY or DD.MM.YY) from section_1_text.
        """
        try:
            if not section_1_text:
                return datetime.now()

            lines = section_1_text.split("\n")
            if not lines:
                return datetime.now()

            last_line = lines[-1]
            match = re.search(r"(\d{2}\.\d{2}\.\d{4}|\d{2}\.\d{2}\.\d{2})", last_line)
            if match:
                date_str = match.group(1)
                try:
                    if len(date_str) == 10:
                        return datetime.strptime(date_str, "%d.%m.%Y")
                    elif len(date_str) == 8:
                        yr = int(date_str[-2:])
                        yr += 2000 if yr < 50 else 1900
                        return datetime.strptime(f"{date_str[:6]}{yr}", "%d.%m.%Y")
                except ValueError:
                    pass
            return datetime.now()
        except Exception as e:
            print(f"[LOG] Error extracting shipment date: {e}")
            return datetime.now()
