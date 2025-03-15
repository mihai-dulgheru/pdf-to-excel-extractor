import concurrent.futures
import re
from collections import defaultdict
from datetime import datetime

import pandas as pd
import pdfplumber

from config import Constants
from functions import calculate_coordinates, get_country_code_from_address, get_bnr_exchange_rate, \
    get_delivery_location, get_previous_workday, parse_mixed_number


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
                        results.extend(result)
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

                is_credit_note = "CREDIT NOTE" in s1.upper() if s1 else False
                is_debit_note = "DEBIT NOTE" in s1.upper() if s1 else False

                company = InvoiceProcessor._extract_company(s2)
                invoice_number = InvoiceProcessor._extract_invoice_number(s1)
                nc8_data = InvoiceProcessor._extract_nc8_codes(pdf, first_page, page_width, page_height, is_credit_note)
                origin = InvoiceProcessor._extract_origin(s4)

                destination_field = InvoiceProcessor._extract_field(s3, r"Invoiced to\s*:\s*(.+?)\nCredit transfer",
                                                                    "Unknown", re.DOTALL)
                destination = get_country_code_from_address(destination_field)

                invoice_value_eur, invoice_value_ron, currency = InvoiceProcessor._extract_invoice_values(last_page,
                                                                                                          page_width,
                                                                                                          page_height,
                                                                                                          is_credit_note)

                if nc8_data == [("REFERENCE; INTERNAL ORDER", 0)]:
                    total_net_weight = 0
                else:
                    total_net_weight = InvoiceProcessor._extract_net_weight(last_page, is_credit_note or is_debit_note)

                shipment_date = InvoiceProcessor._extract_shipment_date(s1)
                exchange_rate = get_bnr_exchange_rate(get_previous_workday(shipment_date))
                vat_number = InvoiceProcessor._extract_field(s3, r"Tax number\s*:\s*(\w+)", "Unknown", re.IGNORECASE)
                delivery_location = get_delivery_location(s1)
                delivery_condition = InvoiceProcessor._extract_field(s2, r"Incoterms\s*:\s*(\w+)", "Unknown",
                                                                     re.IGNORECASE)

                if len(nc8_data) == 1:
                    return [{"company": company, "invoice_number": invoice_number, "nc8_code": nc8_data[0][0],
                             "origin": origin, "destination": destination, "invoice_value_eur": invoice_value_eur,
                             "net_weight": total_net_weight, "shipment_date": shipment_date.strftime("%d.%m.%Y"),
                             "exchange_rate": exchange_rate, "value_ron": invoice_value_ron, "vat_number": vat_number,
                             "delivery_location": delivery_location, "delivery_condition": delivery_condition}]
                else:
                    proportional_weights = InvoiceProcessor._calculate_proportional_weights(nc8_data, total_net_weight)

                    results = []
                    for (nc8_code, partial_value), (_, proportional_weight) in zip(nc8_data, proportional_weights):
                        if currency == "EUR":
                            invoice_value_eur_for_code = round(partial_value, 2)
                            invoice_value_ron_for_code = round(partial_value * exchange_rate, 2)
                        elif currency == "RON":
                            invoice_value_ron_for_code = round(partial_value, 2)
                            invoice_value_eur_for_code = round(partial_value / exchange_rate, 2) if exchange_rate else 0
                        else:
                            invoice_value_eur_for_code = 0
                            invoice_value_ron_for_code = 0

                        result = {"company": company, "invoice_number": invoice_number, "nc8_code": nc8_code,
                                  "origin": origin, "destination": destination,
                                  "invoice_value_eur": invoice_value_eur_for_code, "net_weight": proportional_weight,
                                  "shipment_date": shipment_date.strftime("%d.%m.%Y"), "exchange_rate": exchange_rate,
                                  "value_ron": invoice_value_ron_for_code, "vat_number": vat_number,
                                  "delivery_location": delivery_location, "delivery_condition": delivery_condition}
                        results.append(result)

                    return results
        except Exception as e:
            print(f"[LOG] Error processing PDF {pdf_path}: {e}")
            return []

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
        Extracts NC8 codes along with their corresponding values from the PDF.
        """
        if is_credit_note:
            return [("Credit Note", 0)]

        s4 = InvoiceProcessor._extract_section_text(first_page, "section_4", page_width, page_height)
        if s4 and "REFERENCE" in s4 and "INTERNAL ORDER" in s4:
            return [("REFERENCE; INTERNAL ORDER", 0)]

        currency_pattern = re.compile(r"\b(EUR|RON)\b\s+([\d.,]+)\s+([\d.,]+)")
        specialized_pattern = re.compile(
            r"^[A-Za-z0-9]+\s+PER\s+(?:\d{1,3}(?:[.,]\d{3})+|\d+)\s+PC\s+\d+PC\s+([\d.,]+)\s+([\d.,]+)$")
        code_pattern = re.compile(r"Commodity Code\s*:\s*(\d+)")

        lines_all = []
        for page in pdf.pages:
            text = page.extract_text() or ""
            page_lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
            lines_all.extend(page_lines)

        pairs = []
        current_value = None
        i = 0

        while i < len(lines_all):
            line = lines_all[i]

            currency_match = currency_pattern.search(line)
            if currency_match:
                raw_val = currency_match.group(2)
                current_value = parse_mixed_number(raw_val)
                i += 1
                continue

            sp_match = specialized_pattern.match(line)
            if sp_match and (i + 4) < len(lines_all):
                line_purch = lines_all[i + 2]
                line_code = lines_all[i + 3]
                line_country = lines_all[i + 4]

                cmatch = code_pattern.search(line_code)

                if (line_purch.lower().startswith("purch. order no.") and cmatch and line_country.lower().startswith(
                        "country of origin")):
                    raw_main_val = sp_match.group(2)
                    current_value = parse_mixed_number(raw_main_val)

                    nc8_code = cmatch.group(1)
                    pairs.append((nc8_code, current_value))

                    current_value = None
                    i += 5
                    continue

            cmatch = code_pattern.search(line)
            if cmatch:
                nc8_code = cmatch.group(1)
                if current_value is not None:
                    pairs.append((nc8_code, current_value))
                    current_value = None
                else:
                    pairs.append((nc8_code, 0.0))
                i += 1
                continue

            i += 1

        if not pairs:
            return [("Unknown", 0)]

        merged = defaultdict(float)
        for code, val in pairs:
            merged[code] += val

        return [(cd, total_val) for cd, total_val in merged.items()]

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

            amount = parse_mixed_number(amount_str)
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
    def _extract_net_weight(last_page, is_credit_or_debit_note):
        """
        Extract net weight if not a credit or debit note.
        Returns the total net weight in kg.
        """
        if is_credit_or_debit_note:
            return 0
        text = last_page.extract_text()
        if not text:
            return 0

        match = re.search(r"Net weight\s+([\d.,]+)\s+KG", text, re.IGNORECASE)
        if match:
            raw_val = match.group(1)
            num_val = parse_mixed_number(raw_val)
            return round(num_val)
        return 0

    @staticmethod
    def _calculate_proportional_weights(nc8_data, total_weight):
        """
        Calculate proportional weights for each NC8 code based on their values.
        Returns a list of tuples (nc8_code, proportional_weight).
        """
        if not nc8_data or total_weight <= 0:
            return [(code, 0) for code, _ in nc8_data]

        total_value = sum(value for _, value in nc8_data)
        if total_value <= 0:
            return [(code, 0) for code, _ in nc8_data]

        return [(code, round(total_weight * value / total_value)) for code, value in nc8_data]

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
