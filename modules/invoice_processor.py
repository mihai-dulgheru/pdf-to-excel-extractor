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
    This class is responsible for extracting information from PDF invoices
    and storing that data into a pandas DataFrame.
    """
    __slots__ = ("input_paths", "df", "progress_callback")

    def __init__(self, paths, progress_callback=None):
        """
        Initialize the InvoiceProcessor with paths to PDF files and an optional progress callback.
        :param paths: List of paths to the input PDF files.
        :param progress_callback: A callable to report progress (processed, total).
        """
        self.input_paths = paths
        self.df = pd.DataFrame(columns=Constants.COLUMNS)
        self.progress_callback = progress_callback

    def process_invoices(self):
        """
        Read each PDF file, parse relevant information, and store the results in a pandas DataFrame.
        If a progress_callback is provided, it is called after processing each file.
        """
        total_invoices = len(self.input_paths)
        for idx, pdf_path in enumerate(self.input_paths):
            with pdfplumber.open(pdf_path) as pdf:
                first_page = pdf.pages[0]
                page_width, page_height = first_page.width, first_page.height

                # Extract text from predefined sections
                section_2_text = self._extract_section_text(first_page, "section_2", page_width, page_height)
                company = self._extract_field(section_2_text, r"Our payment address\n(.+)", "Unknown")
                print(f"[LOG] Company: {company}")  # In English (log message)

                section_1_text = self._extract_section_text(first_page, "section_1", page_width, page_height)
                invoice_number = self._extract_invoice_number(section_1_text)
                print(f"[LOG] Invoice Number: {invoice_number}")

                nc8_codes = self._extract_nc8_codes(pdf)
                print(f"[LOG] NC8 Codes: {nc8_codes}")

                origin = get_country_code_from_address(
                    self._extract_field(section_2_text, r"Our payment address\n(.+?)\nPayment date", "Unknown",
                                        re.DOTALL))
                print(f"[LOG] Origin: {origin}")

                section_3_text = self._extract_section_text(first_page, "section_3", page_width, page_height)
                destination = get_country_code_from_address(
                    self._extract_field(section_3_text, r"Invoiced to : (.+?)\nCredit transfer", "Unknown", re.DOTALL))
                print(f"[LOG] Destination: {destination}")

                invoice_value_eur, invoice_value_ron = self._extract_invoice_values(pdf.pages[-1], page_width,
                                                                                    page_height)
                print(f"[LOG] Invoice Value EUR: {invoice_value_eur}")
                print(f"[LOG] Invoice Value RON: {invoice_value_ron}")

                net_weight = self._extract_net_weight(pdf.pages[-1])
                print(f"[LOG] Net Weight: {net_weight}")

                # Extract shipment date
                raw_text_last_page = pdf.pages[-1].extract_text()
                shipment_date_str = self._extract_field(raw_text_last_page,
                                                        r"Transportation date: (\d{2}\.\d{2}\.\d{4})", None)
                shipment_date = datetime.strptime(shipment_date_str,
                                                  "%d.%m.%Y") if shipment_date_str else datetime.now()
                print(f"[LOG] Shipment Date: {shipment_date.strftime('%d.%m.%Y')}")

                exchange_rate = get_bnr_exchange_rate(shipment_date, "RON")
                print(f"[LOG] Exchange Rate: {exchange_rate}")

                delivery_location = get_delivery_location(section_1_text)
                print(f"[LOG] Delivery Location: {delivery_location}")

                vat_number = self._extract_field(section_3_text, r"Tax number : (\w+)", "Unknown")
                print(f"[LOG] VAT Number: {vat_number}")

                delivery_condition = self._extract_field(section_2_text, r"Incoterms : (\w+)", "Unknown")
                print(f"[LOG] Delivery Condition: {delivery_condition}")

                # Populate the DataFrame
                self.df.loc[idx] = [company, invoice_number, ", ".join(nc8_codes), origin, destination,
                                    invoice_value_eur, net_weight, shipment_date.strftime("%d.%m.%Y"), exchange_rate,
                                    invoice_value_ron, vat_number, delivery_location, delivery_condition]

            # Report progress if callback is available
            if self.progress_callback:
                self.progress_callback(idx + 1, total_invoices)

    @staticmethod
    def _extract_section_text(page, section_name, page_width, page_height):
        """
        Extract text from a specific section in the PDF page based on predefined proportions.
        :param page: The pdfplumber page object.
        :param section_name: The section key used to look up proportions in Constants.PROPORTIONS.
        :param page_width: The width of the page.
        :param page_height: The height of the page.
        :return: The extracted text from the specified section.
        """
        from functions.calculate_coordinates import calculate_coordinates  # Local import to avoid circular references
        section_coords = calculate_coordinates(page_width, page_height, Constants.PROPORTIONS[section_name])
        return page.within_bbox(section_coords).extract_text()

    @staticmethod
    def _extract_field(text, pattern, default, flags=0):
        """
        Extract a value from text using a regex pattern.
        If no match is found, return the default value.
        :param text: The input text to search.
        :param pattern: The regex pattern.
        :param default: Default value if pattern is not found.
        :param flags: Regex flags.
        :return: The extracted field or the default value.
        """
        if not text:
            return default
        match = re.search(pattern, text, flags)
        return match.group(1).strip() if match else default

    @staticmethod
    def _extract_invoice_number(section_1_text):
        """
        Extract the invoice number from a given section text.
        If the text is empty or invalid, return 'Unknown'.
        :param section_1_text: The text containing invoice number information.
        :return: The extracted invoice number or 'Unknown'.
        """
        if not section_1_text:
            return "Unknown"
        lines = section_1_text.split("\n")
        if lines:
            return lines[-1].split(" ")[0]
        return "Unknown"

    @staticmethod
    def _extract_nc8_codes(pdf):
        """
        Extract NC8 commodity codes from all pages of the PDF.
        :param pdf: The pdfplumber PDF object.
        :return: A list of found NC8 codes.
        """
        codes = []
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                codes += re.findall(r"Commodity Code : (\d+)", text)
        return codes

    @staticmethod
    def _extract_invoice_values(last_page, page_width, page_height):
        """
        Extract invoice values (EUR and RON) from a specific bounding box on the last page.
        :param last_page: The last pdfplumber page object.
        :param page_width: The page width.
        :param page_height: The page height.
        :return: A tuple (invoice_value_eur, invoice_value_ron).
        """
        from functions.calculate_coordinates import calculate_coordinates
        bbox_ratio = (1125 / 1653, 2054 / 2339, 1552 / 1653, 2182 / 2339)
        bbox_coords = calculate_coordinates(page_width, page_height, bbox_ratio)
        bbox_text = last_page.within_bbox(bbox_coords).extract_text()
        if not bbox_text:
            return 0, 0

        lines = bbox_text.splitlines()
        if not lines:
            return 0, 0
        last_line = lines[-1]

        if "*" in last_line:
            parts = last_line.split("*")
            if len(parts) >= 2:
                currency = parts[0].strip().lower()
                raw_value = parts[-1].strip()
                cleaned_value = raw_value.replace(".", "").replace(",", "")
                try:
                    if len(cleaned_value) > 2:
                        numeric_value = float(cleaned_value[:-2] + "." + cleaned_value[-2:])
                    else:
                        numeric_value = float("0." + cleaned_value.zfill(2))
                except ValueError:
                    return 0, 0

                if currency == "eur":
                    return numeric_value, 0
                elif currency == "ron":
                    return 0, numeric_value
        return 0, 0

    @staticmethod
    def _extract_net_weight(last_page):
        """
        Extract the net weight from the last page text.
        :param last_page: The pdfplumber page object for the last page.
        :return: The float value of the net weight, or 0 if not found.
        """
        text = last_page.extract_text()
        match = re.search(r"Net weight\s+(.+?)\s+KG", text) if text else None
        if match:
            # Attempt to convert to float
            val = match.group(1).replace(",", ".").replace(".", "")
            try:
                return float(val)
            except ValueError:
                pass
        return 0
