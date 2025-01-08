import re
from datetime import datetime

import pandas as pd
import pdfplumber

from functions.calculate_coordinates import calculate_coordinates
from functions.get_bnr_exchange_rate import get_bnr_exchange_rate
from functions.get_country_code_from_address import get_country_code_from_address
from functions.get_delivery_location import get_delivery_location


class Constants:
    # Proportions for invoice sections
    PROPORTIONS = {"section_1": (0.0, 0.0, 1.0, 0.16), "section_2": (0.0, 0.16, 0.46, 0.54),
                   "section_3": (0.46, 0.16, 1.0, 0.54), "section_4": (0.0, 0.54, 1.0, 0.93),
                   "section_5": (0.0, 0.93, 1.0, 1.0), }


class InvoiceProcessor:
    def __init__(self, paths):
        self.input_paths = paths
        self.df = pd.DataFrame(
            columns=["company", "invoice_number", "nc8_code", "origin", "destination", "invoice_value_eur",
                     "net_weight", "shipment_date", "exchange_rate", "value_ron", "vat_number", "delivery_location",
                     "delivery_condition"])

    def process_invoices(self):
        for idx, pdf_path in enumerate(self.input_paths):
            with pdfplumber.open(pdf_path) as pdf:
                first_page = pdf.pages[0]
                page_width, page_height = first_page.width, first_page.height

                self._extract_sections_as_images(first_page, page_width, page_height)

                section_2_text = self._extract_section_text(first_page, "section_2", page_width, page_height)
                company = self._extract_field(section_2_text, r"Our payment address\n(.+)", "Unknown")
                print("Company:", company)

                section_1_text = self._extract_section_text(first_page, "section_1", page_width, page_height)
                invoice_number = self._extract_invoice_number(section_1_text)
                print("Invoice Number:", invoice_number)

                nc8_codes = self._extract_nc8_codes(pdf)
                print("NC8 Codes:", nc8_codes)

                origin = get_country_code_from_address(
                    self._extract_field(section_2_text, r"Our payment address\n(.+?)\nPayment date", "Unknown",
                                        re.DOTALL))
                print("Origin:", origin)

                section_3_text = self._extract_section_text(first_page, "section_3", page_width, page_height)
                destination = get_country_code_from_address(
                    self._extract_field(section_3_text, r"Invoiced to : (.+?)\nCredit transfer", "Unknown", re.DOTALL))
                print("Destination:", destination)

                invoice_value_eur, invoice_value_ron = self._extract_invoice_values(pdf.pages[-1], page_width,
                                                                                    page_height)
                print("Invoice Value EUR:", invoice_value_eur)
                print("Invoice Value RON:", invoice_value_ron)

                net_weight = self._extract_net_weight(pdf.pages[-1])
                print("Net Weight:", net_weight)

                shipment_date = self._extract_field(pdf.pages[-1].extract_text(),
                                                    r"Transportation date: (\d{2}\.\d{2}\.\d{4})", None)
                shipment_date = datetime.strptime(shipment_date, "%d.%m.%Y") if shipment_date else datetime.now()
                print("Shipment Date:", shipment_date.strftime("%d.%m.%Y"))

                exchange_rate = get_bnr_exchange_rate(shipment_date, "RON")
                print("Exchange Rate:", exchange_rate)

                delivery_location = get_delivery_location(section_1_text)
                print("Delivery Location:", delivery_location)

                vat_number = self._extract_field(section_3_text, r"Tax number : (\w+)", "Unknown")
                print("VAT Number:", vat_number)

                delivery_condition = self._extract_field(section_2_text, r"Incoterms : (\w+)", "Unknown")
                print("Delivery Condition:", delivery_condition)

                self.df.loc[idx] = [company, invoice_number, ", ".join(nc8_codes), origin, destination,
                                    invoice_value_eur, net_weight, shipment_date.strftime("%d.%m.%Y"), exchange_rate,
                                    invoice_value_ron, vat_number, delivery_location, delivery_condition]

    @staticmethod
    def _extract_sections_as_images(page, page_width, page_height):
        for section_name, proportion in Constants.PROPORTIONS.items():
            section_coords = calculate_coordinates(page_width, page_height, proportion)
            section_image = page.within_bbox(section_coords).to_image()
            image_path = f"output/{section_name}.png"
            section_image.save(image_path)

    @staticmethod
    def _extract_section_text(page, section_name, page_width, page_height):
        section_coords = calculate_coordinates(page_width, page_height, Constants.PROPORTIONS[section_name])
        return page.within_bbox(section_coords).extract_text()

    @staticmethod
    def _extract_field(text, pattern, default, flags=0):
        match = re.search(pattern, text, flags)
        return match.group(1).strip() if match else default

    @staticmethod
    def _extract_invoice_number(section_1_text):
        lines = section_1_text.split("\n")
        return lines[-1].split(" ")[0] if lines else "Unknown"

    @staticmethod
    def _extract_nc8_codes(pdf):
        codes = []
        for page in pdf.pages:
            codes += re.findall(r"Commodity Code : (\d+)", page.extract_text())
        return codes

    @staticmethod
    def _extract_invoice_values(last_page, page_width, page_height):
        bbox_ratio = (1125 / 1653, 2054 / 2339, 1552 / 1653, 2182 / 2339)
        bbox_coords = calculate_coordinates(page_width, page_height, bbox_ratio)
        bbox_text = last_page.within_bbox(bbox_coords).extract_text()
        lines = bbox_text.splitlines() if bbox_text else []
        last_line = lines[-1] if lines else ""
        if "*" in last_line:
            currency, currency_value = last_line.split("*")[0].strip().lower(), last_line.split("*")[-1].strip()
            if currency == "eur":
                return float(currency_value.replace(",", "").replace(".", ".")), 0
            elif currency == "ron":
                return 0, float(currency_value.replace(".", "").replace(",", "."))
        return 0, 0

    @staticmethod
    def _extract_net_weight(last_page):
        match = re.search(r"Net weight\s+(.+?)\s+KG", last_page.extract_text())
        return float(match.group(1).replace(",", ".").replace(".", "")) if match else 0


if __name__ == "__main__":
    input_paths = ["pdfs/0912626530_C015000044_ZSM0_001.PDF", "pdfs/9610169997_2007000000_ZSM0_001.PDF"]
    processor = InvoiceProcessor(input_paths)
    processor.process_invoices()
    print(processor.df)
