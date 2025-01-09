import os
import re
from dataclasses import dataclass
from datetime import datetime

import pandas as pd
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

from functions.calculate_coordinates import calculate_coordinates
from functions.convert_to_date import convert_to_date
from functions.format_nc8_code import format_nc8_code
from functions.get_bnr_exchange_rate import get_bnr_exchange_rate
from functions.get_country_code_from_address import get_country_code_from_address
from functions.get_delivery_location import get_delivery_location


@dataclass(frozen=True)
class Constants:
    __slots__ = ()

    COLUMN_FORMATS = {"nr_crt": "General", "company": "General", "invoice_number": "0", "nc8_code": "00 00 0000",
                      "origin": "General", "destination": "General", "invoice_value_eur": "#,##0.00",
                      "net_weight": "#,##0.00", "shipment_date": "dd.mmm", "exchange_rate": "#,##0.0000;-#,##0.0000",
                      "value_ron": "#,##0;-#,##0", "vat_number": "General", "delivery_location": "General",
                      "delivery_condition": "General", "percentage": "0.00", "transport": "#,##0.00",
                      "statistic": "#,##0;-#,##0", }

    COLUMNS = ["company", "invoice_number", "nc8_code", "origin", "destination", "invoice_value_eur", "net_weight",
               "shipment_date", "exchange_rate", "value_ron", "vat_number", "delivery_location", "delivery_condition"]

    FONT_SIZE = 12

    HEADERS = {"nr_crt": "Nr Crt", "company": "Firma", "invoice_number": "Nr Factura Marfa", "nc8_code": "Cod NC8",
               "origin": "Origine", "destination": "Destinatie", "invoice_value_eur": "Val Fact Euro",
               "net_weight": "Greutate Neta", "shipment_date": "Data Expeditiei", "exchange_rate": "Curs Valutar",
               "value_ron": "Valoare Ron", "vat_number": "Vat Cumparator", "delivery_location": "Loc Livrare",
               "delivery_condition": "Conditie Livrare", "percentage": "%", "transport": "Transport",
               "statistic": "Statistica", }

    PROPORTIONS = {"section_1": (0.0, 0.0, 1.0, 0.16), "section_2": (0.0, 0.16, 0.46, 0.54),
                   "section_3": (0.46, 0.16, 1.0, 0.54), "section_4": (0.0, 0.54, 1.0, 0.93),
                   "section_5": (0.0, 0.93, 1.0, 1.0), }

    SCALING_FACTOR = 1.2


class InvoiceProcessor:
    __slots__ = ("input_paths", "df")

    def __init__(self, paths):
        self.input_paths = paths
        self.df = pd.DataFrame(columns=Constants.COLUMNS)

    def process_invoices(self):
        for idx, pdf_path in enumerate(self.input_paths):
            with pdfplumber.open(pdf_path) as pdf:
                first_page = pdf.pages[0]
                page_width, page_height = first_page.width, first_page.height

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
            currency, raw_value = last_line.split("*")[0].strip().lower(), last_line.split("*")[-1].strip()
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
        match = re.search(r"Net weight\s+(.+?)\s+KG", last_page.extract_text())
        return float(match.group(1).replace(",", ".").replace(".", "")) if match else 0


class ExcelGenerator:
    __slots__ = "data"

    def __init__(self, data):
        self.data = data

    def generate_excel(self, path):
        self._prepare_data()
        self._add_totals(["net_weight", "value_ron", "statistic"], group_by="vat_number")
        self._add_excel_formulas()

        wb = Workbook()
        ws = wb.active
        ws.title = "Invoices"

        self._write_headers(ws)
        self._write_rows(ws)

        wb.save(path)
        print(f"Excel file saved to {path}")

    def _prepare_data(self):
        self.data = self.data.sort_values(by=["vat_number", "shipment_date"]).reset_index(drop=True)
        self.data.insert(0, "nr_crt", range(1, len(self.data) + 1))
        self.data["percentage"] = 0.6
        self.data["transport"] = ""
        self.data["statistic"] = ""
        self.data["shipment_date"] = self.data["shipment_date"].apply(convert_to_date)
        self.data["invoice_number"] = self.data["invoice_number"].fillna(0).astype(int)
        self.data["nc8_code"] = self.data["nc8_code"].str.split(", ").str[0].apply(format_nc8_code)

    def _add_totals(self, columns_to_total, group_by=None):
        current_row_offset = 1

        if group_by:
            grouped = self.data.groupby(group_by, sort=False)
            results = []
            for name, group in grouped:
                group = group.reset_index(drop=True)
                results.append(group)

                total_row = {col: "" for col in self.data.columns}
                total_row[group_by] = f"Total {name}"
                total_row["nr_crt"] = ""

                for col in columns_to_total:
                    col_letter = chr(65 + self.data.columns.get_loc(col))
                    start_row = current_row_offset + 1
                    end_row = start_row + len(group) - 1
                    total_row[col] = f"=SUM({col_letter}{start_row}:{col_letter}{end_row})"

                current_row_offset += len(group) + 1

                results.append(pd.DataFrame([total_row]))

            self.data = pd.concat(results, ignore_index=True)
        else:
            total_row = {col: "" for col in self.data.columns}
            total_row["nr_crt"] = "Total"

            for col in columns_to_total:
                col_letter = chr(65 + self.data.columns.get_loc(col))
                start_row = 2
                end_row = len(self.data) + 1
                total_row[col] = f"=SUM({col_letter}{start_row}:{col_letter}{end_row})"

            self.data = pd.concat([self.data, pd.DataFrame([total_row])], ignore_index=True)

    def _add_excel_formulas(self):
        for i, row in self.data.iterrows():
            excel_row = i + 2

            cell_exchange_rate = f"J{excel_row}"
            cell_net_weight = f"H{excel_row}"
            cell_value_eur = f"G{excel_row}"
            cell_value_ron = f"K{excel_row}"
            cell_percentage = f"O{excel_row}"

            if pd.isna(row["nr_crt"]) or str(row["nr_crt"]).strip() == "":
                continue

            if row["value_ron"] == 0:
                self.data.at[i, "value_ron"] = f"={cell_value_eur}*{cell_exchange_rate}"
            if row["invoice_value_eur"] == 0:
                self.data.at[i, "invoice_value_eur"] = f"={cell_value_ron}/{cell_exchange_rate}"

            if pd.notna(row["net_weight"]) and pd.notna(row["exchange_rate"]):
                self.data.at[i, "transport"] = f"=28000*{cell_exchange_rate}/147000*{cell_net_weight}"
            else:
                self.data.at[i, "transport"] = ""

            if pd.notna(row["value_ron"]):
                self.data.at[i, "statistic"] = f"=ROUND({cell_value_ron}+{cell_percentage}*{cell_exchange_rate}, 0)"
            else:
                self.data.at[i, "statistic"] = ""

    @staticmethod
    def _write_headers(ws):
        header_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
        for col_num, header in enumerate(Constants.HEADERS.values(), 1):
            cell = ws.cell(row=1, column=col_num, value=header)
            cell.font = Font(bold=True, size=Constants.FONT_SIZE, name="Arial")
            cell.alignment = Alignment(horizontal="left", vertical="bottom", wrap_text=True)
            cell.fill = header_fill

        ws.freeze_panes = "A2"

    def _write_rows(self, ws):
        headers_keys = list(Constants.HEADERS.keys())

        for i, row in self.data.iterrows():
            is_falsy = pd.isna(row["nr_crt"]) or str(row["nr_crt"]).strip() == ""
            for col_num, cell_value in enumerate(row, 1):
                cell = ws.cell(row=i + 2, column=col_num, value=cell_value)
                header = headers_keys[col_num - 1]
                if is_falsy:
                    cell.font = Font(bold=True, size=12, name="Arial")
                else:
                    cell.font = Font(size=12, name="Arial")
                if header in Constants.COLUMN_FORMATS:
                    cell.number_format = Constants.COLUMN_FORMATS[header]

        for col_num, header in enumerate(Constants.HEADERS.values(), 1):
            column_letter = get_column_letter(col_num)
            max_length = len(header)
            for row in ws.iter_rows(min_col=col_num, max_col=col_num, min_row=2, values_only=True):
                value = row[0]
                if value is not None:
                    max_length = max(max_length, len(str(value)))
            adjusted_width = max_length * (Constants.FONT_SIZE / 10) * Constants.SCALING_FACTOR
            ws.column_dimensions[column_letter].width = adjusted_width + 2


if __name__ == "__main__":
    input_paths = [os.path.join("pdfs", file) for file in os.listdir("pdfs") if file.lower().endswith(".pdf")]
    processor = InvoiceProcessor(input_paths)
    processor.process_invoices()

    current_date = datetime.now().strftime("%d-%m-%Y")
    output_path = f"output/{current_date}-EXP.xlsx"

    excel_generator = ExcelGenerator(processor.df)
    excel_generator.generate_excel(output_path)
