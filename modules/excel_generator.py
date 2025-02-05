import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

from config import Constants
from functions.convert_to_date import convert_to_date
from functions.format_nc8_code import format_nc8_code


class ExcelGenerator:
    """
    This class handles the creation and formatting of an Excel file
    from the processed invoice data in a pandas DataFrame.
    """
    __slots__ = ("data", "percentage")

    def __init__(self, data, percentage=0.6):
        """
        Initialize the ExcelGenerator with a pandas DataFrame and a percentage value.
        :param data: pandas DataFrame containing the invoice information.
        :param percentage: A float percentage used in calculations.
        """
        self.data = data
        self.percentage = percentage

    def generate_excel(self, path):
        """
        Main method that prepares data, calculates totals, adds formulas,
        and writes everything to an Excel file at the specified path.
        :param path: The output Excel file path.
        """
        self._prepare_data()
        self._add_totals(["net_weight", "value_ron", "statistic"], group_by="vat_number")
        self._add_excel_formulas()

        wb = Workbook()
        ws = wb.active
        ws.title = "Invoices"

        self._write_headers(ws)
        self._write_rows(ws)

        wb.save(path)
        print(f"[LOG] Excel file saved to {path}")  # English log message

    def _prepare_data(self):
        """
        Prepare the DataFrame by sorting, inserting extra columns, and cleaning up data.
        """
        self.data = self.data.sort_values(by=["vat_number", "shipment_date"]).reset_index(drop=True)
        self.data.insert(0, "nr_crt", range(1, len(self.data) + 1))
        self.data["percentage"] = self.percentage
        self.data["transport"] = ""
        self.data["statistic"] = ""
        self.data["shipment_date"] = self.data["shipment_date"].apply(convert_to_date)
        self.data["invoice_number"] = self.data["invoice_number"].fillna(0)

        # Convert invoice_number to integer if possible
        try:
            self.data["invoice_number"] = self.data["invoice_number"].astype(int)
        except ValueError:
            pass

        self.data["nc8_code"] = self.data["nc8_code"].str.split(", ").str[0].apply(format_nc8_code)

    def _add_totals(self, columns_to_total, group_by=None):
        """
        Add totals (formulas) at the end of each group or at the end of the entire DataFrame.
        :param columns_to_total: A list of column names to sum up.
        :param group_by: If provided, totals are inserted after each group.
        """
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
        """
        Add in-cell formulas to the DataFrame for columns like value_ron, invoice_value_eur, transport, and statistic.
        """
        for i, row in self.data.iterrows():
            excel_row = i + 2

            cell_exchange_rate = f"J{excel_row}"
            cell_net_weight = f"H{excel_row}"
            cell_value_eur = f"G{excel_row}"
            cell_value_ron = f"K{excel_row}"
            cell_percentage = f"O{excel_row}"

            # Skip total rows or empty rows
            if pd.isna(row["nr_crt"]) or str(row["nr_crt"]).strip() == "":
                continue

            # If RON is 0, calculate it
            if row["value_ron"] == 0:
                self.data.at[i, "value_ron"] = f"={cell_value_eur}*{cell_exchange_rate}"

            # If EUR is 0, calculate it
            if row["invoice_value_eur"] == 0:
                self.data.at[i, "invoice_value_eur"] = f"={cell_value_ron}/{cell_exchange_rate}"

            # Transport formula
            if pd.notna(row["net_weight"]) and pd.notna(row["exchange_rate"]):
                self.data.at[i, "transport"] = f"=28000*{cell_exchange_rate}/147000*{cell_net_weight}"
            else:
                self.data.at[i, "transport"] = ""

            # Statistic formula
            if pd.notna(row["value_ron"]):
                self.data.at[i, "statistic"] = f"=ROUND({cell_value_ron}+{cell_percentage}*{cell_exchange_rate}, 0)"
            else:
                self.data.at[i, "statistic"] = ""

    @staticmethod
    def _write_headers(ws):
        """
        Write the column headers into the worksheet, apply styling,
        and freeze the header row.
        """
        header_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
        for col_num, header in enumerate(Constants.HEADERS.values(), 1):
            cell = ws.cell(row=1, column=col_num, value=header)
            cell.font = Font(bold=True, size=Constants.FONT_SIZE, name="Arial")
            cell.alignment = Alignment(horizontal="left", vertical="bottom", wrap_text=True)
            cell.fill = header_fill
        ws.freeze_panes = "A2"

    def _write_rows(self, ws):
        """
        Write the data rows into the worksheet, including any formulas,
        and adjust column widths for readability.
        """
        headers_keys = list(Constants.HEADERS.keys())

        # Write cell values
        for i, row in self.data.iterrows():
            is_total_row = pd.isna(row["nr_crt"]) or str(row["nr_crt"]).strip() == ""
            for col_num, cell_value in enumerate(row, 1):
                cell = ws.cell(row=i + 2, column=col_num, value=cell_value)
                header_key = headers_keys[col_num - 1]

                # Styling for totals
                if is_total_row:
                    cell.font = Font(bold=True, size=12, name="Arial")
                else:
                    cell.font = Font(size=12, name="Arial")

                # Apply number format if defined in Constants.COLUMN_FORMATS
                if header_key in Constants.COLUMN_FORMATS:
                    cell.number_format = Constants.COLUMN_FORMATS[header_key]

        # Auto-adjust column widths
        for col_num, header in enumerate(Constants.HEADERS.values(), 1):
            column_letter = get_column_letter(col_num)
            max_length = len(header)
            for row in ws.iter_rows(min_col=col_num, max_col=col_num, min_row=2, values_only=True):
                value = row[0]
                if value is not None:
                    max_length = max(max_length, len(str(value)))
            adjusted_width = max_length * (Constants.FONT_SIZE / 10) * Constants.SCALING_FACTOR
            ws.column_dimensions[column_letter].width = adjusted_width + 2
