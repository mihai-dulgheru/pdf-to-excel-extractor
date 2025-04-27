import bisect
import os

import openpyxl
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

from config import Constants
from functions import convert_to_date, format_nc8_code


class ExcelGenerator:
    """
    Creates and formats an Excel file from a pandas DataFrame.
    """

    __slots__ = ("data", "percentage")

    def __init__(self, data, percentage=0.6):
        """
        :param data: DataFrame with invoice information.
        :param percentage: Float percentage for calculations.
        """
        self.data = data
        self.percentage = percentage

    def generate_excel(self, path, existing_excel=None):
        """
        Prepare data, add totals, add formulas, then either append to an existing workbook
        or create a new one at 'path'.
        """
        self._prepare_data()
        if existing_excel and os.path.exists(existing_excel):
            self._append_new_invoices_to_workbook(existing_excel, path)
            return

        self._add_excel_formulas()
        self._add_totals(["net_weight", "value_ron", "statistic"], group_by="vat_number")

        wb = Workbook()
        ws = wb.active
        ws.title = "Invoices"

        self._write_headers(ws)
        self._write_rows(ws)
        wb.save(path)

    def _prepare_data(self):
        """
        Sort, clean, and prepare data for Excel output.
        """
        self.data["shipment_date"] = self.data["shipment_date"].apply(convert_to_date)
        self.data["invoice_number"] = pd.to_numeric(self.data["invoice_number"], errors="coerce").fillna(0).astype(int)
        self.data = self.data.sort_values(by=["vat_number", "shipment_date", "invoice_number"]).reset_index(drop=True)
        self.data.insert(0, "nr_crt", self.data.groupby("vat_number").cumcount() + 1)
        self.data["percentage"] = self.percentage
        self.data["transport"] = self.data.get("transport", "")
        self.data["statistic"] = self.data.get("statistic", "")
        self.data["nc8_code"] = self.data["nc8_code"].apply(format_nc8_code)

    def _append_new_invoices_to_workbook(self, src_path, dest_path):
        """
        Append non-duplicate invoice records into a copy of the workbook at src_path
        and save results to dest_path, preserving formatting and formulas.
        """
        if not os.path.exists(src_path) or self.data.empty:
            return

        wb = openpyxl.load_workbook(src_path)
        ws = wb.active
        formats = Constants.COLUMN_FORMATS

        header_map = self._map_headers(ws)
        struct = self._find_data_block(ws, header_map)
        existing = self._collect_existing_keys(ws, struct, header_map)

        df = (self.data.assign(invoice_number=lambda x: pd.to_numeric(x.invoice_number, errors='coerce')).dropna(
            subset=['invoice_number']).assign(invoice_number=lambda x: x.invoice_number.astype(int),
                                              vat_number=lambda x: x.vat_number.astype(str),
                                              nc8_code=lambda x: x.nc8_code.astype(str)))

        new_records = [r for _, r in df.iterrows() if (r.vat_number, r.invoice_number, r.nc8_code) not in existing]
        new_records.sort(key=lambda r: (r.vat_number, r.shipment_date, r.invoice_number))

        initial_vats = set(struct['vat_groups'])
        new_vats = {}

        for rec in new_records:
            row_idx = self._find_insert_rows(ws, rec, struct, header_map)
            ws.insert_rows(row_idx)
            self._write_record(ws, row_idx, rec, header_map)
            self._advance_structure(struct, row_idx)

            formulas = self._build_formulas_for_row(row_idx, rec)
            for field, col in header_map.items():
                cell = ws.cell(row=row_idx, column=col)
                if field in formulas:
                    cell.value = formulas[field]
                else:
                    val = rec.get(field)
                    if val is not None:
                        cell.value = val
                if field in formats:
                    cell.number_format = formats[field]
            v = rec.vat_number
            if v not in initial_vats and v not in new_vats:
                new_vats[v] = row_idx

        self._rebuild_totals(ws, struct)

        for vat, start in new_vats.items():
            count = sum(1 for r in new_records if r.vat_number == vat)
            end = start + count - 1
            tot_row = end + 1
            ws.insert_rows(tot_row)

            idx = header_map['vat_number']
            header_cell = ws.cell(row=tot_row, column=idx, value=f"Total {vat}")
            header_cell.font = Font(bold=True, size=12, name='Arial')

            for field in ('net_weight', 'value_ron', 'statistic'):
                col = header_map[field]
                letter = get_column_letter(col)
                sum_cell = ws.cell(row=tot_row, column=col, value=f"=SUM({letter}{start}:{letter}{end})")
                sum_cell.font = Font(bold=True, size=12, name='Arial')

        wb.save(dest_path)

    def _add_excel_formulas(self):
        """
        Add in-cell formulas for computed columns (value_ron, invoice_value_eur, transport, statistic).
        """
        formula_cols = ["value_ron", "invoice_value_eur", "transport", "statistic"]
        self.data = self.data.astype({col: object for col in formula_cols})

        for i, row in self.data.iterrows():
            excel_row = i + 2
            formulas = self._build_formulas_for_row(excel_row, row.to_dict())
            for col, formula in formulas.items():
                self.data.at[i, col] = formula

    def _add_totals(self, columns_to_total, group_by=None):
        """
        Insert total rows after each group or at the end of the DataFrame.
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

    @staticmethod
    def _write_headers(ws):
        """
        Write column headers, apply style, freeze header row.
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
        Write DataFrame rows, including formulas, and adjust column widths.
        """
        headers_keys = list(Constants.HEADERS.keys())
        for i, row in self.data.iterrows():
            is_total_row = pd.isna(row["nr_crt"]) or str(row["nr_crt"]).strip() == ""
            for col_num, cell_value in enumerate(row, 1):
                cell = ws.cell(row=i + 2, column=col_num, value=cell_value)
                header_key = headers_keys[col_num - 1]

                if is_total_row:
                    cell.font = Font(bold=True, size=12, name="Arial")
                else:
                    cell.font = Font(size=12, name="Arial")

                if header_key in Constants.COLUMN_FORMATS:
                    cell.number_format = Constants.COLUMN_FORMATS[header_key]

        for col_num, header in enumerate(Constants.HEADERS.values(), 1):
            col_letter = get_column_letter(col_num)
            max_len = len(header)
            for row in ws.iter_rows(min_col=col_num, max_col=col_num, min_row=2, values_only=True):
                val = row[0]
                if val is not None:
                    max_len = max(max_len, len(str(val)))
            adjusted_width = max_len * (Constants.FONT_SIZE / 10) * Constants.SCALING_FACTOR
            ws.column_dimensions[col_letter].width = adjusted_width + 2

    @staticmethod
    def _map_headers(ws):
        """Map header titles to column indices."""
        rev = {v: k for k, v in Constants.HEADERS.items()}
        return {rev[cell.value]: i + 1 for i, cell in enumerate(ws[1]) if cell.value in rev}

    @staticmethod
    def _find_data_block(ws, cmap):
        """Identify data range and VAT grouping positions."""
        struct = {'data_start': 2, 'data_end': None, 'blank_row': None, 'vat_groups': {}}
        current_vat = None
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if all(not v or (isinstance(v, str) and not v.strip()) for v in row):
                struct['blank_row'], struct['data_end'] = row_idx, row_idx - 1
                break
            vat = row[cmap['vat_number'] - 1]
            num = row[cmap['nr_crt'] - 1]
            if isinstance(vat, str) and vat.startswith('Total '):
                label = vat.split(' ', 1)[1]
                group = struct['vat_groups'].setdefault(label, {})
                group['end'] = row_idx
            elif vat and num and vat != current_vat:
                current_vat = vat
                struct['vat_groups'][vat] = {'start': row_idx, 'end': None}
        if struct['blank_row'] is None:
            struct['blank_row'], struct['data_end'] = ws.max_row + 1, ws.max_row
        return struct

    @staticmethod
    def _collect_existing_keys(ws, struct, cmap):
        """Collect existing invoice keys to avoid duplicates."""
        keys = set()
        for r in range(struct['data_start'], struct['data_end'] + 1):
            vat = ws.cell(r, cmap['vat_number']).value
            inv = ws.cell(r, cmap['invoice_number']).value
            nc = ws.cell(r, cmap['nc8_code']).value
            if vat and inv and nc and not str(vat).startswith('Total '):
                keys.add((str(vat), int(inv), str(nc)))
        return keys

    @staticmethod
    def _find_insert_rows(ws, rec, struct, cmap):
        """Determine the correct row for inserting a new record."""
        grp = struct['vat_groups'].get(rec.vat_number)
        if not grp or not grp.get('end'):
            return struct['data_end'] + 1
        start, end = grp['start'], grp['end']
        rows = list(range(start, end))
        pairs = [(ws.cell(r, cmap['shipment_date']).value, int(ws.cell(r, cmap['invoice_number']).value), r) for r in
                 rows]
        pos = bisect.bisect_left([(d, inv) for d, inv, _ in pairs], (rec.shipment_date, rec.invoice_number))
        return pairs[pos][2] if pos < len(pairs) else end

    @staticmethod
    def _write_record(ws, row, rec, cmap):
        """Write invoice record into worksheet row."""
        for field, col in cmap.items():
            val = rec.get(field)
            if field == 'delivery_location' and str(val) == '0':
                alt = next((v for v in rec.delivery_location if v != '0'), None)
                if alt:
                    val = alt
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = Font(size=12, name='Arial')

    @staticmethod
    def _advance_structure(struct, ins):
        """Adjust data block indices after row insertion."""
        struct['data_end'] += 1
        struct['blank_row'] += 1
        for g in struct['vat_groups'].values():
            if g.get('start', 0) >= ins: g['start'] += 1
            if g.get('end', 0) and g['end'] >= ins: g['end'] += 1

    @staticmethod
    def _build_formulas_for_row(row_idx, row):
        """
        Generate Excel formulas for a given row index and data row.
        Returns a dict mapping column names to formula strings.
        """
        excel_row = row_idx
        cell_exchange_rate = f"J{excel_row}"
        cell_value_eur = f"G{excel_row}"
        cell_value_ron = f"K{excel_row}"
        cell_net_weight = f"H{excel_row}"
        cell_percentage = f"O{excel_row}"
        cell_transport = f"P{excel_row}"

        formulas = {}
        if pd.isna(row.get("nr_crt")) or str(row.get("nr_crt")).strip() == "":
            return formulas

        val_eur = row.get("invoice_value_eur", 0)
        val_ron = row.get("value_ron", 0)

        if (val_eur == 0 and val_ron == 0) or (val_eur != 0 and val_ron != 0) or (val_eur != 0 and val_ron == 0):
            formulas["value_ron"] = f"={cell_value_eur}*{cell_exchange_rate}"
        elif val_eur == 0 and val_ron != 0:
            formulas["invoice_value_eur"] = f"={cell_value_ron}/{cell_exchange_rate}"

        if pd.notna(row.get("net_weight")) and pd.notna(row.get("exchange_rate")):
            formulas["transport"] = f"=28000*{cell_exchange_rate}/147000*{cell_net_weight}"
        else:
            formulas["transport"] = ""

        if pd.notna(val_ron):
            formulas["statistic"] = f"=ROUND({cell_value_ron}+{cell_percentage}*{cell_transport}, 0)"
        else:
            formulas["statistic"] = ""

        return formulas

    @staticmethod
    def _rebuild_totals(ws, struct):
        """Update Total rows formulas to match new data ranges."""
        for grp in struct['vat_groups'].values():
            start = grp.get('start', 0)
            end_row = grp.get('end', 0) - 1
            total_row = grp.get('end')
            if start and end_row >= start and total_row:
                for cell in ws[total_row]:
                    f = cell.value
                    if not (isinstance(f, str) and f.startswith('=SUM(') and f.endswith(')')):
                        continue
                    inside = f[f.find('(') + 1: f.rfind(')')]
                    if ':' not in inside:
                        continue
                    start_ref, end_ref = inside.split(':', 1)
                    col = ''.join(ch for ch in start_ref if ch.isalpha())
                    cell.value = f"=SUM({col}{start}:{col}{end_row})"
