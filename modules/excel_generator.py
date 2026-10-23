import os
import re
import shutil
from datetime import datetime, date
from functools import lru_cache

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

    This class handles the generation of Excel files from invoice data,
    including formatting, formula generation, and appending to existing files.
    """

    __slots__ = ("data", "percentage", "appended")

    def __init__(self, data, percentage=0.6):
        """
        Initialize the ExcelGenerator with invoice data.
        """
        self.data = data
        self.percentage = percentage
        self.appended = pd.DataFrame(columns=getattr(data, "columns", None))  # rows actually written in append mode

    def generate_excel(self, path, existing_excel=None):
        """
        Generate an Excel file from the invoice data.

        Prepare data, add totals, add formulas, then either append to an existing workbook
        or create a new one at 'path'.
        """
        self._prepare_data()

        if existing_excel and os.path.exists(existing_excel):
            self._append_new_invoices_to_workbook(existing_excel, path)
            return

        self._add_totals(["net_weight", "value_ron", "statistic"], group_by="vat_number")
        self._add_excel_formulas()

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
        self.data["vat_number"] = self.data["vat_number"].astype(str)
        self.data["invoice_number"] = pd.to_numeric(self.data["invoice_number"], errors="coerce").fillna(0).astype(int)

        self.data["shipment_date"] = self.data["shipment_date"].apply(convert_to_date)
        self.data["nc8_code"] = self.data["nc8_code"].apply(format_nc8_code)

        self.data = self.data.sort_values(by=["vat_number", "shipment_date", "invoice_number"],
                                          kind="stable").reset_index(drop=True)

        # Nr Crt counts invoices inside a VAT group (rows of one invoice share the number), as in the customer's files
        self.data.insert(0, "nr_crt", self.data.groupby("vat_number")["invoice_number"]
                         .transform(lambda s: s.ne(s.shift()).cumsum()))
        self.data["percentage"] = self.percentage

        for col in ["transport", "statistic"]:
            if col not in self.data.columns:
                self.data[col] = ""
            else:
                self.data[col] = self.data[col].fillna("")

    def _append_new_invoices_to_workbook(self, src_path, dest_path):
        """
        Append the invoices not yet in the workbook into a copy of it (dest_path), keeping the customer's rows,
        formats and formulas. Every row insert/delete is recorded as an event; at the end each pre-existing
        formula is re-pointed by replaying those events, so formulas outside the table (e.g. a hand-made grand
        total under it) stay correct. The rows actually written are left in self.appended.
        """
        wb = openpyxl.load_workbook(src_path)
        ws = self._sheet_with_header(wb)
        header_map = self._map_headers(ws)
        formats = Constants.COLUMN_FORMATS

        events = []  # ("insert" | "delete", row), in the order applied to ws
        snapshot = [(c.row, c.column, c.value) for row in ws.iter_rows() for c in row
                    if isinstance(c.value, str) and c.value.startswith("=")]

        old_totals = self._remove_total_rows(ws, header_map, events)
        struct = self._find_data_block(ws, header_map)
        existing = self._collect_existing_keys(ws, struct, header_map)

        df = self.data.dropna(subset=['invoice_number'])
        # an invoice already in the workbook is left untouched, whatever its rows look like now
        new_records = [row for _, row in df.iterrows() if (row.vat_number, row.invoice_number) not in existing]
        self.appended = pd.DataFrame(new_records, columns=self.data.columns)
        if not new_records:
            shutil.copyfile(src_path, dest_path)  # nothing to add: the output is the input, unchanged
            return

        inserted = []
        for rec in new_records:
            row_idx = self._find_insert_rows(ws, rec, struct, header_map)
            ws.insert_rows(row_idx)
            events.append(("insert", row_idx))
            inserted.append((len(events), row_idx, rec))
            self._write_record(ws, row_idx, rec, header_map)

        struct = self._find_data_block(ws, header_map)
        self._recompute_nr_crt(ws, struct, header_map)
        new_totals = self._insert_total_rows(ws, struct, header_map, events)

        # formulas of the new rows: written once the structure is final, so they reference their own row
        for done, row_idx, rec in inserted:
            final_row = self._replay_row(row_idx, events[done:])
            formulas = self._build_formulas_for_row(final_row, rec)
            for field, col in header_map.items():
                cell = ws.cell(row=final_row, column=col)
                if field in formulas:
                    cell.value = formulas[field]
                if field in formats:
                    cell.number_format = formats[field]

        # pre-existing formulas: move with their cell and follow the rows they point at
        total_moves = {old_row: new_totals.get(vat) for old_row, vat in old_totals.items()}
        for row, col, formula in snapshot:
            if row in old_totals:
                continue  # total rows are regenerated
            ws.cell(row=self._replay_row(row, events), column=col).value = self._repoint(formula, events, total_moves)

        wb.active = wb.worksheets.index(ws)  # open on the sheet that received the rows
        wb.save(dest_path)

    def _sheet_with_header(self, wb):
        """The worksheet that carries the invoice table header, whichever sheet is active."""
        for ws in wb.worksheets:
            if {"vat_number", "invoice_number", "nr_crt"} <= set(self._map_headers(ws)):
                return ws
        raise ValueError("Fișierul Excel ales nu are nicio foaie cu antetul așteptat "
                         "(Nr Crt, Firma, Nr Factura Marfa, ...).")

    @staticmethod
    def _replay_row(row, events, range_start=False):
        """Where a row that existed before `events` sits after them (callers skip deleted rows).
        A range starting at the table top (row 2) stays put when a row is inserted there, so a SUM over the
        whole table grows; every other reference moves exactly as Excel would move it."""
        for kind, pos in events:
            if kind == "insert":
                if row > pos or (row == pos and not (range_start and pos == 2)):
                    row += 1
            elif row > pos:
                row -= 1
        return row

    # ponytail: A1 references only; sheet-qualified refs are left alone, quoted text is skipped as a whole
    _CELL_REF = re.compile(r'("(?:[^"]|"")*")|'
                           r"(?<![A-Za-z0-9_.!])(\$?[A-Z]{1,3}\$?)(\d+)(?![0-9(])(?::(\$?[A-Z]{1,3}\$?)(\d+))?")

    @classmethod
    def _repoint(cls, formula, events, total_moves):
        """Shift every row reference of a pre-existing formula through `events`; refs to a removed total row
        follow that VAT group's re-created total row."""

        def moved(row, range_start=False):
            return total_moves.get(row) or cls._replay_row(row, events, range_start)

        def shift(match):
            if match.group(1) is not None:
                return match.group(1)  # string literal, e.g. a VAT number in a SUMIF criterion
            col1, row1, col2, row2 = match.groups()[1:]
            if col2 is None:
                return f"{col1}{moved(int(row1))}"
            return f"{col1}{moved(int(row1), range_start=True)}:{col2}{moved(int(row2))}"

        return cls._CELL_REF.sub(shift, formula)

    def _add_excel_formulas(self):
        """
        Add in-cell formulas for computed columns.
        """
        for i, row in self.data.iterrows():
            excel_row = i + 2

            formulas = self._build_formulas_for_row(excel_row, row.to_dict())

            for col, formula in formulas.items():
                self.data.at[i, col] = formula

    def _add_totals(self, columns_to_total, group_by=None):
        """
        Insert total rows after each group or at the end of the DataFrame.
        """
        if not self.data.empty:
            if group_by:
                total_rows = []
                current_row_offset = 1

                grouped = self.data.groupby(group_by, sort=False, as_index=False)

                for name, group in grouped:
                    group_size = len(group)

                    total_row = {col: "" for col in self.data.columns}
                    total_row[group_by] = f"Total {name}"
                    total_row["nr_crt"] = ""

                    for col in columns_to_total:
                        if col in self.data.columns:
                            col_letter = chr(65 + self.data.columns.get_loc(col))
                            start_row = current_row_offset + 1
                            end_row = start_row + group_size - 1
                            total_row[col] = f"=SUM({col_letter}{start_row}:{col_letter}{end_row})"

                    current_row_offset += group_size + 1
                    total_rows.append(total_row)

                result_df = pd.DataFrame()
                for (_, group), total_row in zip(grouped, total_rows):
                    result_df = pd.concat([result_df, group.reset_index(drop=True), pd.DataFrame([total_row])],
                                          ignore_index=True)

                self.data = result_df
            else:
                total_row = {col: "" for col in self.data.columns}
                total_row["nr_crt"] = "Total"

                for col in columns_to_total:
                    if col in self.data.columns:
                        col_letter = chr(65 + self.data.columns.get_loc(col))
                        start_row = 2
                        end_row = len(self.data) + 1
                        total_row[col] = f"=SUM({col_letter}{start_row}:{col_letter}{end_row})"

                self.data = pd.concat([self.data, pd.DataFrame([total_row])], ignore_index=True)

    @staticmethod
    def _write_headers(ws):
        """
        Write column headers, apply style, and freeze the header row.
        """
        header_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
        header_font = Font(bold=True, size=Constants.FONT_SIZE, name="Arial")
        header_alignment = Alignment(horizontal="left", vertical="bottom", wrap_text=True)

        for col_num, header in enumerate(Constants.HEADERS.values(), 1):
            cell = ws.cell(row=1, column=col_num, value=header)
            cell.font = header_font
            cell.alignment = header_alignment
            cell.fill = header_fill

        ws.freeze_panes = "A2"

    def _write_rows(self, ws):
        """
        Write DataFrame rows, including formulas, and adjust column widths.
        """
        headers_keys = list(Constants.HEADERS.keys())
        formats = Constants.COLUMN_FORMATS
        regular_font = Font(size=12, name="Arial")
        bold_font = Font(bold=True, size=12, name="Arial")

        col_max_lengths = {col_num: len(header) for col_num, header in enumerate(Constants.HEADERS.values(), 1)}

        for i, row in self.data.iterrows():
            is_total_row = pd.isna(row["nr_crt"]) or str(row["nr_crt"]).strip() == ""
            excel_row = i + 2

            for col_num, cell_value in enumerate(row, 1):
                if cell_value is not None:
                    col_max_lengths[col_num] = max(col_max_lengths[col_num], len(str(cell_value)))

                cell = ws.cell(row=excel_row, column=col_num, value=cell_value)
                header_key = headers_keys[col_num - 1]

                cell.font = bold_font if is_total_row else regular_font

                if header_key in formats:
                    cell.number_format = formats[header_key]

        scaling = Constants.FONT_SIZE / 10 * Constants.SCALING_FACTOR
        for col_num, max_len in col_max_lengths.items():
            col_letter = get_column_letter(col_num)
            adjusted_width = max_len * scaling + 2
            ws.column_dimensions[col_letter].width = adjusted_width

    @staticmethod
    def _map_headers(ws):
        """
        Map header titles to column indices.
        """
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
                if current_vat and current_vat in struct['vat_groups'] and struct['vat_groups'][current_vat].get(
                        'end') is None:
                    struct['vat_groups'][current_vat]['end'] = row_idx
                current_vat = vat
                struct['vat_groups'][vat] = {'start': row_idx, 'end': None}

        if struct['blank_row'] is None:
            struct['blank_row'], struct['data_end'] = ws.max_row + 1, ws.max_row

        data_end = struct['data_end']
        for vat, group in struct['vat_groups'].items():
            if group.get('end') is None:
                group['end'] = data_end + 1

        return struct

    @staticmethod
    def _collect_existing_keys(ws, struct, cmap):
        """Collect the (vat, invoice) pairs already in the workbook."""
        keys = set()
        for r in range(struct['data_start'], struct['data_end'] + 1):
            vat = ws.cell(r, cmap['vat_number']).value
            inv = ws.cell(r, cmap['invoice_number']).value
            if vat and inv and not str(vat).startswith('Total '):
                keys.add((str(vat), int(inv)))
        return keys

    @staticmethod
    def _find_insert_rows(ws, rec, struct, cmap):
        """
        Determine the correct row for inserting a new record, updating struct accordingly.
        """
        vat = rec.vat_number

        target_date = ExcelGenerator._normalize_date(rec.shipment_date)
        target = (target_date, rec.invoice_number)

        grp = struct['vat_groups'].get(vat)
        if not grp or not grp.get('end'):
            insert_row = ExcelGenerator._find_new_vat_position(vat, struct)

            ExcelGenerator._update_group_positions(struct, insert_row, vat)

            struct['vat_groups'][vat] = {'start': insert_row, 'end': insert_row + 1}

            struct['data_end'] += 1
            struct['blank_row'] = struct['data_end'] + 1

            return insert_row

        start, end = grp['start'], grp['end']
        insert_row = ExcelGenerator._find_position_in_group(ws, cmap, start, end, target)

        if insert_row is None:
            insert_row = end

        ExcelGenerator._update_group_positions(struct, insert_row, vat)

        struct['vat_groups'][vat]['end'] += 1

        struct['data_end'] += 1
        struct['blank_row'] = struct['data_end'] + 1

        return insert_row

    @staticmethod
    def _normalize_date(date_value):
        """Convert various date formats to a standard date object."""
        if isinstance(date_value, (pd.Timestamp, datetime)):
            return date_value.date()
        elif isinstance(date_value, date):
            return date_value
        return date_value

    @staticmethod
    def _find_new_vat_position(vat, struct):
        """Find the position for a new VAT group."""
        vats = [(v, g['start']) for v, g in struct['vat_groups'].items() if g.get('start') is not None]

        vats.sort(key=lambda x: x[0])

        for v_existing, start_row in vats:
            if v_existing > vat:
                return start_row

        return struct['data_end'] + 1

    @staticmethod
    def _update_group_positions(struct, insert_row, current_vat=None):
        """Update group positions after insertion."""
        for other_v, other_grp in struct['vat_groups'].items():
            if current_vat and other_v == current_vat:
                continue

            if other_grp.get('start') is not None and other_grp['start'] >= insert_row:
                other_grp['start'] += 1

            if other_grp.get('end') is not None and other_grp['end'] >= insert_row:
                other_grp['end'] += 1

    @staticmethod
    def _find_position_in_group(ws, cmap, start, end, target):
        """Find the correct position within a VAT group based on date and invoice number."""
        target_date, target_invoice = target

        for r in range(start, end):
            date_val = ExcelGenerator._get_cell_date(ws, r, cmap['shipment_date'])
            if date_val is None:
                continue

            try:
                inv = int(ws.cell(r, cmap['invoice_number']).value or 0)
            except (ValueError, TypeError):
                continue

            if (date_val, inv) > (target_date, target_invoice):  # rows of one invoice stay in extraction order
                return r

        return None

    @staticmethod
    def _get_cell_date(ws, row, col):
        """Extract and normalize a date value from a cell."""
        raw = ws.cell(row, col).value

        if isinstance(raw, datetime):
            return raw.date()
        elif isinstance(raw, date):
            return raw
        elif isinstance(raw, str):
            try:
                return date.fromisoformat(raw)
            except ValueError:
                pass

        return None

    @staticmethod
    def _write_record(ws, row, rec, cmap):
        """
        Write invoice record into worksheet row.
        """
        regular_font = Font(size=12, name='Arial')

        for field, col in cmap.items():
            val = rec.get(field)

            if field == 'delivery_location' and str(val) == '0':
                alt = next((v for v in rec.delivery_location if v != '0'), None)
                if alt:
                    val = alt

            cell = ws.cell(row=row, column=col, value=val)
            cell.font = regular_font

    @staticmethod
    @lru_cache(maxsize=128)
    def _get_cell_references(row_idx):
        """
        Get Excel cell references for a given row index.
        """
        return {"exchange_rate": f"J{row_idx}", "value_eur": f"G{row_idx}", "value_ron": f"K{row_idx}",
                "net_weight": f"H{row_idx}", "percentage": f"O{row_idx}", "transport": f"P{row_idx}"}

    @staticmethod
    def _build_formulas_for_row(row_idx, row):
        """
        Generate Excel formulas for a given row index and data row.
        """
        if pd.isna(row.get("nr_crt")) or str(row.get("nr_crt")).strip() == "":
            return {}

        cells = ExcelGenerator._get_cell_references(row_idx)

        formulas = {}

        val_eur = row.get("invoice_value_eur", 0)
        val_ron = row.get("value_ron", 0)

        if val_ron == 0 or (val_eur != 0 and val_ron != 0):
            formulas["value_ron"] = f"={cells['value_eur']}*{cells['exchange_rate']}"
        elif val_eur == 0 and val_ron != 0:
            formulas["invoice_value_eur"] = f"={cells['value_ron']}/{cells['exchange_rate']}"

        if pd.notna(row.get("net_weight")) and pd.notna(row.get("exchange_rate")):
            formulas["transport"] = f"=28000*{cells['exchange_rate']}/147000*{cells['net_weight']}"
        else:
            formulas["transport"] = ""

        if pd.notna(val_ron):
            formulas["statistic"] = f"=ROUND({cells['value_ron']}+{cells['transport']}*{cells['percentage']}, 0)"
        else:
            formulas["statistic"] = ""

        return formulas

    @staticmethod
    def _remove_total_rows(ws, header_map, events):
        """Remove all 'Total {vat_number}' rows; return {original row: vat} and record the deletions."""
        vat_col = header_map['vat_number']
        totals = {}

        for row in range(2, ws.max_row + 1):
            cell_value = ws.cell(row=row, column=vat_col).value
            if isinstance(cell_value, str) and cell_value.startswith('Total '):
                totals[row] = cell_value.split(' ', 1)[1]

        for row in sorted(totals, reverse=True):
            ws.delete_rows(row)
            events.append(("delete", row))
        return totals

    @staticmethod
    def _recompute_nr_crt(ws, struct, header_map):
        """Recompute and rewrite the nr_crt sequence for each VAT block, starting at 1 (one number per invoice)."""
        nr_crt_col = header_map['nr_crt']
        vat_col = header_map['vat_number']
        invoice_col = header_map['invoice_number']

        for vat, group in struct['vat_groups'].items():
            start = group.get('start')
            end = group.get('end')

            if not start or not end:
                continue

            counter, previous_invoice = 0, object()

            for row in range(start, end):
                cell_value = ws.cell(row=row, column=vat_col).value
                if isinstance(cell_value, str) and cell_value.startswith('Total '):
                    continue

                invoice = ws.cell(row=row, column=invoice_col).value
                if invoice != previous_invoice:
                    counter += 1
                    previous_invoice = invoice
                ws.cell(row=row, column=nr_crt_col).value = counter

    @staticmethod
    def _insert_total_rows(ws, struct, header_map, events):
        """Insert a 'Total {vat_number}' row after each VAT block with SUM formulas; return {vat: row}."""
        vat_col = header_map['vat_number']
        created = {}

        vat_items = [(vat, grp['start'], grp['end']) for vat, grp in struct['vat_groups'].items() if
                     grp.get('start') and grp.get('end')]
        vat_items.sort(key=lambda x: x[1])

        total_inserted = 0

        for vat, orig_start, orig_end in vat_items:
            new_start = orig_start + total_inserted
            new_end = orig_end + total_inserted

            ws.insert_rows(new_end)
            events.append(("insert", new_end))
            created[vat] = new_end
            total_inserted += 1

            label_cell = ws.cell(row=new_end, column=vat_col, value=f"Total {vat}")
            label_cell.font = Font(bold=True, size=12, name='Arial')

            for field in ('net_weight', 'value_ron', 'statistic'):
                if field in header_map:
                    col_idx = header_map[field]
                    col_letter = get_column_letter(col_idx)
                    sum_formula = f"=SUM({col_letter}{new_start}:{col_letter}{new_end - 1})"
                    sum_cell = ws.cell(row=new_end, column=col_idx, value=sum_formula)
                    sum_cell.font = Font(bold=True, size=12, name='Arial')
                    if field in Constants.COLUMN_FORMATS:
                        sum_cell.number_format = Constants.COLUMN_FORMATS[field]
        return created
