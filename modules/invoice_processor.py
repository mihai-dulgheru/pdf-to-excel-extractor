import concurrent.futures
import logging
import os
import re
from datetime import datetime

import pandas as pd
import pdfplumber

from config import Constants
from functions import ExchangeRateUnavailable, calculate_coordinates, get_country_code_from_address, \
    get_bnr_exchange_rate, get_delivery_location, get_previous_workday, parse_mixed_number, get_logger

logger = get_logger("pdf_to_excel.invoice_processor")


class _ProblemCollector(logging.Handler):
    """Keeps every WARNING+ message emitted during one run, so the UI can show them."""

    def __init__(self):
        super().__init__(level=logging.WARNING)
        self.messages = []

    def emit(self, record):
        self.messages.append(record.getMessage())


class InvoiceProcessor:
    """
    Extracts invoice data from PDF files into a pandas DataFrame.
    """

    __slots__ = ("input_paths", "df", "progress_callback", "problems", "fatal_error")

    # English and German (Porsche "Gutschrift") labels of the same fields
    VAT_PATTERN = r"(?:Tax number|KUNDEN UST-IDNR)\s*:\s*(\w+)"
    DESTINATION_PATTERN = r"(?:Invoiced to|Fakturiert zu)\s*:\s*(.+?)\n(?:Credit transfer|Bankverbindung)"
    INCOTERMS_PATTERN = r"Incoterms?\s*:\s*(\w+)"

    def __init__(self, paths, progress_callback=None):
        """
        :param paths: List of PDF file paths.
        :param progress_callback: Callable or None to report progress.
        """
        self.input_paths = paths
        self.df = pd.DataFrame(columns=Constants.COLUMNS)
        self.progress_callback = progress_callback
        self.problems = []
        self.fatal_error = None

    def process_invoices(self):
        """
        Parse each PDF file in parallel, then aggregate it results in self.df.
        Per-file problems are collected in self.problems; a missing BNR rate aborts the run (self.fatal_error).
        """
        collector = _ProblemCollector()
        logger.addHandler(collector)
        try:
            total = len(self.input_paths)
            results = []
            processed = 0
            owner_of = {}  # (vat, invoice) -> first file that produced it
            duplicates = {}  # (file, first file) -> invoices seen twice

            with concurrent.futures.ThreadPoolExecutor() as executor:
                future_map = {executor.submit(self._process_single_invoice, path): path for path in self.input_paths}

                for future in concurrent.futures.as_completed(future_map):
                    path = future_map[future]
                    try:
                        for row in future.result():
                            # unreadable files all say "Unknown": never treat two of them as one invoice
                            key = path if row["invoice_number"] == "Unknown" else (row["vat_number"], row["invoice_number"])
                            first = owner_of.setdefault(key, path)
                            if first == path:
                                results.append(row)
                            else:
                                duplicates.setdefault((path, first), set()).add(row["invoice_number"])
                    except ExchangeRateUnavailable as e:
                        self.fatal_error = str(e)
                        logger.error(self.fatal_error)
                        executor.shutdown(wait=False, cancel_futures=True)
                        break
                    except Exception as e:
                        logger.error(f"{os.path.basename(path)}: {e}")
                    processed += 1
                    if self.progress_callback:
                        self.progress_callback(processed, total)

            for (path, first), invoices in duplicates.items():
                logger.warning(f"{os.path.basename(path)}: factura {', '.join(sorted(invoices))} apare și în "
                               f"{os.path.basename(first)}; a fost păstrată o singură dată")

            if self.fatal_error is None:
                self.df = pd.DataFrame(results, columns=Constants.COLUMNS)
        except Exception as e:
            logger.error(f"Error in process_invoices: {e}")
        finally:
            logger.removeHandler(collector)
            self.problems = collector.messages

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
                if not first_page.chars:
                    raise ValueError("PDF scanat sau fără text; are nevoie de OCR")

                s1 = InvoiceProcessor._extract_section_text(first_page, "section_1", page_width, page_height)
                s2 = InvoiceProcessor._extract_section_text(first_page, "section_2", page_width, page_height)
                s3 = InvoiceProcessor._extract_section_text(first_page, "section_3", page_width, page_height)
                s4 = InvoiceProcessor._extract_section_text(first_page, "section_4", page_width, page_height)

                is_credit_note = InvoiceProcessor._is_credit_note(s1)
                is_debit_note = "DEBIT NOTE" in s1.upper() if s1 else False

                company = InvoiceProcessor._extract_company(s2)
                invoice_number = InvoiceProcessor._extract_invoice_number(s1)
                origin = InvoiceProcessor._extract_origin(s4)
                nc8_data = InvoiceProcessor._extract_nc8_codes(pdf, first_page, page_width, page_height, is_credit_note,
                                                               origin)

                destination_field = InvoiceProcessor._extract_field(s3, InvoiceProcessor.DESTINATION_PATTERN, "Unknown",
                                                                    re.DOTALL)
                destination = get_country_code_from_address(destination_field)

                invoice_value_eur, invoice_value_ron, currency = InvoiceProcessor._extract_invoice_values(last_page,
                                                                                                          page_width,
                                                                                                          page_height,
                                                                                                          is_credit_note)

                collapsed = InvoiceProcessor._collapse_unparsed(
                    nc8_data, invoice_value_eur if currency == "EUR" else invoice_value_ron)
                if collapsed is not nc8_data:
                    logger.warning(f"{os.path.basename(pdf_path)}: valorile pe articole nu au putut fi citite; "
                                   f"factura a fost scrisă pe un singur rând")
                    nc8_data = collapsed

                if nc8_data[0][0] == "INTERNAL ORDER":
                    total_net_weight = 0
                else:
                    total_net_weight = InvoiceProcessor._extract_net_weight(last_page, is_credit_note or is_debit_note)

                shipment_date = InvoiceProcessor._extract_shipment_date(s1)
                exchange_rate = get_bnr_exchange_rate(get_previous_workday(shipment_date))
                vat_number = InvoiceProcessor._extract_field(s3, InvoiceProcessor.VAT_PATTERN, "Unknown", re.IGNORECASE)
                delivery_location = get_delivery_location(s1)
                delivery_condition = InvoiceProcessor._extract_field(s2, InvoiceProcessor.INCOTERMS_PATTERN, "Unknown",
                                                                     re.IGNORECASE)

                if len(nc8_data) == 1:
                    code, _, code_origin = nc8_data[0]
                    return [
                        {"company": company, "invoice_number": invoice_number, "nc8_code": code, "origin": code_origin,
                         "destination": destination, "invoice_value_eur": invoice_value_eur,
                         "net_weight": total_net_weight, "shipment_date": shipment_date.strftime("%d.%m.%Y"),
                         "exchange_rate": exchange_rate, "value_ron": invoice_value_ron, "vat_number": vat_number,
                         "delivery_location": delivery_location, "delivery_condition": delivery_condition}]
                else:
                    per_code_total = sum(val for _, val, _ in nc8_data)
                    invoice_total = invoice_value_eur if currency == "EUR" else invoice_value_ron
                    if abs(per_code_total - invoice_total) > 0.01:
                        logger.warning(f"{os.path.basename(pdf_path)}: valorile pe cod însumează "
                                       f"{per_code_total:.2f}, totalul facturii este {invoice_total:.2f}")
                    values_only = [(code, val) for (code, val, _) in nc8_data]
                    proportional_weights = InvoiceProcessor._calculate_proportional_weights(values_only,
                                                                                            total_net_weight)

                    results = []
                    for (code, partial_value, code_origin), (_, proportional_weight) in zip(nc8_data,
                                                                                            proportional_weights):
                        if currency == "EUR":
                            invoice_value_eur_for_code = partial_value
                            invoice_value_ron_for_code = partial_value * exchange_rate
                        elif currency == "RON":
                            invoice_value_ron_for_code = partial_value
                            invoice_value_eur_for_code = partial_value / exchange_rate if exchange_rate else 0
                        else:
                            invoice_value_eur_for_code = 0
                            invoice_value_ron_for_code = 0

                        result = {"company": company, "invoice_number": invoice_number, "nc8_code": code,
                                  "origin": code_origin, "destination": destination,
                                  "invoice_value_eur": invoice_value_eur_for_code, "net_weight": proportional_weight,
                                  "shipment_date": shipment_date.strftime("%d.%m.%Y"), "exchange_rate": exchange_rate,
                                  "value_ron": invoice_value_ron_for_code, "vat_number": vat_number,
                                  "delivery_location": delivery_location, "delivery_condition": delivery_condition}
                        results.append(result)

                    return results
        except ExchangeRateUnavailable:
            raise
        except Exception as e:
            logger.error(f"{os.path.basename(pdf_path)}: fișierul nu a putut fi citit ({e})")
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
            logger.error(f"Error extracting section {section_name}: {e}")
            return ""

    @staticmethod
    def _collapse_unparsed(nc8_data, invoice_total):
        """
        No item value could be read although the document has a total: return one row with the codes joined,
        so the printed total is kept instead of several rows of 0. Otherwise return nc8_data unchanged.
        """
        if len(nc8_data) > 1 and abs(sum(val for _, val, _ in nc8_data)) < 0.01 and abs(invoice_total) > 0.01:
            return [("; ".join(dict.fromkeys(code for code, _, _ in nc8_data)), 0, nc8_data[0][2])]
        return nc8_data

    @staticmethod
    def _is_credit_note(section_1_text):
        """True for an English "Credit Note" or a German "Gutschrift" document."""
        text = (section_1_text or "").upper()
        return "CREDIT NOTE" in text or "GUTSCHRIFT" in text

    @staticmethod
    def _extract_company(text):
        """
        Return the company name found after the 'ORDERED BY' (German: 'Bestellt von') line, otherwise 'Unknown'.
        A trailing comma is dropped: it only announces the address printed on the next line.
        """
        if not text:
            return "Unknown"
        lines = text.split("\n")
        for i, line in enumerate(lines):
            if ("ORDERED BY" in line or "Bestellt von" in line) and i + 1 < len(lines):
                return lines[i + 1].strip().rstrip(" ,")
        return "Unknown"

    @staticmethod
    def _extract_field(text, pattern, default, flags=0):
        """
        Extract a single value from a text using a regex pattern.
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
            logger.error(f"Error extracting invoice number: {e}")
            return "Unknown"

    @staticmethod
    def _extract_nc8_codes(pdf, first_page, page_width, page_height, is_credit_note, default_origin):
        """
        Extract NC-8 codes together with the corresponding value for each code and origin.
        Credit notes are split per code like invoices; their values are negated.
        """
        if not is_credit_note:
            s4 = InvoiceProcessor._extract_section_text(first_page, "section_4", page_width, page_height)
            if "REFERENCE" in s4 and "INTERNAL ORDER" in s4:
                return [("INTERNAL ORDER", 0, default_origin)]

        # Only "<currency> <net value>" is required (pdfplumber may push the VAT amount onto a third line).
        # The currency token may come out glued to the unit ("PCEUR", "PECUR"); the net value may group thousands
        # with a space ("1 041,40") or with ,/. ("1,041.40", "1.041,40") and may lose its last digit to a space
        # ("2.988,6 4" = 2.988,64).
        currency_pattern = re.compile(r"\b(PCEUR|PECUR|EUR|RON)\b\s+"
                                      r"(\d{1,3}(?: \d{3})+[.,]\d{2}|[\d.,]+[.,]\d \d(?=\s|$)|[\d.,]+)", re.IGNORECASE)
        specialized_pattern = re.compile(
            r"^[A-Za-z0-9]+\s+PER\s+(?:\d{1,3}(?:[.,]\d{3})+|\d+)\s+PC\s+\d+PC\s+([\d.,]+)\s+([\d.,]+)$")
        code_pattern = re.compile(r"Commodity Code\s*:\s*(\d*)", re.IGNORECASE)  # value may be blank
        purch_pattern = re.compile(r"purch\. order no\.", re.IGNORECASE)
        origin_pattern = re.compile(r"Country of origin\s*:[ \t]*([A-Za-z]{2})\b", re.IGNORECASE)

        lines = [ln.strip() for page in pdf.pages for ln in (page.extract_text() or "").splitlines() if ln.strip()]

        pairs, current_val = [], None
        i, n = 0, len(lines)

        def origin_after(idx):
            """Origin printed for the item whose 'Commodity Code' line is lines[idx] (None when blank)."""
            for following in lines[idx + 1:idx + 4]:
                if code_pattern.search(following):
                    return None
                if m_origin := origin_pattern.search(following):
                    return m_origin.group(1).upper()
            return None

        while i < n:
            line = lines[i]
            joined = f"{line} {lines[i + 1]}" if i + 1 < n else line

            if m := currency_pattern.search(joined):
                current_val = parse_mixed_number(m.group(2))
                i += 1
                continue

            if m := specialized_pattern.match(line):
                current_val = parse_mixed_number(m.group(2))

                purch_seen = False
                nc8_code = None
                for j in range(i + 1, min(i + 12, n)):
                    l = lines[j]
                    if purch_pattern.search(l):
                        purch_seen = True
                    if not nc8_code and (cm := code_pattern.search(l)):
                        nc8_code = cm.group(1) or "-"

                    if purch_seen and nc8_code:
                        pairs.append((nc8_code, current_val, origin_after(j)))
                        current_val = None
                        i = j
                        break

            elif cm := code_pattern.search(line):
                # an item printed without a code keeps its value under "-", so the invoice total still closes
                pairs.append((cm.group(1) or "-", current_val if current_val is not None else 0.0, origin_after(i)))
                current_val = None

            i += 1

        if not pairs:
            return [("Credit Note" if is_credit_note else "Unknown", 0, default_origin)]

        first_origin = next((origin for _, _, origin in pairs if origin), default_origin)
        sign = -1 if is_credit_note else 1
        by_code = {}  # code (first seen) -> origin (first seen) -> value; keeps a code's origin rows together
        for code, val, origin in pairs:
            origins = by_code.setdefault(code, {})
            origin = origin or first_origin
            origins[origin] = origins.get(origin, 0.0) + sign * val

        if "-" in by_code and len(by_code) == 2:
            # a single real code: the items printed without one belong to it (what the customer does by hand);
            # with several codes the "-" row stays visible so he can place it
            real = next(code for code in by_code if code != "-")
            for origin, val in by_code.pop("-").items():
                by_code[real][origin] = by_code[real].get(origin, 0.0) + val

        return [(code, val, origin) for code, origins in by_code.items() for origin, val in origins.items()]

    @staticmethod
    def _extract_origin(section_4_text):
        """
        Extract the country code from 'Country of origin', or '-' if not found.
        """
        if not section_4_text:
            return "-"
        match = re.search(r"Country of origin\s*:[ \t]*([A-Za-z]{2})\b", section_4_text, re.IGNORECASE)
        return match.group(1).upper() if match else "-"

    @staticmethod
    def _extract_invoice_values(last_page, page_width, page_height, is_credit_note):
        """
        Extract invoice values (EUR, RON) and currency from the last page bounding box.
        Returns (value_eur, value_ron, "EUR"/"RON"/None).
        """
        try:
            coords = calculate_coordinates(page_width, page_height, (0.52, 0.88, 0.94, 0.94))
            text = (last_page.within_bbox(coords).extract_text() or "").strip()
            if not text:
                return 0, 0, None

            lines = [line.strip().replace('*', ' ') for line in text.splitlines() if
                     any(cur in line for cur in ("EUR", "RON"))]
            if not lines:
                return 0, 0, None

            normalized_line = lines[0].replace(" ", "")

            currency_str = None
            amount_str = ""

            for c in ("EUR", "RON"):
                if c in normalized_line:
                    currency_str = c.lower()
                    _, _, amount_str = normalized_line.partition(c)
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
            else:
                return 0, 0, None

        except Exception as e:
            logger.error(f"Error extracting invoice values: {e}")
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
            # "1.035 KG" beside a gross weight "4.218,300 KG" is 1,035 kg: when the net weight is printed without
            # decimals, the gross weight on the same page tells which separator groups thousands
            gross = re.search(r"Gross Weight\s+([\d.,]+)\s+KG", text, re.IGNORECASE)
            if gross and "." in gross.group(1) and "," in gross.group(1) and re.fullmatch(r"\d+[.,]\d{3}", raw_val):
                thousands_sep = "," if gross.group(1).rfind(".") > gross.group(1).rfind(",") else "."
                raw_val = raw_val.replace(thousands_sep, "")
            num_val = parse_mixed_number(raw_val)
            return round(num_val)
        return 0

    @staticmethod
    def _calculate_proportional_weights(nc8_data, total_weight):
        """
        Calculate proportional weights for each NC8 code based on their values.
        Returns a list of tuples (nc8_code, proportional_weight) whose weights add up to total_weight.
        """
        if not nc8_data or total_weight <= 0:
            return [(code, 0) for code, _ in nc8_data]

        total_value = sum(abs(value) for _, value in nc8_data)
        if total_value <= 0:
            return [(code, 0) for code, _ in nc8_data]

        # largest remainder: whole kilograms first, the leftover ones to the biggest fractions, so the rows add up
        # to the invoice weight and no row goes negative
        shares = [total_weight * abs(value) / total_value for _, value in nc8_data]
        weights = [int(share) for share in shares]
        leftover = max(0, round(total_weight) - sum(weights))
        for i in sorted(range(len(shares)), key=lambda i: weights[i] - shares[i])[:leftover]:
            weights[i] += 1
        return [(code, weight) for (code, _), weight in zip(nc8_data, weights)]

    @staticmethod
    def _extract_shipment_date(section_1_text):
        """
        Extract the shipment date (DD.MM.YYYY or DD.MM.YY) from section_1_text.
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
            logger.error(f"Error extracting shipment date: {e}")
            return datetime.now()
