import re
from functools import lru_cache

import camelot
import openpyxl
import pandas as pd
import pdfplumber
from pdfplumber import open as open_pdf

from functions.get_country_code_from_address import get_country_code_from_address
from functions.get_loc_livrare import get_loc_livrare


def extract_pdf_data_to_excel(pdf_paths, excel_output_path):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Invoices"

    # Create headers
    headers = ["Nr. Crt.", "Firma", "Nr. Factura marfa", "Cod NC8", "Origine", "Destinatie", "Val. Fact. EURO",
               "Greutate neta", "Data expeditiei", "Curs valutar", "Valoare RON", "Vat/cumparator", "Loc livrare",
               "Conditie livrare", "%", "Transport", "Statistica"]
    sheet.append(headers)

    for idx, pdf_path in enumerate(pdf_paths):
        with open_pdf(pdf_path) as pdf:
            text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())
            print(text)

            # Extract fields based on specified patterns
            firma = re.search(r"Our payment address\n(.+)", text).group(1).strip() if re.search(
                r"Our payment address\n(.+)", text) else "Unknown"
            nr_factura = re.search(r"Number Date\n(.+)", text).group(1).strip() if re.search(r"Number Date\n(.+)",
                                                                                             text) else "Unknown"
            cod_nc8 = re.findall(r"Commodity Code : (\d+)", text)
            net_weight = re.search(r"Net weight (\d+.\d+) KG", text).group(1).strip() if re.search(
                r"Net weight (\d+.\d+) KG", text) else "0"
            net_to_pay_match = re.search(r"NET TO PAY\n(.+)", text)
            net_to_pay = net_to_pay_match.group(1).strip() if net_to_pay_match else "0 EUR"
            # Extract other fields as needed...

            # Handle currency conversion if necessary
            date_match = re.search(r"Number Date\n(\d{2}\.\d{2}\.\d{4})", text)
            date_of_invoice = datetime.strptime(date_match.group(1), "%d.%m.%Y") if date_match else datetime.now()
            exchange_rate = get_bnr_exchange_rate(date_of_invoice - timedelta(days=1))
            val_fact_eur = convert_to_eur(net_to_pay, exchange_rate)
            val_ron = float(val_fact_eur) * exchange_rate

            # Append data row to Excel
            sheet.append([idx + 1, firma, nr_factura, ", ".join(cod_nc8), "RO", "Destination Placeholder", val_fact_eur,
                          net_weight, date_of_invoice.strftime("%d.%m.%Y"), exchange_rate, val_ron, "Vat Placeholder",
                          "LOC LIVRARE Placeholder", "Conditie Placeholder", 0.64,
                          "=3200*{}*{}/18000*{}".format(exchange_rate, val_ron, net_weight),
                          '=ROUND({}+{}*{};0)'.format(val_ron, val_ron, 0.64)])

    workbook.save(excel_output_path)


import requests
from datetime import datetime, timedelta


@lru_cache(maxsize=128)
def get_bnr_exchange_rate(data_expeditiei, valuta="EUR"):
    """
    Obține cursul valutar RON-EUR sau RON altă monedă de la ECB pentru o anumită dată.
    Dacă nu există curs pentru data specificată, caută în zilele precedente.

    :param data_expeditiei: Data pentru care vrei să obții cursul (tip datetime).
    :param valuta: Moneda pentru care se dorește cursul (implicit "EUR").
    :return: Cursul valutar (float) sau None dacă nu se găsește.
    """
    max_attempts = 5  # Numărul maxim de zile în care să caute în trecut
    attempts = 0

    while attempts < max_attempts:
        data_anterioara = (data_expeditiei - timedelta(days=attempts + 1)).strftime("%Y-%m-%d")
        print(f"Fetching exchange rate for {valuta} on {data_anterioara}")

        try:
            # URL-ul către serviciul ECB pentru datele valutare
            ecb_api_url = "https://sdw-wsrest.ecb.europa.eu/service/data/EXR/D..EUR.SP00.A"
            query_params = {"startPeriod": data_anterioara, "endPeriod": data_anterioara, "format": "csvdata"}
            response = requests.get(ecb_api_url, params=query_params)

            if response.status_code == 200:
                # Procesăm răspunsul ca fișier CSV
                csv_response = response.text
                exchange_rate_lines = csv_response.splitlines()

                # Iterăm prin linii ignorând antetul (prima linie)
                for line in exchange_rate_lines[1:]:
                    columns = line.split(",")
                    currency = columns[2]  # Coloana `CURRENCY`
                    rate = columns[7]  # Coloana `OBS_VALUE`

                    if currency == valuta:
                        return float(rate)

                print(f"Exchange rate for {valuta} not found on {data_anterioara}")
            else:
                print(f"Error fetching exchange rate: {response.status_code} {response.reason}")
        except Exception as e:
            print(f"Exception occurred while fetching exchange rate: {e}")

        attempts += 1

    print(f"Could not fetch exchange rate for {valuta} after {max_attempts} attempts.")
    return None


def get_bnr_exchange_rate_dummy(data_expeditiei, valuta="EUR"):
    print(f"Using fallback exchange rate for {valuta} on {data_expeditiei}")
    return 4.9  # Exemplu pentru EUR


def convert_to_eur(value, exchange_rate):
    if "RON" in value:
        amount = float(re.sub(r"[^\d.]", "", value))
        return amount / exchange_rate
    return float(re.sub(r"[^\d.]", "", value))


def extract_tables_from_pdf(pdf_paths):
    result = []

    for pdf_path in pdf_paths:
        t = camelot.read_pdf(pdf_path, pages='all', flavor='stream')
        for table in t:
            result.append(table)

    return result


"""
    Format: A4
    -----------------------------------------
    |                                       |
    |               Section 1               |
    |                                       |
    -----------------------------------------
    |                   |                   |
    |                   |                   |
    |                   |                   |
    |                   |                   |
    |     Section 2     |     Section 3     |
    |                   |                   |
    |                   |                   |
    |                   |                   |
    |                   |                   |
    -----------------------------------------
    |                                       |
    |                                       |
    |                                       |
    |                                       |
    |               Section 4               |
    |                                       |
    |                                       |
    |                                       |
    |                                       |
    -----------------------------------------
    |               Section 5               |
    |                                       |
    -----------------------------------------
"""
# Define proportions for each section of the invoice
proportions = {"section_1": (0.0, 0.0, 1.0, 0.16), "section_2": (0.0, 0.16, 0.46, 0.54),
               "section_3": (0.46, 0.16, 1.0, 0.54), "section_4": (0.0, 0.54, 1.0, 0.93),
               "section_5": (0.0, 0.93, 1.0, 1.0), }


def calculate_coordinates(page_width, page_height, proportion):
    """
    Calculate coordinates based on page width, height and proportion
    :param page_width: The width of the page
    :param page_height: The height of the page
    :param proportion: The proportion of the page to calculate coordinates for
    :return: A tuple of coordinates (x0, y0, x1, y1)
    """
    x0 = proportion[0] * page_width
    y0 = proportion[1] * page_height
    x1 = proportion[2] * page_width
    y1 = proportion[3] * page_height
    return x0, y0, x1, y1


def round_to_n_decimals(number, n=2):
    """
    Round a number to n decimals
    :param number: The number to round
    :param n: The number of decimals to round to
    :return: The rounded number
    """
    return round(number, n)


if __name__ == '__main__':
    input_paths = ["pdfs/{}".format(pdf) for pdf in
                   ["0912626530_C015000044_ZSM0_001.PDF", "9610169997_2007000000_ZSM0_001.PDF"]]
    # output_path = "output/invoices.xlsx"
    # extract_pdf_data_to_excel(input_paths, output_path)
    # tables = extract_tables_from_pdf(input_paths)
    # pprint([table.df for table in tables])
    # for idx, table in enumerate(tables):
    #     table.to_csv("output/table_{}.csv".format(idx))

    df = pd.DataFrame(columns=[  # "nr_crt",
        "firma", "nr_factura_marfa", "cod_nc8", "origine", "destinatie", "val_fact_euro", "greutate_neta",
        "data_expeditiei", "curs_valutar", "valoare_ron", "vat_cumparator", "loc_livrare", "conditie_livrare",
        # "procent",
        # "transport",
        # "statistica"
    ])
    for idx, pdf_path in enumerate(input_paths):
        with pdfplumber.open(pdf_path) as pdf:
            first_page = pdf.pages[0]
            page_width = first_page.width
            page_height = first_page.height

            # Extract sections from the invoice based on the defined proportions and save them as images
            for section_name, proportion in proportions.items():
                section_coords = calculate_coordinates(page_width, page_height, proportion)
                section_image = first_page.within_bbox(section_coords).to_image()
                image_path = f"output/{section_name}.png"
                section_image.save(image_path)

            section_2_coords = calculate_coordinates(page_width, page_height, proportions["section_2"])
            section_2 = first_page.within_bbox(section_2_coords)
            section_2_text = section_2.extract_text()

            firma_match = re.search(r"Our payment address\n(.+)", section_2_text)
            firma = firma_match.group(1).strip() if firma_match else "Unknown"
            print("Firma:", firma)

            section_1_coords = calculate_coordinates(page_width, page_height, proportions["section_1"])
            section_1 = first_page.within_bbox(section_1_coords)
            section_1_text = section_1.extract_text()

            nr_factura_marfa = section_1_text.split("\n")[-1].split(" ")[0]
            print("Nr. Factura marfa:", nr_factura_marfa)

            cod_nc8 = []
            for page in pdf.pages:
                text = page.extract_text()
                cod_nc8 += re.findall(r"Commodity Code : (\d+)", text)
            print("Cod NC8:", cod_nc8)

            our_payment_address = re.search(r"Our payment address\n(.+?)\nPayment date", section_2_text,
                                            re.DOTALL).group(1).strip()
            origine = get_country_code_from_address(our_payment_address)
            print("Origine:", origine)

            section_3_coords = calculate_coordinates(page_width, page_height, proportions["section_3"])
            section_3 = first_page.within_bbox(section_3_coords)
            section_3_text = section_3.extract_text()
            invoiced_to = re.search(r"Invoiced to : (.+?)\nCredit transfer", section_3_text, re.DOTALL).group(1).strip()
            destinatie = get_country_code_from_address(invoiced_to)
            print("Destinatie:", destinatie)

            invoice_value_eur = 0
            invoice_value_ron = 0
            last_page = pdf.pages[-1]

            error_margin = 0.01
            bbox_ratio = (
                round_to_n_decimals(1125 / 1653 - error_margin), round_to_n_decimals(2054 / 2339 - error_margin),
                round_to_n_decimals(1552 / 1653 + error_margin), round_to_n_decimals(2182 / 2339 + error_margin),)
            bbox_coordinates = calculate_coordinates(page_width, page_height, bbox_ratio)

            # Extract text from the defined bounding box
            bbox_text = last_page.within_bbox(bbox_coordinates).extract_text()

            # Split the text into lines and process the last relevant line
            lines = bbox_text.splitlines()
            if lines:
                last_line = lines[-1]  # Assume the last line contains currency and value

                # Split the last line based on "*", if present
                if "*" in last_line:
                    currency = last_line.split("*")[0].strip().lower()
                    currency_value = last_line.split("*")[-1].strip()
                else:
                    print("No * symbol found in the last line:", last_line)
                    currency = None
                    currency_value = None

                # Assign values based on currency
                if currency == "eur":
                    # For EUR: Replace "," (thousands separator) with "", and "." (decimal separator) with "."
                    currency_value = currency_value.replace(",", "").replace(".", ".")
                    invoice_value_eur = float(currency_value)
                elif currency == "ron":
                    # For RON: Replace "." (thousands separator) with "", and "," (decimal separator) with "."
                    currency_value = currency_value.replace(".", "").replace(",", ".")
                    invoice_value_ron = float(currency_value)
                else:
                    print("Unknown currency:", currency)
            else:
                print("No text found in the bounding box.")

            print("Valoare factura in EUR:", invoice_value_eur)
            print("Valoare factura in RON:", invoice_value_ron)

            greutate_neta = 0
            last_page_text = last_page.extract_text()
            greutate_neta_match = re.search(r"Net weight\s+(.+?)\s+KG", last_page_text)
            if greutate_neta_match:
                raw_weight = greutate_neta_match.group(1).strip()
                if currency == "eur":
                    processed_weight = raw_weight.replace(",", "")
                elif currency == "ron":
                    processed_weight = raw_weight.replace(".", "").replace(",", ".")
                else:
                    processed_weight = raw_weight  # Dacă nu e definită moneda, folosește direct valoarea brută
                greutate_neta = float(processed_weight)

            print("Greutate neta:", greutate_neta)

            data_expeditiei_match = re.search(r"Transportation date: (\d{2}\.\d{2}\.\d{4})", last_page_text)
            data_expeditiei = datetime.strptime(data_expeditiei_match.group(1),
                                                "%d.%m.%Y") if data_expeditiei_match else datetime.now()
            print("Data expeditiei:", data_expeditiei.strftime("%d.%m.%Y"))

            exchange_rate = get_bnr_exchange_rate(data_expeditiei, "RON")
            print("Curs valutar:", exchange_rate)

            if exchange_rate:
                valoare_in_eur = round_to_n_decimals(
                    invoice_value_eur if invoice_value_eur else invoice_value_ron / exchange_rate)
                valoare_in_ron = round_to_n_decimals(
                    invoice_value_ron if invoice_value_ron else invoice_value_eur * exchange_rate)
                print(f"Valoare în RON: {valoare_in_ron}")
                print(f"Valoare în EUR: {valoare_in_eur}")
            else:
                print("Could not fetch exchange rate.")

            vat_cumparator_match = re.search(r"Tax number : (\w+)", section_3_text)
            vat_cumparator = vat_cumparator_match.group(1) if vat_cumparator_match else "Unknown"
            print("VAT cumparator:", vat_cumparator)

            loc_livrare = get_loc_livrare(section_1_text)
            print("Loc livrare:", loc_livrare)

            conditie_livrare_match = re.search(r"Incoterms : (\w+)", section_2_text)
            conditie_livrare = conditie_livrare_match.group(1) if conditie_livrare_match else "Unknown"
            print("Conditie livrare:", conditie_livrare)
