from datetime import timedelta
from functools import lru_cache

import requests
from bs4 import BeautifulSoup


@lru_cache(maxsize=128)
def get_bnr_exchange_rate(invoice_date, currency="EUR"):
    """
    Retrieves the RON-to-EUR exchange rate from BNR data for a specific date.
    If the rate is not available for the given date, it searches for prior days.
    """

    max_attempts = 5
    for attempt in range(max_attempts):
        date_to_check = (invoice_date - timedelta(days=attempt)).strftime("%Y-%m-%d")

        try:
            bnr_url = f"https://www.cursbnr.ro/arhiva-curs-bnr-{date_to_check}"
            response = requests.get(bnr_url)

            if response.status_code == 200:
                soup = BeautifulSoup(response.text, "html.parser")

                currency_row = soup.select_one("#table-currencies tbody tr:nth-child(1) td:nth-child(3)")

                if currency_row:
                    exchange_rate = float(currency_row.text.replace(",", "."))
                    return exchange_rate
                else:
                    print(f"[LOG] Exchange rate not found for {date_to_check}.")

            else:
                print(f"[LOG] Error fetching exchange rate from BNR: {response.status_code} {response.reason}")

        except Exception as e:
            print(f"[LOG] Exception while fetching exchange rate: {e}")

    print(f"[LOG] Could not fetch exchange rate for {currency} after {max_attempts} attempts.")
    return None
