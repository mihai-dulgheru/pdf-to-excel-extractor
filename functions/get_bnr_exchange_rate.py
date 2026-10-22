import logging
from datetime import timedelta
from functools import lru_cache

import requests
from bs4 import BeautifulSoup

logger = logging.getLogger("pdf_to_excel.exchange_rate")

REQUEST_TIMEOUT_SECONDS = 15


class ExchangeRateUnavailable(RuntimeError):
    """
    No BNR rate could be fetched for a date. Raised (not returned as None) so that lru_cache does not
    remember the failure: a later call in the same session can succeed once the network is back.
    """


@lru_cache(maxsize=128)
def get_bnr_exchange_rate(invoice_date, currency="EUR"):
    """
    Retrieves the RON-to-EUR exchange rate from BNR data for a specific date.
    If the rate is not available for the given date, it searches for prior days.
    Raises ExchangeRateUnavailable when nothing could be fetched.
    """

    max_attempts = 5
    for attempt in range(max_attempts):
        date_to_check = (invoice_date - timedelta(days=attempt)).strftime("%Y-%m-%d")

        try:
            bnr_url = f"https://www.cursbnr.ro/arhiva-curs-bnr-{date_to_check}"
            response = requests.get(bnr_url, timeout=REQUEST_TIMEOUT_SECONDS)

            if response.status_code == 200:
                soup = BeautifulSoup(response.text, "html.parser")

                currency_row = soup.select_one("#table-currencies tbody tr:nth-child(1) td:nth-child(3)")

                if currency_row:
                    exchange_rate = float(currency_row.text.replace(",", "."))
                    return exchange_rate
                else:
                    logger.warning(f"Exchange rate not found for {date_to_check}.")

            else:
                logger.warning(f"Error fetching exchange rate from BNR: {response.status_code} {response.reason}")

        except Exception as e:
            logger.warning(f"Exception while fetching exchange rate: {e}")

    raise ExchangeRateUnavailable(f"Cursul BNR pentru {invoice_date:%d.%m.%Y} nu a putut fi obținut după "
                                  f"{max_attempts} încercări. Verifică conexiunea la internet și reîncearcă.")
