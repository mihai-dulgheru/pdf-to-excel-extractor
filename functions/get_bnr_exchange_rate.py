from datetime import timedelta
from functools import lru_cache

import requests


@lru_cache(maxsize=128)
def get_bnr_exchange_rate(expedition_date, currency="EUR"):
    """
    Retrieves the RON-to-currency exchange rate from ECB data for a specific date.
    If not available for the given date, it searches for prior days up to a maximum number of attempts.
    """
    max_attempts = 5
    for attempt in range(max_attempts):
        date_to_check = (expedition_date - timedelta(days=attempt + 1)).strftime("%Y-%m-%d")
        print(f"Fetching exchange rate for {currency} on {date_to_check}")

        try:
            ecb_api_url = "https://sdw-wsrest.ecb.europa.eu/service/data/EXR/D..EUR.SP00.A"
            query_params = {"startPeriod": date_to_check, "endPeriod": date_to_check, "format": "csvdata"}
            response = requests.get(ecb_api_url, params=query_params)

            if response.status_code == 200:
                csv_response = response.text
                exchange_rate_lines = csv_response.splitlines()

                for line in exchange_rate_lines[1:]:
                    columns = line.split(",")
                    if columns[2] == currency:  # Column `CURRENCY`
                        return float(columns[7])  # Column `OBS_VALUE`

                print(f"Exchange rate for {currency} not found on {date_to_check}")
            else:
                print(f"Error fetching exchange rate: {response.status_code} {response.reason}")
        except Exception as e:
            print(f"Exception while fetching exchange rate: {e}")

    print(f"Could not fetch exchange rate for {currency} after {max_attempts} attempts.")
    return None
