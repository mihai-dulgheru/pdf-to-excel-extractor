import re


def get_country_code_from_address(address):
    """
    Extracts the country ISO 3166-1 Alpha-2 code from an address.
    """
    country_codes = {"RO": "Romania", "DE": "Germany", "FR": "France", "IT": "Italy", "ES": "Spain", "HU": "Hungary",
                     "PL": "Poland", "BG": "Bulgaria", "CZ": "Czech Republic", "SK": "Slovakia", "AT": "Austria",
                     "NL": "Netherlands", "BE": "Belgium", "LU": "Luxembourg", "UK": "United Kingdom",
                     "US": "United States", }

    # Check for VAT number format (e.g., RO15599111 -> RO)
    tax_match = re.search(r"\b([A-Z]{2})\d+", address)
    if tax_match:
        return tax_match.group(1)

    # Check for country name in the address
    for code, country in country_codes.items():
        if re.search(rf"\b{country}\b", address, re.IGNORECASE):
            return code

    return "Unknown"
