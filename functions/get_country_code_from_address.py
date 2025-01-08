import re


def get_country_code_from_address(address):
    """
    Identifică codul țării (ISO 3166-1 Alpha-2) dintr-o adresă.

    :param address: Textul adresei (string).
    :return: Codul țării (string) sau 'Unknown' dacă nu poate fi identificat.
    """
    # Dictionar mapare țări -> coduri ISO
    country_codes = {"RO": "Romania", "DE": "Germany", "FR": "France", "IT": "Italy", "ES": "Spain", "HU": "Hungary",
        "PL": "Poland", "BG": "Bulgaria", "CZ": "Czech Republic", "SK": "Slovakia", "AT": "Austria",
        "NL": "Netherlands", "BE": "Belgium", "LU": "Luxembourg", "UK": "United Kingdom", "US": "United States", }

    # 1. Caută cod fiscal (ex. RO15599111 -> RO)
    tax_match = re.search(r"\b([A-Z]{2})\d+", address)
    if tax_match:
        return tax_match.group(1)

    # 2. Verifică dacă numele țării este prezent
    for code, country in country_codes.items():
        if re.search(rf"\b{country}\b", address, re.IGNORECASE):
            return code

    # 3. Fallback: Returnează "Unknown" dacă țara nu poate fi identificată
    return "Unknown"
