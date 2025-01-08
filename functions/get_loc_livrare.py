import re


def get_loc_livrare(section_1_text):
    """
    Determină `loc_livrare` pe baza textului din `section_1_text`.

    :param section_1_text: Textul din secțiunea 1 a facturii.
    :return: Valoarea pentru `loc_livrare` (string) sau 'Unknown' dacă nu se poate determina.
    """
    # 1. Verifică dacă există `Delivering plant : <ADRESA>`
    delivering_plant_match = re.search(r"Delivering plant : (.+?)$", section_1_text, re.MULTILINE)
    if delivering_plant_match:
        address = delivering_plant_match.group(1).strip()
        # Returnăm ultima parte a adresei (presupunând că este locația sucursalei)
        return address.split()[-1]

    # 2. Verifică dacă există `Our BAU Code : <CODE>`
    bau_code_match = re.search(r"Our BAU Code : ([A-Z0-9_]+)", section_1_text)
    if bau_code_match:
        bau_code = bau_code_match.group(1)
        # Maparea codurilor BAU către locații
        bau_code_to_location = {"RO03_E_CRA_1593": "CRAIOVA",  # Adaugă aici alte coduri BAU în viitor
                                }
        return bau_code_to_location.get(bau_code, "Unknown")

    # 3. Fallback: Returnează 'Unknown' dacă nu se găsește nimic
    return "Unknown"
