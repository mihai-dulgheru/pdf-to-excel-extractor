import re


def get_delivery_location(section_text):
    """
    Determines the delivery location code based on the content of the section.

    :param section_text: Extracted text from the document section.
    :return: The location code as an integer.
    """
    # Check for "Delivering plant : <ADDRESS>" (allowing variations in spacing)
    delivering_plant_match = re.search(r"Delivering\s+plant\s*:\s*(.+?)$", section_text, re.MULTILINE)
    if delivering_plant_match:
        address = delivering_plant_match.group(1).strip()
        location_mapping = {"BUDESTI": 1759, "CATEASCA": 1826, "CRAIOVA": 1593, }
        return location_mapping.get(address.upper(), 2093)  # Default to 2093 if not found

    # Check for "Our BAU Code : <CODE>" (allowing variations in spacing)
    bau_code_match = re.search(r"Our\s+BAU\s+Code\s*:\s*([A-Z0-9_]+)", section_text)
    if bau_code_match:
        bau_code = bau_code_match.group(1).strip()
        bau_code_to_location = {"RO03_E_CRA_1593": 1593,  # CRAIOVA
                                # Add more BAU codes as needed
                                }
        return bau_code_to_location.get(bau_code, 2093)  # Default to 2093 if not found

    return 2093  # Default code
