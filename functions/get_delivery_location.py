import re


def get_delivery_location(section_text):
    """
    Determines the delivery location based on the content of the section.
    """
    # Check for "Delivering plant : <ADDRESS>"
    delivering_plant_match = re.search(r"Delivering plant : (.+?)$", section_text, re.MULTILINE)
    if delivering_plant_match:
        address = delivering_plant_match.group(1).strip()
        return address.split()[-1]  # Assume the last word is the location

    # Check for "Our BAU Code : <CODE>"
    bau_code_match = re.search(r"Our BAU Code : ([A-Z0-9_]+)", section_text)
    if bau_code_match:
        bau_code_to_location = {"RO03_E_CRA_1593": "CRAIOVA",  # Add more BAU codes as needed
                                }
        return bau_code_to_location.get(bau_code_match.group(1), "Unknown")

    return "Unknown"
