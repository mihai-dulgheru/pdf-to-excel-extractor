import re

from config import Constants


def get_delivery_location(section_text):
    """
    Determines the delivery location code based on the content of the provided text section.

    This function checks the text for the following patterns:
    1. 'Delivering plant : <ADDRESS>'
       - Extracts the last word from <ADDRESS>, converts it to uppercase,
         and attempts to match it in Constants.LOCATION_MAPPING. If no match
         is found or if parsing fails, returns Constants.DEFAULT_CODE.
    2. 'Our BAU Code : <CODE>'
       - Extracts <CODE>, converts it to uppercase, splits by underscore ('_'),
         and attempts to parse the last chunk as an integer. If parsing fails,
         returns Constants.DEFAULT_CODE.
    3. If neither pattern matches or if the text is empty/None, returns
       Constants.DEFAULT_CODE.

    :param section_text: str
        The extracted text from the document where expressions are searched.
    :return: int
        A location code (e.g., 1759, 1593, etc.) or Constants.DEFAULT_CODE if
        no valid match is found.
    """
    # If the text is empty or None, return the default code immediately
    if not section_text:
        return Constants.DEFAULT_CODE

    # Search for the pattern 'Delivering plant : <ADDRESS>'
    delivering_plant_regex = re.compile(r"Delivering\s+plant\s*:\s*(.+?)$", re.MULTILINE)
    delivering_plant_match = delivering_plant_regex.search(section_text)
    if delivering_plant_match:
        try:
            # Extract the last word in the matched address and convert to uppercase
            address = delivering_plant_match.group(1).strip().upper().split()[-1]
            return Constants.LOCATION_MAPPING.get(address, Constants.DEFAULT_CODE)
        except IndexError:
            # In case splitting fails or there is no last word
            return Constants.DEFAULT_CODE

    # Search for the pattern 'Our BAU Code : <CODE>'
    bau_code_regex = re.compile(r"Our\s+BAU\s+Code\s*:\s*([A-Z0-9_]+)", re.MULTILINE)
    bau_code_match = bau_code_regex.search(section_text)
    if bau_code_match:
        try:
            # Split the code by underscore, take the last part, and convert to integer
            bau_code_parts = bau_code_match.group(1).strip().upper().split('_')
            return int(bau_code_parts[-1])
        except (IndexError, ValueError):
            # If the code format isn't valid or can't be converted to int
            return Constants.DEFAULT_CODE

    # If no patterns matched, return the default code
    return Constants.DEFAULT_CODE
