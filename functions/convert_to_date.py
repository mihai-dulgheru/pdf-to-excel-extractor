from datetime import datetime


def convert_to_date(value):
    """
    Converts a string value to a datetime object using the '%d.%m.%Y' format.
    """
    try:
        return datetime.strptime(value, "%d.%m.%Y")
    except ValueError:
        return None
