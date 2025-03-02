from datetime import datetime

import pandas as pd


def convert_to_date(value):
    """
    Converts a value to a datetime object.
    Handles string dates in '%d.%m.%Y' format, existing datetime objects, and Excel numeric dates.
    """
    if pd.isna(value):
        return None

    # If it's already a datetime, return it
    if isinstance(value, datetime):
        return value

    # If it's a string, try to parse it
    if isinstance(value, str):
        try:
            return datetime.strptime(value, "%d.%m.%Y")
        except ValueError:
            try:
                # Try alternative formats
                return pd.to_datetime(value)
            except ValueError:
                return None

    # If it's a float or int (Excel numeric date), convert it
    try:
        return pd.to_datetime(value, unit='D', origin='1899-12-30')
    except ValueError:
        return None
