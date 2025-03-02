from datetime import timedelta


def get_previous_workday(input_date):
    """
    Returns the last working day before the given date.

    :param input_date: datetime object representing the invoice date.
    :return: datetime object representing the previous working day.
    """
    previous_date = input_date - timedelta(days=1)

    while previous_date.weekday() in (5, 6):
        previous_date -= timedelta(days=1)

    return previous_date
