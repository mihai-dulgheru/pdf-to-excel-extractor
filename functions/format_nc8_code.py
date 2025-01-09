def format_nc8_code(value):
    """
    Format the NC8 code into the pattern "XX XX XXXX" if it contains exactly 8 digits.
    If the input is invalid or does not meet the criteria, return the original value.
    """
    if not value:
        return value

    value = "".join(char for char in str(value) if char.isdigit())
    if len(value) == 8:
        return f"{value[:2]} {value[2:4]} {value[4:]}"
    return value
