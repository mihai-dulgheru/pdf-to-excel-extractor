def format_nc8_code(value):
    """
    Format the NC8 code into the pattern "XX XX XXXX" if it contains exactly 8 digits.
    If the input is invalid or does not meet the criteria, return the original value.
    """
    if not value:
        return value

    original_value = value
    numeric_value = "".join(char for char in str(value) if char.isdigit())

    if len(numeric_value) == 8:
        return f"{numeric_value[:2]} {numeric_value[2:4]} {numeric_value[4:]}"

    return original_value
