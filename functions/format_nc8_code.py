def format_nc8_code(value):
    """
    Formats the NC8 code into the pattern "XX XX XXXX" if it contains exactly 8 digits.
    If the input is invalid or does not meet the criteria, returns the original value.
    If the code contains multiple codes separated by ";", formatting is applied to each one.
    """
    if not value:
        return value

    codes = [code.strip() for code in str(value).split(";")]
    formatted_codes = []

    for code in codes:
        digits = "".join(character for character in code if character.isdigit())
        if len(digits) == 8:
            formatted_code = f"{digits[:2]} {digits[2:4]} {digits[4:]}"
            formatted_codes.append(formatted_code)
        else:
            formatted_codes.append(code)

    return "; ".join(formatted_codes)
