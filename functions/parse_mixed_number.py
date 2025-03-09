def parse_mixed_number(num_str):
    """
    Helper function that attempts to convert a string containing '.' and/or ','
    into a coherent decimal format. The basic logic is:
     - Remove spaces.
     - If both '.' and ',' exist, we treat the last one as the decimal separator
       and remove the other (thousands separator).
     - If only '.' or only ',' or none are present, interpret accordingly.
    """
    temp = num_str.replace(" ", "")
    dot_index = temp.rfind(".")
    comma_index = temp.rfind(",")

    if dot_index != -1 and comma_index != -1:
        # Both '.' and ',' are present; see which one appears last
        if dot_index > comma_index:
            # '.' is the decimal separator => remove ','
            temp = temp.replace(",", "")
        else:
            # ',' is the decimal separator => remove '.'
            temp = temp.replace(".", "")
            # replace ',' with '.'
            temp = temp.replace(",", ".")
    else:
        # Only '.' or only ',' or none
        temp = temp.replace(",", ".")  # unify on '.'

    try:
        return float(temp)
    except ValueError:
        return 0.0
