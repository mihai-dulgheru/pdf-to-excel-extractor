def calculate_coordinates(page_width, page_height, proportion):
    """
    Calculate coordinates based on page width, height and proportion.
    """
    x0 = proportion[0] * page_width
    y0 = proportion[1] * page_height
    x1 = proportion[2] * page_width
    y1 = proportion[3] * page_height
    return x0, y0, x1, y1
