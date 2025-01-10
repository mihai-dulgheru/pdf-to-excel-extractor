from functions.convert_to_date import convert_to_date
from functions.format_nc8_code import format_nc8_code


def test_format_nc8_code():
    assert format_nc8_code("87082990") == "87 08 2990"
    assert format_nc8_code("12345678") == "12 34 5678"
    assert format_nc8_code("invalid") == ""


def test_convert_to_date():
    assert convert_to_date("01.01.2025").strftime("%d.%m.%Y") == "01.01.2025"
    assert convert_to_date("invalid") is None
