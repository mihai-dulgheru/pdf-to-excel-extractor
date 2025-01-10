import os

import pandas as pd
import pytest

from modules.excel_generator import ExcelGenerator


@pytest.fixture
def sample_data():
    return pd.DataFrame(
        {"company": ["Company A", "Company B"], "invoice_number": ["123", "456"], "nc8_code": ["87082990", "94019080"],
         "origin": ["RO", "DE"], "destination": ["DE", "FR"], "invoice_value_eur": [1000.00, 2000.00],
         "net_weight": [100, 200], "shipment_date": ["01.01.2025", "02.01.2025"], "exchange_rate": [4.95, 5.00],
         "value_ron": [4950.00, 10000.00], "vat_number": ["VAT123", "VAT456"],
         "delivery_location": ["Location A", "Location B"], "delivery_condition": ["FCA", "DAP"]})


@pytest.fixture
def excel_generator(sample_data):
    return ExcelGenerator(sample_data)


def test_generate_excel(excel_generator, tmpdir):
    output_path = os.path.join(tmpdir, "output.xlsx")
    excel_generator.generate_excel(output_path)
    assert os.path.exists(output_path)
