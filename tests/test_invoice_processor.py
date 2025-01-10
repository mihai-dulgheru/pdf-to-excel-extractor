import os

import pytest

from modules.invoice_processor import InvoiceProcessor


@pytest.fixture
def sample_paths():
    return [os.path.join("../pdfs", file) for file in os.listdir("../pdfs") if file.lower().endswith(".pdf")]


@pytest.fixture
def processor(sample_paths):
    return InvoiceProcessor(sample_paths)


def test_process_invoices(processor):
    processor.process_invoices()
    assert len(processor.df) > 0  # Ensure data is processed
