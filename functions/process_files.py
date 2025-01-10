import json
import os
import sys
from datetime import datetime
from tempfile import TemporaryDirectory

from modules.excel_generator import ExcelGenerator
from modules.invoice_processor import InvoiceProcessor


def process_invoices(file_paths, percentage):
    with TemporaryDirectory() as temp_dir:
        input_paths = []
        for file_path in file_paths:
            temp_path = os.path.join(temp_dir, os.path.basename(file_path))
            os.rename(file_path, temp_path)
            input_paths.append(temp_path)

        processor = InvoiceProcessor(input_paths)
        processor.process_invoices()

        current_date = datetime.now().strftime("%m-%Y")
        output_path = os.path.join(temp_dir, f"{current_date}-EXP.xlsx")
        excel_generator = ExcelGenerator(processor.df, percentage)
        excel_generator.generate_excel(output_path)

        return output_path


if __name__ == "__main__":
    output_file = process_invoices(json.loads(sys.argv[1]), float(sys.argv[2]))
