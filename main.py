import os
from datetime import datetime

from modules import InvoiceProcessor, ExcelGenerator

if __name__ == "__main__":
    input_paths = [os.path.join("pdfs", file) for file in os.listdir("pdfs") if file.lower().endswith(".pdf")]
    processor = InvoiceProcessor(input_paths)
    processor.process_invoices()

    current_date = datetime.now().strftime("%m-%Y")
    output_path = f"output/{current_date}-EXP.xlsx"

    excel_generator = ExcelGenerator(processor.df)
    excel_generator.generate_excel(output_path)
