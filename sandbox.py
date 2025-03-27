import os
from datetime import datetime

from functions import get_all_pdf_files, show_progress_bar, merge_existing_with_new
from modules.excel_generator import ExcelGenerator
from modules.invoice_processor import InvoiceProcessor

if __name__ == "__main__":
    input_directory = "input"
    input_paths = get_all_pdf_files(input_directory)
    existing_excel = ""

    if not input_paths:
        print("[LOG] No PDF files found in input directory.")
    else:
        print(f"[LOG] Found {len(input_paths)} PDF files to process.")

        processor = InvoiceProcessor(input_paths, progress_callback=show_progress_bar)
        processor.process_invoices()

        if existing_excel and os.path.exists(existing_excel):
            processor.df = merge_existing_with_new(existing_excel, processor.df)

        current_date = datetime.now().strftime("%d-%m-%Y")
        output_path = f"output/{current_date}-EXP.xlsx"

        excel_generator = ExcelGenerator(processor.df)
        excel_generator.generate_excel(output_path)

        print(f"[LOG] Excel generated: {output_path}")
