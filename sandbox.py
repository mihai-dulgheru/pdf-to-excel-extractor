from datetime import datetime

from functions import get_all_pdf_files, show_progress_bar
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

        current_date = datetime.now().strftime("%d-%m-%Y")
        output_path = f"output/{current_date}-EXP-2.xlsx"

        excel_generator = ExcelGenerator(processor.df)
        excel_generator.generate_excel(output_path, existing_excel)

        print(f"[LOG] Excel generated: {output_path}")
