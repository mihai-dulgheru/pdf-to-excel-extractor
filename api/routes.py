import os
from datetime import datetime
from tempfile import TemporaryDirectory

from fastapi import APIRouter, UploadFile, Form
from fastapi.responses import FileResponse

from modules.excel_generator import ExcelGenerator
from modules.invoice_processor import InvoiceProcessor

router = APIRouter()


@router.post("/process-invoices/", response_class=FileResponse)
async def process_invoices(files: list[UploadFile], percentage: float = Form(...)):
    """
    Process a list of uploaded PDF files and generate an Excel file.
    """
    # Use a temporary directory to save uploaded files
    with TemporaryDirectory() as temp_dir:
        input_paths = []

        # Save uploaded files to temporary directory
        for file in files:
            temp_path = os.path.join(temp_dir, file.filename)
            with open(temp_path, "wb") as f:
                f.write(await file.read())
            input_paths.append(temp_path)

        # Process invoices and generate Excel
        processor = InvoiceProcessor(input_paths)
        processor.process_invoices()

        # Generate output path and create Excel
        current_date = datetime.now().strftime("%m-%Y")
        output_path = f"output/{current_date}-EXP.xlsx"
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        excel_generator = ExcelGenerator(processor.df, percentage)
        excel_generator.generate_excel(output_path)

        # Return the Excel file
        return FileResponse(output_path, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            filename=f"{current_date}-EXP.xlsx", )
