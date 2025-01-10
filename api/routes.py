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
    print(f"Received files: {[file.filename for file in files]}")
    print(f"Received percentage: {percentage}")

    with TemporaryDirectory() as temp_dir:
        input_paths = []

        for file in files:
            file_path = os.path.join(temp_dir, file.filename)
            with open(file_path, "wb") as f:
                f.write(await file.read())
            input_paths.append(file_path)

        processor = InvoiceProcessor(input_paths)
        processor.process_invoices()

        current_date = datetime.now().strftime("%m-%Y")
        file_path = os.path.join("output", f"{current_date}-EXP.xlsx")
        os.makedirs(os.path.dirname(file_path), exist_ok=True)

        excel_generator = ExcelGenerator(processor.df, percentage)
        excel_generator.generate_excel(file_path)

        return FileResponse(file_path, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            filename=f"{current_date}-EXP.xlsx", )
