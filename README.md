# PDF to Excel Extractor

This project processes invoices in PDF format and generates an Excel file with structured data.

## Features

- Extract data from PDF invoices.
- Process and format extracted data.
- Generate Excel files with formatted columns and calculated fields.

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/your-repo/pdf-to-excel-extractor.git
   ```
2. Navigate to the project directory:
   ```bash
   cd pdf-to-excel-extractor
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

Run the main script to process PDFs and generate an Excel file:

```bash
python main.py
```

## Testing

Run tests with `pytest`:

```bash
pytest tests/
```

## Project Structure

```
├── api/                  # API logic and server
├── config/               # Configuration files for application settings and environment variables
├── functions/            # Utility functions
├── modules/              # Core modules (InvoiceProcessor, ExcelGenerator)
├── output/               # Generated Excel files
├── pdfs/                 # Input PDFs
├── postman/              # Postman collection for API testing
├── tests/                # Unit tests
├── main.py               # Main script
├── README.md             # Project documentation
└── requirements.txt      # Project dependencies
```

## License

This project is licensed under the MIT License.
