# PDF to Excel Extractor

This project processes invoices in PDF format and generates an Excel file with structured data. It includes a **PyQt6
desktop application** for user interaction.

## Features

- Extract data from PDF invoices
- Process and format extracted data into a pandas DataFrame
- Generate Excel files with formatted columns and calculated fields
- Desktop application using PyQt6 for a user-friendly interface

## Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/your-repo/pdf-to-excel-extractor.git
   ```
2. **Navigate to the project directory**:
   ```bash
   cd pdf-to-excel-extractor
   ```
3. **(Optional) Create and activate a virtual environment**:
   ```bash
   python -m venv .venv
   source .venv/bin/activate  # On Windows, use .venv\Scripts\activate
   ```
4. **Install Python dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. **Run the main application**:
   ```bash
   python main.py
   ```
2. **Select PDF files** and set the percentage for processing.
3. **Generate and save the Excel file** once processing is complete.

## Project Structure

Below is a brief overview of the relevant directories and files:

```
pdf-to-excel-extractor/
├── config/               # Configuration (e.g., proportions, constants)
│   └── ...
├── functions/            # Utility functions (e.g., coordinates, exchange rate)
│   └── ...
├── modules/              # Core modules (e.g., InvoiceProcessor, ExcelGenerator)
│   └── ...
├── qt_ui.py              # PyQt6 user interface code
├── main.py               # Main Python script for launching the app
├── requirements.txt      # Python dependencies
├── setup.py              # Setup or packaging script
└── README.md             # Project documentation
```

## License

This project is licensed under the **MIT License**. Feel free to modify and adapt it to your needs.
