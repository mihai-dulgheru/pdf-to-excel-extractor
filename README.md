# PDF to Excel Extractor

A powerful desktop application that processes PDF invoices and generates structured Excel files. The application
extracts detailed invoice data, including multipage support with proportional weight distribution across NC8 codes.

## Overview

This application is designed to streamline the process of extracting data from PDF invoices into organized Excel
spreadsheets. It features a user-friendly PyQt6-based interface and supports complex data extraction scenarios,
including:

- Multipage invoice processing
- Proportional weight distribution across NC8 codes
- Automatic currency conversion (EUR/RON)
- Intelligent data extraction and formatting

## Features

- **PDF Processing**

    - Extract data from single and multipage invoices
    - Support for various invoice formats and layouts
    - Intelligent text extraction using pdfplumber
    - Parallel processing for multiple files

- **Data Extraction**

    - Company information and invoice numbers
    - NC8 codes with corresponding values
    - Origin and destination countries
    - Net weights with proportional distribution
    - Currency values (EUR/RON) with exchange rates
    - Shipment dates and delivery conditions

- **Excel Generation**

    - Formatted output with proper column types
    - Automatic calculations and formulas
    - Support for merging with existing Excel files
    - Customizable percentage calculations

- **User Interface**
    - Modern PyQt6-based GUI
    - Progress tracking for file processing
    - Drag-and-drop of PDF files and folders onto the window
    - Error handling and user feedback

## Installation

### Development Setup

1. **Clone the repository**:

   ```bash
   git clone https://github.com/your-repo/pdf-to-excel-extractor.git
   cd pdf-to-excel-extractor
   ```

2. **Set up Python environment** (Python 3.10+ required):

   ```bash
   python -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```

3. **Install dependencies**:

   ```bash
   pip install -r requirements.txt
   ```

### Portable Executable Generation

To create a standalone Windows executable:

1. **Install PyInstaller**:

   ```bash
   pip install pyinstaller
   ```

2. **Generate the executable**:

   ```bash
   python setup.py
   ```

The application will be created in the `dist/PDFToExcelApp` directory as a multi-file distribution with the following
structure:

```plaintext
dist/PDFToExcelApp/
├── PDFToExcelApp.exe     # Main executable
├── _internal/            # Dependencies, libraries and the packaged Python modules
├── config.json          # Last used folder (written by the app)
└── logs/                # pdf_to_excel.log, rotated at 1 MB (written by the app)
```

Key features of the build:

- Multi-file distribution for better performance and easier updates
- No console window (runs in windowed mode only)
- No bytecode optimization, UPX compression or stripped binaries (see `PDFToExcelApp.spec`)

## Usage

### Running from Source

1. **Start the application**:

   ```bash
   python main.py
   ```

2. **Process invoices**:
    - Click "Selectează fișiere PDF" or "Selectează dosar cu fișiere PDF", or drag PDF files and folders onto
      the window; each selection adds to the list
    - Set the processing percentage (default: 60%)
    - Click "Procesează fișierele PDF și generează Excel" to start extraction
    - Monitor progress in the status bar
    - Choose where to save the generated Excel file in the dialog shown at the end

### Using the Portable Executable

1. **Launch the application**: Run `PDFToExcelApp.exe`
2. **Select PDF files** or a folder with the file selector
3. **Configure processing** options if needed
4. **Generate the Excel file** and choose where to save it; `config.json` and `logs/` are written next to the
   executable

## Configuration

The application's configuration lives in the `config` directory:

- `constants.py`: Column formats, headers, PDF section proportions, application paths and other constants

The application icon is drawn in `assets/icon.svg`; `python tools/make_icon.py` re-renders `assets/icon.png` and
`assets/icon.ico` from it (run on Windows, where the Segoe UI font used for the labels is available).

## Dependencies

Major dependencies include:

- `pandas`: Data processing and Excel file generation
- `pdfplumber`: PDF text extraction
- `PyQt6`: Graphical user interface
- `openpyxl`: Excel file handling
- `requests` & `beautifulsoup4`: Exchange rate fetching
- Additional utilities: `locale`, `concurrent.futures`

## Project Structure

```plaintext
pdf-to-excel-extractor/
├── assets/               # Application assets (icons, etc.)
├── config/              # Configuration files
├── functions/           # Utility functions
│   ├── calculate_coordinates.py
│   ├── get_bnr_exchange_rate.py
│   └── ...
├── modules/             # Core application modules
│   ├── invoice_processor.py
│   ├── excel_generator.py
│   └── ...
├── main.py             # Application entry point
├── qt_ui.py            # User interface code
├── setup.py            # Build configuration
└── requirements.txt    # Python dependencies
```

## Portable Executable Features

The standalone executable:

- Runs on Windows 10/11 without Python installation
- Includes all required dependencies
- Preserves full functionality of the source version
- No additional installation or configuration needed

## License

This project is licensed under the **MIT License**. See the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues for improvements and bug fixes.
