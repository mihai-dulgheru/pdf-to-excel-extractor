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
    - Drag-and-drop file support
    - Error handling and user feedback

## Installation

### Development Setup

1. **Clone the repository**:

   ```bash
   git clone https://github.com/your-repo/pdf-to-excel-extractor.git
   cd pdf-to-excel-extractor
   ```

2. **Set up Python environment** (Python 3.8+ required):

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
├── assets/               # Application assets
├── config/              # Configuration files
├── functions/           # Utility functions
├── modules/             # Core modules
├── _internal/           # Dependencies and libraries
├── input/              # Place input PDF files here
└── output/             # Generated Excel files
```

Key features of the build:

- Multi-file distribution for better performance and easier updates
- Optimized code with level 2 optimization
- No console window (runs in windowed mode only)
- UPX compression enabled for smaller file size
- Stripped binaries for reduced size

## Usage

### Running from Source

1. **Start the application**:

   ```bash
   python main.py
   ```

2. **Process invoices**:
    - Click "Select Files" or drag-and-drop PDF files
    - Set the processing percentage (default: 0.6)
    - Click "Process" to start extraction
    - Monitor progress in the status bar
    - Access the generated Excel file in the output directory

### Using the Portable Executable

1. **Launch the application**: Run `PDFToExcelApp.exe`
2. **Place PDF files** in the `input` directory or use the file selector
3. **Configure processing** options if needed
4. **Generate Excel file** which will appear in the `output` directory

## Configuration

The application uses several configuration files in the `config` directory:

- `constants.py`: Defines column formats, headers, and other constants
- `proportions.py`: Configures PDF section extraction coordinates

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
- Maintains organized input/output directory structure
- Preserves full functionality of the source version
- No additional installation or configuration needed

## License

This project is licensed under the **MIT License**. See the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues for improvements and bug fixes.
