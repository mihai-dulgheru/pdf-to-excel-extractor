# PDF to Excel Extractor

This project processes invoices in PDF format and generates an Excel file with structured data. It includes both an API
for processing and an Electron-based desktop application for user interaction.

## Features

- Extract data from PDF invoices.
- Process and format extracted data.
- Generate Excel files with formatted columns and calculated fields.
- API for backend processing.
- Desktop application using Electron for a user-friendly interface.

## Installation

### Backend Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/your-repo/pdf-to-excel-extractor.git
   ```
2. Navigate to the project directory:
   ```bash
   cd pdf-to-excel-extractor
   ```
3. Install Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```
4. Start the API server:
   ```bash
   uvicorn api.app:app --reload
   ```

### Electron Desktop Application Setup

1. Navigate to the `electron` directory:
   ```bash
   cd electron
   ```
2. Install Node.js dependencies:
   ```bash
   npm install
   ```
3. Start the Electron application:
   ```bash
   npm start
   ```

## Usage

### Using the API

1. Start the backend API server:
   ```bash
   uvicorn api.app:app --reload
   ```
2. Use the included Postman collection in the `postman/` directory to test the API endpoints.

### Using the Electron Application

1. Start the Electron application:
   ```bash
   npm start
   ```
2. Use the interface to upload PDF files and set the percentage for processing. The application will generate and allow
   you to save the Excel file.

## Testing

Run tests with `pytest` for the backend and unit tests for the Electron application.

### Backend Tests

```bash
pytest tests/
```

### Electron Tests

1. Navigate to the `electron` directory:
   ```bash
   cd electron
   ```
2. Run the tests:
   ```bash
   npm test
   ```

## Project Structure

```
├── api/                  # API logic and server
│   ├── app.py            # Main FastAPI application
│   └── routes/           # API routes
├── config/               # Configuration files for application settings and environment variables
├── electron/             # Electron-based desktop application
│   ├── main.js           # Main Electron process
│   ├── preload.js        # Preload script for secure context isolation
│   ├── renderer/         # Frontend HTML, CSS, and JS files
│   └── package.json      # Node.js dependencies and scripts
├── functions/            # Utility functions
├── modules/              # Core modules (InvoiceProcessor, ExcelGenerator)
├── output/               # Generated Excel files
├── pdfs/                 # Input PDFs
├── postman/              # Postman collection for API testing
├── tests/                # Unit tests for the backend
├── main.py               # Main script for command-line usage
├── README.md             # Project documentation
└── requirements.txt      # Python dependencies
```

## License

This project is licensed under the MIT License.