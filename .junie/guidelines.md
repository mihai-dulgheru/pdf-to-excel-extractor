# PDF to Excel Extractor - Developer Guidelines

This document provides specific information for developers working on the PDF to Excel Extractor project. It includes
build/configuration instructions, testing information, and additional development details.

## Build/Configuration Instructions

### Development Environment Setup

1. **Python Environment**:
    - Python 3.8+ is required
    - Create a virtual environment:
      ```powershell
      python -m venv .venv
      .\.venv\Scripts\activate
      pip install -r requirements.txt
      ```

2. **Running the Application in Development Mode**:
   ```powershell
   python main.py
   ```

### Building the Portable Executable

The application can be packaged as a standalone Windows executable using PyInstaller:

1. **Generate the executable**:
   ```powershell
   python setup.py
   ```

2. **Build Configuration**:
    - The build configuration is defined in `setup.py` and `PDFToExcelApp.spec`
    - The executable is created in the `dist/PDFToExcelApp` directory
    - The build includes:
        - Multi-file distribution for better performance
        - Level 2 optimization
        - UPX compression for smaller file size
        - No console window (windowed mode only)

3. **Distribution Structure**:
   ```
   dist/PDFToExcelApp/
   ├── PDFToExcelApp.exe     # Main executable
   ├── assets/               # Application assets
   ├── config/               # Configuration files
   ├── functions/            # Utility functions
   ├── modules/              # Core modules
   ├── _internal/            # Dependencies and libraries
   ├── input/                # Place input PDF files here
   └── output/               # Generated Excel files
   ```

## Testing Information

### Running Tests

1. **Running All Tests**:
   ```powershell
   python -m unittest discover tests
   ```

2. **Running a Specific Test File**:
   ```powershell
   python -m unittest tests/test_convert_to_date.py
   ```

3. **Running a Specific Test Case**:
   ```powershell
   python -m unittest tests.test_convert_to_date.TestConvertToDate.test_string_format
   ```

### Adding New Tests

1. **Test File Structure**:
    - Create test files in the `tests` directory
    - Name test files with the prefix `test_` followed by the name of the module being tested
    - Use the `unittest` framework

2. **Test Class Structure**:
    - Create a test class that inherits from `unittest.TestCase`
    - Name the test class with the prefix `Test` followed by the name of the function or class being tested
    - Include descriptive docstrings for each test method

3. **Test Method Structure**:
    - Name test methods with the prefix `test_` followed by a description of what is being tested
    - Include assertions to verify expected behavior
    - Use `self.subTest` for testing multiple cases within a single test method

4. **Example Test**:
   ```python
   import unittest
   from functions import round_to_n_decimals

   class TestRoundToNDecimals(unittest.TestCase):
       def test_default_rounding(self):
           """Test rounding with default precision (2 decimals)."""
           self.assertEqual(round_to_n_decimals(3.14159), 3.14)
           self.assertEqual(round_to_n_decimals(2.999), 3.0)
   
   if __name__ == "__main__":
       unittest.main()
   ```

5. **Running the Example Test**:
   ```powershell
   python -m unittest tests/test_round_to_n_decimals.py
   ```

## Additional Development Information

### Project Structure

The project follows a modular structure:

- **modules/**: Core application modules
    - `invoice_processor.py`: Handles PDF data extraction
    - `excel_generator.py`: Generates Excel files from extracted data

- **functions/**: Utility functions
    - Each function is in its own file for better maintainability
    - Functions are imported directly from the package (e.g., `from functions import convert_to_date`)

- **config/**: Configuration files
    - `constants.py`: Defines column formats, headers, and other constants

- **assets/**: Application assets (icons, etc.)

### Code Style Guidelines

1. **Naming Conventions**:
    - Use snake_case for functions, methods, and variables
    - Use PascalCase for classes
    - Use UPPER_CASE for constants

2. **Documentation**:
    - Include docstrings for all functions, classes, and methods
    - Use type hints where appropriate

3. **Error Handling**:
    - Use try-except blocks for operations that might fail
    - Log errors with appropriate messages
    - Return None or default values for functions that can't complete their operation

### Performance Considerations

1. **PDF Processing**:
    - PDF processing is done in parallel using `concurrent.futures`
    - Progress is tracked and reported via callbacks

2. **Caching**:
    - Use `@lru_cache` for expensive operations like fetching exchange rates

3. **Excel Generation**:
    - Excel files are generated using pandas and openpyxl
    - Formulas are updated when new rows are inserted
    - Totals are recalculated after modifications

### Debugging Tips

1. **PDF Extraction Issues**:
    - Check the coordinates in `calculate_coordinates.py`
    - Use `pdfplumber.open(pdf_path).pages[0].to_image().save("debug.png")` to visualize the PDF

2. **Excel Generation Issues**:
    - Set breakpoints in `excel_generator.py` to inspect the data
    - Check the column formats in `constants.py`

3. **UI Issues**:
    - Use `print` statements or logging in the PyQt6 event handlers
    - Check the signal-slot connections in `qt_ui.py`