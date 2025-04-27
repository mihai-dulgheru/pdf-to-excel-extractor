# PDF to Excel Extractor - Developer Guidelines

This document provides essential information for developers working on the PDF to Excel Extractor project.

## Build/Configuration Instructions

### Environment Setup

1. **Python Environment**:
    - Python 3.9+ is recommended
    - Create a virtual environment:
      ```
      python -m venv .venv
      .venv\Scripts\activate
      ```
    - Install dependencies:
      ```
      pip install -r requirements.txt
      ```

2. **Development Mode**:
    - Run the application directly with:
      ```
      python main.py
      ```

3. **Building Executable**:
    - The project uses PyInstaller to create standalone executables
    - Build with:
      ```
      python setup.py
      ```
    - The executable will be created in the `dist/PDFToExcelApp` directory
    - The build process automatically includes all necessary assets and dependencies

4. **Configuration**:
    - Application settings are stored in `config.json` and the `config` directory
    - Constants are defined in `config/constants.py`
    - Modify these files to adjust application behavior

## Testing Information

### Running Tests

1. **Running All Tests**:
   ```
   python -m unittest discover tests
   ```

2. **Running Specific Tests**:
   ```
   python -m unittest tests/test_file_name.py
   ```

3. **Running a Specific Test Case**:
   ```
   python -m unittest tests.test_file_name.TestClassName.test_method_name
   ```

### Adding New Tests

1. **Test Structure**:
    - Tests use Python's standard `unittest` framework
    - Place test files in the `tests` directory
    - Name test files with the pattern `test_*.py`
    - Name test classes with the pattern `Test*`
    - Name test methods with the pattern `test_*`

2. **Test Example**:
   ```python
   import unittest
   from datetime import datetime
   
   from functions import convert_to_date
   
   class TestConvertToDate(unittest.TestCase):
       def test_string_format(self):
           """Test conversion of strings in the primary format."""
           date_str = "15.03.2025"
           expected = datetime(2025, 3, 15)
           self.assertEqual(convert_to_date(date_str), expected)
   
   if __name__ == "__main__":
       unittest.main()
   ```

3. **Test Best Practices**:
    - Write tests for all new functionality
    - Use descriptive test method names
    - Include docstrings explaining what each test does
    - Use `self.subTest` for testing multiple cases within a single test method
    - Test edge cases and error conditions

## Additional Development Information

### Code Structure

1. **Project Organization**:
    - `assets/`: Application icons and images
    - `config/`: Configuration files and constants
    - `functions/`: Utility functions used throughout the application
    - `input/`: Directory for input PDF files
    - `modules/`: Core application modules
    - `output/`: Directory for generated Excel files
    - `tests/`: Test files

2. **Key Modules**:
    - `main.py`: Application entry point
    - `qt_ui.py`: PyQt6 user interface
    - `modules/invoice_processor.py`: Core PDF processing logic
    - `modules/excel_generator.py`: Excel file generation

### Code Style

1. **Formatting**:
    - Follow PEP 8 guidelines
    - Use four spaces for indentation
    - Maximum line length of 120 characters

2. **Documentation**:
    - Include docstrings for all functions, classes, and methods
    - Use descriptive variable and function names
    - Comment complex logic

3. **Error Handling**:
    - Use try/except blocks for error-prone operations
    - Log errors with descriptive messages
    - Provide user-friendly error messages in the UI

### Common Development Tasks

1. **Adding New PDF Processing Features**:
    - Extend the `InvoiceProcessor` class in `modules/invoice_processor.py`
    - Add new extraction methods following the existing pattern
    - Update the UI in `qt_ui.py` if needed

2. **Modifying Excel Output**:
    - Update the `ExcelGenerator` class in `modules/excel_generator.py`
    - Modify column formats in `functions/enforce_column_formats.py`

3. **Debugging Tips**:
    - Use the `sandbox.py` file for testing code snippets
    - Check the input/output directories for sample files
    - Run with specific test PDFs to isolate issues

### Known Issues and Improvements

1. **Date Parsing Warning**:
    - The `convert_to_date` function generates a warning when parsing dates in %d/%m/%Y format
    - Consider adding `dayfirst=True` to the `pd.to_datetime()` call

2. **Performance Considerations**:
    - Large PDF files may cause memory issues
    - Consider implementing batch processing for large files