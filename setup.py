from PyInstaller.__main__ import run

options = ["qt_ui.py",  # Main application file
           "--name=PDFToExcelApp",  # Name of the executable
           "--onefile",  # Generate a single executable file
           "--windowed",  # Run as a GUI application (no terminal window)
           "--hidden-import=pandas",  # Include pandas library
           "--hidden-import=openpyxl",  # Include openpyxl for Excel generation
           "--hidden-import=pdfplumber",  # Include pdfplumber for PDF processing
           "--hidden-import=PyQt6",  # Ensure PyQt6 is included
           "--hidden-import=requests",  # Include requests for HTTP requests
           "--add-data=config;config",  # Include config directory
           "--add-data=functions;functions",  # Include functions directory
           "--add-data=modules;modules",  # Include modules directory
           "--icon=assets/icon.ico",  # Set application icon (optional)
           ]

run(options)
