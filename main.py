import sys

from PyQt6.QtWidgets import (QApplication)

from qt_ui import PDFToExcelApp

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFToExcelApp()
    window.show()
    sys.exit(app.exec())
