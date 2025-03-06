import sys

from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import QApplication

from qt_ui import PDFToExcelApp


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Extractor PDF în Excel")
    app.setApplicationVersion("1.0.0")
    app.setOrganizationName("PDF Extractor")

    # Set the main window icon
    app_icon = QIcon("assets/icon.png")
    app.setWindowIcon(app_icon)

    window = PDFToExcelApp()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
