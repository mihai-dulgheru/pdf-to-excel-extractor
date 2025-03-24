import sys

from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import QApplication

from config import Constants
from qt_ui import PDFToExcelApp


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Extractor PDF în Excel")
    app.setApplicationVersion(Constants.APPLICATION_VERSION)
    app.setOrganizationName("PDF Extractor")

    app_icon = QIcon("assets/icon.png")
    app.setWindowIcon(app_icon)

    window = PDFToExcelApp()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
