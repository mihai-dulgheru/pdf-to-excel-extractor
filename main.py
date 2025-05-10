import logging
import sys

from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import QApplication

from config import Constants
from functions import setup_logger
from qt_ui import PDFToExcelApp


def main():
    setup_logger(name="pdf_to_excel", log_level=logging.INFO, log_to_file=True)
    logger = logging.getLogger("pdf_to_excel")
    logger.info("Application starting")

    app = QApplication(sys.argv)
    app.setApplicationName("Extractor PDF în Excel")
    app.setApplicationVersion(Constants.APPLICATION_VERSION)
    app.setOrganizationName("PDF Extractor")

    app_icon = QIcon("assets/icon.png")
    app.setWindowIcon(app_icon)

    window = PDFToExcelApp()
    window.show()

    logger.info("Application UI initialized")

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
