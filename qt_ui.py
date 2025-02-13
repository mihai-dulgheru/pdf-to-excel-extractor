import json
import os
import shutil
from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory

from PyQt6.QtCore import QThread, pyqtSignal
from PyQt6.QtGui import QIcon, QFont
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QProgressBar, QSpinBox,
                             QMessageBox, QScrollArea, QFrame, QAbstractSpinBox)

from config import Constants
from modules.excel_generator import ExcelGenerator
from modules.invoice_processor import InvoiceProcessor


def load_last_directory():
    """
    Load the last used directory from the config file.
    If missing or invalid, return the user's home directory.
    """
    if os.path.exists(Constants.CONFIG_FILE):
        try:
            with open(Constants.CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f).get("last_directory", os.path.expanduser("~"))
        except (json.JSONDecodeError, IOError):
            pass
    return os.path.expanduser("~")


def save_last_directory(directory: str):
    """
    Save the given directory path as the last used directory in the config file.
    Print an error message in case of failure.
    """
    try:
        with open(Constants.CONFIG_FILE, "w", encoding="utf-8") as f:
            f.write(json.dumps({"last_directory": directory}, indent=4))
    except IOError:
        print("Failed to save last directory")


# noinspection PyUnresolvedReferences
class ProcessingThread(QThread):
    """
    Background thread that processes PDF invoices, generates an Excel file,
    and emits signals for progress and completion.
    """
    finished = pyqtSignal(str)
    progress_signal = pyqtSignal(int)

    def __init__(self, pdf_files, percentage):
        """
        :param pdf_files: List of PDF file paths to process.
        :param percentage: Float percentage for calculations.
        """
        super().__init__()
        self.pdf_files = pdf_files
        self.percentage = percentage

    def copy_files_to_temp(self, temp_dir: str):
        """
        Copy PDF files to a temporary directory.
        :return: List of new file paths.
        """
        return [shutil.copy(str(Path(pdf)), os.path.join(temp_dir, Path(pdf).name)) for pdf in self.pdf_files]

    @staticmethod
    def generate_excel_filename():
        """
        Generate an Excel filename based on current month and year.
        """
        return f"{datetime.now().strftime('%m-%Y')}-EXP.xlsx"

    def process_invoices(self):
        """
        Process PDF invoices in a temp directory; create the final Excel file.
        :return: Final path of the generated Excel.
        """
        with TemporaryDirectory() as temp_dir:
            temp_paths = self.copy_files_to_temp(temp_dir)
            processor = InvoiceProcessor(temp_paths, progress_callback=self._progress_callback)
            processor.process_invoices()

            excel_filename = self.generate_excel_filename()
            temp_excel_path = os.path.join(temp_dir, excel_filename)
            generator = ExcelGenerator(processor.df, self.percentage)
            generator.generate_excel(temp_excel_path)

            final_path = os.path.join(os.getcwd(), excel_filename)
            shutil.move(temp_excel_path, final_path)
        return final_path

    def _progress_callback(self, processed, total):
        """
        Calculate integer progress and emit the signal.
        """
        progress_val = int((processed / total) * 100)
        self.progress_signal.emit(progress_val)

    def run(self):
        """
        Run the invoice processing; emit 'finished' with output path.
        """
        try:
            output_path = self.process_invoices()
            self.finished.emit(output_path)
        except Exception as e:
            print(f"[LOG] Error in ProcessingThread: {e}")
            self.finished.emit("")


# noinspection PyUnresolvedReferences
class PDFToExcelApp(QWidget):
    """
    Main window for selecting and processing PDF invoices into Excel.
    """

    def __init__(self):
        super().__init__()
        self.setWindowIcon(QIcon("assets/icon.png"))
        self.resize(800, 600)

        font = QFont("Segoe UI", 11)
        self.setFont(font)
        self.setWindowTitle("Procesor PDF în Excel")

        self.main_layout = QVBoxLayout()
        self.setLayout(self.main_layout)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)

        self.content_frame = QFrame()
        self.content_frame.setObjectName("ContentFrame")
        self.content_layout = QVBoxLayout(self.content_frame)
        self.content_layout.setSpacing(16)

        self.label = QLabel("Selectați fișiere PDF")
        label_font = QFont("Segoe UI", 13, QFont.Weight.Bold)
        self.label.setFont(label_font)
        self.content_layout.addWidget(self.label)

        self.select_button = QPushButton("Alegeți fișiere")
        self.select_button.clicked.connect(self.select_pdf_files)
        self.content_layout.addWidget(self.select_button)

        self.percentage_label = QLabel("Introduceți procentul:")
        self.content_layout.addWidget(self.percentage_label)

        self.percentage_input = QSpinBox()
        self.percentage_input.setRange(0, 99)
        self.percentage_input.setValue(60)
        self.percentage_input.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.content_layout.addWidget(self.percentage_input)

        self.process_button = QPushButton("Încărcați și procesați")
        self.process_button.setEnabled(False)
        self.process_button.clicked.connect(self.start_file_processing)
        self.content_layout.addWidget(self.process_button)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.content_layout.addWidget(self.progress_bar)

        self.status_label = QLabel("")
        status_font = QFont("Segoe UI", 10, QFont.Weight.Normal)
        self.status_label.setFont(status_font)
        self.content_layout.addWidget(self.status_label)

        scroll_area.setWidget(self.content_frame)
        self.main_layout.addWidget(scroll_area)

        self.pdf_files = []
        self.processing_thread = None
        self.last_directory = load_last_directory()

        self.setStyleSheet("""
            QWidget {
                background-color: #F9FAFB;
                margin: 0;
                padding: 0;
                font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif;
                color: #111827;
            }
            #ContentFrame {
                margin: 24px auto;
                padding: 24px;
                border: none;
                border-radius: 8px;
                background-color: #F3F4F6;
            }
            QLabel {
                font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif; 
                color: #1F2937;
                background-color: transparent;
            }
            QSpinBox {
                padding: 8px;
                height: 40px;
                font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif; 
                font-size: 13px;
                color: #111827;
                background-color: #F9FAFB;
                border: 1px solid #D1D5DB;
                border-radius: 4px;
            }
            QProgressBar {
                margin-top: 8px;
                margin-bottom: 8px;
                height: 24px;
                font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif; 
                font-size: 12px;
                color: #111827;
                text-align: center;
                background-color: #F9FAFB;
                border: 1px solid #D1D5DB;
                border-radius: 6px;
            }
            QProgressBar::chunk {
                color: #ffffff;
                background-color: #3B82F6;
                border-radius: 6px;
            }
            QPushButton {
                margin-top: 8px;
                padding: 8px 16px;
                font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif;
                font-size: 14px;
                color: #fff;
                background-color: #2563EB;
                border: 1px solid #1D4ED8;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #3B82F6;
            }
        """)

    def show_message(self, title, message, icon=QMessageBox.Icon.Information):
        """
        Show a message box with a given title, message, and optional icon.
        """
        msg = QMessageBox(self)
        msg.setIcon(icon)
        msg.setWindowTitle(title)
        msg.setText(message)
        msg.exec()

    def select_pdf_files(self):
        """
        Prompt a file dialog to select PDF files; enable processing if any selected.
        """
        selected_files, _ = QFileDialog.getOpenFileNames(self, "Alegeți fișiere PDF", self.last_directory,
                                                         "PDF Files (*.pdf)")
        if selected_files:
            self.pdf_files = selected_files
            self.label.setText(f"{len(selected_files)} fișier(e) selectat(e)")
            self.process_button.setEnabled(True)
            self.last_directory = os.path.dirname(selected_files[0])
            save_last_directory(self.last_directory)

    def start_file_processing(self):
        """
        Initialize the background thread for PDF processing and update the UI.
        """
        if not self.pdf_files:
            self.show_message("Eroare", "Vă rugăm să selectați fișiere PDF!", QMessageBox.Icon.Critical)
            return

        self.process_button.setEnabled(False)
        self.status_label.setText("Procesare în curs...")
        self.progress_bar.setValue(0)

        percentage = self.percentage_input.value() / 100.0
        self.processing_thread = ProcessingThread(self.pdf_files, percentage)
        self.processing_thread.progress_signal.connect(self.handle_progress_update)
        self.processing_thread.finished.connect(self.handle_processing_finished)
        self.processing_thread.start()

    def handle_progress_update(self, value):
        """
        Update the progress bar with current progress value.
        """
        self.progress_bar.setValue(value)

    def handle_processing_finished(self, output_path):
        """
        Handle finishing of processing. Allow user to save or show an error.
        """
        self.progress_bar.setValue(100)
        self.process_button.setEnabled(True)

        if output_path and os.path.exists(output_path):
            default_save_path = os.path.join(self.last_directory, os.path.basename(output_path))
            save_path, _ = QFileDialog.getSaveFileName(self, "Salvați fișierul Excel", default_save_path,
                                                       "Excel Files (*.xlsx)")

            if save_path:
                shutil.move(output_path, save_path)
                self.status_label.setText(f"Fișier salvat la: {save_path}")
                self.last_directory = os.path.dirname(save_path)
                save_last_directory(self.last_directory)
            else:
                os.remove(output_path)
                self.status_label.setText("Salvarea a fost anulată.")
        else:
            self.show_message("Eroare", "A apărut o eroare în timpul procesării!", QMessageBox.Icon.Critical)
