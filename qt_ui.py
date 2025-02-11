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
    Load the last used directory from the configuration file.
    If the config file is missing or invalid, return the user's home directory.
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
    Save the given directory path as the last used directory in the configuration file.
    If writing fails, print an error message in English.
    """
    try:
        with open(Constants.CONFIG_FILE, "w", encoding="utf-8") as f:
            json_data = json.dumps({"last_directory": directory}, indent=4)
            f.write(json_data)
    except IOError:
        print("Failed to save last directory")


# noinspection PyUnresolvedReferences
class ProcessingThread(QThread):
    """
    Background thread that processes PDF invoices, generates Excel files,
    and emits signals for progress and completion.
    """
    finished = pyqtSignal(str)
    progress_signal = pyqtSignal(int)

    def __init__(self, files, percentage):
        """
        Initialize the ProcessingThread with a list of PDF file paths and a percentage value.
        :param files: List of PDF file paths to process.
        :param percentage: The percentage (float) to be applied in calculations.
        """
        super().__init__()
        self.files = files
        self.percentage = percentage

    def copy_files_to_temp(self, temp_dir: str):
        """
        Copy input PDF files to a temporary directory and return their new paths.
        :param temp_dir: The path of the temporary directory.
        :return: A list of new file paths in the temp directory.
        """
        return [shutil.copy(str(Path(file)), os.path.join(temp_dir, Path(file).name)) for file in self.files]

    @staticmethod
    def generate_excel_filename():
        """
        Generate an Excel filename based on the current month and year.
        :return: A string representing the generated filename.
        """
        return f"{datetime.now().strftime('%m-%Y')}-EXP.xlsx"

    def process_invoices(self):
        """
        Process the PDF invoices in a temporary directory and generate the final Excel file.
        :return: The final path of the generated Excel file.
        """
        with TemporaryDirectory() as temp_dir:
            input_paths = self.copy_files_to_temp(temp_dir)

            processor = InvoiceProcessor(input_paths, progress_callback=self._progress_callback)
            processor.process_invoices()

            excel_filename = self.generate_excel_filename()
            output_temp_path = os.path.join(temp_dir, excel_filename)

            excel_generator = ExcelGenerator(processor.df, self.percentage)
            excel_generator.generate_excel(output_temp_path)

            output_final_path = os.path.join(os.getcwd(), excel_filename)
            shutil.move(output_temp_path, output_final_path)

        return output_final_path

    def _progress_callback(self, processed, total):
        """
        Callback function that calculates the current progress as an integer percentage
        and emits it through the progress_signal.
        :param processed: Number of processed files.
        :param total: Total number of files.
        """
        progress_value = int((processed / total) * 100)
        self.progress_signal.emit(progress_value)

    def run(self):
        """
        Entry point for the QThread.
        It calls process_invoices() and emits a 'finished' signal with the resulting path.
        If an exception occurs, it emits an empty string.
        """
        try:
            output_path = self.process_invoices()
            self.finished.emit(output_path)
        except Exception as e:
            print(f"Error in processing thread: {e}")
            self.finished.emit("")


# noinspection PyUnresolvedReferences
class PDFToExcelApp(QWidget):
    """
    Main application window that allows users to select and process PDF invoices into Excel.
    """

    def __init__(self):
        """
        Initialize the main window, UI elements, and default states.
        """
        super().__init__()

        self.setWindowIcon(QIcon("assets/icon.png"))
        self.resize(800, 600)
        font = QFont("Segoe UI", 13)
        self.setFont(font)

        self.setWindowTitle("Procesor PDF în Excel")
        self.layout = QVBoxLayout()

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)

        self.content_frame = QFrame()
        self.content_frame.setObjectName("ContentFrame")
        self.content_layout = QVBoxLayout(self.content_frame)

        self.label = QLabel("Selectați fișiere PDF")
        self.content_layout.addWidget(self.label)

        self.select_button = QPushButton("Alegeți fișiere")
        self.select_button.clicked.connect(self.select_files)
        self.content_layout.addWidget(self.select_button)

        self.percentage_label = QLabel("Introduceți procentul:")
        self.content_layout.addWidget(self.percentage_label)

        self.percentage_input = QSpinBox()
        self.percentage_input.setRange(0, 99)
        self.percentage_input.setValue(60)
        self.percentage_input.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.content_layout.addWidget(self.percentage_input)

        self.process_button = QPushButton("Încărcați și procesați")
        self.process_button.clicked.connect(self.process_files)
        self.process_button.setEnabled(False)
        self.content_layout.addWidget(self.process_button)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.content_layout.addWidget(self.progress_bar)

        self.status_label = QLabel("")
        self.content_layout.addWidget(self.status_label)

        scroll_area.setWidget(self.content_frame)
        self.layout.addWidget(scroll_area)

        self.setLayout(self.layout)
        self.files = []
        self.processing_thread = None
        self.last_directory = load_last_directory()

        self.setStyleSheet("""
            QWidget {
                margin: 0;
                padding: 0;
                background-color: #ffffff;
            }

            #ContentFrame {
                margin: 20px auto;
                padding: 20px;
                border: none;
                border-radius: 8px;
                background-color: #f0f0f0;
            }

            QLabel {
                font-family: 'Segoe UI';
                font-size: 14pt;
                color: #333333;
                background-color: #f0f0f0;
            }

            QSpinBox {
                padding: 8px;
                height: 40px;
                font-family: 'Segoe UI';
                font-size: 14pt;
                color: #333333;
            }

            QProgressBar {
                margin-top: 8px;
                margin-bottom: 8px;
                height: 40px;
                font-family: 'Segoe UI';
                font-size: 14pt;
                color: #333333;
                text-align: center;
                background-color: #ffffff;
                border-radius: 6px;
            }
            QProgressBar::chunk {
                color: #ffffff;
                background-color: #6c9acf;
                border-radius: 6px;
            }

            QPushButton {
                margin-top: 8px;
                padding: 8px 16px;
                font-family: 'Segoe UI';
                font-size: 14pt;
                color: #ffffff;
                background-color: #6c9acf;
                border: 1px solid #5c8bb8;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #5c8bb8;
            }
        """)

    def show_message(self, title, message, icon=QMessageBox.Icon.Information):
        """
        Display a message box with the given title, message, and optional icon.
        :param title: Title of the message box.
        :param message: Text content of the message box.
        :param icon: The icon to display, default is Information.
        """
        msg = QMessageBox(self)
        msg.setIcon(icon)
        msg.setWindowTitle(title)
        msg.setText(message)
        msg.exec()

    def select_files(self):
        """
        Open a file dialog to select PDF files.
        Update the label with the number of selected files and enable the process button.
        """
        selected_files, _ = QFileDialog.getOpenFileNames(self, "Alegeți fișiere PDF", self.last_directory,
                                                         "PDF Files (*.pdf)")
        if selected_files:
            self.files = selected_files
            self.label.setText(f"{len(selected_files)} fișier(e) selectat(e)")
            self.process_button.setEnabled(True)
            self.last_directory = os.path.dirname(selected_files[0])
            save_last_directory(self.last_directory)

    def process_files(self):
        """
        Start the background thread to process the selected PDFs.
        Update UI elements to indicate the processing status.
        """
        if not self.files:
            self.show_message("Eroare", "Vă rugăm să selectați fișiere PDF!", QMessageBox.Icon.Critical)
            return

        self.process_button.setEnabled(False)
        self.status_label.setText("Procesare în curs...")
        self.progress_bar.setValue(0)

        percentage = self.percentage_input.value() / 100.0
        self.processing_thread = ProcessingThread(self.files, percentage)

        self.processing_thread.progress_signal.connect(self.on_progress_update)
        self.processing_thread.finished.connect(self.on_processing_finished)

        self.processing_thread.start()

    def on_progress_update(self, value):
        """
        Update the progress bar value based on the emitted progress signal.
        :param value: An integer between 0 and 100 representing the current progress.
        """
        self.progress_bar.setValue(value)

    def on_processing_finished(self, output_path):
        """
        Handle the finalization of the processing thread.
        Prompt the user to save the resulting Excel file or display errors.
        :param output_path: The final path of the generated Excel file,
                            or empty string if an error occurred.
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
