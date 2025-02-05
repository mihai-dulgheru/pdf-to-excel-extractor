import json
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory

from PyQt6.QtCore import QThread, pyqtSignal
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QProgressBar,
                             QSpinBox, QMessageBox)

from modules.excel_generator import ExcelGenerator
from modules.invoice_processor import InvoiceProcessor

CONFIG_FILE = "config.json"


def load_last_directory():
    """Load the last accessed directory from the configuration file."""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f).get("last_directory", os.path.expanduser("~"))
        except (json.JSONDecodeError, IOError):
            pass
    return os.path.expanduser("~")


def save_last_directory(directory: str):
    """Save the last accessed directory to the configuration file."""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json_data = json.dumps({"last_directory": directory}, indent=4)
            f.write(json_data)
    except IOError:
        print("Failed to save last directory")


class ProcessingThread(QThread):
    finished = pyqtSignal(str)

    def __init__(self, files, percentage):
        super().__init__()
        self.files = files
        self.percentage = percentage

    def copy_files_to_temp(self, temp_dir: str):
        """Copy input PDFs to a temporary directory and return their paths."""
        return [shutil.copy(str(Path(file)), os.path.join(temp_dir, Path(file).name)) for file in self.files]

    @staticmethod
    def generate_excel_filename():
        """Generate Excel filename based on the current month and year."""
        return f"{datetime.now().strftime('%m-%Y')}-EXP.xlsx"

    def process_invoices(self):
        """Process invoices and generate an Excel file."""
        with TemporaryDirectory() as temp_dir:
            input_paths = self.copy_files_to_temp(temp_dir)

            processor = InvoiceProcessor(input_paths)
            processor.process_invoices()

            excel_filename = self.generate_excel_filename()
            output_temp_path = os.path.join(temp_dir, excel_filename)

            excel_generator = ExcelGenerator(processor.df, self.percentage)
            excel_generator.generate_excel(output_temp_path)

            output_final_path = os.path.join(os.getcwd(), excel_filename)
            shutil.move(output_temp_path, output_final_path)

        return output_final_path

    def run(self):
        try:
            output_path = self.process_invoices()
            self.finished.emit(output_path)
        except Exception as e:
            print(f"Error in processing thread: {e}")
            self.finished.emit("")


class PDFToExcelApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Procesor PDF în Excel")
        self.setGeometry(100, 100, 500, 300)
        self.layout = QVBoxLayout()

        self.label = QLabel("Selectați fișiere PDF")
        self.layout.addWidget(self.label)

        self.select_button = QPushButton("Alegeți fișiere")
        self.select_button.clicked.connect(self.select_files)
        self.layout.addWidget(self.select_button)

        self.percentage_label = QLabel("Introduceți procentul:")
        self.layout.addWidget(self.percentage_label)

        self.percentage_input = QSpinBox()
        self.percentage_input.setRange(0, 99)
        self.percentage_input.setValue(60)
        self.layout.addWidget(self.percentage_input)

        self.process_button = QPushButton("Încărcați și procesați")
        self.process_button.clicked.connect(self.process_files)
        self.process_button.setEnabled(False)
        self.layout.addWidget(self.process_button)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.layout.addWidget(self.progress_bar)

        self.status_label = QLabel("")
        self.layout.addWidget(self.status_label)

        self.setLayout(self.layout)
        self.files = []
        self.processing_thread = None
        self.last_directory = load_last_directory()

    def show_message(self, title, message, icon=QMessageBox.Icon.Information):
        """Display a message box with the given title and message."""
        msg = QMessageBox(self)
        msg.setIcon(icon)
        msg.setWindowTitle(title)
        msg.setText(message)
        msg.exec()

    def select_files(self):
        selected_files, _ = QFileDialog.getOpenFileNames(self, "Alegeți fișiere PDF", self.last_directory,
                                                         "PDF Files (*.pdf)")
        if selected_files:
            self.files = selected_files
            self.label.setText(f"{len(selected_files)} fișier(e) selectat(e)")
            self.process_button.setEnabled(True)
            self.last_directory = os.path.dirname(selected_files[0])
            save_last_directory(self.last_directory)

    def process_files(self):
        if not self.files:
            self.show_message("Eroare", "Vă rugăm să selectați fișiere PDF!", QMessageBox.Icon.Critical)
            return

        self.process_button.setEnabled(False)
        self.status_label.setText("Procesare în curs...")
        self.progress_bar.setValue(50)

        percentage = self.percentage_input.value() / 100.0
        self.processing_thread = ProcessingThread(self.files, percentage)
        self.processing_thread.finished.connect(self.on_processing_finished)
        self.processing_thread.start()

    def on_processing_finished(self, output_path):
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


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFToExcelApp()
    window.show()
    sys.exit(app.exec())
