import json
import os
import shutil
from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd
from PyQt6.QtCore import QThread, pyqtSignal, Qt, QSize, QTimer
from PyQt6.QtGui import QIcon, QFont, QPixmap, QCursor
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QProgressBar, QSpinBox,
                             QMessageBox, QScrollArea, QFrame, QAbstractSpinBox, QHBoxLayout, QRadioButton,
                             QButtonGroup, QGroupBox, QCheckBox, QToolButton)

from config import Constants
from functions import merge_existing_with_new
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
    finished = pyqtSignal(str, pd.DataFrame)
    progress_signal = pyqtSignal(int)

    def __init__(self, pdf_files, percentage, existing_excel=None):
        """
        :param pdf_files: List of PDF file paths to process.
        :param percentage: Float percentage for calculations.
        :param existing_excel: Path to existing Excel file to append to, or None.
        """
        super().__init__()
        self.pdf_files = pdf_files
        self.percentage = percentage
        self.existing_excel = existing_excel

    def copy_files_to_temp(self, temp_dir: str):
        """
        Copy PDF files to a temporary directory.
        :return: List of new file paths.
        """
        temp_paths = []
        for pdf in self.pdf_files:
            try:
                temp_path = shutil.copy(str(Path(pdf)), os.path.join(temp_dir, Path(pdf).name))
                temp_paths.append(temp_path)
            except Exception as e:
                print(f"[LOG] Error copying file {pdf}: {e}")
        return temp_paths

    @staticmethod
    def generate_excel_filename():
        """
        Generate an Excel filename based on current month and year.
        """
        return f"{datetime.now().strftime('%m-%Y')}-EXP.xlsx"

    def process_invoices(self):
        """
        Process PDF invoices in a temp directory; create the final Excel file.
        :return: Final path of the generated Excel and DataFrame.
        """
        try:
            with TemporaryDirectory() as temp_dir:
                temp_paths = self.copy_files_to_temp(temp_dir)
                processor = InvoiceProcessor(temp_paths, progress_callback=self._progress_callback)
                processor.process_invoices()

                if self.existing_excel and os.path.exists(self.existing_excel):
                    processor.df = merge_existing_with_new(self.existing_excel, processor.df)

                excel_filename = self.generate_excel_filename()
                temp_excel_path = os.path.join(temp_dir, excel_filename)
                generator = ExcelGenerator(processor.df, self.percentage)
                generator.generate_excel(temp_excel_path)

                final_path = os.path.join(os.getcwd(), excel_filename)
                shutil.move(temp_excel_path, final_path)

                return final_path, processor.df
        except Exception as e:
            print(f"[LOG] Error in process_invoices: {e}")
            return "", pd.DataFrame()

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
            output_path, df = self.process_invoices()
            self.finished.emit(output_path, df)
        except Exception as e:
            print(f"[LOG] Error in ProcessingThread: {e}")
            self.finished.emit("", pd.DataFrame())


# noinspection PyUnresolvedReferences
class PDFToExcelApp(QWidget):
    """
    Main window for selecting and processing PDF invoices into Excel.
    """

    def __init__(self):
        super().__init__()
        self.setWindowIcon(QIcon("assets/icon.png"))
        self.resize(1024, 768)

        font = QFont("Segoe UI", 10)
        self.setFont(font)
        self.setWindowTitle("Extractor PDF în Excel")

        self.append_to_excel = None
        self.auto_open = None
        self.content_frame = None
        self.content_layout = None
        self.create_new_excel = None
        self.excel_file = None
        self.excel_options = None
        self.excel_path_label = None
        self.file_list_layout = None
        self.file_list_widget = None
        self.last_directory = load_last_directory()
        self.pdf_count_label = None
        self.pdf_files = []
        self.percentage_input = None
        self.percentage_label = None
        self.process_button = None
        self.processing_thread = None
        self.progress_bar = None
        self.select_excel_button = None
        self.select_folder_button = None
        self.select_pdf_button = None
        self.status_label = None
        self.status_message = None

        self.main_layout = QVBoxLayout()
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.setLayout(self.main_layout)

        self.create_content_area()
        self.apply_light_theme()

    def create_content_area(self):
        """Create the main content area with all UI elements."""
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)

        self.content_frame = QFrame()
        self.content_layout = QVBoxLayout(self.content_frame)
        self.content_layout.setSpacing(15)
        self.content_layout.setContentsMargins(20, 20, 20, 20)

        self.create_file_section()
        self.create_options_section()
        self.create_action_section()

        scroll_area.setWidget(self.content_frame)
        self.main_layout.addWidget(scroll_area)

        self.create_status_bar()

    def create_file_section(self):
        """Create the file selection section."""
        excel_group = QGroupBox("Fișier Excel")
        excel_group.setObjectName("firstGroupBox")
        excel_layout = QVBoxLayout()
        excel_layout.setSpacing(8)

        excel_options_layout = QVBoxLayout()
        self.excel_options = QButtonGroup(self)
        self.create_new_excel = QRadioButton("Creează fișier Excel nou")
        self.create_new_excel.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.append_to_excel = QRadioButton("Adaugă într-un fișier Excel existent")
        self.append_to_excel.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.excel_options.addButton(self.create_new_excel)
        self.excel_options.addButton(self.append_to_excel)
        self.create_new_excel.setChecked(True)

        excel_options_layout.addWidget(self.create_new_excel)
        excel_options_layout.addWidget(self.append_to_excel)
        excel_layout.addLayout(excel_options_layout)

        excel_file_layout = QHBoxLayout()
        self.excel_path_label = QLabel("Niciun fișier selectat")

        self.select_excel_button = QPushButton("  Selectează fișier Excel")
        self.select_excel_button.setIcon(QIcon("assets/icons/file-excel.svg"))
        self.select_excel_button.setEnabled(False)
        self.select_excel_button.clicked.connect(self.select_excel_file)

        excel_file_layout.addWidget(self.excel_path_label, 1)
        excel_file_layout.addWidget(self.select_excel_button)

        excel_layout.addLayout(excel_file_layout)
        excel_group.setLayout(excel_layout)

        pdf_group = QGroupBox("Fișiere PDF")
        pdf_layout = QVBoxLayout()
        pdf_layout.setSpacing(8)

        pdf_buttons_layout = QHBoxLayout()

        self.select_pdf_button = QPushButton("  Selectează fișiere PDF")
        self.select_pdf_button.setIcon(QIcon("assets/icons/file-pdf-white.svg"))
        self.select_pdf_button.clicked.connect(self.select_pdf_files)
        self.select_pdf_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

        self.select_folder_button = QPushButton("  Selectează dosar cu fișiere PDF")
        self.select_folder_button.setIcon(QIcon("assets/icons/folder-open.svg"))
        self.select_folder_button.clicked.connect(self.select_pdf_folder)
        self.select_folder_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

        pdf_buttons_layout.addWidget(self.select_pdf_button)
        pdf_buttons_layout.addWidget(self.select_folder_button)

        pdf_layout.addLayout(pdf_buttons_layout)

        self.file_list_widget = QScrollArea()
        self.file_list_widget.setWidgetResizable(True)

        file_list_content = QFrame()
        self.file_list_layout = QVBoxLayout(file_list_content)
        self.file_list_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.file_list_widget.setWidget(file_list_content)

        pdf_layout.addWidget(self.file_list_widget)

        self.pdf_count_label = QLabel("Niciun fișier selectat")
        pdf_layout.addWidget(self.pdf_count_label)

        pdf_group.setLayout(pdf_layout)

        self.content_layout.addWidget(excel_group)
        self.content_layout.addWidget(pdf_group)

        self.create_new_excel.toggled.connect(self.toggle_excel_selection)
        self.append_to_excel.toggled.connect(self.toggle_excel_selection)

    def create_options_section(self):
        """Create the options section."""
        options_section = QGroupBox("Opțiuni de procesare")
        options_layout = QVBoxLayout(options_section)
        options_layout.setSpacing(10)

        percentage_layout = QHBoxLayout()
        self.percentage_label = QLabel("Procent de calcul (0–99%):")
        self.percentage_input = QSpinBox()
        self.percentage_input.setRange(0, 99)
        self.percentage_input.setValue(60)
        self.percentage_input.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.percentage_input.wheelEvent = lambda event: None

        percentage_layout.addWidget(self.percentage_label)
        percentage_layout.addWidget(self.percentage_input)

        options_layout.addLayout(percentage_layout)

        self.auto_open = QCheckBox("Deschide automat fișierul după procesare")
        self.auto_open.setChecked(False)
        self.auto_open.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        options_layout.addWidget(self.auto_open)

        self.content_layout.addWidget(options_section)

    def create_action_section(self):
        """Create the action section with process button and progress indicators."""
        action_section = QGroupBox("Procesează")
        action_layout = QVBoxLayout(action_section)
        action_layout.setSpacing(10)

        self.process_button = QPushButton("  Procesează fișierele PDF și generează Excel")
        self.process_button.setIcon(QIcon("assets/icons/circle-play.svg"))
        self.process_button.setIconSize(QSize(20, 20))
        self.process_button.setEnabled(False)
        self.process_button.clicked.connect(self.start_file_processing)
        self.process_button.setMinimumHeight(40)
        action_layout.addWidget(self.process_button)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("%p%")
        action_layout.addWidget(self.progress_bar)

        self.status_label = QLabel("Selectează fișierele PDF și opțiunile, apoi apasă pe butonul de procesare.")
        self.status_label.setWordWrap(True)
        action_layout.addWidget(self.status_label)

        self.content_layout.addWidget(action_section)

    def create_status_bar(self):
        """Create the status bar at the bottom of the window."""
        status_bar = QFrame()
        status_bar.setFixedHeight(25)

        status_layout = QHBoxLayout(status_bar)
        status_layout.setContentsMargins(10, 0, 10, 0)

        self.status_message = QLabel("Gata")
        version_label = QLabel(f"v{Constants.APPLICATION_VERSION}")

        status_layout.addWidget(self.status_message, 1)
        status_layout.addWidget(version_label)

        self.main_layout.addWidget(status_bar)

    def apply_light_theme(self):
        """
        Apply the light theme to the application.
        """
        bg_color = "#F9FAFB"
        text_color = "#111827"
        accent_color = "#2563EB"
        accent_hover = "#3B82F6"
        border_color = "#D1D5DB"
        content_bg = "#FFFFFF"
        disabled_bg = "#9CA3AF"
        disabled_border = "#6B7280"

        self.setStyleSheet(f"""
            QWidget {{
                background-color: {bg_color};
                margin: 0;
                padding: 0;
                font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif;
                color: {text_color};
            }}

            QFrame {{
                background-color: {bg_color};
            }}

            QLabel {{
                background-color: transparent;
            }}

            QGroupBox {{
                font-weight: normal;
                border: 1px solid {border_color};
                border-radius: 8px;
                margin-top: 16px;
                padding-top: 16px;
                background-color: {content_bg};
            }}
            QGroupBox#firstGroupBox {{
                margin-top: 0;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }}

            QPushButton {{
                padding: 8px 16px;
                font-size: 14px;
                color: #FFFFFF;
                background-color: {accent_color};
                border: 1px solid {accent_color};
                border-radius: 6px;
            }}
            QPushButton:hover {{
                background-color: {accent_hover};
            }}
            QPushButton:disabled {{
                background-color: {disabled_bg};
                border: 1px solid {disabled_border};
            }}

            QSpinBox, QComboBox {{
                padding: 8px;
                color: {text_color};
                background-color: {content_bg};
                border: 1px solid {border_color};
                border-radius: 4px;
            }}

            QProgressBar {{
                height: 20px;
                text-align: center;
                color: {text_color};
                background-color: {content_bg};
                border: 1px solid {border_color};
                border-radius: 6px;
            }}
            QProgressBar::chunk {{
                background-color: {accent_color};
                border-radius: 6px;
            }}

            QRadioButton, QCheckBox {{
                background-color: {content_bg};
            }}
            QRadioButton::indicator, QCheckBox::indicator {{
                width: 16px;
                height: 16px;
            }}
            QRadioButton::indicator:unchecked {{
                image: url("assets/icons/circle.svg");
            }}
            QRadioButton::indicator:checked {{
                image: url("assets/icons/circle-dot.svg");
            }}
            QCheckBox::indicator:unchecked {{
                image: url("assets/icons/square.svg");
            }}
            QCheckBox::indicator:checked {{
                image: url("assets/icons/square-check.svg");
            }}

            QScrollArea {{
                border: none;
                background-color: transparent;
            }}

            QToolButton {{
                padding: 2px;
                border: none;
                border-radius: 4px;
                background-color: transparent;
            }}
            QToolButton:hover {{
                background-color: {border_color};
            }}
            QToolButton:disabled {{
                background-color: transparent;
                opacity: 0.6;
            }}
        """)

    def handle_pdf_selection(self, pdf_files):
        """Handle the selection of PDF files."""
        if pdf_files:
            self.pdf_files = pdf_files
            self.update_file_list()
            self.pdf_count_label.setText(f"{len(pdf_files)} fișier(e) PDF selectat(e)")
            self.last_directory = os.path.dirname(pdf_files[0])
            save_last_directory(self.last_directory)
            self.update_process_button_state()

            self.status_message.setText(f"Au fost adăugate {len(pdf_files)} fișier(e) PDF")
            QTimer.singleShot(3000, lambda: self.status_message.setText("Gata"))

    def update_file_list(self):
        """Update the file list widget with the selected PDF files."""
        while self.file_list_layout.count():
            item = self.file_list_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        for pdf_file in self.pdf_files:
            file_item = QFrame()
            file_item_layout = QHBoxLayout(file_item)
            file_item_layout.setContentsMargins(5, 5, 5, 5)

            file_icon = QLabel()
            file_icon.setPixmap(QPixmap("assets/icons/file-pdf.svg").scaled(16, 16, Qt.AspectRatioMode.KeepAspectRatio,
                                                                            Qt.TransformationMode.SmoothTransformation))

            file_name = QLabel(os.path.basename(pdf_file))
            file_name.setToolTip(pdf_file)

            remove_button = QToolButton()
            remove_button.setIcon(QIcon("assets/icons/trash-can.svg"))
            remove_button.setToolTip("Șterge fișierul")
            remove_button.clicked.connect(lambda checked, path=pdf_file, btn=remove_button: self.remove_file(path, btn))
            remove_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

            file_item_layout.addWidget(file_icon)
            file_item_layout.addWidget(file_name, 1)
            file_item_layout.addWidget(remove_button)

            self.file_list_layout.addWidget(file_item)

    def remove_file(self, file_path, btn):
        """
        Remove a file from the list of selected PDF files.
        Temporarily disable the button that triggered this action,
        then re-enable it after the logic is complete.
        """
        btn.setEnabled(False)

        try:
            if file_path in self.pdf_files:
                self.pdf_files.remove(file_path)
                self.update_file_list()
                self.pdf_count_label.setText(f"{len(self.pdf_files)} fișier(e) PDF selectat(e)")
                self.update_process_button_state()

                self.status_message.setText(f"A fost șters fișierul: {os.path.basename(file_path)}")
                QTimer.singleShot(3000, lambda: self.status_message.setText("Gata"))
        finally:
            btn.setEnabled(True)
            btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

    def toggle_excel_selection(self):
        """
        Enable or disable Excel file selection based on radio button state.
        """
        self.select_excel_button.setEnabled(self.append_to_excel.isChecked())
        if self.select_excel_button.isEnabled():
            self.select_excel_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        if self.create_new_excel.isChecked():
            self.excel_file = None
            self.excel_path_label.setText("Niciun fișier selectat")
        self.update_process_button_state()

    def show_message(self, title, message, icon=QMessageBox.Icon.Information):
        """
        Show a message box with a given title, message, and optional icon.
        """
        msg = QMessageBox(self)
        msg.setIcon(icon)
        msg.setWindowTitle(title)
        msg.setText(message)
        msg.exec()

    def select_excel_file(self):
        """
        Prompt a file dialog to select an Excel file.
        """
        selected_file, _ = QFileDialog.getOpenFileName(self, "Selectează fișier Excel existent", self.last_directory,
                                                       "Excel Files (*.xlsx)")

        if selected_file:
            self.excel_file = selected_file
            self.excel_path_label.setText(os.path.basename(selected_file))
            self.last_directory = os.path.dirname(selected_file)
            save_last_directory(self.last_directory)
            self.update_process_button_state()

            self.status_message.setText(f"A fost selectat fișierul Excel: {os.path.basename(selected_file)}")
            QTimer.singleShot(3000, lambda: self.status_message.setText("Gata"))

    def select_pdf_files(self):
        """
        Prompt a file dialog to select PDF files; enable processing if any selected.
        """
        selected_files, _ = QFileDialog.getOpenFileNames(self, "Selectează fișiere PDF pentru procesare",
                                                         self.last_directory, "PDF Files (*.pdf)")

        if selected_files:
            self.handle_pdf_selection(selected_files)

    def select_pdf_folder(self):
        """
        Prompt a folder dialog to select a directory with PDF files.
        """
        selected_folder = QFileDialog.getExistingDirectory(self, "Selectează un dosar care conține fișiere PDF",
                                                           self.last_directory)

        if selected_folder:
            pdf_files = []
            for root, _, files in os.walk(selected_folder):
                for file in files:
                    if file.lower().endswith('.pdf'):
                        pdf_files.append(os.path.join(root, file))

            if pdf_files:
                self.handle_pdf_selection(pdf_files)
            else:
                self.show_message("Niciun fișier PDF găsit", "Niciun fișier PDF nu a fost găsit în dosarul selectat.",
                                  QMessageBox.Icon.Information)

    def update_process_button_state(self):
        """
        Enable the process button if we have PDF files and, if appending, an Excel file.
        """
        has_pdfs = len(self.pdf_files) > 0
        excel_ok = self.create_new_excel.isChecked() or (
                self.append_to_excel.isChecked() and self.excel_file is not None)
        self.process_button.setEnabled(has_pdfs and excel_ok)
        if self.process_button.isEnabled():
            self.process_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

    def start_file_processing(self):
        """
        Initialize the background thread for PDF processing and update the UI.
        """
        if not self.pdf_files:
            self.show_message("Niciun fișier PDF selectat", "Selectează cel puțin un fișier PDF pentru procesare!",
                              QMessageBox.Icon.Critical)
            return

        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Icon.Question)
        msg_box.setWindowTitle("Confirmă procesarea")
        msg_box.setText(f"Doriți să procesați {len(self.pdf_files)} fișier(e) PDF?")
        msg_box.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        msg_box.setDefaultButton(QMessageBox.StandardButton.Yes)

        yes_button = msg_box.button(QMessageBox.StandardButton.Yes)
        yes_button.setText("Da")
        yes_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

        no_button = msg_box.button(QMessageBox.StandardButton.No)
        no_button.setText("Nu")
        no_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

        confirm = msg_box.exec()

        if confirm == QMessageBox.StandardButton.No:
            return

        self.process_button.setEnabled(False)
        self.select_pdf_button.setEnabled(False)
        self.select_folder_button.setEnabled(False)
        self.select_excel_button.setEnabled(False)

        self.status_label.setText("Se procesează... Te rog așteaptă.")
        self.status_message.setText("Procesare în desfășurare")
        self.progress_bar.setValue(0)

        percentage = self.percentage_input.value() / 100.0
        excel_path = self.excel_file if self.append_to_excel.isChecked() else None

        self.processing_thread = ProcessingThread(self.pdf_files, percentage, excel_path)
        self.processing_thread.progress_signal.connect(self.handle_progress_update)
        self.processing_thread.finished.connect(self.handle_processing_finished)
        self.processing_thread.start()

    def handle_progress_update(self, value):
        """
        Update the progress bar with current progress value.
        """
        self.progress_bar.setValue(value)

    def handle_processing_finished(self, output_path, df):
        """
        Handle finishing of processing. Allow user to save or show an error.
        """
        self.process_button.setEnabled(True)
        self.process_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

        self.select_pdf_button.setEnabled(True)
        self.select_pdf_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

        self.select_folder_button.setEnabled(True)
        self.select_folder_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

        self.select_excel_button.setEnabled(self.append_to_excel.isChecked())
        if self.select_excel_button.isEnabled():
            self.select_excel_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

        self.progress_bar.setValue(100)

        if output_path and os.path.exists(output_path):
            num_entries = len(df)

            if self.append_to_excel.isChecked() and self.excel_file:
                try:
                    existing_df = pd.read_excel(self.excel_file)

                    if 'Nr Crt' in existing_df.columns:
                        existing_df = existing_df[existing_df['Nr Crt'].apply(lambda x: pd.notna(x) and (
                                isinstance(x, (int, float)) or (isinstance(x, str) and x.isdigit())))]

                    num_existing = len(existing_df)
                    num_new = num_entries - num_existing

                    if num_new > 0:
                        self.status_label.setText(
                            f"Procesare finalizată cu succes! Au fost adăugate {num_new} înregistrări noi. "
                            f"Total: {num_entries} înregistrări.")
                    else:
                        self.status_label.setText(
                            f"Procesare finalizată! Nicio înregistrare nouă. Total: {num_entries} înregistrări.")
                except Exception as e:
                    print(f"[LOG] Error calculating new entries: {e}")
                    self.status_label.setText(f"Procesare finalizată cu succes! {num_entries} înregistrări totale.")
            else:
                self.status_label.setText(f"Procesare finalizată cu succes! {num_entries} înregistrări.")

            success_message = QMessageBox(self)
            success_message.setIcon(QMessageBox.Icon.Information)
            success_message.setWindowTitle("Procesare finalizată cu succes")
            success_message.setText("PDF-urile au fost procesate cu succes!")
            success_message.setInformativeText(
                f"Au fost procesate {len(self.pdf_files)} fișiere PDF.\nAu fost extrase {num_entries} înregistrări.")

            success_message.setStandardButtons(QMessageBox.StandardButton.Save | QMessageBox.StandardButton.Discard)
            success_message.setDefaultButton(QMessageBox.StandardButton.Save)

            save_button = success_message.button(QMessageBox.StandardButton.Save)
            save_button.setText("Salvează")
            save_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

            discard_button = success_message.button(QMessageBox.StandardButton.Discard)
            discard_button.setText("Renunță")
            discard_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

            result = success_message.exec()

            if result == QMessageBox.StandardButton.Save:
                default_save_path = os.path.join(self.last_directory, os.path.basename(output_path))
                save_path, _ = QFileDialog.getSaveFileName(self, "Salvează fișierul Excel generat", default_save_path,
                                                           "Excel Files (*.xlsx)")

                if save_path:
                    try:
                        shutil.move(output_path, save_path)
                        self.status_label.setText(f"Fișierul Excel a fost salvat cu succes la: {save_path}")
                        self.status_message.setText("Fișier salvat cu succes")
                        self.last_directory = os.path.dirname(save_path)
                        save_last_directory(self.last_directory)

                        if self.auto_open.isChecked():
                            os.startfile(save_path)
                    except Exception as e:
                        print(f"[LOG] Error saving file: {e}")
                        self.show_message("Eroare la salvare", f"Eroare la salvarea fișierului: {str(e)}",
                                          QMessageBox.Icon.Critical)
                else:
                    os.remove(output_path)
                    self.status_label.setText("Salvarea fișierului Excel a fost anulată.")
                    self.status_message.setText("Salvare anulată")
            else:
                os.remove(output_path)
                self.status_label.setText("Fișierul Excel a fost șters.")
                self.status_message.setText("Fișier șters")
        else:
            self.status_label.setText("Procesare eșuată. Verifică fișierele PDF și încearcă din nou.")
            self.status_message.setText("Procesare eșuată")
            self.show_message("Eroare de procesare",
                              "Nu s-au putut procesa fișierele PDF. Verifică validitatea fișierelor și încearcă din nou.",
                              QMessageBox.Icon.Critical)
