import json
import os
import re
import shutil
from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory, gettempdir, mkdtemp

import pandas as pd
from PyQt6.QtCore import QThread, pyqtSignal, Qt, QSize, QTimer
from PyQt6.QtGui import QIcon, QFont, QPixmap, QCursor
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QProgressBar, QSpinBox,
                             QMessageBox, QScrollArea, QFrame, QAbstractSpinBox, QHBoxLayout, QRadioButton,
                             QButtonGroup, QGroupBox, QCheckBox, QToolButton)

from config import Constants
from functions import get_logger
from modules.excel_generator import ExcelGenerator
from modules.invoice_processor import InvoiceProcessor

logger = get_logger("pdf_to_excel.ui")


def asset(name):
    """Absolute path of a packaged asset, independent of the working directory."""
    return str(Constants.ASSETS_DIR / name)


def discard_staging(output_path):
    """
    Remove a generated workbook that is not wanted any more, together with its staging folder, but only when that
    folder is one process_invoices() created under the temp dir. Anywhere else just the workbook goes.
    """
    folder = os.path.dirname(output_path)
    if os.path.basename(folder).startswith("pdf_to_excel_") and \
            os.path.normcase(os.path.dirname(folder)) == os.path.normcase(gettempdir()):
        shutil.rmtree(folder, ignore_errors=True)
    elif os.path.exists(output_path):
        os.remove(output_path)


def collect_pdfs(paths):
    """PDF files among the given files and folders (folders searched recursively), in a stable order."""
    pdfs = []
    for path in paths:
        if os.path.isdir(path):
            for root, _, files in os.walk(path):
                pdfs += [os.path.join(root, f) for f in sorted(files, key=str.lower) if f.lower().endswith(".pdf")]
        elif path.lower().endswith(".pdf"):
            pdfs.append(path)
    return pdfs


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
        logger.error("Failed to save last directory")


def unique_paths(paths):
    """The same file selected twice (or via a differently spelled path) is kept once, in selection order."""
    seen, unique = set(), []
    for path in paths:
        key = os.path.normcase(os.path.abspath(path))
        if key not in seen:
            seen.add(key)
            unique.append(path)
    return unique


# noinspection PyUnresolvedReferences
class ProcessingThread(QThread):
    """
    Background thread that processes PDF invoices, generates an Excel file,
    and emits signals for progress and completion.
    """
    finished = pyqtSignal(str, pd.DataFrame, list)
    progress_signal = pyqtSignal(int)

    def __init__(self, pdf_files, percentage, existing_excel=None):
        """
        :param pdf_files: List of PDF file paths to process.
        :param percentage: Float percentage for calculations.
        :param existing_excel: Path to the existing Excel file to append to, or None.
        """
        super().__init__()
        self.pdf_files = pdf_files
        self.percentage = percentage
        self.existing_excel = existing_excel
        self.temp_names = {}
        self.copy_problems = []

    def copy_files_to_temp(self, temp_dir: str):
        """
        Copy PDF files to a temporary directory.
        :return: List of new file paths.
        """
        temp_paths = []
        names = [Path(pdf).name for pdf in self.pdf_files]
        for i, pdf in enumerate(self.pdf_files):
            # messages name the file as selected; when two selected files share a name, the folder is shown too
            shown = f"{Path(pdf).parent.name}/{Path(pdf).name}" if names.count(Path(pdf).name) > 1 else Path(pdf).name
            try:
                # index prefix: files with the same name in different folders must not overwrite each other
                temp_path = shutil.copy(str(Path(pdf)), os.path.join(temp_dir, f"{i:04d}_{Path(pdf).name}"))
                temp_paths.append(temp_path)
                self.temp_names[os.path.basename(temp_path)] = shown
            except Exception as e:
                logger.error(f"Error copying file {pdf}: {e}")
                self.copy_problems.append(f"{shown}: fișierul nu a putut fi copiat ({e})")
        return temp_paths

    def _with_original_names(self, message):
        """Temp names -> selected names, in one pass so a replaced name is never matched again."""
        if not self.temp_names:
            return message
        pattern = "|".join(map(re.escape, sorted(self.temp_names, key=len, reverse=True)))
        return re.sub(pattern, lambda m: self.temp_names[m.group()], message)

    @staticmethod
    def generate_excel_filename():
        """
        Generate an Excel filename based on the current month and year.
        """
        return f"{datetime.now().strftime('%d-%m-%Y')}-EXP.xlsx"

    def process_invoices(self):
        """
        Process PDF invoices in a temp directory; create the final Excel file.
        :return: Final path of the generated Excel, DataFrame and the list of problem messages.
        """
        try:
            with TemporaryDirectory() as temp_dir:
                temp_paths = self.copy_files_to_temp(temp_dir)
                processor = InvoiceProcessor(temp_paths, progress_callback=self._progress_callback)
                processor.process_invoices()
                if processor.fatal_error:
                    return "", pd.DataFrame(), [processor.fatal_error]
                problems = self.copy_problems + [self._with_original_names(m) for m in processor.problems]

                excel_filename = self.generate_excel_filename()
                temp_excel_path = os.path.join(temp_dir, excel_filename)
                generator = ExcelGenerator(processor.df, self.percentage)
                generator.generate_excel(temp_excel_path, self.existing_excel)

                # staged in a folder of its own until the user picks where to save it: a fixed name in a shared
                # folder could overwrite the workbook saved there earlier today
                final_path = os.path.join(mkdtemp(prefix="pdf_to_excel_"), excel_filename)
                shutil.move(temp_excel_path, final_path)

                written = generator.appended if self.existing_excel else processor.df
                return final_path, written, problems
        except Exception as e:
            logger.error(f"Error in process_invoices: {e}")
            return "", pd.DataFrame(), [f"Eroare la generarea fișierului Excel: {e}"]

    def _progress_callback(self, processed, total):
        """
        Calculate integer progress and emit the signal.
        """
        progress_val = int((processed / total) * 100)
        self.progress_signal.emit(progress_val)

    def run(self):
        """
        Run the invoice processing; emit 'finished' with the output path and DataFrame.
        """
        try:
            output_path, df, problems = self.process_invoices()
            self.finished.emit(output_path, df, problems)
        except Exception as e:
            logger.error(f"Error in ProcessingThread: {e}")
            self.finished.emit("", pd.DataFrame(), [])


# noinspection PyUnresolvedReferences
class PDFToExcelApp(QWidget):
    """
    Main window for selecting and processing PDF invoices into Excel.
    """

    def __init__(self):
        super().__init__()
        self.setWindowIcon(QIcon(asset("icon.png")))
        self.resize(1024, 768)
        self.setAcceptDrops(True)

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
        self.update_file_list()  # shows the empty-state hint

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
        self.select_excel_button.setIcon(QIcon(asset("icons/file-excel.svg")))
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
        self.select_pdf_button.setIcon(QIcon(asset("icons/file-pdf-white.svg")))
        self.select_pdf_button.clicked.connect(self.select_pdf_files)
        self.select_pdf_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

        self.select_folder_button = QPushButton("  Selectează dosar cu fișiere PDF")
        self.select_folder_button.setIcon(QIcon(asset("icons/folder-open.svg")))
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
        self.clear_list_button = QToolButton()
        self.clear_list_button.setText("Golește lista")
        self.clear_list_button.setIcon(QIcon(asset("icons/trash-can.svg")))
        self.clear_list_button.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        self.clear_list_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.clear_list_button.setEnabled(False)
        self.clear_list_button.clicked.connect(self.clear_file_list)
        count_layout = QHBoxLayout()
        count_layout.addWidget(self.pdf_count_label, 1)
        count_layout.addWidget(self.clear_list_button)
        pdf_layout.addLayout(count_layout)

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
        self.percentage_label = QLabel("Procent de calcul:")
        self.percentage_input = QSpinBox()
        self.percentage_input.setRange(0, 99)
        self.percentage_input.setValue(60)
        self.percentage_input.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.percentage_input.wheelEvent = lambda event: None
        self.percentage_input.setSuffix(" %")
        self.percentage_input.setFixedWidth(110)
        self.percentage_input.setToolTip("Coloana „%” din Excel: Statistica = ROUND(Valoare Ron + Transport × %, 0)")

        percentage_layout.addWidget(self.percentage_label)
        percentage_layout.addWidget(self.percentage_input)
        percentage_layout.addStretch()

        options_layout.addLayout(percentage_layout)

        self.auto_open = QCheckBox("Deschide automat fișierul după procesare")
        self.auto_open.setChecked(False)
        self.auto_open.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        options_layout.addWidget(self.auto_open)

        self.content_layout.addWidget(options_section)

    def create_action_section(self):
        """Create the action section with a process button and progress indicators."""
        action_section = QGroupBox("Procesează")
        action_layout = QVBoxLayout(action_section)
        action_layout.setSpacing(10)

        self.process_button = QPushButton("  Procesează fișierele PDF și generează Excel")
        self.process_button.setIcon(QIcon(asset("icons/circle-play.svg")))
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

        icons = Constants.ASSETS_DIR.joinpath("icons").as_posix()

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
            QLabel#emptyHint {{
                color: #6B7280;
                font-style: italic;
                padding: 16px 20px;
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
            QSpinBox:disabled, QRadioButton:disabled, QCheckBox:disabled, QToolButton:disabled {{
                color: {disabled_bg};
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
                image: url("{icons}/circle.svg");
            }}
            QRadioButton::indicator:checked {{
                image: url("{icons}/circle-dot.svg");
            }}
            QCheckBox::indicator:unchecked {{
                image: url("{icons}/square.svg");
            }}
            QCheckBox::indicator:checked {{
                image: url("{icons}/square-check.svg");
            }}
            QRadioButton::indicator:checked:disabled {{
                image: url("{icons}/circle-dot-disabled.svg");
            }}
            QCheckBox::indicator:checked:disabled {{
                image: url("{icons}/square-check-disabled.svg");
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
            }}
        """)

    def handle_pdf_selection(self, pdf_files):
        """Add the selected PDF files to the list; a file already in the list is kept once."""
        if not pdf_files:
            return
        merged = unique_paths(self.pdf_files + list(pdf_files))
        added = len(merged) - len(self.pdf_files)
        self.pdf_files = merged
        self.update_file_list()
        self.last_directory = os.path.dirname(pdf_files[0])
        save_last_directory(self.last_directory)

        self.status_message.setText(f"Au fost adăugate {added} fișier(e) PDF" if added
                                    else "Fișierele selectate erau deja în listă")
        QTimer.singleShot(3000, lambda: self.status_message.setText("Gata"))

    def _is_processing(self):
        return self.processing_thread is not None and self.processing_thread.isRunning()

    def dragEnterEvent(self, event):
        """Accept PDF files and folders dragged from Explorer, except while a run is in progress."""
        paths = [url.toLocalFile() for url in event.mimeData().urls()]
        if not self._is_processing() and any(os.path.isdir(p) or p.lower().endswith(".pdf") for p in paths):
            # always a copy: answering "move" (Shift held) would make Explorer delete the dropped files
            event.setDropAction(Qt.DropAction.CopyAction)
            event.accept()

    def dropEvent(self, event):
        """Runs inside Explorer's drag, so nothing modal here: a refused drop only leaves a status-bar note."""
        pdfs = collect_pdfs([url.toLocalFile() for url in event.mimeData().urls()])
        if not pdfs:
            event.ignore()
            self.status_message.setText("Elementele trase nu conțin fișiere PDF")
            QTimer.singleShot(3000, lambda: self.status_message.setText("Gata"))
            return
        event.setDropAction(Qt.DropAction.CopyAction)
        event.accept()
        self.handle_pdf_selection(pdfs)

    def clear_file_list(self):
        """Empty the list of selected PDF files."""
        self.pdf_files = []
        self.update_file_list()
        self.status_message.setText("Lista de fișiere a fost golită")
        QTimer.singleShot(3000, lambda: self.status_message.setText("Gata"))

    def update_file_list(self):
        """Rebuild the file list widget, the count label and the buttons that depend on the list."""
        while self.file_list_layout.count():
            item = self.file_list_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        if not self.pdf_files:
            hint = QLabel("Trage aici fișiere PDF sau un dosar, ori folosește butoanele de mai sus.")
            hint.setObjectName("emptyHint")
            hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
            hint.setWordWrap(True)
            self.file_list_layout.addWidget(hint)

        for pdf_file in self.pdf_files:
            file_item = QFrame()
            file_item_layout = QHBoxLayout(file_item)
            file_item_layout.setContentsMargins(5, 5, 5, 5)

            file_icon = QLabel()
            file_icon.setPixmap(QPixmap(asset("icons/file-pdf.svg")).scaled(16, 16, Qt.AspectRatioMode.KeepAspectRatio,
                                                                            Qt.TransformationMode.SmoothTransformation))

            file_name = QLabel(os.path.basename(pdf_file))
            file_name.setToolTip(pdf_file)

            remove_button = QToolButton()
            remove_button.setIcon(QIcon(asset("icons/trash-can.svg")))
            remove_button.setToolTip("Șterge fișierul")
            remove_button.clicked.connect(lambda checked, path=pdf_file, btn=remove_button: self.remove_file(path, btn))
            remove_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

            file_item_layout.addWidget(file_icon)
            file_item_layout.addWidget(file_name, 1)
            file_item_layout.addWidget(remove_button)

            self.file_list_layout.addWidget(file_item)

        n = len(self.pdf_files)
        self.pdf_count_label.setText(f"{n} fișier(e) PDF selectat(e)" if n else "Niciun fișier selectat")
        self.clear_list_button.setEnabled(n > 0)
        self.update_process_button_state()

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

                self.status_message.setText(f"A fost șters fișierul: {os.path.basename(file_path)}")
                QTimer.singleShot(3000, lambda: self.status_message.setText("Gata"))
        finally:
            btn.setEnabled(True)
            btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

    def toggle_excel_selection(self):
        """
        Enable or disable Excel file selection based on the radio button state.
        """
        self.select_excel_button.setEnabled(self.append_to_excel.isChecked())
        if self.select_excel_button.isEnabled():
            self.select_excel_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        if self.create_new_excel.isChecked():
            self.excel_file = None
            self.excel_path_label.setText("Niciun fișier selectat")
            self.excel_path_label.setToolTip("")
        self.update_process_button_state()

    def show_message(self, title, message, icon: QMessageBox.Icon = QMessageBox.Icon.Information):
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
            self.excel_path_label.setToolTip(selected_file)
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
            pdf_files = collect_pdfs([selected_folder])
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

        self._set_controls_enabled(False)

        self.status_label.setText("Se procesează... Te rog așteaptă.")
        self.status_message.setText("Procesare în desfășurare")
        self.progress_bar.setValue(0)

        percentage = self.percentage_input.value() / 100.0
        excel_path = self.excel_file if self.append_to_excel.isChecked() else None

        self.processing_thread = ProcessingThread(self.pdf_files, percentage, excel_path)
        self.processing_thread.progress_signal.connect(self.handle_progress_update)
        self.processing_thread.finished.connect(self.handle_processing_finished)
        self.processing_thread.start()

    def _set_controls_enabled(self, enabled):
        """Lock everything the user could change while the processing thread runs; unlock it afterwards."""
        for widget in (self.process_button, self.select_pdf_button, self.select_folder_button,
                       self.create_new_excel, self.append_to_excel, self.percentage_input, self.auto_open,
                       self.file_list_widget, self.clear_list_button):
            widget.setEnabled(enabled)
        self.select_excel_button.setEnabled(enabled and self.append_to_excel.isChecked())
        if enabled:
            self.clear_list_button.setEnabled(bool(self.pdf_files))
            self.update_process_button_state()
            for btn in (self.process_button, self.select_pdf_button, self.select_folder_button,
                        self.select_excel_button):
                btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

    def handle_progress_update(self, value):
        """
        Update the progress bar with the current progress value.
        """
        self.progress_bar.setValue(value)

    def handle_processing_finished(self, output_path, df, problems):
        """
        Handle finishing of processing: re-enable controls, update status,
        show success dialog (with Save/Discard), and handle saving or cleanup.
        """
        self._set_controls_enabled(True)
        excel_enabled = self.append_to_excel.isChecked()

        self.progress_bar.setValue(100)

        if not output_path or not os.path.exists(output_path):
            self.status_label.setText("Procesare eșuată. Verifică fișierele PDF și încearcă din nou.")
            self.status_message.setText("Procesare eșuată")
            self.show_message("Eroare de procesare", problems[0] if problems else
            "Nu s-au putut procesa fișierele PDF. Verifică validitatea fișierelor și încearcă din nou.",
                              QMessageBox.Icon.Critical)
            return

        num_new = len(df)
        if excel_enabled and self.excel_file:
            num_before = self._count_existing_records()
            total = num_before + num_new
            if num_new > 0:
                text = (f"Procesare finalizată cu succes! "
                        f"Au fost adăugate {num_new} înregistrări noi. Total: {total} înregistrări.")
            else:
                text = f"Procesare finalizată! Nicio înregistrare nouă. Total: {total} înregistrări."
        else:
            text = f"Procesare finalizată cu succes! {num_new} înregistrări."
        self.status_label.setText(text)

        dlg = QMessageBox(self)
        if problems:
            dlg.setIcon(QMessageBox.Icon.Warning)
            dlg.setWindowTitle("Procesare finalizată cu avertismente")
            dlg.setText(f"PDF-urile au fost procesate, dar {len(problems)} fișier(e) au probleme.")
            dlg.setDetailedText("\n".join(problems))
        else:
            dlg.setIcon(QMessageBox.Icon.Information)
            dlg.setWindowTitle("Procesare finalizată cu succes")
            dlg.setText("PDF-urile au fost procesate cu succes!")
        dlg.setInformativeText(f"Au fost procesate {len(self.pdf_files)} fișiere PDF.\n"
                               + (f"Au fost adăugate {num_new} înregistrări noi." if excel_enabled and self.excel_file
                                  else f"Au fost extrase {num_new} înregistrări.")
                               + ("\nApasă „Show Details…” pentru lista fișierelor cu probleme." if problems else ""))
        dlg.setStandardButtons(QMessageBox.StandardButton.Save | QMessageBox.StandardButton.Discard)
        dlg.setDefaultButton(QMessageBox.StandardButton.Save)

        for role, text in ((QMessageBox.StandardButton.Save, "Salvează"),
                           (QMessageBox.StandardButton.Discard, "Renunță"),):
            btn = dlg.button(role)
            btn.setText(text)
            btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

        if dlg.exec() == QMessageBox.StandardButton.Save:
            self._save_output(output_path)
        else:
            discard_staging(output_path)
            self.status_label.setText("Fișierul Excel a fost șters.")
            self.status_message.setText("Fișier șters")

    def _count_existing_records(self):
        """
        Read the selected Excel and return the number of valid existing rows.
        """
        try:
            col = Constants.HEADERS["nr_crt"]
            needed = {col, Constants.HEADERS["invoice_number"], Constants.HEADERS["vat_number"]}
            for df in pd.read_excel(self.excel_file, sheet_name=None, dtype=str).values():
                if needed <= set(df.columns):  # the sheet the generator writes to, whichever is active
                    return int(df[col].notna().sum())
            return 0
        except Exception as e:
            logger.error(f"Error counting existing records: {e}")
            return 0

    def _save_output(self, output_path):
        """
        Handle the Save path dialog and actual move (and auto-open).
        """
        default = os.path.join(self.last_directory, os.path.basename(output_path))
        save_path, _ = QFileDialog.getSaveFileName(self, "Salvează fișierul Excel generat", default,
                                                   "Excel Files (*.xlsx)")
        if not save_path:
            discard_staging(output_path)
            self.status_label.setText("Salvarea fișierului Excel a fost anulată.")
            self.status_message.setText("Salvare anulată")
            return
        try:
            shutil.move(output_path, save_path)
        except Exception as e:
            logger.error(f"Error saving file to {save_path}: {e}")
            self.status_label.setText(f"Salvarea la {save_path} a eșuat. Fișierul generat a rămas la: {output_path}")
            self.status_message.setText("Salvare eșuată")
            self.show_message("Eroare la salvare", f"Fișierul nu a putut fi salvat la:\n{save_path}\n\n{e}\n\n"
                                                   "Închide fișierul dacă este deschis în Excel sau alege alt nume, apoi salvează din nou.\n"
                                                   "Dacă renunți la salvare, fișierul generat se șterge.",
                              QMessageBox.Icon.Critical)
            return self._save_output(output_path)  # back to the save dialog: the workbook is still staged
        discard_staging(output_path)
        self.status_label.setText(f"Fișierul Excel a fost salvat cu succes la: {save_path}")
        self.status_message.setText("Fișier salvat cu succes")
        self.last_directory = os.path.dirname(save_path)
        save_last_directory(self.last_directory)
        if self.auto_open.isChecked():
            try:
                os.startfile(save_path)
            except OSError as e:
                logger.error(f"Error opening {save_path}: {e}")
                self.show_message("Fișierul nu a putut fi deschis", f"Fișierul a fost salvat la:\n{save_path}\n\n"
                                                                    f"dar nu a putut fi deschis automat: {e}",
                                  QMessageBox.Icon.Warning)
