import os
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt6.QtWidgets import QApplication

import qt_ui

app = QApplication.instance() or QApplication([])


class TestFileSelection(unittest.TestCase):
    def setUp(self):
        self.window = qt_ui.PDFToExcelApp()

    def test_selecting_again_adds_to_the_list_instead_of_replacing_it(self):
        with patch.object(qt_ui, "save_last_directory"):
            self.window.handle_pdf_selection([r"C:\a\1.pdf"])
            self.window.handle_pdf_selection([r"C:\a\2.pdf", r"C:\a\1.pdf"])
        self.assertEqual([os.path.basename(p) for p in self.window.pdf_files], ["1.pdf", "2.pdf"])
        self.assertEqual(self.window.pdf_count_label.text(), "2 fișier(e) PDF selectat(e)")
        self.assertTrue(self.window.clear_list_button.isEnabled())
        self.window.clear_file_list()
        self.assertEqual(self.window.pdf_files, [])
        self.assertFalse(self.window.clear_list_button.isEnabled())
        self.assertFalse(self.window.process_button.isEnabled())
        self.assertEqual(self.window.pdf_count_label.text(), "Niciun fișier selectat")

    def test_every_control_is_locked_while_processing(self):
        w = self.window
        w._set_controls_enabled(False)
        for widget in (w.percentage_input, w.create_new_excel, w.append_to_excel, w.auto_open, w.file_list_widget,
                       w.clear_list_button, w.select_pdf_button, w.select_folder_button, w.select_excel_button,
                       w.process_button):
            self.assertFalse(widget.isEnabled(), widget)
        w._set_controls_enabled(True)
        self.assertTrue(w.percentage_input.isEnabled())
        self.assertFalse(w.select_excel_button.isEnabled())  # still follows the "append" radio, which is off
        self.assertFalse(w.process_button.isEnabled())  # nothing selected, so still nothing to process

    def test_failed_save_offers_the_save_dialog_again_and_cancel_discards_the_staged_workbook(self):
        """The usual failure is the target being open in Excel: the user must get a second chance, not a %TEMP% path."""
        from tempfile import mkdtemp
        staging = mkdtemp(prefix="pdf_to_excel_")
        staged = os.path.join(staging, "x.xlsx")
        with open(staged, "wb") as f:
            f.write(b"1")
        with patch.object(qt_ui.QFileDialog, "getSaveFileName", side_effect=[(r"C:\\busy\\x.xlsx", ""), ("", "")]) as dialog, \
                patch.object(qt_ui.shutil, "move", side_effect=PermissionError("open in Excel")), \
                patch.object(self.window, "show_message") as shown:
            self.window._save_output(staged)
        self.assertEqual(dialog.call_count, 2)
        self.assertEqual(shown.call_count, 1)
        self.assertIn("Eroare la salvare", shown.call_args.args[0])
        self.assertFalse(os.path.exists(staging))


if __name__ == "__main__":
    unittest.main()
