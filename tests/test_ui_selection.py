import os
import unittest
from tempfile import TemporaryDirectory
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt6.QtCore import QMimeData, QPoint, QPointF, Qt, QUrl
from PyQt6.QtGui import QDragEnterEvent, QDropEvent
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
        with patch.object(qt_ui.QFileDialog, "getSaveFileName",
                          side_effect=[(r"C:\\busy\\x.xlsx", ""), ("", "")]) as dialog, \
                patch.object(qt_ui.shutil, "move", side_effect=PermissionError("open in Excel")), \
                patch.object(self.window, "show_message") as shown:
            self.window._save_output(staged)
        self.assertEqual(dialog.call_count, 2)
        self.assertEqual(shown.call_count, 1)
        self.assertIn("Eroare la salvare", shown.call_args.args[0])
        self.assertFalse(os.path.exists(staging))


def _drag(cls, paths, modifiers=Qt.KeyboardModifier.NoModifier, actions=Qt.DropAction.CopyAction):
    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(p) for p in paths])
    pos = QPoint(1, 1) if cls is QDragEnterEvent else QPointF(1, 1)
    event = cls(pos, actions, mime, Qt.MouseButton.LeftButton, modifiers)
    event.mime = mime  # the event only points at the mime data; keep it alive as long as the event
    event.ignore()
    return event


class TestDragAndDrop(unittest.TestCase):
    def setUp(self):
        self.window = qt_ui.PDFToExcelApp()

    def test_collect_pdfs_takes_files_and_folders_and_skips_everything_else(self):
        with TemporaryDirectory() as td:
            os.mkdir(os.path.join(td, "sub"))
            for name in ("b.PDF", "a.txt", os.path.join("sub", "c.pdf")):
                with open(os.path.join(td, name), "wb") as f:
                    f.write(b"1")
            self.assertEqual(qt_ui.collect_pdfs([td]), [os.path.join(td, "b.PDF"), os.path.join(td, "sub", "c.pdf")])
            self.assertEqual(qt_ui.collect_pdfs([os.path.join(td, "a.txt"), os.path.join(td, "b.PDF")]),
                             [os.path.join(td, "b.PDF")])

    def test_dropped_pdfs_are_added_and_other_drags_are_refused(self):
        w = self.window
        enter = _drag(QDragEnterEvent, [r"C:\\x\\notes.txt"])
        w.dragEnterEvent(enter)
        self.assertFalse(enter.isAccepted())
        enter = _drag(QDragEnterEvent, [r"C:\\x\\1.pdf"])
        w.dragEnterEvent(enter)
        self.assertTrue(enter.isAccepted())
        with patch.object(qt_ui, "save_last_directory"):
            w.dropEvent(_drag(QDropEvent, [r"C:\\x\\1.pdf"]))
        self.assertEqual([os.path.basename(p) for p in w.pdf_files], ["1.pdf"])

    def test_empty_list_shows_a_hint_that_disappears_with_the_first_file(self):
        w = self.window
        hints = lambda: [w.file_list_layout.itemAt(i).widget() for i in range(w.file_list_layout.count())
                         if w.file_list_layout.itemAt(i).widget().objectName() == "emptyHint"]
        self.assertEqual(len(hints()), 1)
        with patch.object(qt_ui, "save_last_directory"):
            w.handle_pdf_selection([r"C:\\x\\1.pdf"])
        self.assertEqual(hints(), [])
        w.clear_file_list()
        self.assertEqual(len(hints()), 1)


class TestDragSafety(unittest.TestCase):
    def setUp(self):
        self.window = qt_ui.PDFToExcelApp()

    def test_a_shift_drag_is_answered_as_a_copy_never_a_move(self):
        """Answering "move" would make Explorer delete the dropped files: the app only stores their paths."""
        both, shift = Qt.DropAction.CopyAction | Qt.DropAction.MoveAction, Qt.KeyboardModifier.ShiftModifier
        enter = _drag(QDragEnterEvent, ["C:/x/1.pdf"], shift, both)
        self.assertEqual(enter.proposedAction(), Qt.DropAction.MoveAction)  # what Explorer proposes with Shift
        self.window.dragEnterEvent(enter)
        self.assertTrue(enter.isAccepted())
        self.assertEqual(enter.dropAction(), Qt.DropAction.CopyAction)
        drop = _drag(QDropEvent, ["C:/x/1.pdf"], shift, both)
        with patch.object(qt_ui, "save_last_directory"):
            self.window.dropEvent(drop)
        self.assertTrue(drop.isAccepted())
        self.assertEqual(drop.dropAction(), Qt.DropAction.CopyAction)
        self.assertEqual([os.path.basename(p) for p in self.window.pdf_files], ["1.pdf"])

    def test_a_drop_without_pdfs_is_refused_without_a_dialog(self):
        """The drop runs inside Explorer's drag: a modal dialog there freezes Explorer until it is closed."""
        with TemporaryDirectory() as td, patch.object(self.window, "show_message") as shown:
            drop = _drag(QDropEvent, [td], Qt.KeyboardModifier.ShiftModifier,
                         Qt.DropAction.CopyAction | Qt.DropAction.MoveAction)
            self.window.dropEvent(drop)
        self.assertFalse(drop.isAccepted())
        self.assertEqual(shown.call_count, 0)
        self.assertIn("nu conțin fișiere PDF", self.window.status_message.text())

    def test_drags_are_refused_while_processing(self):
        with patch.object(self.window, "_is_processing", return_value=True):
            enter = _drag(QDragEnterEvent, ["C:/x/1.pdf"])
            self.window.dragEnterEvent(enter)
        self.assertFalse(enter.isAccepted())

    def test_excel_tooltip_is_cleared_when_switching_back_to_a_new_file(self):
        w = self.window
        w.append_to_excel.setChecked(True)
        with patch.object(qt_ui.QFileDialog, "getOpenFileName", return_value=("C:/x/Export.xlsx", "")), \
                patch.object(qt_ui, "save_last_directory"):
            w.select_excel_file()
        self.assertEqual(w.excel_path_label.toolTip(), "C:/x/Export.xlsx")
        w.create_new_excel.setChecked(True)
        self.assertEqual(w.excel_path_label.toolTip(), "")
        self.assertIsNone(w.excel_file)


if __name__ == "__main__":
    unittest.main()
