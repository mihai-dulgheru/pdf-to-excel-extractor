import logging
import os
import sys
import tempfile
import unittest
from unittest.mock import patch
from logging.handlers import RotatingFileHandler
from pathlib import Path

from config import Constants
from functions import get_logger, setup_logger


class TestPathsAreAnchoredToTheApp(unittest.TestCase):
    def test_assets_config_and_data_dirs_do_not_depend_on_the_working_directory(self):
        self.assertTrue(Path(Constants.BASE_DIR).is_absolute())
        self.assertTrue(Path(Constants.ASSETS_DIR).is_absolute())
        self.assertTrue(Path(Constants.DATA_DIR).is_absolute())
        self.assertTrue((Constants.ASSETS_DIR / "icon.png").exists())
        self.assertTrue((Constants.ASSETS_DIR / "icons" / "trash-can.svg").exists())
        self.assertTrue(os.path.isabs(Constants.CONFIG_FILE))
        self.assertEqual(Path(Constants.CONFIG_FILE).parent, Path(Constants.DATA_DIR))

    def test_column_formats_for_ron_and_statistic_are_whole_numbers(self):
        """The customer's template and every file he edited format K and Q as '#,##0'."""
        self.assertEqual(Constants.COLUMN_FORMATS["value_ron"], "#,##0")
        self.assertEqual(Constants.COLUMN_FORMATS["statistic"], "#,##0")
        self.assertEqual(Constants.COLUMN_FORMATS["invoice_value_eur"], "#,##0.00")


class TestLogging(unittest.TestCase):
    def test_child_loggers_get_no_handlers_of_their_own(self):
        """Module loggers only propagate to the one configured in main(): no extra files, no duplicate lines."""
        child = get_logger("pdf_to_excel.test_child_logger")
        self.assertEqual(child.handlers, [])
        self.assertTrue(child.propagate)

    def test_file_logging_rotates_a_single_fixed_file(self):
        with tempfile.TemporaryDirectory() as td:
            name = "pdf_to_excel_test_rotation"
            logger = setup_logger(name=name, log_level=logging.INFO, log_to_file=True, logs_dir=td)
            try:
                rotating = [h for h in logger.handlers if isinstance(h, RotatingFileHandler)]
                self.assertEqual(len(rotating), 1)
                self.assertEqual(Path(rotating[0].baseFilename).name, f"{name}.log")
                self.assertGreater(rotating[0].backupCount, 0)
            finally:
                for h in logger.handlers[:]:
                    h.close()
                    logger.removeHandler(h)

    def test_unwritable_logs_folder_falls_back_to_the_temp_folder(self):
        """The exe may sit in a read-only folder (Program Files, a share): the log then goes to the temp folder."""
        with tempfile.NamedTemporaryFile(delete=False) as blocker:
            pass
        logger = setup_logger(name="pdf_to_excel_test_unwritable", log_to_file=True,
                              logs_dir=os.path.join(blocker.name, "logs"))
        try:
            rotating = [h for h in logger.handlers if isinstance(h, RotatingFileHandler)]
            self.assertEqual(len(rotating), 1)
            self.assertEqual(Path(rotating[0].baseFilename).parent,
                             Path(tempfile.gettempdir()) / "pdf_to_excel" / "logs")
        finally:
            for h in logger.handlers[:]:
                h.close()
                logger.removeHandler(h)
            for h in rotating:
                os.remove(h.baseFilename)
            os.remove(blocker.name)

    def test_no_console_handler_when_there_is_no_console(self):
        """The windowed exe has sys.stderr = None; a StreamHandler on it would only swallow every record."""
        with patch.object(sys, "stderr", None):
            logger = setup_logger(name="pdf_to_excel_test_no_console", log_to_file=False)
        self.assertEqual(logger.handlers, [])

    def test_rollover_blocked_by_another_instance_keeps_logging(self):
        """Two instances share one log file; on Windows the rename of a rollover then fails. No record may be lost."""
        with tempfile.TemporaryDirectory() as td:
            name = "pdf_to_excel_test_blocked_rollover"
            logger = setup_logger(name=name, log_to_file=True, logs_dir=td)
            try:
                handler = next(h for h in logger.handlers if isinstance(h, RotatingFileHandler))
                handler.maxBytes = 200
                with open(os.path.join(td, f"{name}.log"), "a"):  # the other instance holding the file
                    for _ in range(20):
                        logger.warning("x" * 50)
                written = sum(open(os.path.join(td, f), encoding="utf-8").read().count("x" * 50)
                              for f in os.listdir(td))
                self.assertEqual(written, 20)
            finally:
                for h in logger.handlers[:]:
                    h.close()
                    logger.removeHandler(h)


if __name__ == "__main__":
    unittest.main()
