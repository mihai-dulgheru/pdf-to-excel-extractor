import logging
import sys
import tempfile
from logging.handlers import RotatingFileHandler
from pathlib import Path

from config import Constants


class _RotatingFileHandler(RotatingFileHandler):
    """
    Keeps appending when the rollover rename fails (Windows: a second instance of the app holds the same log file)
    instead of dropping every record from then on, which is what the stock handler does.
    """

    def doRollover(self):
        try:
            super().doRollover()
        except OSError:
            if self.stream is None:
                self.stream = self._open()


def setup_logger(name="pdf_to_excel", log_level=logging.INFO, log_to_file=True, logs_dir=None):
    """
    Configure the application logger: the console (when there is one) plus one rotating file,
    <logs_dir>/<name>.log (1 MB, 5 backups). Module loggers named "<name>.<module>" propagate to it and need no
    handlers of their own.

    Args:
        name (str): The name of the logger
        log_level (int): The logging level (e.g., logging.DEBUG, logging.INFO)
        log_to_file (bool): Whether to log to a file in addition to the console
        logs_dir (str | Path | None): Folder of the log file; defaults to <data dir>/logs, and falls back to the
            temp folder when that one is not writable (the exe may sit in Program Files or on a read-only share)

    Returns:
        logging.Logger: The configured logger
    """
    logger = logging.getLogger(name)
    logger.setLevel(log_level)

    for handler in logger.handlers[:]:
        handler.close()
        logger.removeHandler(handler)

    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    if sys.stderr is not None:  # the windowed exe has no console
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)

    if log_to_file:
        folders = [Path(logs_dir) if logs_dir else Constants.DATA_DIR / "logs",
                   Path(tempfile.gettempdir()) / "pdf_to_excel" / "logs"]
        failed = []
        for folder in folders:
            try:
                folder.mkdir(parents=True, exist_ok=True)
                file_handler = _RotatingFileHandler(folder / f"{name}.log", maxBytes=1_000_000, backupCount=5,
                                                    encoding="utf-8")
            except OSError as e:
                failed.append(f"{folder}: {e}")
                continue
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
            break
        for message in failed:
            logger.warning(f"Log file could not be opened in {message}")

    return logger


def get_logger(name="pdf_to_excel"):
    """
    Return the named logger. Its records propagate to the logger configured by setup_logger().
    """
    return logging.getLogger(name)


def set_log_level(level):
    """
    Set the log level of the pdf_to_excel logger and its handlers.
    """
    logger = get_logger()
    logger.setLevel(level)
    for handler in logger.handlers:
        handler.setLevel(level)
