import logging
from datetime import datetime
from pathlib import Path


def setup_logger(name="pdf_to_excel", log_level=logging.INFO, log_to_file=True):
    """
    Set up and configure a logger with the specified name and log level.
    
    Args:
        name (str): The name of the logger
        log_level (int): The logging level (e.g., logging.DEBUG, logging.INFO)
        log_to_file (bool): Whether to log to a file in addition to the console
        
    Returns:
        logging.Logger: The configured logger
    """
    logger = logging.getLogger(name)
    logger.setLevel(log_level)

    for handler in logger.handlers[:]:
        logger.removeHandler(handler)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)

    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    console_handler.setFormatter(formatter)

    logger.addHandler(console_handler)

    if log_to_file:
        logs_dir = Path("logs")
        logs_dir.mkdir(exist_ok=True)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = logs_dir / f"{name}_{timestamp}.log"

        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(log_level)
        file_handler.setFormatter(formatter)

        logger.addHandler(file_handler)

    return logger


def get_logger(name="pdf_to_excel"):
    """
    Get an existing logger or create a new one if it doesn't exist.
    
    Args:
        name (str): The name of the logger
        
    Returns:
        logging.Logger: The requested logger
    """
    logger = logging.getLogger(name)

    if not logger.handlers:
        logger = setup_logger(name)

    return logger


def set_log_level(level):
    """
    Set the log level for all handlers of the pdf_to_excel logger.
    
    Args:
        level (int): The logging level (e.g., logging.DEBUG, logging.INFO)
    """
    logger = get_logger()
    logger.setLevel(level)

    for handler in logger.handlers:
        handler.setLevel(level)
