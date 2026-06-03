"""
PACE Payroll Validator — Logging Setup
=======================================
Configures application-wide logging to both console and a log file.
"""

import logging
import sys
from config import LOG_FILE


def setup_logger(name: str = "pace_validator") -> logging.Logger:
    """
    Create and return a configured logger.

    Outputs:
        - Console  : INFO and above
        - Log file : DEBUG and above (full detail)
    """
    logger = logging.getLogger(name)

    # Prevent duplicate handlers if called multiple times
    if logger.handlers:
        return logger

    logger.setLevel(logging.DEBUG)

    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # --- Console handler ---
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # --- File handler ---
    file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    return logger
