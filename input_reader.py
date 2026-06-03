"""
PACE Payroll Validator — Input Reader
======================================
Reads Work Order IDs (WOIDs) from a TXT or Excel file.
"""

from pathlib import Path
from typing import List

import pandas as pd

from logger_setup import setup_logger

logger = setup_logger()


def read_woids(file_path: str) -> List[str]:
    """
    Read WOIDs from a .txt or .xlsx file.

    Parameters
    ----------
    file_path : str
        Path to the input file containing WOIDs.

    Returns
    -------
    List[str]
        Deduplicated list of WOID strings.

    Raises
    ------
    FileNotFoundError
        If the input file does not exist.
    ValueError
        If the file extension is not supported.
    """
    path = Path(file_path)

    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {file_path}")

    ext = path.suffix.lower()

    if ext == ".txt":
        woids = _read_txt(path)
    elif ext in (".xlsx", ".xls"):
        woids = _read_excel(path)
    else:
        raise ValueError(
            f"Unsupported input file format '{ext}'. Use .txt or .xlsx"
        )

    # Deduplicate while preserving order
    seen = set()
    unique_woids: List[str] = []
    for woid in woids:
        if woid not in seen:
            seen.add(woid)
            unique_woids.append(woid)

    logger.info("Loaded %d unique WOID(s) from %s", len(unique_woids), path.name)
    return unique_woids


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------

def _read_txt(path: Path) -> List[str]:
    """Parse WOIDs from a plain-text file (one per line)."""
    woids: List[str] = []
    with open(path, "r", encoding="utf-8") as fh:
        for line in fh:
            stripped = line.strip()
            if stripped:
                woids.append(stripped)
    return woids


def _read_excel(path: Path) -> List[str]:
    """Parse WOIDs from the first column of an Excel file."""
    df = pd.read_excel(path, engine="openpyxl", dtype=str)
    if df.empty:
        return []

    # Use the first column regardless of its header name
    first_col = df.iloc[:, 0]
    woids = first_col.dropna().astype(str).str.strip().tolist()
    return [w for w in woids if w]
