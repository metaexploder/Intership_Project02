"""
PACE Payroll Validator — Payroll Extractor
============================================
Reads payroll Excel files, identifies pay-period columns using
case/space-insensitive matching, and extracts unique payroll periods.
"""

from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

from logger_setup import setup_logger

logger = setup_logger()

# Canonical (normalized) column names we search for
_TARGET_START = "payperiodstart"
_TARGET_END = "payperiodend"


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def extract_payroll_periods(
    file_paths: List[Path], woid: str
) -> Dict:
    """
    Extract unique payroll periods from all files and sheets.

    Parameters
    ----------
    file_paths : List[Path]
        Paths to the payroll Excel files for this WOID.
    woid : str
        Work Order ID (for logging context).

    Returns
    -------
    Dict with keys:
        - ``periods``: List[Tuple[date, date]] — unique (start, end) pairs
        - ``sheets_processed``: int — number of sheets that had valid columns
        - ``files_count``: int — number of files provided
    """
    all_periods: List[Tuple[date, date]] = []
    sheets_processed = 0

    for fpath in file_paths:
        try:
            xl = pd.ExcelFile(fpath, engine="openpyxl")
        except Exception as exc:
            logger.error(
                "WOID %s — Cannot open file %s: %s", woid, fpath.name, exc
            )
            continue

        for sheet_name in xl.sheet_names:
            result = _process_sheet(xl, sheet_name, fpath.name, woid)
            if result is not None:
                all_periods.extend(result)
                sheets_processed += 1

    # Deduplicate
    unique_periods = list(set(all_periods))
    unique_periods.sort(key=lambda t: t[0])

    logger.info(
        "WOID %s — Extracted %d unique payroll period(s) from %d file(s), "
        "%d sheet(s) processed.",
        woid, len(unique_periods), len(file_paths), sheets_processed,
    )

    return {
        "periods": unique_periods,
        "sheets_processed": sheets_processed,
        "files_count": len(file_paths),
    }


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _normalize(col_name: str) -> str:
    """Lowercase and strip all whitespace from a column name."""
    return "".join(col_name.lower().split())


def _find_column(df: pd.DataFrame, target: str) -> Optional[str]:
    """
    Find the original column name in *df* that matches *target*
    after normalization.
    """
    for col in df.columns:
        if _normalize(str(col)) == target:
            return col
    return None


def _process_sheet(
    xl: pd.ExcelFile,
    sheet_name: str,
    file_name: str,
    woid: str,
) -> Optional[List[Tuple[date, date]]]:
    """
    Process a single sheet.  Returns a list of (start, end) tuples,
    or None if the sheet is skipped.
    """
    try:
        df = xl.parse(sheet_name, dtype=str)
    except Exception as exc:
        logger.warning(
            "WOID %s — Could not read sheet '%s' in %s: %s",
            woid, sheet_name, file_name, exc,
        )
        return None

    if df.empty:
        logger.debug(
            "WOID %s — Sheet '%s' in %s is empty, skipping.",
            woid, sheet_name, file_name,
        )
        return None

    start_col = _find_column(df, _TARGET_START)
    end_col = _find_column(df, _TARGET_END)

    if start_col is None or end_col is None:
        logger.debug(
            "WOID %s — Sheet '%s' in %s missing pay-period columns, skipping.",
            woid, sheet_name, file_name,
        )
        return None

    # Extract the two columns
    sub = df[[start_col, end_col]].copy()
    sub.columns = ["start", "end"]

    # Drop blanks / nulls
    sub = sub.dropna(subset=["start", "end"])
    sub = sub[sub["start"].astype(str).str.strip() != ""]
    sub = sub[sub["end"].astype(str).str.strip() != ""]

    # Deduplicate
    sub = sub.drop_duplicates()

    periods: List[Tuple[date, date]] = []
    for _, row in sub.iterrows():
        s = _parse_date(row["start"])
        e = _parse_date(row["end"])
        if s is not None and e is not None:
            periods.append((s, e))
        else:
            logger.warning(
                "WOID %s — Unparseable dates in sheet '%s' of %s: "
                "start=%r, end=%r",
                woid, sheet_name, file_name, row["start"], row["end"],
            )

    logger.debug(
        "WOID %s — Sheet '%s' in %s yielded %d period(s).",
        woid, sheet_name, file_name, len(periods),
    )
    return periods


def _parse_date(value) -> Optional[date]:
    """Best-effort date parsing from various string formats."""
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    if isinstance(value, datetime):
        return value.date()

    raw = str(value).strip()
    if not raw:
        return None

    # Common formats
    for fmt in (
        "%Y-%m-%d %H:%M:%S.%f",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%m-%d-%Y",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y/%m/%d",
    ):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            continue

    # Pandas fallback
    try:
        return pd.to_datetime(raw).date()
    except Exception:
        return None
