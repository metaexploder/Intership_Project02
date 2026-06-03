"""
PACE Payroll Validator — Payroll Extractor
============================================
Reads payroll Excel files, identifies pay-period columns using
case/space-insensitive matching, and extracts unique payroll periods.

Optimized: uses openpyxl read_only mode for faster reads and only
loads the columns we need.
"""

from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl

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

    Returns Dict with keys: periods, sheets_processed, files_count.
    """
    all_periods: List[Tuple[date, date]] = []
    sheets_processed = 0

    for fpath in file_paths:
        try:
            wb = openpyxl.load_workbook(str(fpath), read_only=True, data_only=True)
        except Exception as exc:
            logger.error(
                "WOID %s - Cannot open file %s: %s", woid, fpath.name, exc
            )
            continue

        try:
            for sheet_name in wb.sheetnames:
                result = _process_sheet_fast(wb[sheet_name], sheet_name, fpath.name, woid)
                if result is not None:
                    all_periods.extend(result)
                    sheets_processed += 1
        finally:
            wb.close()

    # Deduplicate
    unique_periods = list(set(all_periods))
    unique_periods.sort(key=lambda t: t[0])

    logger.info(
        "WOID %s - Extracted %d unique payroll period(s) from %d file(s), "
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
    return "".join(str(col_name).lower().split())


def _process_sheet_fast(
    ws,
    sheet_name: str,
    file_name: str,
    woid: str,
) -> Optional[List[Tuple[date, date]]]:
    """
    Process a single sheet using openpyxl read_only worksheet.
    Much faster than pandas for large files since we only need 2 columns.
    """
    rows_iter = ws.iter_rows()

    # Read header row
    try:
        header_row = next(rows_iter)
    except StopIteration:
        logger.debug("WOID %s - Sheet '%s' in %s is empty, skipping.",
                      woid, sheet_name, file_name)
        return None

    # Find column indices for pay period start/end
    start_idx = None
    end_idx = None
    for idx, cell in enumerate(header_row):
        val = cell.value
        if val is None:
            continue
        normalized = _normalize(str(val))
        if normalized == _TARGET_START:
            start_idx = idx
        elif normalized == _TARGET_END:
            end_idx = idx

    if start_idx is None or end_idx is None:
        logger.debug(
            "WOID %s - Sheet '%s' in %s missing pay-period columns, skipping.",
            woid, sheet_name, file_name,
        )
        return None

    # Read only the two columns we need
    seen = set()
    periods: List[Tuple[date, date]] = []

    for row in rows_iter:
        try:
            raw_start = row[start_idx].value
            raw_end = row[end_idx].value
        except IndexError:
            continue

        # Skip blanks/nulls
        if raw_start is None or raw_end is None:
            continue

        start_str = str(raw_start).strip()
        end_str = str(raw_end).strip()
        if not start_str or not end_str:
            continue

        # Dedup key
        key = (start_str, end_str)
        if key in seen:
            continue
        seen.add(key)

        s = _parse_date(raw_start)
        e = _parse_date(raw_end)
        if s is not None and e is not None:
            periods.append((s, e))
        else:
            logger.warning(
                "WOID %s - Unparseable dates in sheet '%s' of %s: "
                "start=%r, end=%r",
                woid, sheet_name, file_name, raw_start, raw_end,
            )

    if periods:
        logger.debug(
            "WOID %s - Sheet '%s' in %s yielded %d period(s).",
            woid, sheet_name, file_name, len(periods),
        )

    return periods if periods else None


def _parse_date(value) -> Optional[date]:
    """Best-effort date parsing."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    raw = str(value).strip()
    if not raw:
        return None

    # Common formats — try most likely first
    for fmt in (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%Y-%m-%d %H:%M:%S.%f",
        "%m-%d-%Y",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y/%m/%d",
    ):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            continue

    return None
