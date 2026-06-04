"""
PACE Payroll Validator — Report Generator
============================================
Produces the final Excel validation report using openpyxl.
"""

from pathlib import Path
from typing import Dict, List

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from logger_setup import setup_logger

logger = setup_logger()

# ---------------------------------------------------------------------------
# Colour palette
# ---------------------------------------------------------------------------
_HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
_HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
_TITLE_FONT = Font(name="Calibri", bold=True, size=14, color="1F4E79")
_METRIC_LABEL_FONT = Font(name="Calibri", bold=True, size=11)
_METRIC_VALUE_FONT = Font(name="Calibri", size=11)
_STATUS_FONTS = {
    "FULL": Font(name="Calibri", bold=True, color="217346"),      # green
    "PARTIAL": Font(name="Calibri", bold=True, color="BF8F00"),   # amber
    "NO": Font(name="Calibri", bold=True, color="C00000"),        # red
}
_STATUS_FILLS = {
    "FULL": PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),
    "PARTIAL": PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"),
    "NO": PatternFill(start_color="FCE4EC", end_color="FCE4EC", fill_type="solid"),
}
_THIN_BORDER = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
_CENTER = Alignment(horizontal="center", vertical="center")
_LEFT = Alignment(horizontal="left", vertical="center")


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def generate_report(results: List[Dict], output_path: Path) -> None:
    """
    Generate ``PACE_PayPeriod_Validation_Report.xlsx``.

    Parameters
    ----------
    results : List[Dict]
        Each dict contains:
            - woid: str
            - status: str ("FULL" / "PARTIAL" / "NO")
            - missing_gaps: List[Tuple[date, date]]
            - sheets_processed: int
            - files_count: int
    output_path : Path
        Destination file path for the report.
    """
    wb = Workbook()
    _build_summary_sheet(wb, results)
    wb.save(str(output_path))
    logger.info("Report saved -> %s", output_path)


# ---------------------------------------------------------------------------
# Sheet builders
# ---------------------------------------------------------------------------

def _build_summary_sheet(wb: Workbook, results: List[Dict]) -> None:
    """Create Sheet 1 — Summary."""
    ws = wb.active
    ws.title = "Summary"

    # Freeze column widths
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14

    # --- Counts ---
    total = len(results)
    full_count = sum(1 for r in results if r["status"] == "FULL")
    partial_count = sum(1 for r in results if r["status"] == "PARTIAL")
    no_count = sum(1 for r in results if r["status"] == "NO")

    # Title row
    ws.merge_cells("A1:D1")
    cell = ws["A1"]
    cell.value = "PayPeriod Validation Summary"
    cell.font = _TITLE_FONT
    cell.alignment = Alignment(horizontal="left", vertical="center")

    # Metrics block  (rows 3-6)
    metrics = [
        ("Total WOIDs Processed:", total),
        ("FULL:", full_count),
        ("PARTIAL:", partial_count),
        ("NO:", no_count),
    ]
    for idx, (label, value) in enumerate(metrics, start=3):
        lbl_cell = ws.cell(row=idx, column=1, value=label)
        lbl_cell.font = _METRIC_LABEL_FONT
        val_cell = ws.cell(row=idx, column=2, value=value)
        val_cell.font = _METRIC_VALUE_FONT
        val_cell.alignment = _CENTER

    # --- Data Table ---
    header_row = 8
    headers = ["WOID", "Coverage", "Sheets", "Files"]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=header)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = _CENTER
        cell.border = _THIN_BORDER

    # Data rows
    for row_offset, res in enumerate(results, start=1):
        row_num = header_row + row_offset
        status = res["status"]

        woid_cell = ws.cell(row=row_num, column=1, value=res["woid"])
        woid_cell.alignment = _CENTER
        woid_cell.border = _THIN_BORDER

        status_cell = ws.cell(row=row_num, column=2, value=status)
        status_cell.font = _STATUS_FONTS.get(status, Font())
        status_cell.fill = _STATUS_FILLS.get(status, PatternFill())
        status_cell.alignment = _CENTER
        status_cell.border = _THIN_BORDER

        sheets_label = "Multiple" if res.get("sheets_processed", 1) > 1 else "Single"
        sheets_cell = ws.cell(row=row_num, column=3, value=sheets_label)
        sheets_cell.alignment = _CENTER
        sheets_cell.border = _THIN_BORDER

        files_label = "Multiple" if res.get("files_count", 1) > 1 else "Single"
        files_cell = ws.cell(row=row_num, column=4, value=files_label)
        files_cell.alignment = _CENTER
        files_cell.border = _THIN_BORDER

