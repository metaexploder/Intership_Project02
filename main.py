"""
PACE Payroll Period Coverage Validator (Optimized)
====================================================
Entry point — orchestrates the full validation pipeline.

Performance optimizations:
    - Batch SQL queries (2 total instead of 2 x N)
    - ThreadPoolExecutor for parallel file retrieval + Excel parsing
    - openpyxl read_only mode for fast Excel reads
"""

import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, List, Optional, Tuple
from datetime import date, timedelta

from config import BUFFER_DAYS, DATASET_DIR, OUTPUT_REPORT
from db_client import fetch_all_file_info, fetch_all_policy_periods, get_connection
from file_retriever import retrieve_files
from input_reader import read_woids
from logger_setup import setup_logger
from payroll_extractor import extract_payroll_periods
from report_generator import generate_report
from validator import validate_coverage

logger = setup_logger()

# Max parallel workers for file I/O + Excel parsing
MAX_WORKERS = 8


def main(input_file: str) -> None:
    """
    Run the full PACE payroll validation pipeline (optimized).
    """
    start_time = time.time()
    logger.info("=" * 70)
    logger.info("PACE Payroll Period Coverage Validator - Starting")
    logger.info("=" * 70)

    # ------------------------------------------------------------------
    # 1. Read WOIDs
    # ------------------------------------------------------------------
    woids = read_woids(input_file)
    if not woids:
        logger.error("No WOIDs found in input file. Exiting.")
        sys.exit(1)

    logger.info("Processing %d WOID(s) ...", len(woids))

    # ------------------------------------------------------------------
    # 2. Connect to SQL Server
    # ------------------------------------------------------------------
    try:
        conn = get_connection()
    except Exception as exc:
        logger.error("Failed to connect to SQL Server: %s", exc)
        sys.exit(1)

    # ------------------------------------------------------------------
    # 3. BATCH fetch all data in 2 SQL round-trips
    # ------------------------------------------------------------------
    t1 = time.time()
    all_policies = fetch_all_policy_periods(conn, woids)
    all_file_info = fetch_all_file_info(conn, woids)
    logger.info("SQL queries completed in %.1f seconds.", time.time() - t1)

    # Close connection early — no longer needed
    try:
        conn.close()
        logger.info("Database connection closed.")
    except Exception:
        pass

    # ------------------------------------------------------------------
    # 4. Process WOIDs in parallel (file I/O + Excel parsing)
    # ------------------------------------------------------------------
    t2 = time.time()
    results: List[Dict] = []

    with ThreadPoolExecutor(max_workers=min(MAX_WORKERS, len(woids))) as executor:
        future_to_woid = {}
        for woid in woids:
            policy = all_policies.get(woid)
            file_info = all_file_info.get(woid, [])
            future = executor.submit(_process_woid, woid, policy, file_info)
            future_to_woid[future] = woid

        for future in as_completed(future_to_woid):
            woid = future_to_woid[future]
            try:
                result = future.result()
                results.append(result)
            except Exception as exc:
                logger.error("WOID %s - Unexpected error: %s", woid, exc)
                results.append({
                    "woid": woid,
                    "status": "NO",
                    "missing_gaps": [],
                    "sheets_processed": 0,
                    "files_count": 0,
                    "policy_start": None,
                    "policy_end": None,
                    "period_start": None,
                    "period_end": None,
                    "start_buffer": None,
                    "end_buffer": None,
                })

    # Sort results back to input order
    woid_order = {w: i for i, w in enumerate(woids)}
    results.sort(key=lambda r: woid_order.get(r["woid"], 999999))

    logger.info("File processing completed in %.1f seconds.", time.time() - t2)

    # ------------------------------------------------------------------
    # 5. Generate report
    # ------------------------------------------------------------------
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    report_filename = f"PACE_PayPeriod_Validation_Report_{timestamp}.xlsx"
    output_report_path = OUTPUT_REPORT.parent / report_filename
    generate_report(results, output_report_path)

    # ------------------------------------------------------------------
    # 6. Summary
    # ------------------------------------------------------------------
    full = sum(1 for r in results if r["status"] == "FULL")
    partial = sum(1 for r in results if r["status"] == "PARTIAL")
    no_cov = sum(1 for r in results if r["status"] == "NO")
    elapsed = time.time() - start_time

    logger.info("=" * 70)
    logger.info("DONE - Total: %d | FULL: %d | PARTIAL: %d | NO: %d",
                len(results), full, partial, no_cov)
    logger.info("Report: %s", output_report_path)
    logger.info("Elapsed: %.1f seconds", elapsed)
    logger.info("=" * 70)


# ---------------------------------------------------------------------------
# Per-WOID processing (runs in thread pool)
# ---------------------------------------------------------------------------

def _process_woid(
    woid: str,
    policy,
    file_info: List[Dict[str, str]],
) -> Dict:
    """
    Process a single WOID: retrieve files, extract periods, validate.
    This function is thread-safe (no shared mutable state).
    """
    logger.info("WOID %s - Processing ...", woid)

    # --- Policy period ---
    if policy is None:
        logger.warning("WOID %s - No policy period; marking as NO.", woid)
        return {
            "woid": woid,
            "status": "NO",
            "missing_gaps": [],
            "sheets_processed": 0,
            "files_count": 0,
            "policy_start": None,
            "policy_end": None,
            "period_start": None,
            "period_end": None,
            "start_buffer": None,
            "end_buffer": None,
        }
    policy_start, policy_end = policy

    # --- File info ---
    if not file_info:
        logger.warning("WOID %s - No file info; marking as NO.", woid)
        return {
            "woid": woid,
            "status": "NO",
            "missing_gaps": [(policy_start, policy_end)],
            "sheets_processed": 0,
            "files_count": 0,
            "policy_start": policy_start,
            "policy_end": policy_end,
            "period_start": None,
            "period_end": None,
            "start_buffer": None,
            "end_buffer": None,
        }

    # --- Retrieve files ---
    copied_files = retrieve_files(woid, file_info, DATASET_DIR)
    if not copied_files:
        logger.warning("WOID %s - No files copied; marking as NO.", woid)
        return {
            "woid": woid,
            "status": "NO",
            "missing_gaps": [(policy_start, policy_end)],
            "sheets_processed": 0,
            "files_count": len(file_info),
            "policy_start": policy_start,
            "policy_end": policy_end,
            "period_start": None,
            "period_end": None,
            "start_buffer": None,
            "end_buffer": None,
        }

    # --- Extract payroll periods ---
    extraction = extract_payroll_periods(copied_files, woid)
    payroll_periods = extraction["periods"]

    if not payroll_periods:
        logger.warning("WOID %s - No payroll periods extracted; marking as NO.", woid)
        return {
            "woid": woid,
            "status": "NO",
            "missing_gaps": [(policy_start, policy_end)],
            "sheets_processed": extraction["sheets_processed"],
            "files_count": extraction["files_count"],
            "policy_start": policy_start,
            "policy_end": policy_end,
            "period_start": None,
            "period_end": None,
            "start_buffer": None,
            "end_buffer": None,
        }

    # --- Compute PeriodStart / PeriodEnd from OVERLAPPING periods only ---
    # Only periods whose date range intersects [policy_start, policy_end] are relevant.
    # Periods from other policy years are excluded so they don't skew buffer values.
    overlapping = [
        (s, e) for s, e in payroll_periods
        if e >= policy_start and s <= policy_end
    ]

    if overlapping:
        period_start: date = min(s for s, _ in overlapping)
        period_end:   date = max(e for _, e in overlapping)
        start_buffer: int = (policy_start - period_start).days
        end_buffer:   int = (period_end   - policy_end).days
    else:
        # No overlapping periods — validator will return NO
        period_start = None
        period_end   = None
        start_buffer = None
        end_buffer   = None

    logger.info(
        "WOID %s - PeriodStart=%s PeriodEnd=%s StartBuffer=%s EndBuffer=%s",
        woid, period_start, period_end,
        f"{start_buffer:+d}" if start_buffer is not None else "N/A",
        f"{end_buffer:+d}"   if end_buffer   is not None else "N/A",
    )

    # --- Validate coverage ---
    status, gaps = validate_coverage(policy_start, policy_end, payroll_periods)

    if gaps:
        for g_start, g_end in gaps:
            logger.info("WOID %s - Missing gap: %s -> %s", woid, g_start, g_end)

    logger.info("WOID %s - Validation result: %s", woid, status)

    return {
        "woid": woid,
        "status": status,
        "missing_gaps": gaps,
        "sheets_processed": extraction["sheets_processed"],
        "files_count": extraction["files_count"],
        "policy_start": policy_start,
        "policy_end": policy_end,
        "period_start": period_start,
        "period_end": period_end,
        "start_buffer": start_buffer,
        "end_buffer": end_buffer,
    }


# ---------------------------------------------------------------------------
# Configuration — Set your input file path here
# ---------------------------------------------------------------------------
INPUT_FILE = r"woids.txt"  # Change this to your WOID file path (.txt or .xlsx)


if __name__ == "__main__":
    main(INPUT_FILE)
