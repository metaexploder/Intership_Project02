"""
PACE Payroll Period Coverage Validator
=======================================
Entry point — orchestrates the full validation pipeline.

Optimized with:
    - Batch SQL queries (2 round-trips instead of 2*N)
    - Parallel file retrieval & Excel processing via ThreadPoolExecutor
"""

import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, List, Optional, Tuple
from datetime import date

from config import DATASET_DIR, OUTPUT_REPORT
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
    Run the full PACE payroll validation pipeline.
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
    # 3. Batch-fetch ALL policy periods & file info (2 SQL calls total)
    # ------------------------------------------------------------------
    logger.info("Fetching all policy periods (batch) ...")
    all_policies = fetch_all_policy_periods(conn, woids)

    logger.info("Fetching all file info (batch) ...")
    all_file_info = fetch_all_file_info(conn, woids)

    # Close DB connection early — no longer needed
    try:
        conn.close()
        logger.info("Database connection closed.")
    except Exception:
        pass

    # ------------------------------------------------------------------
    # 4. Process WOIDs in parallel (file retrieval + extraction + validation)
    # ------------------------------------------------------------------
    results: List[Dict] = []
    workers = min(MAX_WORKERS, len(woids))

    logger.info("Processing WOIDs with %d parallel worker(s) ...", workers)

    with ThreadPoolExecutor(max_workers=workers) as executor:
        future_to_woid = {}
        for woid in woids:
            policy = all_policies.get(woid)
            file_info = all_file_info.get(woid, [])
            future = executor.submit(
                _process_woid, woid, policy, file_info
            )
            future_to_woid[future] = woid

        for future in as_completed(future_to_woid):
            woid = future_to_woid[future]
            try:
                result = future.result()
                results.append(result)
            except Exception as exc:
                logger.error(
                    "WOID %s - Unexpected error, skipping: %s", woid, exc,
                    exc_info=True,
                )
                results.append({
                    "woid": woid,
                    "status": "NO",
                    "missing_gaps": [],
                    "sheets_processed": 0,
                    "files_count": 0,
                })

    # Preserve original WOID order for the report
    woid_order = {w: i for i, w in enumerate(woids)}
    results.sort(key=lambda r: woid_order.get(r["woid"], 0))

    # ------------------------------------------------------------------
    # 5. Generate report
    # ------------------------------------------------------------------
    generate_report(results, OUTPUT_REPORT)

    # ------------------------------------------------------------------
    # 6. Summary
    # ------------------------------------------------------------------
    full = sum(1 for r in results if r["status"] == "FULL")
    partial = sum(1 for r in results if r["status"] == "PARTIAL")
    no_cov = sum(1 for r in results if r["status"] == "NO")
    elapsed = time.time() - start_time

    logger.info("=" * 70)
    logger.info("DONE  -  Total: %d | FULL: %d | PARTIAL: %d | NO: %d",
                len(results), full, partial, no_cov)
    logger.info("Report: %s", OUTPUT_REPORT)
    logger.info("Elapsed: %.1f seconds", elapsed)
    logger.info("=" * 70)


# ---------------------------------------------------------------------------
# Per-WOID processing (runs in thread pool)
# ---------------------------------------------------------------------------

def _process_woid(
    woid: str,
    policy: Optional[Tuple[date, date]],
    file_info: List[Dict[str, str]],
) -> Dict:
    """
    Run file retrieval, extraction, and validation for a single WOID.
    Policy and file_info are pre-fetched from batch queries.
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
        }

    # --- Extract payroll periods ---
    extraction = extract_payroll_periods(copied_files, woid)
    payroll_periods = extraction["periods"]

    if not payroll_periods:
        logger.warning(
            "WOID %s - No payroll periods extracted; marking as NO.", woid
        )
        return {
            "woid": woid,
            "status": "NO",
            "missing_gaps": [(policy_start, policy_end)],
            "sheets_processed": extraction["sheets_processed"],
            "files_count": extraction["files_count"],
        }

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
    }


# ---------------------------------------------------------------------------
# Configuration — Set your input file path here
# ---------------------------------------------------------------------------
INPUT_FILE = r"woids.txt"  # Change this to your WOID file path (.txt or .xlsx)


if __name__ == "__main__":
    main(INPUT_FILE)
