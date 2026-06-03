"""
PACE Payroll Period Coverage Validator
=======================================
Entry point — orchestrates the full validation pipeline.

Usage:
    python main.py <input_file>

    <input_file>  Path to a .txt or .xlsx file containing WOIDs.
"""

import sys
import time
from typing import Dict, List

from config import DATASET_DIR, OUTPUT_REPORT
from db_client import fetch_file_info, fetch_policy_period, get_connection
from file_retriever import retrieve_files
from input_reader import read_woids
from logger_setup import setup_logger
from payroll_extractor import extract_payroll_periods
from report_generator import generate_report
from validator import validate_coverage

logger = setup_logger()


def main(input_file: str) -> None:
    """
    Run the full PACE payroll validation pipeline.

    Parameters
    ----------
    input_file : str
        Path to a .txt or .xlsx file containing WOIDs.
    """
    start_time = time.time()
    logger.info("=" * 70)
    logger.info("PACE Payroll Period Coverage Validator — Starting")
    logger.info("=" * 70)

    # ------------------------------------------------------------------
    # 1. Read WOIDs
    # ------------------------------------------------------------------
    woids = read_woids(input_file)
    if not woids:
        logger.error("No WOIDs found in input file. Exiting.")
        sys.exit(1)

    logger.info("Processing %d WOID(s) …", len(woids))

    # ------------------------------------------------------------------
    # 2. Connect to SQL Server
    # ------------------------------------------------------------------
    try:
        conn = get_connection()
    except Exception as exc:
        logger.error("Failed to connect to SQL Server: %s", exc)
        sys.exit(1)

    # ------------------------------------------------------------------
    # 3. Process each WOID
    # ------------------------------------------------------------------
    results: List[Dict] = []

    for idx, woid in enumerate(woids, start=1):
        logger.info("-" * 50)
        logger.info("WOID %s  [%d / %d]", woid, idx, len(woids))
        logger.info("-" * 50)

        try:
            result = _process_woid(conn, woid)
            results.append(result)
        except Exception as exc:
            logger.error(
                "WOID %s — Unexpected error, skipping: %s", woid, exc,
                exc_info=True,
            )
            results.append({
                "woid": woid,
                "status": "NO",
                "missing_gaps": [],
                "sheets_processed": 0,
                "files_count": 0,
            })

    # ------------------------------------------------------------------
    # 4. Close connection
    # ------------------------------------------------------------------
    try:
        conn.close()
        logger.info("Database connection closed.")
    except Exception:
        pass

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
    logger.info("DONE  —  Total: %d | FULL: %d | PARTIAL: %d | NO: %d",
                len(results), full, partial, no_cov)
    logger.info("Report: %s", OUTPUT_REPORT)
    logger.info("Elapsed: %.1f seconds", elapsed)
    logger.info("=" * 70)


# ---------------------------------------------------------------------------
# Per-WOID processing
# ---------------------------------------------------------------------------

def _process_woid(conn, woid: str) -> Dict:
    """
    Run the full pipeline for a single WOID and return a result dict.
    """
    # --- Policy period ---
    policy = fetch_policy_period(conn, woid)
    if policy is None:
        logger.warning("WOID %s — No policy period; marking as NO.", woid)
        return {
            "woid": woid,
            "status": "NO",
            "missing_gaps": [],
            "sheets_processed": 0,
            "files_count": 0,
        }
    policy_start, policy_end = policy

    # --- File info ---
    file_info = fetch_file_info(conn, woid)
    if not file_info:
        logger.warning("WOID %s — No file info; marking as NO.", woid)
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
        logger.warning("WOID %s — No files copied; marking as NO.", woid)
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
            "WOID %s — No payroll periods extracted; marking as NO.", woid
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
            logger.info(
                "WOID %s - Missing gap: %s -> %s", woid, g_start, g_end
            )

    logger.info("WOID %s — Validation result: %s", woid, status)

    return {
        "woid": woid,
        "status": status,
        "missing_gaps": gaps,
        "sheets_processed": extraction["sheets_processed"],
        "files_count": extraction["files_count"],
    }


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage:  python main.py <input_file>")
        print("  <input_file>  .txt or .xlsx file containing WOIDs")
        sys.exit(1)

    main(sys.argv[1])
