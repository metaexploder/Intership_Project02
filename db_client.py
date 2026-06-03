"""
PACE Payroll Validator — Database Client
==========================================
Handles SQL Server connectivity and query execution via pyodbc.
Supports both single-WOID and batch queries for performance.
"""

from datetime import date, datetime
from typing import Dict, List, Optional, Tuple

import pyodbc

from config import QUERY_FILE_INFO, QUERY_POLICY_PERIOD, get_connection_string
from logger_setup import setup_logger

logger = setup_logger()


def get_connection() -> pyodbc.Connection:
    """
    Establish and return a pyodbc connection to SQL Server.
    """
    conn_str = get_connection_string()
    logger.debug("Connecting to SQL Server ...")
    conn = pyodbc.connect(conn_str, timeout=30)
    logger.info("SQL Server connection established.")
    return conn


# ---------------------------------------------------------------------------
# Batch queries — fetch all WOIDs in one round-trip
# ---------------------------------------------------------------------------

def fetch_all_policy_periods(
    conn: pyodbc.Connection, woids: List[str]
) -> Dict[str, Tuple[date, date]]:
    """
    Fetch policy periods for ALL WOIDs in a single query (or batched chunks).

    Returns
    -------
    Dict[str, Tuple[date, date]]
        Mapping of woid -> (policy_start, policy_end).
        Missing WOIDs are not included.
    """
    result: Dict[str, Tuple[date, date]] = {}

    # Build base query by stripping the WHERE clause and rebuilding with IN
    base_query = QUERY_POLICY_PERIOD.strip()

    for chunk in _chunked(woids, 500):
        placeholders = ",".join(["?" for _ in chunk])
        # Replace "WHERE woid = ?" or "WHERE WOID = ?" with IN clause
        # We rebuild the query to select woid as well
        batch_query = _build_batch_policy_query(placeholders)

        try:
            cursor = conn.cursor()
            cursor.execute(batch_query, chunk)
            rows = cursor.fetchall()
            cursor.close()

            for row in rows:
                woid = str(row[0]).strip()
                inception = _to_date(row[1])
                expiration = _to_date(row[2])
                if inception and expiration:
                    result[woid] = (inception, expiration)
                    logger.info("WOID %s - Policy period: %s -> %s", woid, inception, expiration)
                else:
                    logger.error("WOID %s - Could not parse policy dates: %s, %s", woid, row[1], row[2])

        except pyodbc.Error as exc:
            logger.error("Batch policy query error: %s", exc)
            # Fallback: query one by one
            for woid in chunk:
                try:
                    single = fetch_policy_period(conn, woid)
                    if single:
                        result[woid] = single
                except Exception:
                    pass

    logger.info("Fetched policy periods for %d / %d WOIDs.", len(result), len(woids))
    return result


def fetch_all_file_info(
    conn: pyodbc.Connection, woids: List[str]
) -> Dict[str, List[Dict[str, str]]]:
    """
    Fetch file info for ALL WOIDs in a single query (or batched chunks).

    Returns
    -------
    Dict[str, List[Dict[str, str]]]
        Mapping of woid -> list of {"DocName": ..., "ReposSpec": ...}.
    """
    result: Dict[str, List[Dict[str, str]]] = {}

    for chunk in _chunked(woids, 500):
        placeholders = ",".join(["?" for _ in chunk])
        batch_query = _build_batch_file_query(placeholders)

        try:
            cursor = conn.cursor()
            cursor.execute(batch_query, chunk)
            rows = cursor.fetchall()
            cursor.close()

            for row in rows:
                woid = str(row[0]).strip()
                doc_name = str(row[1]).strip() if row[1] else ""
                repos_spec = str(row[2]).strip() if row[2] else ""
                if doc_name and repos_spec:
                    if woid not in result:
                        result[woid] = []
                    result[woid].append({"DocName": doc_name, "RecosSpec": repos_spec})

        except pyodbc.Error as exc:
            logger.error("Batch file info query error: %s", exc)
            # Fallback: query one by one
            for woid in chunk:
                try:
                    single = fetch_file_info(conn, woid)
                    if single:
                        result[woid] = single
                except Exception:
                    pass

    logger.info("Fetched file info for %d / %d WOIDs.", len(result), len(woids))
    return result


# ---------------------------------------------------------------------------
# Single-WOID queries (used as fallback)
# ---------------------------------------------------------------------------

def fetch_policy_period(
    conn: pyodbc.Connection, woid: str
) -> Optional[Tuple[date, date]]:
    """Execute Query 1 for a single WOID."""
    try:
        cursor = conn.cursor()
        cursor.execute(QUERY_POLICY_PERIOD, (woid,))
        row = cursor.fetchone()
        cursor.close()

        if row is None:
            logger.warning("WOID %s - No policy period found.", woid)
            return None

        inception = _to_date(row[0])
        expiration = _to_date(row[1])

        if inception is None or expiration is None:
            logger.error("WOID %s - Could not parse policy dates: %s, %s", woid, row[0], row[1])
            return None

        return (inception, expiration)

    except pyodbc.Error as exc:
        logger.error("WOID %s - SQL error fetching policy period: %s", woid, exc)
        raise


def fetch_file_info(
    conn: pyodbc.Connection, woid: str
) -> List[Dict[str, str]]:
    """Execute Query 2 for a single WOID."""
    try:
        cursor = conn.cursor()
        cursor.execute(QUERY_FILE_INFO, (woid,))
        rows = cursor.fetchall()
        cursor.close()

        if not rows:
            return []

        results: List[Dict[str, str]] = []
        for row in rows:
            doc_name = str(row[0]).strip() if row[0] else ""
            repos_spec = str(row[1]).strip() if row[1] else ""
            if doc_name and repos_spec:
                results.append({"DocName": doc_name, "RecosSpec": repos_spec})
        return results

    except pyodbc.Error as exc:
        logger.error("WOID %s - SQL error fetching file info: %s", woid, exc)
        raise


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _build_batch_policy_query(placeholders: str) -> str:
    """Build a batch policy query selecting woid + dates dynamically from config."""
    import re
    q = QUERY_POLICY_PERIOD.strip()
    # Ensure woid is selected
    q_sub = re.sub(r'(?i)^\s*select\s+', 'SELECT woid, ', q)
    # Replace woid = ? with woid IN (...)
    q_final = re.sub(r'(?i)\bwoid\s*=\s*\?', f'woid IN ({placeholders})', q_sub)
    return q_final


def _build_batch_file_query(placeholders: str) -> str:
    """Build a batch file-info query selecting PrimaryIndex + columns dynamically from config."""
    import re
    q = QUERY_FILE_INFO.strip()
    # Ensure PrimaryIndex is selected
    q_sub = re.sub(r'(?i)^\s*select\s+', 'SELECT PrimaryIndex, ', q)
    # Replace PrimaryIndex = ? with PrimaryIndex IN (...)
    q_final = re.sub(r'(?i)\bPrimaryIndex\s*=\s*\?', f'PrimaryIndex IN ({placeholders})', q_sub)
    return q_final


def _chunked(lst: List, size: int):
    """Yield successive chunks of *size* from *lst*."""
    for i in range(0, len(lst), size):
        yield lst[i : i + size]


def _to_date(value) -> Optional[date]:
    """Convert a value to a datetime.date."""
    if isinstance(value, date):
        return value if not isinstance(value, datetime) else value.date()
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, str):
        for fmt in ("%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
            try:
                return datetime.strptime(value.strip(), fmt).date()
            except ValueError:
                continue
    return None
