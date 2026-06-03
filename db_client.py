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
) -> Dict[str, Optional[Tuple[date, date]]]:
    """
    Batch-fetch policy periods for all WOIDs in one query.

    Returns dict mapping woid -> (policy_start, policy_end) or None.
    """
    result: Dict[str, Optional[Tuple[date, date]]] = {}

    # Process in chunks of 500 to avoid SQL parameter limits
    for chunk in _chunked(woids, 500):
        placeholders = ",".join(["?" for _ in chunk])
        # Dynamically build batch query from the single-WOID query pattern
        query = f"""
            SELECT woid, InceptionDate, ExpirationDate
            FROM osi..WOPolicy (nolock)
            WHERE woid IN ({placeholders})
        """
        try:
            cursor = conn.cursor()
            cursor.execute(query, chunk)
            rows = cursor.fetchall()
            cursor.close()

            for row in rows:
                woid_val = str(row[0]).strip()
                inception = _to_date(row[1])
                expiration = _to_date(row[2])
                if inception and expiration:
                    result[woid_val] = (inception, expiration)
                    logger.info("WOID %s - Policy period: %s -> %s",
                                woid_val, inception, expiration)
                else:
                    logger.error("WOID %s - Could not parse policy dates: %s, %s",
                                 woid_val, row[1], row[2])
                    result[woid_val] = None

        except pyodbc.Error as exc:
            logger.error("Batch policy query error: %s", exc)
            # Fall back to individual queries for this chunk
            for woid in chunk:
                result[woid] = fetch_policy_period(conn, woid)

    # Mark missing WOIDs
    for woid in woids:
        if woid not in result:
            logger.warning("WOID %s - No policy period found.", woid)
            result[woid] = None

    logger.info("Batch policy query complete: %d/%d found.",
                sum(1 for v in result.values() if v is not None), len(woids))
    return result


def fetch_all_file_info(
    conn: pyodbc.Connection, woids: List[str]
) -> Dict[str, List[Dict[str, str]]]:
    """
    Batch-fetch file info for all WOIDs in one query.

    Returns dict mapping woid -> list of {DocName, ReposSpec}.
    """
    result: Dict[str, List[Dict[str, str]]] = {w: [] for w in woids}

    for chunk in _chunked(woids, 500):
        placeholders = ",".join(["?" for _ in chunk])
        query = f"""
            SELECT PrimaryIndex, DocName, ReposSpec
            FROM Docrepository..Documents (nolock)
            WHERE PrimaryIndex IN ({placeholders})
              AND DocDesc LIKE '%Payroll%'
        """
        try:
            cursor = conn.cursor()
            cursor.execute(query, chunk)
            rows = cursor.fetchall()
            cursor.close()

            for row in rows:
                woid_val = str(row[0]).strip()
                doc_name = str(row[1]).strip() if row[1] else ""
                repos_spec = str(row[2]).strip() if row[2] else ""
                if doc_name and repos_spec and woid_val in result:
                    result[woid_val].append({"DocName": doc_name, "ReposSpec": repos_spec})

        except pyodbc.Error as exc:
            logger.error("Batch file-info query error: %s", exc)
            # Fall back to individual queries for this chunk
            for woid in chunk:
                result[woid] = fetch_file_info(conn, woid)

    for woid in woids:
        count = len(result[woid])
        if count:
            logger.info("WOID %s - Found %d payroll file record(s).", woid, count)
        else:
            logger.warning("WOID %s - No file info rows returned.", woid)

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
            return None

        inception = _to_date(row[0])
        expiration = _to_date(row[1])

        if inception is None or expiration is None:
            return None

        return (inception, expiration)

    except pyodbc.Error as exc:
        logger.error("WOID %s - SQL error fetching policy period: %s", woid, exc)
        return None


def fetch_file_info(
    conn: pyodbc.Connection, woid: str
) -> List[Dict[str, str]]:
    """Execute Query 2 for a single WOID."""
    try:
        cursor = conn.cursor()
        cursor.execute(QUERY_FILE_INFO, (woid,))
        rows = cursor.fetchall()
        cursor.close()

        results: List[Dict[str, str]] = []
        for row in rows:
            doc_name = str(row[0]).strip() if row[0] else ""
            repos_spec = str(row[1]).strip() if row[1] else ""
            if doc_name and repos_spec:
                results.append({"DocName": doc_name, "ReposSpec": repos_spec})
        return results

    except pyodbc.Error as exc:
        logger.error("WOID %s - SQL error fetching file info: %s", woid, exc)
        return []


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _to_date(value) -> Optional[date]:
    """Convert a value to a ``datetime.date``."""
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


def _chunked(lst: List, size: int):
    """Yield successive chunks of *size* from *lst*."""
    for i in range(0, len(lst), size):
        yield lst[i : i + size]
