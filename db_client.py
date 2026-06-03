"""
PACE Payroll Validator — Database Client
==========================================
Handles SQL Server connectivity and query execution via pyodbc.
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

    Returns
    -------
    pyodbc.Connection
        Active database connection.

    Raises
    ------
    pyodbc.Error
        If the connection cannot be established.
    """
    conn_str = get_connection_string()
    logger.debug("Connecting to SQL Server …")
    conn = pyodbc.connect(conn_str, timeout=30)
    logger.info("SQL Server connection established.")
    return conn


def fetch_policy_period(
    conn: pyodbc.Connection, woid: str
) -> Optional[Tuple[date, date]]:
    """
    Execute Query 1 to retrieve the policy period for a WOID.

    Parameters
    ----------
    conn : pyodbc.Connection
    woid : str

    Returns
    -------
    Optional[Tuple[date, date]]
        (policy_start, policy_end) or None if no row found.
    """
    try:
        cursor = conn.cursor()
        cursor.execute(QUERY_POLICY_PERIOD, (woid,))
        row = cursor.fetchone()
        cursor.close()

        if row is None:
            logger.warning("WOID %s — No policy period found.", woid)
            return None

        inception = _to_date(row[0])
        expiration = _to_date(row[1])

        if inception is None or expiration is None:
            logger.error(
                "WOID %s — Could not parse policy dates: %s, %s",
                woid, row[0], row[1],
            )
            return None

        logger.info(
            "WOID %s - Policy period: %s -> %s", woid, inception, expiration
        )
        return (inception, expiration)

    except pyodbc.Error as exc:
        logger.error("WOID %s — SQL error fetching policy period: %s", woid, exc)
        raise


def fetch_file_info(
    conn: pyodbc.Connection, woid: str
) -> List[Dict[str, str]]:
    """
    Execute Query 2 to retrieve payroll file information for a WOID.

    Parameters
    ----------
    conn : pyodbc.Connection
    woid : str

    Returns
    -------
    List[Dict[str, str]]
        Each dict has keys ``DocName`` and ``RecosSpec``.
    """
    try:
        cursor = conn.cursor()
        cursor.execute(QUERY_FILE_INFO, (woid,))
        rows = cursor.fetchall()
        cursor.close()

        if not rows:
            logger.warning("WOID %s — No file info rows returned.", woid)
            return []

        results: List[Dict[str, str]] = []
        for row in rows:
            doc_name = str(row[0]).strip() if row[0] else ""
            recos_spec = str(row[1]).strip() if row[1] else ""
            if doc_name and recos_spec:
                results.append({"DocName": doc_name, "RecosSpec": recos_spec})
            else:
                logger.warning(
                    "WOID %s — Skipping row with empty DocName/RecosSpec: %s",
                    woid, row,
                )

        logger.info(
            "WOID %s — Found %d payroll file record(s).", woid, len(results)
        )
        return results

    except pyodbc.Error as exc:
        logger.error("WOID %s — SQL error fetching file info: %s", woid, exc)
        raise


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
