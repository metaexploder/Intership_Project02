"""
PACE Payroll Validator — Configuration
=======================================
Centralized configuration for database connections, SQL queries,
file paths, and application settings.

Update the placeholders below with actual values before running.
"""

import os
from pathlib import Path

# ---------------------------------------------------------------------------
# Project Paths
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DATASET_DIR = BASE_DIR / "dataset"
OUTPUT_REPORT = BASE_DIR / "PACE_PayPeriod_Validation_Report.xlsx"
LOG_FILE = BASE_DIR / "pace_validator.log"

# ---------------------------------------------------------------------------
# SQL Server Connection
# ---------------------------------------------------------------------------
def get_connection_string() -> str:
    """Build a pyodbc connection string from DB_CONFIG."""
    return (
        'DRIVER=ODBC DRIVER 17 FOR SQL Server;'
        'SERVER=OSIRPTS01;'
        'Trusted_Connection=yes;'
        'ApplicationIntent=Readonly;'
    )


# ---------------------------------------------------------------------------
# SQL Queries  (use '?' as the WOID parameter placeholder)
# ---------------------------------------------------------------------------
# Query 1 — Policy Period
QUERY_POLICY_PERIOD = """
    SELECT InceptionDate, ExpirationDate
    FROM osi..WOPolicy (nolock)
    WHERE woid = ?
"""

# Query 2 — Payroll File Information
QUERY_FILE_INFO = """
    SELECT DocName, ReposSpec
    FROM Docrepository..Documents (nolock)
    WHERE PrimaryIndex = ?
      AND DocDesc = 'PACE Extracted'
"""

