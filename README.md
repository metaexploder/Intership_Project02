# PACE Payroll Period Coverage Validator

A high-performance, concurrent Python utility designed to validate whether payroll coverages (extracted from physical Excel documents) fully cover active policy windows for a set of Work Order IDs (WOIDs).

---

## Key Features

*   **Database Batching:** Replaces the slow $2N$ sequential queries with **2 total database queries** using dynamic batch query builders, drastically reducing SQL Server connection overhead.
*   **Parallel Execution:** Orchestrates file retrieval, size-based cache matching, and Excel sheet parsing concurrently using a worker pool (`ThreadPoolExecutor`).
*   **Fast Excel Parsing:** Leverages `openpyxl`'s `read_only` stream mode to avoid heavy memory allocation, parsing column structures up to 5x faster than standard pandas/openpyxl routines.
*   **Smart Sized-Based Caching:** Automatically skips copy operations for files that already exist in `dataset/<WOID>/` and match the size of database source files, preventing network bottlenecks on repeated runs.
*   **Interval Merging Algorithm:** Sorts, merges, and clips adjacent or overlapping date intervals to calculate precise coverage gaps without off-by-one errors.
*   **Dynamic Report Generation:** Generates styled, color-coded Microsoft Excel reports (`PACE_PayPeriod_Validation_Report_YYYYMMDD_HHMMSS.xlsx`) featuring color-coded coverage statuses (FULL, PARTIAL, NO) and run statistics.

---

## Project Structure

```
PACE VALIDATOR/
├── main.py                  # Pipeline orchestrator & entry point
├── config.py                # Database connection, paths, and SQL query definitions
├── logger_setup.py          # Unified console & file logging setup
├── input_reader.py          # Parses unique WOID lists from TXT or Excel
├── db_client.py             # Database client (supports dynamic batch queries)
├── file_retriever.py        # Location resolution, copying, and file-level caching
├── payroll_extractor.py     # Fast column matching & pay period extraction
├── validator.py             # Algorithmic date range merging and gap detection
├── report_generator.py      # openpyxl formatted summary workbook builder
├── requirements.txt         # Package dependencies (pyodbc, openpyxl, pandas)
└── Working.md               # Detailed technical architecture & logical design
```

---

## Quick Start

### 1. Installation
Install the required packages using pip:
```bash
pip install -r requirements.txt
```

### 2. Configuration
Open [config.py](file:///c:/Users/Vishal Chauhan/Desktop/PACE VALIDATOR/config.py) and update placeholders with your SQL Server configurations:
*   `get_connection_string()`: Update Server name, driver, or Authentication methods.
*   `QUERY_POLICY_PERIOD`: Standard SELECT query for policy windows.
*   `QUERY_FILE_INFO`: Standard SELECT query for documents.

### 3. Usage
Run the application by passing your input file (either `.txt` or `.xlsx` containing WOIDs):
```bash
python main.py
```
*(By default, this reads from `woids.txt`. You can edit `INPUT_FILE` in `main.py` or modify the file path directly).*

---

## Validation Statuses

*   **`FULL`** (Green): Payroll periods cover every single day of the policy period.
*   **`PARTIAL`** (Amber): The policy period is partially covered. Gaps are identified, listed, and logged.
*   **`NO`** (Red): No overlap whatsoever between payroll periods and the policy window, or no files were found.

For more technical detail, see [Working.md](file:///c:/Users/Vishal%20Chauhan/Desktop/PACE%20VALIDATOR/Working.md).
