# PACE Payroll Period Coverage Validator — Technical Architecture & File Logic

This document provides a detailed breakdown of the components, algorithms, and design choices implemented in the PACE Payroll Period Coverage Validator.

---

## Technical Overview
The system validates whether the payroll period coverage extracted from Excel documents (stored in a document repository) fully covers a policy's active window (stored in a policy database) for a list of Work Order IDs (WOIDs). 

The pipeline is optimized for high throughput using:
- **Batch SQL queries** to reduce database round-trips from $2N$ to 2.
- **Concurrent Worker Pool** (`ThreadPoolExecutor`) for parallel file retrieval and Excel sheet parsing.
- **`openpyxl` Read-Only Mode** to bypass heavy `pandas` sheet-loading overhead.
- **Smart Sized-Based Caching** to avoid redundant copying of existing files.

---

## Component Breakdown

### 1. `config.py`
**Purpose:** Centralizes database queries, connection formats, and project path structures.
*   **Key Path Constants:** Defines `BASE_DIR`, `DATASET_DIR` (where downloaded payroll sheets are saved), and `LOG_FILE`.
*   **`get_connection_string()`**: Builds the `pyodbc` connection string for server `OSIRPTS01` using `Trusted_Connection=yes` (Windows Authentication) and `ApplicationIntent=Readonly` (read-only mode to prevent DB locks).
*   **`QUERY_POLICY_PERIOD`**: SQL statement to fetch `InceptionDate` and `ExpirationDate` for a WOID from `osi..WOPolicy`.
*   **`QUERY_FILE_INFO`**: SQL statement to fetch document names (`DocName`) and paths (`ReposSpec`) from `Docrepository..Documents` matching the description `'PACE Extracted'`.

---

### 2. `logger_setup.py`
**Purpose:** Sets up system-wide unified logging.
*   **Log Sinks:** Registers two output handlers:
    1.  **Console Handler:** Prints clean `INFO` level logs to `sys.stdout`.
    2.  **File Handler:** Writes verbose `DEBUG` level log traces (including stack traces) to `pace_validator.log` with `utf-8` encoding.
*   **Format:** Matches the format `YYYY-MM-DD HH:MM:SS | LEVEL | MESSAGE`.

---

### 3. `input_reader.py`
**Purpose:** Reads and normalizes input WOIDs from plain text or Excel spreadsheets.
*   **`read_woids(file_path)`**:
    - **TXT parser (`.txt`):** Reads the file line-by-line, strips leading/trailing whitespaces, and ignores blank lines.
    - **Excel parser (`.xlsx` / `.xls`):** Uses pandas to read the first column of the spreadsheet, strips whitespaces, and filters out nulls.
    - **Deduplication:** Filters out duplicate WOIDs while strictly preserving the initial sequence of the input list.

---

### 4. `db_client.py`
**Purpose:** Handles low-level DB execution, batches queries to avoid round-trip network lag, and parses dates.
*   **Batch Policy/File Queries (`fetch_all_policy_periods` / `fetch_all_file_info`)**:
    - Queries up to 500 WOIDs at a time using `IN (?, ?, ...)` clauses.
    - **Dynamic Query Construction:** Uses regular expressions to read queries defined in `config.py`, dynamically appends the group columns (`woid` and `PrimaryIndex`), and replaces the single-param equals sign (`= ?`) with the chunked placeholder `IN (...)`. This guarantees that if query filters in `config.py` change, the optimized batch queries update automatically.
    - **Fallback Handler:** If the batch query fails, the script gracefully falls back to querying the database one-by-one for each WOID to prevent a total crash.
*   **`_to_date(value)`**: Normalizes various DB date formats (datetime, date, strings) into standard Python `datetime.date` objects to avoid timezone offset shifts.

---

### 5. `file_retriever.py`
**Purpose:** Resolves UNC paths or disk specifications, copies payroll documents to the local project dataset, and manages caches.'

*   **`_locate_source_file(recos_spec, woid)`**:
    - Resolves raw paths without file extensions (e.g., `D:\PACE\DATA\12345`).
    - Checks for exact matching files, or searches the parent directory for files whose stem matches (e.g., `12345.xlsx` or `12345.xls`).
*   **`retrieve_files(woid, file_info_list)`**:
    - Creates target folder `dataset/<WOID>/`.
    - **Smart File-Level Caching:** Checks if the destination file already exists and matches the source file's size exactly (`dest_path.stat().st_size == source_file.stat().st_size`). If both conditions are met, it skips the file copy operation to save disk and network IO, but still appends the file path to be validated.
    - Copies missing or altered files via `shutil.copy2` to preserve file metadata.

---

### 6. `payroll_extractor.py`
**Purpose:** Rapidly opens Excel spreadsheets, identifies relevant data columns, and extracts pay periods.
*   **`extract_payroll_periods(file_paths, woid)`**:
    - Iterates over all files for a WOID.
    - Uses `openpyxl.load_workbook(..., read_only=True, data_only=True)`:
        - `read_only=True` avoids loading styles/formulas into memory, making file loads up to 5x faster.
        - `data_only=True` extracts the calculated values rather than original formulas.
*   **`_process_sheet_fast(...)`**:
    - Inspects the first row (header row) and normalizes cell texts (converts to lowercase, removes all spaces/tabs).
    - Looks for columns matching target headers: `"payperiodstart"` and `"payperiodend"`.
    - If found, it reads only those two columns row-by-row, skipping blanks and duplicates.
    - Uses `_parse_date` to parse string date formats or pandas timestamps to `datetime.date`.

---

### 7. `validator.py`
**Purpose:** Algorithmic validation of whether payroll intervals cover the policy period.
*   **`validate_coverage(policy_start, policy_end, payroll_periods)`**:
    - **Inverted Range Checks:** Ignores any payroll period where `start > end` and logs a warning. Checks if `policy_start > policy_end` and flags errors immediately.
    - **`_merge_periods(periods)`**: Sorts the list of intervals and merges overlapping or adjacent ranges. Two intervals are adjacent if they share a boundary or have a gap of exactly 1 day (e.g., `Jan 1 - Jan 14` and `Jan 15 - Jan 31` merge into `Jan 1 - Jan 31`).
    - **`_clip_to_policy(merged, ...)`**: Clips merged intervals to the policy start and end boundaries, ignoring ranges entirely outside.
    - **`_find_gaps(clipped, ...)`**: Locates gaps:
        - Gap before the first interval (if it starts after policy start).
        - Gaps between consecutive intervals.
        - Gap after the last interval (if it ends before policy end).
    - Returns Status:
        - **`FULL`**: Gaps list is empty.
        - **`PARTIAL`**: Gaps exist, but there is some overlap.
        - **`NO`**: Zero overlap between payroll files and policy window.

---

### 8. `report_generator.py`
**Purpose:** Formats and writes the final Microsoft Excel validation spreadsheet.
*   **`generate_report(results, output_path)`**:
    - Initializes an `openpyxl.Workbook` and configures the `Summary` worksheet.
    - Formats column widths and styles (Calibri font, thin grey borders).
    - **Metrics Section:** Computes and writes total processed count, number of FULL, PARTIAL, and NO statuses.
    - **Summary Table:** Color-codes validation results:
        - **FULL**: Green (`#217346` font / `#E2EFDA` background)
        - **PARTIAL**: Amber (`#BF8F00` font / `#FFF2CC` background)
        - **NO**: Red (`#C00000` font / `#FCE4EC` background)
    - Automatically marks sheets/files columns as "Single" or "Multiple" based on counts.

---

### 9. `main.py`
**Purpose:** Entry point that coordinates the application's runtime flow.
*   **Execution Flow:**
    1.  Reads WOIDs from the input file.
    2.  Establishes a database connection.
    3.  Fetches all policy periods and file definitions in **two batch queries**.
    4.  Spins up a `ThreadPoolExecutor` (using up to 8 threads).
    5.  Processes each WOID in parallel:
        - Locates and copies files (utilizing smart size-based caching).
        - Extracts periods using read-only openpyxl.
        - Runs the validation algorithm.
    6.  Sorts results back to matching input order.
    7.  Builds the unique timestamped output report name (`PACE_PayPeriod_Validation_Report_YYYYMMDD_HHMMSS.xlsx`).
    8.  Generates the Excel report, prints summary stats to the console, and updates logs.
