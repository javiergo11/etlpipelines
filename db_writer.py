"""
db_writer.py
============
Writes database-ready CSV files from extracted property data.

Each table lives in a single CSV file inside the 'db_ready/' folder.
Files accumulate across extraction runs — each row carries a run_id so
you can identify which extraction produced it.

TABLE OVERVIEW
--------------
fact_rolling_is.csv           — Income statement (1 row per account per month per run)
fact_unit_rate_snapshot.csv   — Unit occupancy metrics (1 row per metric per run)
fact_ops_sum.csv              — Rental activity (1 row per metric per month per run)
fact_rent_roll_snapshot.csv   — Rent roll (1 row per unit per run)
etl_processing_log.csv        — Extraction audit trail (1 row per log event per run)

LOADING INTO SQLITE
-------------------
  import sqlite3, pandas as pd
  conn = sqlite3.connect("property_data.db")
  for table in ["fact_rolling_is", "fact_unit_rate_snapshot",
                "fact_ops_sum", "fact_rent_roll_snapshot", "etl_processing_log"]:
      df = pd.read_csv(f"db_ready/{table}.csv")
      df.to_sql(table, conn, if_exists="append", index=False)

LOADING INTO PANDAS
-------------------
  import pandas as pd
  df = pd.read_csv("db_ready/fact_rolling_is.csv")

SCHEMA STABILITY NOTE
---------------------
  Column names and order in each SCHEMA_* list are the stable contract.
  You can add columns at the END of a list without breaking existing data.
  Never rename, remove, or reorder existing columns — that will misalign
  data rows written in previous runs.
"""

import csv
import os
from datetime import datetime


# ---------------------------------------------------------------------------
# OUTPUT FOLDER
# ---------------------------------------------------------------------------

DB_READY_FOLDER = "db_ready"


# ---------------------------------------------------------------------------
# SCHEMA DEFINITIONS
# One list per table — these are the CSV column headers in column order.
# ---------------------------------------------------------------------------

SCHEMA_FACT_ROLLING_IS = [
    "run_id",             # extraction run identifier (YYYYMMDD_HHMMSS)
    "property_name",      # user-entered property name
    "property_number",    # numeric code from EXR sheet name (e.g. "7214")
    "managed_by",         # management company (e.g. "Extra Space Storage")
    "source_file",        # original input filename
    "reporting_period",   # most recent month in file (e.g. "Feb 2026")
    "line_item",          # EXR account label (e.g. "Rental Income (4000)")
    "coa",                # mapped P-Builder COA line item (nullable)
    "coa2",               # mapped rollup category (nullable)
    "account_type",       # Income / Expense / EXR_Rollup (nullable)
    "coa_confidence",     # float 0.0–1.0 from COA mapper (nullable)
    "coa_match_method",   # exact_approved / normalized / alias / fuzzy / no_match
    "coa_review_required",# 1 = needs review, 0 = auto-accepted, blank = no mapping run
    "month",              # integer month number (1–12)
    "year",               # four-digit year (e.g. 2026)
    "period_date",        # ISO date of the 1st of that month (YYYY-MM-01)
    "amount",             # dollar amount (negative = expense or contra)
    "extracted_at",       # ISO datetime of this extraction
]

SCHEMA_FACT_UNIT_RATE_SNAPSHOT = [
    "run_id",
    "property_name",
    "property_number",
    "managed_by",
    "source_file",
    "reporting_period",   # date context for the snapshot
    "metric",             # e.g. "Units Available", "Sq Ft Rented"
    "value",              # numeric value
    "extracted_at",
]

SCHEMA_FACT_OPS_SUM = [
    "run_id",
    "property_name",
    "property_number",
    "managed_by",
    "source_file",
    "metric",             # e.g. "Rentals During Month", "Vacates During Month"
    "month",
    "year",
    "period_date",
    "value",
    "extracted_at",
]

SCHEMA_FACT_RENT_ROLL_SNAPSHOT = [
    "run_id",
    "property_name",
    "property_number",
    "managed_by",
    "source_file",
    "reporting_period",
    "tenant_account",     # tenant account identifier
    "unit_number",        # unit number/ID
    "move_in_date",       # ISO date (YYYY-MM-DD)
    "rent_rate",          # current in-place rent
    "street_rate",        # current advertised street rate
    "paid_thru_date",     # ISO date (YYYY-MM-DD)
    "status",             # e.g. "Rented", "Delinquent"
    "size",               # unit size string (e.g. "10X13")
    "unit_type",          # unit type code
    "sq_ft",              # calculated square footage (width * depth)
    "extracted_at",
]

SCHEMA_ETL_PROCESSING_LOG = [
    "run_id",
    "property_name",
    "managed_by",
    "source_file",
    "sheet_name",         # which sheet was being processed
    "status",             # OK / WARNING / ERROR / SKIP
    "message",            # human-readable description
    "extracted_at",
]


# ---------------------------------------------------------------------------
# INTERNAL HELPERS
# ---------------------------------------------------------------------------

def _ensure_folder():
    """Create the db_ready folder if it doesn't exist."""
    if not os.path.exists(DB_READY_FOLDER):
        os.makedirs(DB_READY_FOLDER)
        print(f"  Created folder: {DB_READY_FOLDER}/")


def _csv_path(table_name):
    """Return the full path for a table's CSV file."""
    return os.path.join(DB_READY_FOLDER, f"{table_name}.csv")


def _fmt_date(value):
    """
    Format a value as ISO date string YYYY-MM-DD.
    Handles datetime objects; passes strings and None through as-is.
    """
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    return value if value is not None else ""


def _append_rows(table_name, schema, rows):
    """
    Append rows to a table CSV file.

    Writes the header row only if the file does not exist or is empty.
    Rows must be a list of dicts keyed by schema column names.
    Extra keys are silently ignored (extrasaction='ignore').

    Returns the number of rows written.
    """
    if not rows:
        return 0

    _ensure_folder()
    path      = _csv_path(table_name)
    write_hdr = not os.path.exists(path) or os.path.getsize(path) == 0

    with open(path, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=schema, extrasaction="ignore")
        if write_hdr:
            writer.writeheader()
        writer.writerows(rows)

    return len(rows)


def _parse_period(date_str):
    """
    Convert a 'Feb 2025' date string to (month int, year int, 'YYYY-MM-01' str).
    Returns (None, None, None) on failure.
    """
    try:
        dt = datetime.strptime(date_str, "%b %Y")
        return dt.month, dt.year, dt.strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        return None, None, None


# ---------------------------------------------------------------------------
# TABLE WRITERS
# One function per table. Each takes the raw extracted data and appends rows.
# ---------------------------------------------------------------------------

def write_fact_rolling_is(run_id, property_name, prop_num, managed_by,
                           source_file, reporting_period,
                           dates, rows, coa_lookup, extracted_at):
    """
    Append Rolling IS data to fact_rolling_is.csv.

    dates:      list of "Mon YYYY" strings
    rows:       list of {"label": str, "values": [amount, ...]}
    coa_lookup: dict {label: coa_result} from COAMapper (may be empty)
    """
    output_rows = []
    for row_data in rows:
        label   = row_data["label"]
        mapping = coa_lookup.get(label, {})

        for i, date_str in enumerate(dates):
            amount      = row_data["values"][i] if i < len(row_data["values"]) else None
            month, year, period_date = _parse_period(date_str)

            # coa_review_required: 1/0 when mapper ran; blank when it didn't
            if mapping:
                review_val = 1 if mapping.get("review_required") else 0
            else:
                review_val = ""

            output_rows.append({
                "run_id":              run_id,
                "property_name":       property_name,
                "property_number":     prop_num,
                "managed_by":          managed_by,
                "source_file":         source_file,
                "reporting_period":    reporting_period,
                "line_item":           label,
                "coa":                 mapping.get("coa", ""),
                "coa2":                mapping.get("coa2", ""),
                "account_type":        mapping.get("account_type", ""),
                "coa_confidence":      mapping.get("confidence", ""),
                "coa_match_method":    mapping.get("match_method", ""),
                "coa_review_required": review_val,
                "month":               month,
                "year":                year,
                "period_date":         period_date,
                "amount":              amount if amount is not None else 0,
                "extracted_at":        extracted_at,
            })

    return _append_rows("fact_rolling_is", SCHEMA_FACT_ROLLING_IS, output_rows)


def write_fact_unit_rate_snapshot(run_id, property_name, prop_num, managed_by,
                                   source_file, reporting_period,
                                   metrics, extracted_at):
    """
    Append Unit Rate metrics to fact_unit_rate_snapshot.csv.

    metrics: dict {label: numeric_value}
    """
    output_rows = [
        {
            "run_id":           run_id,
            "property_name":    property_name,
            "property_number":  prop_num,
            "managed_by":       managed_by,
            "source_file":      source_file,
            "reporting_period": reporting_period,
            "metric":           label,
            "value":            value,
            "extracted_at":     extracted_at,
        }
        for label, value in metrics.items()
    ]
    return _append_rows("fact_unit_rate_snapshot",
                        SCHEMA_FACT_UNIT_RATE_SNAPSHOT, output_rows)


def write_fact_ops_sum(run_id, property_name, prop_num, managed_by,
                        source_file, dates, rows, extracted_at):
    """
    Append Ops Sum data to fact_ops_sum.csv.

    dates: list of "Mon YYYY" strings
    rows:  list of {"label": str, "values": [value, ...]}
    """
    output_rows = []
    for row_data in rows:
        for i, date_str in enumerate(dates):
            value = row_data["values"][i] if i < len(row_data["values"]) else None
            month, year, period_date = _parse_period(date_str)
            output_rows.append({
                "run_id":          run_id,
                "property_name":   property_name,
                "property_number": prop_num,
                "managed_by":      managed_by,
                "source_file":     source_file,
                "metric":          row_data["label"],
                "month":           month,
                "year":            year,
                "period_date":     period_date,
                "value":           value if value is not None else 0,
                "extracted_at":    extracted_at,
            })
    return _append_rows("fact_ops_sum", SCHEMA_FACT_OPS_SUM, output_rows)


def write_fact_rent_roll_snapshot(run_id, property_name, prop_num, managed_by,
                                   source_file, reporting_period,
                                   headers, data_rows, extracted_at):
    """
    Append Rent Roll data to fact_rent_roll_snapshot.csv.

    headers:   list of column name strings (e.g. RENT_ROLL_HEADERS + ["Sq Ft"])
    data_rows: list of lists — one list per tenant/unit, values aligned to headers
    """
    # Map the header names we care about to their column indices
    WANT = {
        "Tenant Account": "tenant_account",
        "Unit #":         "unit_number",
        "Move-In Date":   "move_in_date",
        "Rent Rate":      "rent_rate",
        "Street Rate":    "street_rate",
        "Paid-Thru Date": "paid_thru_date",
        "Status":         "status",
        "Size":           "size",
        "Type":           "unit_type",
        "Sq Ft":          "sq_ft",
    }
    date_cols = {"Move-In Date", "Paid-Thru Date"}
    idx       = {col: headers.index(col) for col in WANT if col in headers}

    def _get(row, col):
        i = idx.get(col)
        if i is None or i >= len(row):
            return ""
        val = row[i]
        return _fmt_date(val) if col in date_cols else (val if val is not None else "")

    output_rows = [
        {
            "run_id":           run_id,
            "property_name":    property_name,
            "property_number":  prop_num,
            "managed_by":       managed_by,
            "source_file":      source_file,
            "reporting_period": reporting_period,
            "tenant_account":   _get(row, "Tenant Account"),
            "unit_number":      _get(row, "Unit #"),
            "move_in_date":     _get(row, "Move-In Date"),
            "rent_rate":        _get(row, "Rent Rate"),
            "street_rate":      _get(row, "Street Rate"),
            "paid_thru_date":   _get(row, "Paid-Thru Date"),
            "status":           _get(row, "Status"),
            "size":             _get(row, "Size"),
            "unit_type":        _get(row, "Type"),
            "sq_ft":            _get(row, "Sq Ft"),
            "extracted_at":     extracted_at,
        }
        for row in data_rows
    ]
    return _append_rows("fact_rent_roll_snapshot",
                        SCHEMA_FACT_RENT_ROLL_SNAPSHOT, output_rows)


def write_etl_processing_log(run_id, property_name, managed_by,
                              source_file, log_entries, extracted_at):
    """
    Append the processing log to etl_processing_log.csv.

    log_entries: list of dicts with keys: sheet, status, message
    """
    output_rows = [
        {
            "run_id":        run_id,
            "property_name": property_name,
            "managed_by":    managed_by,
            "source_file":   source_file,
            "sheet_name":    e.get("sheet", ""),
            "status":        e.get("status", ""),
            "message":       e.get("message", ""),
            "extracted_at":  extracted_at,
        }
        for e in log_entries
    ]
    return _append_rows("etl_processing_log",
                        SCHEMA_ETL_PROCESSING_LOG, output_rows)


# ---------------------------------------------------------------------------
# CONVENIENCE: write all tables for one extraction run
# ---------------------------------------------------------------------------

def write_all(run_id, property_name, managed_by, source_file, result, extracted_at):
    """
    Write all database-ready CSV files for one extraction run.

    result: the dict returned by process_workbook() from extractor_core.py
    Skips any table for which data is None (sheet was missing in the source).

    Returns a dict: {table_name: row_count} for each table written.
    """
    rolling_is_data = result.get("rolling_is_data")
    unit_rate_data  = result.get("unit_rate_data")
    ops_sum_data    = result.get("ops_sum_data")
    rent_roll_data  = result.get("rent_roll_data")
    coa_lookup      = result.get("coa_lookup", {})
    log_entries     = result.get("log", [])

    # Use the most recent month in the Rolling IS as the reporting period
    reporting_period = ""
    if rolling_is_data and rolling_is_data.get("dates"):
        reporting_period = rolling_is_data["dates"][-1]   # e.g. "Feb 2026"

    counts = {}

    if rolling_is_data:
        counts["fact_rolling_is"] = write_fact_rolling_is(
            run_id, property_name, rolling_is_data["prop_num"],
            managed_by, source_file, reporting_period,
            rolling_is_data["dates"], rolling_is_data["rows"],
            coa_lookup, extracted_at,
        )

    if unit_rate_data:
        counts["fact_unit_rate_snapshot"] = write_fact_unit_rate_snapshot(
            run_id, property_name, unit_rate_data["prop_num"],
            managed_by, source_file, reporting_period,
            unit_rate_data["metrics"], extracted_at,
        )

    if ops_sum_data:
        counts["fact_ops_sum"] = write_fact_ops_sum(
            run_id, property_name, ops_sum_data["prop_num"],
            managed_by, source_file,
            ops_sum_data["dates"], ops_sum_data["rows"], extracted_at,
        )

    if rent_roll_data:
        counts["fact_rent_roll_snapshot"] = write_fact_rent_roll_snapshot(
            run_id, property_name, rent_roll_data["prop_num"],
            managed_by, source_file, reporting_period,
            rent_roll_data["headers"], rent_roll_data["data_rows"], extracted_at,
        )

    counts["etl_processing_log"] = write_etl_processing_log(
        run_id, property_name, managed_by, source_file, log_entries, extracted_at,
    )

    return counts
