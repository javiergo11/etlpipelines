"""
extract_owner_financials.py
===========================
Command-line script for batch processing owner financial workbooks.

This is a thin wrapper around extractor_core.py.  It handles:
  - folder scanning (input/)
  - console prompts (property name)
  - file saving (output/)
  - progress printing

All extraction and output logic lives in extractor_core.py, which is
shared with the Streamlit webapp.

USAGE:
  - Drop .xlsx files into the "input" folder
  - Run:  python extract_owner_financials.py
  - Follow the prompts to name each output file
  - Check the "output" folder for results

This script is 100% backward compatible with the original workflow.
"""

import os
import sys
from datetime import datetime

# Import shared logic from the core module
from extractor_core import (
    find_sheet_by_prefix,
    guess_property_name,
    make_safe_filename,
    process_workbook,
    MANAGED_BY_OPTIONS,
    DEFAULT_MANAGED_BY,
)

# DB writer is optional — if missing, skip CSV output silently
try:
    import db_writer
    _DB_WRITER_AVAILABLE = True
except ImportError:
    _DB_WRITER_AVAILABLE = False


# ---------------------------------------------------------------------------
# CLI-ONLY FUNCTIONS (console I/O that doesn't belong in the core module)
# ---------------------------------------------------------------------------

def setup_folders():
    """Create 'input' and 'output' folders if they don't already exist."""
    for folder in ("input", "output"):
        if not os.path.exists(folder):
            os.makedirs(folder)
            print(f"  Created folder: {folder}/")


def find_xlsx_files(folder):
    """Return a list of .xlsx file paths inside the given folder."""
    files = []
    for name in sorted(os.listdir(folder)):
        if name.lower().endswith(".xlsx") and not name.startswith("~$"):
            files.append(os.path.join(folder, name))
    return files


def prompt_for_managed_by():
    """
    Ask which management company operates this property.
    Defaults to Extra Space Storage — just press Enter to accept.
    Phase 1: metadata only; does not change extraction logic.
    """
    print(f"\n  Management company [{' / '.join(MANAGED_BY_OPTIONS)}]")
    print(f"  Managed by [default: {DEFAULT_MANAGED_BY}]: ", end="")

    try:
        user_input = input().strip()
    except EOFError:
        user_input = ""

    chosen = user_input if user_input else DEFAULT_MANAGED_BY

    if chosen not in MANAGED_BY_OPTIONS:
        print(f"  NOTE: '{chosen}' is not in the recognized list. "
              f"Stored as-is; extraction uses EXR defaults.")

    return chosen


def prompt_for_property_name(filename):
    """
    Ask the user to enter a property name for the output file.
    Shows a default guess based on the input filename.
    """
    default = guess_property_name(filename)
    print(f"\n  Enter property name for output file [default: {default}]: ", end="")

    try:
        user_input = input().strip()
    except EOFError:
        user_input = ""

    return user_input if user_input else default


# ---------------------------------------------------------------------------
# MAIN PROCESSING LOOP
# ---------------------------------------------------------------------------

def process_file(filepath):
    """
    Process a single Excel file using the shared core logic.
    Handles prompting, printing, and saving to disk.
    """
    filename = os.path.basename(filepath)

    print(f"\n{'=' * 50}")
    print(f"  File: {filename}")
    print(f"{'=' * 50}")

    # Prompt for management company and property name (CLI-only)
    managed_by    = prompt_for_managed_by()
    property_name = prompt_for_property_name(filename)

    # Unique ID for this extraction run — used to link all DB rows together
    run_id       = datetime.now().strftime("%Y%m%d_%H%M%S")
    extracted_at = datetime.now().isoformat()

    # Call the shared core function (no printing, no file I/O)
    result = process_workbook(filepath, property_name, managed_by=managed_by)

    # Print the log entries to console
    for entry in result["log"]:
        status = entry["status"]
        sheet  = entry["sheet"]
        msg    = entry["message"]
        if status == "OK":
            print(f"  OK: {sheet} -> {msg}")
        elif status == "WARNING":
            print(f"  WARNING: {msg}" + (f" in {sheet}" if sheet else ""))
        elif status == "SKIP":
            print(f"  SKIP: {msg}")
        else:
            print(f"  ERROR: {msg}")

    # Save Excel workbook to disk
    if result["output_bytes"] is None:
        return None

    output_path = os.path.join("output", result["output_filename"])
    with open(output_path, "wb") as f:
        f.write(result["output_bytes"])

    print(f"\n  -> Saved: {output_path}")

    # Write database-ready CSV files to db_ready/
    if _DB_WRITER_AVAILABLE:
        counts = db_writer.write_all(
            run_id, property_name, managed_by, filename, result, extracted_at
        )
        total_rows = sum(counts.values())
        tables     = ", ".join(f"{t}: {n}" for t, n in counts.items())
        print(f"  -> DB:    {total_rows} rows written to db_ready/ ({tables})")
    else:
        print("  -> DB:    db_writer.py not found — skipping CSV output")

    return output_path


def main():
    """Main entry point."""
    print("=" * 60)
    print("  EXR Owner Financials Extractor  v3.0")
    print("=" * 60)

    setup_folders()

    files = find_xlsx_files("input")
    if not files:
        print("\n  No .xlsx files found in the 'input' folder.")
        print("  Drop your owner financial files there and run again.")
        print()
        sys.exit(0)

    print(f"\n  Found {len(files)} file(s) to process.")

    results = []
    for filepath in files:
        output_path = process_file(filepath)
        if output_path:
            results.append(output_path)

    print(f"\n{'=' * 60}")
    print(f"  Done!  {len(results)} output file(s) created in the 'output' folder:")
    for path in results:
        print(f"    {os.path.basename(path)}")
    print("=" * 60)


if __name__ == "__main__":
    main()
