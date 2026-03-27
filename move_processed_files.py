"""
move_processed_files.py
=======================
Moves files after you have reviewed your extracted output and confirmed
it looks correct.

WHAT IT DOES:
  1. Moves every file from 'output' into 'completed'
  2. Moves every file from 'input' into 'archive'
  3. If a file with the same name already exists in the destination,
     it adds a numeric suffix like _1, _2, etc. (never overwrites)

USAGE:
  - Review your output files first
  - Then run:  python move_processed_files.py
  - Or double-click:  archive_files.bat

This script does NOT touch your extraction script or any other files.
It only moves files between the four folders listed above.
"""

import os
import sys
import shutil


# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------

# Each tuple is (source_folder, destination_folder, description)
MOVE_RULES = [
    ("output", "completed", "extracted output files"),
    ("input",  "archive",   "original input files"),
]


# ---------------------------------------------------------------------------
# HELPER FUNCTIONS
# ---------------------------------------------------------------------------

def get_safe_destination(dest_folder, filename):
    """
    Return a filepath in dest_folder that does not already exist.

    If 'completed/Chattanooga_extracted.xlsx' already exists, this will
    return 'completed/Chattanooga_extracted_1.xlsx', then _2, _3, etc.
    """
    dest_path = os.path.join(dest_folder, filename)

    # If no conflict, use the original name
    if not os.path.exists(dest_path):
        return dest_path

    # Split the filename into name and extension
    name, ext = os.path.splitext(filename)

    # Try _1, _2, _3, ... until we find one that doesn't exist
    counter = 1
    while True:
        new_filename = f"{name}_{counter}{ext}"
        dest_path = os.path.join(dest_folder, new_filename)
        if not os.path.exists(dest_path):
            return dest_path
        counter += 1


def get_files_in_folder(folder):
    """
    Return a list of filenames (not paths) in the given folder.
    Skips hidden files (starting with .) and temp Excel files (~$...).
    Only returns files, not subfolders.
    """
    if not os.path.exists(folder):
        return []

    files = []
    for name in sorted(os.listdir(folder)):
        full_path = os.path.join(folder, name)
        # Skip directories, hidden files, and Excel temp files
        if os.path.isfile(full_path) and not name.startswith(".") and not name.startswith("~$"):
            files.append(name)
    return files


def move_files(source_folder, dest_folder, description):
    """
    Move all files from source_folder to dest_folder.
    Returns the number of files moved.
    """
    # Make sure the destination folder exists
    if not os.path.exists(dest_folder):
        os.makedirs(dest_folder)
        print(f"  Created folder: {dest_folder}/")

    files = get_files_in_folder(source_folder)

    if not files:
        print(f"  No {description} to move from '{source_folder}/'")
        return 0

    moved = 0
    for filename in files:
        source_path = os.path.join(source_folder, filename)
        dest_path = get_safe_destination(dest_folder, filename)
        dest_filename = os.path.basename(dest_path)

        shutil.move(source_path, dest_path)
        moved += 1

        # Show what happened (note if the name changed due to a conflict)
        if dest_filename != filename:
            print(f"  Moved: {filename}  ->  {dest_folder}/{dest_filename}  (renamed to avoid overwrite)")
        else:
            print(f"  Moved: {filename}  ->  {dest_folder}/")

    return moved


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main():
    print("=" * 55)
    print("  Move Processed Files")
    print("=" * 55)

    # Quick preview of what will be moved
    total_files = 0
    for source, dest, description in MOVE_RULES:
        count = len(get_files_in_folder(source))
        total_files += count
        print(f"  {source}/ -> {dest}/  ({count} file(s))")

    if total_files == 0:
        print("\n  Nothing to move. Both folders are empty.")
        print("=" * 55)
        sys.exit(0)

    # Confirm before moving
    print(f"\n  Ready to move {total_files} file(s) total.")
    print("  Press Enter to continue, or Ctrl+C to cancel: ", end="")
    try:
        input()
    except (EOFError, KeyboardInterrupt):
        print("\n  Cancelled.")
        sys.exit(0)

    # Move files
    total_moved = 0
    for source, dest, description in MOVE_RULES:
        print(f"\n  --- {description} ---")
        moved = move_files(source, dest, description)
        total_moved += moved

    print(f"\n{'=' * 55}")
    print(f"  Done!  {total_moved} file(s) moved.")
    print("=" * 55)


if __name__ == "__main__":
    main()
