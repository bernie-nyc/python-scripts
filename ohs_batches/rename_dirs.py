#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
rename_dirs.py  —  Meticulously commented directory renamer for non-coders
================================================================================

GOAL
----
Rename folders whose names are legacy IDs so that each folder name becomes:
    <LegacyID>_<PersonID>_<Full Name>

EXAMPLE
-------
Before:  12345678
After:   12345678_104942_Abdulraheem, Hanan

DATA INPUTS (expected to be in the SAME folder where you run this script)
-----------------------------------------------------------------------
1) folder_legacyid.txt
   - One legacy ID per line.
   - Each legacy ID should match an existing folder name that you want to rename.

2) Person Query CSV (example filename below)
   - Must contain EXACT column headers:
        "Person ID", "Full Name", "Legacy Person ID"
   - Example filename used in this script:
        "Person Query (1582 records) 2025-08-23.csv"
   - If your CSV has a different name, change PERSON_CSV_FILE below.

SAFETY
------
- Default mode is DRY RUN (no changes). It will only PRINT what it WOULD do.
- Use the --apply switch to actually rename folders.
- The script NEVER overwrites an existing folder:
    If the target name already exists, it adds "_1", "_2", ... until unique.

WINDOWS NAME SAFETY
-------------------
- Removes characters that Windows does not allow in folder names.
- Removes illegal trailing dots/spaces.
- Avoids reserved device names (e.g., CON, NUL, COM1).

USAGE
-----
Open PowerShell (or Command Prompt) in the folder that contains:
  - this script
  - folder_legacyid.txt
  - the Person Query CSV
Then run:

  1) DRY RUN (no changes):
        python rename_dirs.py

  2) APPLY CHANGES:
        python rename_dirs.py --apply

WHAT IF SOMETHING FAILS?
------------------------
- The script prints an error and continues with the next folder.
- Nothing is deleted. Only folder names are changed.

================================================================================
"""

# ------------------------------
# IMPORTS: built-in Python tools
# ------------------------------
import argparse          # Reads command-line options like --apply
import csv               # Reads the Person Query CSV file
import os                # Used only for environment basics (Path does the heavy lifting)
import re                # Regular expressions for name sanitization
from pathlib import Path # Safer path handling than raw strings (works on Windows)
from typing import Dict, Tuple, Set

# ------------------------------
# USER-ADJUSTABLE FILENAMES
# ------------------------------
# Name of the text file with one legacy ID per line
LEGACY_LIST = "folder_legacyid.txt"

# Name of the CSV file with the required columns
# If your file has a different name, change this string to match your file.
PERSON_CSV_FILE = "Person Query (1582 records) 2025-08-23.csv"


# ==============================================================================
# HELPER FUNCTIONS — small, focused tasks with ultra-clear comments
# ==============================================================================

# ------------------------------
# Windows filename sanitization
# ------------------------------
# 1) Define a set of characters that Windows forbids in file/folder names.
#    <>:"/\|?* are illegal. Control chars \x00-\x1F are also invalid.
WIN_ILLEGAL = r'<>:"/\\|?*\x00-\x1F'

# 2) Pre-compile a regular expression that finds any of the illegal characters.
_illegal_re = re.compile(f"[{re.escape(WIN_ILLEGAL)}]")

# 3) Trailing dots or spaces are not allowed on Windows folder names.
_trailing_re = re.compile(r"[\. ]+$")

# 4) Windows has reserved device names that cannot be used as names, case-insensitive.
_reserved: Set[str] = {
    "CON", "PRN", "AUX", "NUL",
    *{f"COM{i}" for i in range(1, 10)},
    *{f"LPT{i}" for i in range(1, 10)},
}

def sanitize_component(name: str) -> str:
    """
    Make a single folder name 'component' safe for Windows.

    Steps:
    - Replace every illegal character with underscore "_".
    - Remove any trailing dots or spaces.
    - If the result is empty, use "_" so it's still a valid name.
    - If the name (before any dot) is a reserved device (like "CON"), append "_".
    """
    # Replace illegal characters with underscore
    out = _illegal_re.sub("_", name)

    # Remove trailing dots/spaces that Windows forbids
    out = _trailing_re.sub("", out)

    # If after cleaning nothing is left, fall back to "_"
    if not out:
        out = "_"

    # If the part before the first dot is a reserved device name, add "_" to avoid conflict
    base_before_dot = out.split(".", 1)[0]
    if base_before_dot.upper() in _reserved:
        out = f"{out}_"

    return out


def ensure_unique_path(dst: Path) -> Path:
    """
    Ensure 'dst' (the desired new folder path) does NOT already exist.

    If it does not exist:
        - Return 'dst' unchanged.

    If it DOES exist:
        - Append a numeric suffix to the folder name:
            <name>, <name>_1, <name>_2, ...
          Stop at the first name that does not exist and return that.

    This guarantees we never overwrite or collide with an existing folder.
    """
    # If the desired destination name is free, use it as-is.
    if not dst.exists():
        return dst

    # Split name and extension-like tail. For folders, ".suffix" is usually empty,
    # but we handle it generically.
    name = dst.name
    stem, dot, ext = name.partition(".")

    # Try adding _1, _2, ... until a free name is found.
    n = 1
    while True:
        candidate = dst.with_name(f"{stem}_{n}{dot}{ext}")
        if not candidate.exists():
            return candidate
        n += 1


def load_legacy_list(path: Path) -> Set[str]:
    """
    Read 'folder_legacyid.txt' and return a SET of legacy IDs.

    - We use a set to avoid duplicates automatically.
    - Empty lines are ignored.
    - Leading/trailing spaces on each line are removed.
    """
    ids: Set[str] = set()  # Use a set so each ID is unique
    with path.open("r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            s = line.strip()  # Remove leading/trailing whitespace/newlines
            if s:              # Skip blank lines
                ids.add(s)     # Add to the set (duplicates auto-ignored)
    return ids


def load_person_map(csv_path: Path) -> Dict[str, Tuple[str, str]]:
    """
    Read the Person Query CSV and return a DICTIONARY mapping:
        Legacy Person ID  -->  (Person ID, Full Name)

    Why a dictionary?
        - So we can very quickly look up the (Person ID, Full Name)
          for a given Legacy Person ID (which is the folder name).
    """
    mapping: Dict[str, Tuple[str, str]] = {}

    # Open the CSV safely, ignoring weird characters if present.
    with csv_path.open("r", encoding="utf-8", errors="ignore", newline="") as f:
        reader = csv.DictReader(f)  # Reads rows using column names

        # We expect columns exactly named:
        #   "Person ID", "Full Name", "Legacy Person ID"
        for row in reader:
            legacy = (row.get("Legacy Person ID") or "").strip()
            pid    = (row.get("Person ID")        or "").strip()
            fname  = (row.get("Full Name")        or "").strip()

            # If legacy is missing, we cannot map that row
            if not legacy:
                continue

            # Store the two useful pieces keyed by legacy
            mapping[legacy] = (pid, fname)

    return mapping


# ==============================================================================
# MAIN PROGRAM — drives the whole process
# ==============================================================================

def main() -> None:
    """
    Orchestrates the entire rename process, step by step:

    1) Parse the --apply flag to decide DRY RUN vs APPLY.
    2) Locate the working directory (where the script is run).
    3) Load 'folder_legacyid.txt' into a set of legacy IDs.
    4) Load the person CSV into a mapping: legacy ID -> (person id, full name).
    5) For each legacy ID that also exists as a folder:
         a) Build the new folder name: LegacyID_PersonID_FullName
         b) Sanitize the new name (Windows safe)
         c) Ensure uniqueness to avoid collisions
         d) Print what would happen (DRY RUN) or perform the rename (APPLY)
    6) Print a summary of results.
    """

    # -------------------------------------------------------------------------
    # STEP 1: Read command-line options
    # -------------------------------------------------------------------------
    parser = argparse.ArgumentParser(
        description="Rename folders from <LegacyID> to <LegacyID>_<PersonID>_<Full Name>."
    )
    # --apply means "actually do it". If omitted, we run in DRY RUN mode.
    parser.add_argument(
        "--apply",
        action="store_true",
        help="Perform renames. Omit this flag to preview (dry run)."
    )
    args = parser.parse_args()
    apply_changes: bool = args.apply

    # -------------------------------------------------------------------------
    # STEP 2: Determine the 'root' folder (current working directory)
    # -------------------------------------------------------------------------
    root: Path = Path.cwd()  # Folder where you typed: python rename_dirs.py
    print(f"Working folder: {root}")
    print(f"Mode: {'APPLY (will rename)' if apply_changes else 'DRY RUN (no changes)'}\n")

    # -------------------------------------------------------------------------
    # STEP 3: Resolve the two input files and check they exist
    # -------------------------------------------------------------------------
    legacy_file: Path = root / LEGACY_LIST
    person_csv:  Path = root / PERSON_CSV_FILE

    # If the legacy list is missing, we cannot proceed.
    if not legacy_file.exists():
        print(f"ERROR: Required file not found: {legacy_file.name}")
        return

    # If the CSV is missing, we cannot map legacy IDs to person info.
    if not person_csv.exists():
        print(f"ERROR: Required file not found: {person_csv.name}")
        print("TIP: If your CSV has a different name, edit PERSON_CSV_FILE at the top of this script.")
        return

    # -------------------------------------------------------------------------
    # STEP 4: Load the legacy IDs and the person mapping
    # -------------------------------------------------------------------------
    legacy_ids: Set[str] = load_legacy_list(legacy_file)
    print(f"Loaded {len(legacy_ids)} legacy ID(s) from {legacy_file.name}")

    person_map: Dict[str, Tuple[str, str]] = load_person_map(person_csv)
    print(f"Loaded {len(person_map)} mapping row(s) from {person_csv.name}\n")

    # -------------------------------------------------------------------------
    # STEP 5: Plan the renames
    # -------------------------------------------------------------------------
    # We will create a list of "tasks", each task is a pair:
    #   (source_folder_path, destination_folder_path)
    tasks: list[tuple[Path, Path]] = []

    # Sort legacy IDs so output is stable and predictable
    for legacy in sorted(legacy_ids):
        # The "source" folder we hope to rename must exist right here under root
        src_dir: Path = root / legacy

        # If the folder doesn't exist, we skip it quietly (you can print if desired)
        if not src_dir.exists() or not src_dir.is_dir():
            # Uncomment the next line if you want to be notified about missing folders
            # print(f"SKIP (folder not found): {legacy}")
            continue

        # We need a row in the CSV where "Legacy Person ID" equals this folder name
        if legacy not in person_map:
            print(f"SKIP (no CSV match): {legacy}")
            continue

        # Extract Person ID and Full Name from the mapping dictionary
        person_id, full_name = person_map[legacy]

        # Sanitize each piece so it's safe on Windows:
        # - legacy is already "folder name", but sanitize anyway for consistency
        safe_legacy   = sanitize_component(legacy)
        safe_personid = sanitize_component(person_id)
        safe_fullname = sanitize_component(full_name)

        # Build the NEW folder name we want:
        #   <LegacyID>_<PersonID>_<Full Name>
        new_name: str = f"{safe_legacy}_{safe_personid}_{safe_fullname}"

        # Create a full path object for the destination folder
        dst_dir: Path = root / new_name

        # If that destination already exists, find a unique variant with _1/_2/etc.
        dst_dir = ensure_unique_path(dst_dir)

        # IMPORTANT:
        # If the computed destination name equals the current name (already renamed),
        # we should skip to avoid a no-op (or confusing logs).
        if dst_dir.name == src_dir.name:
            # Already in desired format (or equivalent). Skip.
            continue

        # Add this planned operation to our task list
        tasks.append((src_dir, dst_dir))

    # Print a preview of what we will do
    print(f"Planned renames: {len(tasks)}")
    for src, dst in tasks:
        print(f" - {src.name}  ->  {dst.name}")

    # If there is nothing to do, stop here
    if not tasks:
        print("\nNothing to rename.")
        return

    # -------------------------------------------------------------------------
    # STEP 6: Execute the plan (or just preview if DRY RUN)
    # -------------------------------------------------------------------------
    if not apply_changes:
        # DRY RUN: Only print what WOULD happen. No changes made.
        print("\nDry run complete. Re-run with --apply to commit changes.")
        return

    # APPLY mode: actually rename folders now.
    print("\nApplying renames...")
    ok = 0   # count of successful renames
    fail = 0 # count of failed renames

    # Iterate over each planned (source, destination) pair
    for src, dst in tasks:
        try:
            # Perform the rename on disk
            src.rename(dst)
            ok += 1
        except Exception as e:
            # If something goes wrong (permissions, locked by another program, etc.),
            # report the error and continue with the next one.
            print(f"ERROR: {src.name} -> {dst.name}: {e}")
            fail += 1

    # -------------------------------------------------------------------------
    # STEP 7: Final summary
    # -------------------------------------------------------------------------
    print(f"\nDone. Successfully renamed: {ok}. Failed: {fail}.")

# ------------------------------------------------------------------------------
# Python entry point — this runs main() when you execute the script directly.
# ------------------------------------------------------------------------------
if __name__ == "__main__":
    main()
