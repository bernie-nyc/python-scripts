#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
PURPOSE
================================================================================
Rename each TOP-LEVEL folder in the current directory when its name equals a
Legacy Person ID (8 digits) found in a CSV. The new folder name format is:

    <PersonID>_<Full Name>_<LegacyID>

Examples
--------
Before:  12345678
After:   104942_Abdulraheem, Hanan_12345678

Key Points
----------
- ONLY folders directly under the folder where you run this script are considered.
- The script MATCHES a folder by normalizing its name to digits and requiring 8.
- The CSV must have columns exactly:
      "Person ID", "Full Name", "Legacy Person ID"
- The script trims spaces:
    * Leading/trailing spaces are removed from CSV fields.
    * Multiple internal whitespace characters are collapsed to a single space.
- Windows-safe naming:
    * Illegal characters are replaced with underscores.
    * Trailing dots/spaces are removed.
    * Reserved device names (CON, NUL, COM1, etc.) are avoided.
- The script NEVER overwrites: it adds "_1", "_2", ... if a collision occurs.
- Default is DRY RUN (no changes). Add --apply to actually rename.

USAGE
-----
1) Open PowerShell or CMD in the directory containing your target folders and CSV.
2) Preview (no changes):
       python rename_dirs_from_csv_prefix.py --csv "Person Query (1582 records) 2025-08-23.csv"
3) Apply changes:
       python rename_dirs_from_csv_prefix.py --csv "Person Query (1582 records) 2025-08-23.csv" --apply

================================================================================
"""

# ==============================================================================
# IMPORTS: standard-library only; no external dependencies
# ==============================================================================
import argparse      # Parse command-line options like --csv and --apply
import csv           # Read the CSV file using column names
import re            # Regular expressions for cleaning/sanitizing text
from pathlib import Path  # Robust path handling across operating systems
from typing import Dict, Tuple, Iterable  # Type hints for clarity (optional)


# ==============================================================================
# LOW-LEVEL TEXT UTILITIES
# ==============================================================================

# A. Characters that Windows forbids in file/folder names:
#    < > : " / \ | ? * and ASCII control chars (0x00-0x1F).
WIN_ILLEGAL = r'<>:"/\\|?*\x00-\x1F'

# Precompile regex that finds any illegal character (performance + clarity).
_ILLEGAL_RE = re.compile(f"[{re.escape(WIN_ILLEGAL)}]")

# Windows forbids trailing spaces or dots in names. This regex removes them.
_TRAILING_RE = re.compile(r"[\. ]+$")

# Windows reserved device base names (case-insensitive). "COM1".."COM9", "LPT1".."LPT9".
_RESERVED = {
    "CON", "PRN", "AUX", "NUL",
    *{f"COM{i}" for i in range(1, 10)},
    *{f"LPT{i}" for i in range(1, 10)},
}


def digits_only(s: str) -> str:
    """
    PURPOSE
    -------
    Remove everything except digits 0-9.

    WHY
    ---
    Folder names should match "Legacy Person ID" as 8 digits. Some names might
    contain stray characters or spaces; this normalizes them.

    EXAMPLE
    -------
    " 12 345 678 "  -> "12345678"
    """
    return re.sub(r"\D", "", s or "")


def normalize_whitespace(s: str) -> str:
    """
    PURPOSE
    -------
    Trim leading/trailing whitespace AND collapse all internal whitespace runs
    (spaces, tabs, newlines) to a single space.

    EXAMPLES
    --------
    "  John   Q.   Public " -> "John Q. Public"
    "Jane\tDoe"             -> "Jane Doe"
    """
    s = s or ""
    # Strip ends
    s = s.strip()
    # Collapse any run of whitespace to a single space
    s = re.sub(r"\s+", " ", s)
    return s


def sanitize_component(name: str) -> str:
    """
    PURPOSE
    -------
    Make a path "component" (one folder name) safe on Windows.

    STEPS
    -----
    1) Replace illegal characters with underscores.
    2) Remove any trailing dots/spaces at the end.
    3) If the result is empty, use underscore.
    4) If the base (before first dot) is a reserved device name (e.g., "CON"),
       append underscore to avoid conflicts.

    NOTE
    ----
    We KEEP spaces internally (Windows allows spaces), but we ensure they are
    normalized beforehand via normalize_whitespace().
    """
    # Replace illegal characters
    out = _ILLEGAL_RE.sub("_", name)

    # Remove trailing dots/spaces
    out = _TRAILING_RE.sub("", out)

    # Avoid empty component
    if not out:
        out = "_"

    # Avoid reserved device names (consider the part before the first dot)
    base_before_dot = out.split(".", 1)[0]
    if base_before_dot.upper() in _RESERVED:
        out = out + "_"

    return out


def ensure_unique_path(dst: Path) -> Path:
    """
    PURPOSE
    -------
    Guarantee that the returned path does NOT already exist. If 'dst' is free,
    return it; otherwise append _1, _2, ... until a free name is found.

    WHY
    ---
    Prevent accidental overwriting or collisions with existing folders.

    EXAMPLE
    -------
    "104942_John Doe_12345678"
    "104942_John Doe_12345678_1"
    "104942_John Doe_12345678_2"
    """
    if not dst.exists():
        return dst

    # Work with the folder name. Folders typically have no extensions, but we
    # handle "name.ext" generically just in case.
    name = dst.name
    if "." in name:
        stem, ext = name.rsplit(".", 1)
        dot = "."
    else:
        stem, ext, dot = name, "", ""

    n = 1
    while True:
        candidate = dst.with_name(f"{stem}_{n}{dot}{ext}")
        if not candidate.exists():
            return candidate
        n += 1


# ==============================================================================
# CSV LOADING
# ==============================================================================

def load_person_map(csv_path: Path) -> Dict[str, Tuple[str, str]]:
    """
    PURPOSE
    -------
    Read the "Person Query" CSV and build a fast lookup (dictionary) keyed by
    normalized 8-digit "Legacy Person ID". The value for each key is a tuple:
        (Person ID, Full Name)

    BEHAVIOR
    --------
    - We normalize "Legacy Person ID" using digits_only() and REQUIRE exactly
      8 digits to keep the entry.
    - We trim/collapse spaces in "Person ID" and "Full Name" using
      normalize_whitespace().
    - We DO NOT change case or punctuation of "Full Name"; we only sanitize
      later when building the final folder name.

    RETURNS
    -------
    Dict[str, Tuple[str, str]]
      Mapping: "12345678" -> ("104942", "Abdulraheem, Hanan")
    """
    mapping: Dict[str, Tuple[str, str]] = {}

    with csv_path.open("r", encoding="utf-8", errors="ignore", newline="") as f:
        reader = csv.DictReader(f)  # Access columns by name

        # Validate the expected columns exist. If they don't, DictReader
        # will return None for missing keys; our get() handles that safely.
        for row in reader:
            legacy_raw = row.get("Legacy Person ID", "")
            pid_raw    = row.get("Person ID", "")
            name_raw   = row.get("Full Name", "")

            # Normalize each field:
            legacy = digits_only(legacy_raw)                 # keep only digits
            pid    = normalize_whitespace(pid_raw)           # trim + collapse spaces
            name   = normalize_whitespace(name_raw)          # trim + collapse spaces

            # Keep ONLY rows with an 8-digit legacy ID
            if len(legacy) != 8:
                continue

            mapping[legacy] = (pid, name)

    return mapping


# ==============================================================================
# DIRECTORY ENUMERATION
# ==============================================================================

def top_level_dirs(root: Path) -> Iterable[Path]:
    """
    PURPOSE
    -------
    Yield every immediate subdirectory of 'root'. Files are ignored.

    NOTE
    ----
    We DO NOT recurse. Only folders directly under 'root' are considered.
    """
    for p in root.iterdir():
        if p.is_dir():
            yield p


# ==============================================================================
# MAIN ORCHESTRATION
# ==============================================================================

def main() -> None:
    """
    CONTROL FLOW (step-by-step)
    ---------------------------
    1) Parse command-line arguments:
         --csv   : required path to the Person Query CSV
         --apply : optional flag to actually perform renames
    2) Resolve the current working directory ("root") where the script runs.
    3) Load the CSV mapping of legacy -> (person id, full name).
    4) Scan immediate subfolders under 'root'.
    5) For each subfolder:
         a) Normalize its name to digits only.
         b) If exactly 8 digits and present in CSV mapping, build the NEW name:
               <PersonID>_<Full Name>_<LegacyID>
            with all components trimmed and Windows-sanitized.
         c) If the destination name already exists, add _1, _2, ...
         d) Add this pair (src -> dst) to a plan list.
    6) Report the plan (how many will be renamed; show first 50).
    7) If --apply was given, perform the rename operations and report success/fail.
       Otherwise, end after the dry-run preview.
    """
    # 1) Parse CLI args
    parser = argparse.ArgumentParser(
        description="Prefix top-level folders with PersonID and Full Name from CSV, keeping LegacyID at the end."
    )
    parser.add_argument(
        "--csv",
        required=True,
        help="Path to the Person Query CSV with columns: 'Person ID', 'Full Name', 'Legacy Person ID'"
    )
    parser.add_argument(
        "--apply",
        action="store_true",
        help="Actually perform the renames. If omitted, this is a dry run."
    )
    args = parser.parse_args()

    # 2) Determine 'root' (where the script is run) and CSV path
    root = Path.cwd()
    csv_path = Path(args.csv)

    # 2a) Basic existence check for the CSV
    if not csv_path.exists():
        print(f"ERROR: CSV not found: {csv_path}")
        return

    # 3) Load the mapping from CSV
    person_map = load_person_map(csv_path)

    # High-level status banner
    print(f"Working folder: {root}")
    print(f"Usable CSV mappings (8-digit legacy keys): {len(person_map)}")
    print(f"Mode: {'APPLY (will rename)' if args.apply else 'DRY RUN (no changes)'}\n")

    # Statistics counters for transparency
    planned = []            # list of (Path src, Path dst) to rename
    skipped_non8 = 0        # folders whose names (digits-only) are not 8 digits
    skipped_missing = 0     # folders with 8-digit names that are not in the CSV
    skipped_already = 0     # folders already matching the intended name

    # 4) Enumerate top-level subdirectories
    for d in top_level_dirs(root):
        # Normalize the folder name to just digits to test if it is an 8-digit legacy
        legacy = digits_only(d.name)

        if len(legacy) != 8:
            # Not an 8-digit "Legacy Person ID"; ignore this folder
            skipped_non8 += 1
            continue

        if legacy not in person_map:
            # We have a plausible legacy folder, but there is no CSV match
            skipped_missing += 1
            continue

        # 5) Build the new name using CSV data
        person_id_raw, full_name_raw = person_map[legacy]

        # Trim and collapse spaces in both pieces (safety + cleanliness)
        person_id_norm = normalize_whitespace(person_id_raw)
        full_name_norm = normalize_whitespace(full_name_raw)

        # Sanitize for Windows filesystem (illegal chars, trailing dots/spaces, reserved names)
        safe_pid  = sanitize_component(person_id_norm)
        safe_name = sanitize_component(full_name_norm)

        # New folder name FORMAT:
        #   <PersonID>_<Full Name>_<LegacyID>
        new_name = f"{safe_pid}_{safe_name}_{legacy}"

        # Compute a destination path object under the same root
        dst = ensure_unique_path(root / new_name)

        # If the computed destination name equals the current, nothing to do
        if dst.name == d.name:
            skipped_already += 1
            continue

        # Add to the plan list
        planned.append((d, dst))

    # 6) Report the plan for verification
    print(f"Planned renames: {len(planned)}")
    print(f"  Skipped (folder name not 8 digits): {skipped_non8}")
    print(f"  Skipped (no CSV match):           {skipped_missing}")
    print(f"  Skipped (already correct):        {skipped_already}\n")

    # Show up to first 50 planned renames for quick inspection
    for src, dst in planned[:50]:
        print(f" - {src.name}  ->  {dst.name}")
    if len(planned) > 50:
        print(f"... and {len(planned)-50} more")

    # If nothing to do, exit now
    if not planned:
        print("\nNothing to rename.")
        return

    # If this is DRY RUN, stop after preview
    if not args.apply:
        print("\nDry run complete. Re-run with --apply to commit changes.")
        return

    # 7) APPLY MODE: perform the renames
    print("\nApplying renames...")
    ok = 0
    fail = 0

    for src, dst in planned:
        try:
            # The actual rename on disk happens here
            src.rename(dst)
            ok += 1
        except Exception as e:
            # If something goes wrong (locked folder, permissions, etc.), record it and keep going
            print(f"ERROR: {src.name} -> {dst.name}: {e}")
            fail += 1

    print(f"\nDone. Renamed: {ok}. Failed: {fail}.")


# ==============================================================================
# ENTRY POINT
# ==============================================================================
if __name__ == "__main__":
    main()
