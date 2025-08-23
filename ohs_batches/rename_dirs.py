#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
WHAT THIS SCRIPT DOES (PLAIN ENGLISH)
================================================================================
Goal:
  Rename folders whose names are 8 digits (a "Legacy Person ID") using a CSV.

New folder name format:
  <PersonID6>_<First-Last>_<LegacyID8>

Example:
  A folder named 12305549 becomes 109303_Bleu-Anderson_12305549

Rules (strict):
  1) Process the CSV row by row from top to bottom.
  2) For each CSV row, look for a TOP-LEVEL folder whose name is EXACTLY the
     8-digit Legacy Person ID for that row. If it exists, rename it.
  3) If the same Legacy ID appears again later in the CSV, the second time will
     SKIP because the first rename already “consumed” that folder.
  4) Person ID must be a 6-digit number (like 109303). If not, that row is skipped.
  5) Full Name in the CSV is "Last, First". The script converts this to "First-Last".
     Spaces are preserved.
  6) The script NEVER adds suffixes (_1, _2, …). If the target name already exists,
     it SKIPS instead of guessing a new name.
  7) The script reports what it did and lists the CSV row indices for rows that were
     skipped and why (bad legacy, bad person ID, no matching folder, or target exists).

Safety:
  - DRY RUN by default (no changes). Use --apply to actually rename.
  - Windows-safe naming: illegal characters are replaced with underscores.
  - Robust CSV reading: handles UTF-8 and UTF-16 exports and removes hidden control
    characters that can corrupt IDs.

How to run:
  1) Put this script in the folder ABOVE all the 8-digit folders you want to rename.
     Example structure:
       C:\Data   <-- run the script here
         ├─ 12305549
         ├─ 12305383
         └─ Person Query (...).csv
  2) Dry run (no changes):
       python rename_dirs.py --csv "Person Query (1582 records) 2025-08-23.csv"
  3) Apply changes:
       python rename_dirs.py --csv "Person Query (1582 records) 2025-08-23.csv" --apply

What the “CSV row index” means in the report:
  - The CSV has a header line. The FIRST data row is index 2.
  - If the report says “row 51 skipped (no matching dir)”, it means the 51st line
    in the CSV file (counting the header as line 1) did not find a folder to rename.

================================================================================
"""

# =========================
# IMPORTS (built-in only)
# =========================
import argparse      # reads command-line flags like --csv and --apply
import csv           # reads the CSV file with column headers
import io            # in-memory text stream for decoded CSV
import re            # regular expressions for cleaning text and finding digits
import sys           # prints error messages to standard error
from pathlib import Path  # handles file/folder paths safely on Windows


# =============================================================================
# WINDOWS-SAFE NAME HELPERS
# These functions make sure each piece of the folder name works on Windows.
# =============================================================================

# Windows forbids these characters in file/folder names:  < > : " / \ | ? *
# Also control characters (ASCII 0–31) are invalid. We will replace them with "_".
WIN_ILLEGAL = r'<>:"/\\|?*\x00-\x1F'

# Precompile find/replace patterns once for speed and clarity.
_ILLEGAL_RE   = re.compile(f"[{re.escape(WIN_ILLEGAL)}]")
# Windows does not allow folder names to end with a dot or space.
_TRAILING_RE  = re.compile(r"[\. ]+$")
# Windows reserved device names that cannot be used as folder names (any case).
_RESERVED     = {
    "CON", "PRN", "AUX", "NUL",
    *{f"COM{i}" for i in range(1, 10)},
    *{f"LPT{i}" for i in range(1, 10)},
}

def sanitize_component(text: str) -> str:
    """
    Make a SINGLE path component safe on Windows.

    Steps:
      1) Replace all illegal characters with "_".
      2) Remove ending dots/spaces.
      3) If empty afterwards, use "_".
      4) If the part before the first dot is a reserved device name, append "_".
    """
    # Replace forbidden characters
    safe = _ILLEGAL_RE.sub("_", text or "")
    # Remove trailing dots/spaces
    safe = _TRAILING_RE.sub("", safe)
    # Avoid empty component
    if not safe:
        safe = "_"
    # Avoid reserved device names
    base = safe.split(".", 1)[0]
    if base.upper() in _RESERVED:
        safe += "_"
    return safe


# =============================================================================
# TEXT CLEANING HELPERS
# These functions normalize CSV values: strip control characters, trim spaces,
# pull out digit sequences, and reformat names.
# =============================================================================

# Control characters (ASCII 0–31). These can sneak in from UTF-16 or bad exports.
_CTRL_RE = re.compile(r"[\x00-\x1F]+")

def strip_ctrl(s: str) -> str:
    """Remove hidden control characters. Prevents weird underscores later."""
    return _CTRL_RE.sub("", s or "")

def norm_ws(s: str) -> str:
    """
    Trim ends and collapse interior whitespace to a single space.
    Example: "  Jane   Q.   Public  " -> "Jane Q. Public"
    """
    return re.sub(r"\s+", " ", (s or "").strip())

def digits_only(s: str) -> str:
    """Keep digits 0-9 only. Example: 'ID: 123-456' -> '123456'."""
    return re.sub(r"\D", "", s or "")

def pid_six(s: str) -> str:
    """
    Extract a STRICT 6-digit person id.
    Steps:
      - Remove all non-digits.
      - Take the first run of 6 digits.
    If none found, return empty string.
    """
    d = digits_only(s)
    m = re.search(r"\d{6}", d)
    return m.group(0) if m else ""

def first_dash_last(name: str) -> str:
    """
    Convert 'Last, First' into 'First-Last' with spaces preserved.
    If no ", " found, just normalize whitespace and return as-is.
    """
    n = norm_ws(strip_ctrl(name))
    if ", " in n:
        last, first = n.split(", ", 1)
        return f"{first}-{last}"
    return n


# =============================================================================
# ROBUST CSV DECODING
# Many school or SIS exports are UTF-16. This handles UTF-8 and UTF-16 safely.
# =============================================================================

def decode_csv(csv_path: Path) -> io.StringIO:
    """
    Read raw bytes from disk. Detect UTF-16 by BOM or by the presence of NUL bytes.
    Decode to text. Wrap in a StringIO so csv.DictReader can read it.
    """
    raw = csv_path.read_bytes()

    # UTF-16 with BOM (byte-order mark)
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        text = raw.decode("utf-16")
    # UTF-16 without BOM often shows lots of NUL bytes
    elif raw.count(b"\x00") > 0:
        try:
            text = raw.decode("utf-16-le")
        except UnicodeDecodeError:
            text = raw.decode("utf-16", errors="strict")
    # Otherwise assume UTF-8 (with or without BOM)
    else:
        text = raw.decode("utf-8-sig")

    return io.StringIO(text)

def hdrmap(fields):
    """
    Build a case/space-insensitive map: "legacy person id" -> actual header text.
    This tolerates minor header differences like extra spaces.
    """
    m = {}
    for h in fields or []:
        key = re.sub(r"\s+", " ", (h or "")).strip().lower()
        m[key] = h
    return m


# =============================================================================
# CSV ROW ITERATOR (WITH 1-BASED INDICES)
# Yields tuples describing each CSV row after validation/cleanup.
# =============================================================================

def iter_rows(csv_path: Path):
    """
    Yield a 4-tuple per CSV data row:
      (row_index, legacy8 | None, pid6 | None, first_dash_last | None)

    Notes:
      - row_index counts the CSV header as line 1, so first data row is 2.
      - If legacy is invalid (not exactly 8 digits), return (idx, None, None, None).
      - If pid is invalid (no 6-digit run), return (idx, legacy, None, None).
      - Otherwise return cleaned values for all three.
    """
    f = decode_csv(csv_path)
    r = csv.DictReader(f)

    # Map flexible header keys to actual header names
    h = hdrmap(r.fieldnames)
    col_legacy = h.get("legacy person id")
    col_pid    = h.get("person id")
    col_name   = h.get("full name")

    if not (col_legacy and col_pid and col_name):
        print("ERROR: CSV must have headers: Person ID, Full Name, Legacy Person ID", file=sys.stderr)
        return

    row_idx = 1  # header line
    for row in r:
        row_idx += 1

        # Clean and validate LEGACY (must be exactly 8 digits)
        legacy = digits_only(strip_ctrl(row.get(col_legacy, "")))
        if len(legacy) != 8:
            # Invalid legacy; cannot proceed with this row
            yield (row_idx, None, None, None)
            continue

        # Clean and validate PID (must be exactly 6 digits)
        pid6 = pid_six(row.get(col_pid, ""))
        if len(pid6) != 6:
            # Invalid PID; report legacy but mark PID as bad
            yield (row_idx, legacy, None, None)
            continue

        # Reformat full name into "First-Last"
        name = first_dash_last(row.get(col_name, ""))

        # Return cleaned results
        yield (row_idx, legacy, pid6, name)


# =============================================================================
# MAIN WORKFLOW
# =============================================================================

def main():
    # -----------------------------
    # 1) Read command-line options
    # -----------------------------
    parser = argparse.ArgumentParser(
        description="Rename <Legacy8> → <PID6>_<First-Last>_<Legacy8> using a CSV, with strict 1:1 matching and a detailed skip report."
    )
    parser.add_argument(
        "--csv",
        required=True,
        help="Path to the Person Query CSV (must contain: Person ID, Full Name, Legacy Person ID)."
    )
    parser.add_argument(
        "--apply",
        action="store_true",
        help="Actually perform the renames. If omitted, the script does a DRY RUN."
    )
    args = parser.parse_args()

    # -----------------------------
    # 2) Identify working paths
    # -----------------------------
    root = Path.cwd()              # the folder where you run the script
    csv_path = Path(args.csv)      # the CSV file path

    if not csv_path.exists():
        print(f"ERROR: CSV not found: {csv_path}")
        return

    print(f"Working folder: {root}")
    print(f"Mode: {'APPLY' if args.apply else 'DRY RUN (no changes)'}\n")

    # -----------------------------
    # 3) Snapshot available folders
    #    We only consider top-level directories named EXACTLY 8 digits.
    #    We store their names in a set so we can “consume” them as we go.
    # -----------------------------
    available = {
        p.name
        for p in root.iterdir()
        if p.is_dir() and re.fullmatch(r"\d{8}", p.name)
    }

    # Keep track of new names we plan to create (to prevent collisions in DRY RUN).
    planned_names = set()

    # -----------------------------
    # 4) Plan and/or perform renames
    #    We also collect detailed skip reasons with CSV row indices.
    # -----------------------------
    plan = []  # list of (src_path, dst_path, row_index)

    rows_total = 0
    skip_invalid_legacy = []   # CSV rows with bad Legacy Person ID
    skip_invalid_pid    = []   # CSV rows with bad Person ID
    skip_no_dir         = []   # CSV rows with no matching folder to rename
    skip_target_exists  = []   # CSV rows where the target name already exists

    for row_index, legacy, pid6, name in iter_rows(csv_path):
        rows_total += 1

        # Case A: Legacy invalid (not 8 digits). Record and skip.
        if legacy is None:
            skip_invalid_legacy.append(row_index)
            continue

        # Case B: PID invalid (not 6 digits). Record and skip.
        if pid6 is None:
            skip_invalid_pid.append(row_index)
            continue

        # Build the exact source folder path. Strict rule: must exist by this name now.
        src = root / legacy
        if legacy not in available:
            # No folder exists named exactly this legacy ID at this moment.
            # Could be because a previous row already renamed it.
            skip_no_dir.append(row_index)
            continue

        # Build the destination folder name. Keep spaces. Sanitize illegal chars.
        dst_name = f"{pid6}_{sanitize_component(name)}_{legacy}"
        dst = root / dst_name

        # If a folder already exists with the target name, skip.
        # Also skip if we already planned to create that same target (dry run case).
        if dst.exists() or dst_name in planned_names:
            skip_target_exists.append(row_index)
            continue

        # Record the plan and mark the legacy folder as “consumed” so later rows skip.
        plan.append((src, dst, row_index))
        planned_names.add(dst_name)
        available.remove(legacy)

        # If we are in APPLY mode, perform the rename immediately.
        if args.apply:
            try:
                src.rename(dst)
            except Exception as e:
                # If rename fails, roll back our bookkeeping so a later row could try again.
                print(f"ERROR (row {row_index}): {src.name} -> {dst.name}: {e}")
                planned_names.discard(dst_name)
                available.add(legacy)

    # -----------------------------
    # 5) Report
    # -----------------------------
    print(f"CSV rows read:              {rows_total}")
    print(f"Planned renames:            {len(plan)}")
    print(f"Skipped (invalid legacy):   {len(skip_invalid_legacy)} -> {skip_invalid_legacy}")
    print(f"Skipped (invalid personid): {len(skip_invalid_pid)}    -> {skip_invalid_pid}")
    print(f"Skipped (no matching dir):  {len(skip_no_dir)}         -> {skip_no_dir}")
    print(f"Skipped (target exists):    {len(skip_target_exists)}  -> {skip_target_exists}\n")

    # Show a preview of the first 50 planned changes so you can spot-check.
    for src, dst, idx in plan[:50]:
        print(f"row {idx}: {src.name} -> {dst.name}")
    if len(plan) > 50:
        print(f"...and {len(plan)-50} more")

    # -----------------------------
    # 6) Finish
    # -----------------------------
    if not plan:
        print("\nNothing to rename.")
        return

    if args.apply:
        print("\nDone.")
    else:
        print("\nDry run complete. Re-run with --apply to commit.")


# =============================================================================
# STANDARD PYTHON ENTRY POINT
# =============================================================================
if __name__ == "__main__":
    main()
