#!/usr/bin/env python3
"""
==============================================================
Two-Pass Filename Normalizer using Folder's First 7 Characters
==============================================================

What this script does
---------------------
You run this script inside a TOP folder. It scans EVERY subfolder under it
(files directly in TOP are ignored).

For each file in each subfolder, it performs TWO PASSES:

PASS 1: REMOVE
  - Take the first 7 characters of the file's *immediate* parent folder name.
    That 7-char string is called KEY.
  - Remove ALL occurrences of KEY from the *filename stem* (the name without
    the extension). This collapses repeated embedded copies of KEY.

PASS 2: ENFORCE PREFIX
  - Ensure the filename now STARTS with exactly KEY.
  - If it doesn’t, we add KEY in front of the stem (no extra separators).
  - Result: each filename contains KEY exactly once, at the very beginning.

Safety
------
- The script never overwrites an existing file. If the target name exists,
  it auto-appends _1, _2, etc., before the extension.
- Default is DRY RUN (no renames). Add --apply to actually rename files.

Usage
-----
Preview (no changes):
    python normalize_filenames_with_key.py

Apply changes:
    python normalize_filenames_with_key.py --apply
"""

import os
import sys
import argparse
from pathlib import Path
from typing import List, Tuple

# ----------------------------
# Console progress bar utility
# ----------------------------
def render_progress(done: int, total: int, phase: str, extra: str = "", width: int = 40) -> None:
    """
    Draw a single-line progress bar:
        [#####........] 12/100 | phase=PASS1 | removed=34

    Parameters
    ----------
    done : int
        Number of items processed so far.
    total : int
        Total number of items to process.
    phase : str
        A short label indicating which pass is running (e.g., PASS1, PASS2).
    extra : str
        Additional short status text to display to the right.
    width : int
        Visual width of the bar.
    """
    total = max(total, 1)  # avoid division by zero
    ratio = min(max(done / total, 0.0), 1.0)
    filled = int(ratio * width)
    bar = "#" * filled + "." * (width - filled)
    msg = f"\r[{bar}] {done}/{total} | phase={phase}"
    if extra:
        msg += f" | {extra}"
    sys.stdout.write(msg)
    sys.stdout.flush()
    if done >= total:
        sys.stdout.write("\n")

# --------------------------
# Safe unique destination path
# --------------------------
def unique_path(dst: Path) -> Path:
    """
    Ensure we do not overwrite. If 'dst' exists, append _1, _2, ...
    before the extension until unique, then return that path.
    """
    if not dst.exists():
        return dst
    stem, ext = dst.stem, dst.suffix
    n = 1
    while True:
        candidate = dst.with_name(f"{stem}_{n}{ext}")
        if not candidate.exists():
            return candidate
        n += 1

# -------------------------------------
# Helpers for string and name processing
# -------------------------------------
SEP_SET = {"_", "-", " ", "."}  # characters we strip if they lead a name

def strip_leading_separators(s: str) -> str:
    """
    Remove leading separator characters repeatedly: _, -, space, or dot.
    Example: "__- file" -> "file"
    """
    i = 0
    while i < len(s) and s[i] in SEP_SET:
        i += 1
    return s[i:]

def remove_all_key_occurrences(stem: str, key: str) -> Tuple[str, int]:
    """
    Remove ALL non-overlapping occurrences of 'key' from 'stem'.
    Return (new_stem, removed_count).
    """
    if not key:
        return stem, 0
    count = stem.count(key)
    if count == 0:
        return stem, 0
    new_stem = stem.replace(key, "")  # remove every occurrence
    return new_stem, count

# ------------------------
# Workload construction
# ------------------------
def collect_files_under_subfolders(root: Path) -> List[Path]:
    """
    Build a list of Path objects for every file that resides in a subfolder
    of 'root'. Files directly in 'root' are skipped.
    """
    files: List[Path] = []
    for dirpath, dirnames, filenames in os.walk(root):
        parent = Path(dirpath)
        if parent == root:
            # Skip top-level files. Only process files inside subfolders.
            continue
        for fname in filenames:
            files.append(parent / fname)
    return files

def folder_key_for(file_path: Path) -> str:
    """
    Derive KEY = the first 7 characters of the file's *immediate* parent folder name.
    If the folder name is shorter than 7, use what exists.
    """
    return file_path.parent.name[:7]

# ------------------------
# Main two-pass procedure
# ------------------------
def pass1_remove(root: Path, files: List[Path], apply_changes: bool) -> Tuple[int, int]:
    """
    PASS 1: For each file, remove ALL occurrences of KEY from its stem, then
    strip leading separators that may remain. If the name changes, rename it.

    Returns (changed_count, total_key_occurrences_removed)
    """
    changed = 0
    removed_total = 0
    total = len(files)
    done = 0
    render_progress(done, total, phase="PASS1", extra="removed=0")

    for src in files:
        stem, ext = src.stem, src.suffix
        key = folder_key_for(src)

        # Remove all instances of KEY from the filename stem
        new_stem, removed = remove_all_key_occurrences(stem, key)

        # Clean leftover leading separators after removal
        new_stem = strip_leading_separators(new_stem)

        if removed > 0 and new_stem != stem:
            dst = src.with_name(f"{new_stem}{ext}")
            dst = unique_path(dst)
            if apply_changes:
                try:
                    src.rename(dst)
                    changed += 1
                    removed_total += removed
                except Exception as e:
                    sys.stdout.write(f"\nERROR (PASS1): {src.name} -> {dst.name}: {e}\n")
            else:
                changed += 1
                removed_total += removed

        done += 1
        render_progress(done, total, phase="PASS1", extra=f"removed={removed_total}")

    return changed, removed_total

def pass2_prefix(root: Path, files: List[Path], apply_changes: bool) -> int:
    """
    PASS 2: Ensure the filename starts with KEY exactly once.
    After PASS 1, KEY should no longer appear elsewhere in the name.
    If the name does not start with KEY, we prefix KEY (no extra separators).

    Returns changed_count for PASS 2.
    """
    changed = 0
    total = len(files)
    done = 0
    render_progress(done, total, phase="PASS2")

    for src in files:
        stem, ext = src.stem, src.suffix
        key = folder_key_for(src)

        # If already starts with KEY, nothing to do.
        if key and not stem.startswith(key):
            final_stem = f"{key}{stem}"  # enforce KEY at start
            dst = src.with_name(f"{final_stem}{ext}")
            dst = unique_path(dst)
            if apply_changes:
                try:
                    src.rename(dst)
                    changed += 1
                except Exception as e:
                    sys.stdout.write(f"\nERROR (PASS2): {src.name} -> {dst.name}: {e}\n")
            else:
                changed += 1

        done += 1
        render_progress(done, total, phase="PASS2")

    return changed

# -------------
# Entry point
# -------------
def main():
    # Parse the --apply flag. Default is dry run.
    ap = argparse.ArgumentParser(
        description=("Two-pass normalization of filenames using the first 7 "
                     "characters of each file's parent folder: "
                     "PASS1 removes all occurrences; PASS2 enforces as prefix."))
    ap.add_argument("--apply", action="store_true", help="Perform renames. Omit for dry run.")
    args = ap.parse_args()
    apply_changes = args.apply

    root = Path.cwd()
    print(f"Root: {root}")
    print(f"Mode: {'APPLY' if apply_changes else 'DRY RUN'}")
    print("Building workload...")

    files = collect_files_under_subfolders(root)
    total = len(files)
    if total == 0:
        print("No files found under subfolders. Nothing to do.")
        return

    print(f"Files discovered: {total}\n")

    # PASS 1: remove all occurrences of KEY within stems
    p1_changed, p1_removed = pass1_remove(root, files, apply_changes)
    print(f"\nPASS 1 summary: changed={p1_changed}, key-occurrences-removed={p1_removed}")

    # Refresh file list because some paths may have changed after PASS 1
    files = collect_files_under_subfolders(root)

    # PASS 2: enforce KEY as the filename prefix (exactly once)
    p2_changed = pass2_prefix(root, files, apply_changes)
    print(f"\nPASS 2 summary: changed={p2_changed}")

    # Final recap
    total_changes = p1_changed + p2_changed
    if apply_changes:
        print(f"\nAll done. Applied changes: {total_changes}")
    else:
        print(f"\nDry run complete. Would apply changes: {total_changes}")
        print("Re-run with --apply to commit.")

if __name__ == "__main__":
    main()
