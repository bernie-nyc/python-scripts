#!/usr/bin/env python3
"""
==============================================================
De-duplicate repeated folder-key substrings in filenames
==============================================================

WHAT THIS DOES (PLAIN LANGUAGE)
- You run this script inside a "top" folder.
- The script looks at every subfolder inside that top folder (not the top itself).
- For each file it finds:
    1) It takes the first 7 characters of that file's immediate parent folder name.
       Example: folder "ADMISSI_extra" -> key "ADMISSI"
    2) It looks at the file's "stem" (the filename without the extension).
    3) If that 7-character key appears more than once in the stem, the script
       reduces it so the key appears exactly once. The first occurrence stays;
       all additional occurrences are removed.
    4) The file extension is preserved.
    5) If the new name already exists, the script avoids overwriting by adding
       "_1", "_2", etc. before the extension.

SAFETY / HOW TO RUN
- Default is DRY RUN (no changes). You see what *would* be renamed.
- To actually rename files, add --apply on the command line.

Examples:
    Preview only (no changes):
        python dedupe_folder_key_in_filenames.py

    Apply changes:
        python dedupe_folder_key_in_filenames.py --apply
"""

# ----------------------------
# IMPORTS: ready-made utilities
# ----------------------------
import os                  # Directory walking across subfolders
import sys                 # Low-level console output for smooth progress bar
import argparse            # Command-line argument parsing (--apply)
from pathlib import Path   # Convenient and safe path handling
from typing import Tuple   # Type hints for clarity (not required to run)

# ----------------------------------------
# TEXT UTILITY: ensure a destination is unique
# ----------------------------------------
def unique_path(dst: Path) -> Path:
    """
    If `dst` does not exist, return it unchanged.
    If it exists, append _1, _2, ... before the extension until unique.

    Why: We must never overwrite someone else's file accidentally.
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

# ----------------------------------------------------
# CORE TEXT TRANSFORM: collapse repeated key occurrences
# ----------------------------------------------------
def collapse_to_single_occurrence(s: str, key: str) -> Tuple[str, int]:
    """
    Input:
      s   = the file's stem (name without extension)
      key = the 7-character folder key

    Output:
      (new_string, removed_count)

    Behavior:
      - Keep the *first* occurrence of `key` exactly where it appears.
      - Remove all subsequent occurrences.
      - Count how many were removed.
    """
    if not key:
        return s, 0  # Edge case: empty key means nothing to do

    i = 0
    seen_first = False
    out_chars = []
    removed = 0
    klen = len(key)

    while i < len(s):
        # Look ahead klen characters to see if they match key
        if s[i:i + klen] == key:
            if not seen_first:
                # Preserve the very first occurrence
                out_chars.append(key)
                seen_first = True
            else:
                # Skip this repeated occurrence
                removed += 1
            i += klen
        else:
            # Keep ordinary characters
            out_chars.append(s[i])
            i += 1

    return "".join(out_chars), removed

# ----------------------------------------------
# PROGRESS BAR: compact, in-place console display
# ----------------------------------------------
def render_progress(done: int, total: int, width: int,
                    extra: str = "") -> None:
    """
    Draw a single-line progress bar such as:
        Progress [#####........] 12/100 | extra info

    - `done`  : how many units completed
    - `total` : total units
    - `width` : number of characters in the bar (visual width)
    - `extra` : short status text shown to the right (optional)
    """
    total = max(total, 1)              # Avoid divide-by-zero
    ratio = min(max(done / total, 0), 1)
    filled = int(ratio * width)
    bar = "#" * filled + "." * (width - filled)
    msg = f"\rProgress [{bar}] {done}/{total}"
    if extra:
        msg += f" | {extra}"
    sys.stdout.write(msg)
    sys.stdout.flush()
    if done >= total:
        sys.stdout.write("\n")         # Newline at 100%

# ------------------------
# MAIN WORKFLOW CONTROLLER
# ------------------------
def main():
    # Parse command-line arguments
    parser = argparse.ArgumentParser(
        description=(
            "Scan all subfolders. If a file's name contains the parent folder's "
            "first 7 characters multiple times, reduce it to a single occurrence."
        )
    )
    parser.add_argument(
        "--apply",
        action="store_true",
        help="Actually perform the renames. Omit for a dry run."
    )
    args = parser.parse_args()
    apply_changes = args.apply

    # Define the top folder as "where you run the script"
    root = Path.cwd()
    print(f"Root: {root}")
    print(f"Mode: {'APPLY (renaming files)' if apply_changes else 'DRY RUN (no changes)'}")

    # ---------------------------
    # DISCOVER WORK: gather files
    # ---------------------------
    # We consider *only* files inside subfolders, not files directly in root.
    workload = []  # list of (folder_path, file_name, key)
    for dirpath, dirnames, filenames in os.walk(root):
        parent = Path(dirpath)
        if parent == root:
            continue  # skip files at the top level
        key = parent.name[:7]  # 7-character substring from the folder name
        for fname in filenames:
            workload.append((parent, fname, key))

    total_files = len(workload)
    if total_files == 0:
        print("No files found in subfolders. Nothing to do.")
        return

    # -----------------------------------------
    # FIRST PASS: quick measurement for context
    # -----------------------------------------
    # Count how many files contain the key at least twice (i.e., need changes)
    needs_change = 0
    total_extra_occurrences = 0  # number of occurrences to be removed in total

    for parent, fname, key in workload:
        stem = (parent / fname).stem
        # Count raw occurrences in the stem
        count = 0
        if key:
            # Simple non-overlapping count is acceptable since key length is fixed
            i = 0
            klen = len(key)
            while i <= len(stem) - klen:
                if stem[i:i + klen] == key:
                    count += 1
                    i += klen
                else:
                    i += 1

        if count >= 2:
            needs_change += 1
            total_extra_occurrences += (count - 1)

    print(f"Files scanned: {total_files}")
    print(f"Files needing change: {needs_change}")
    print(f"Extra occurrences to remove: {total_extra_occurrences}\n")

    # --------------------------------
    # SECOND PASS: do the work with UI
    # --------------------------------
    changed = 0
    unchanged = 0
    removed_total = 0

    # Visual bar parameters
    bar_width = 40
    done = 0

    # Initial draw of the progress bar
    render_progress(done, total_files, bar_width, extra="removed=0")

    for parent, fname, key in workload:
        src = parent / fname
        stem, ext = src.stem, src.suffix

        # Transform the stem: collapse repeated key to a single occurrence
        new_stem, removed = collapse_to_single_occurrence(stem, key)

        if removed == 0:
            unchanged += 1
        else:
            # Build the destination and ensure uniqueness
            dst = parent / f"{new_stem}{ext}"
            dst = unique_path(dst)

            # Report action (dry-run vs apply)
            if apply_changes:
                try:
                    src.rename(dst)
                    changed += 1
                except Exception as e:
                    # If rename fails, treat as unchanged but print error
                    print(f"\nERROR: failed to rename '{src.name}': {e}")
            else:
                # Dry-run: we only count and display; no rename occurs
                changed += 1

            removed_total += removed

        # Update and redraw progress bar after each file
        done += 1
        render_progress(
            done, total_files, bar_width,
            extra=f"removed={removed_total}"
        )

    # -------------
    # FINAL SUMMARY
    # -------------
    print("\nSummary")
    print(f"  Changed:           {changed} {'(applied)' if apply_changes else '(would apply)'}")
    print(f"  Unchanged/skipped: {unchanged}")
    print(f"  Occurrences removed: {removed_total}")

# ---------------
# SCRIPT ENTRYPOINT
# ---------------
if __name__ == "__main__":
    main()
