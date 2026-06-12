#!/usr/bin/env python3
"""
==============================================================
Rename Files by Prepending Parent Folder Prefix
==============================================================

What it does
------------
- Starts in the folder where you run this script.
- Goes through every subfolder (and its subfolders).
- For each file inside:
    * Takes the subfolder's name.
    * Extracts the first 7 characters (or fewer if shorter).
    * Puts those characters at the **front** of the original filename.
    * Keeps the original file extension.
- Does NOT add extra underscores.

Example
-------
Folder:   C:\Base\ADMISSI_
File:     resume.docx
Result:   ADMISSI_resume.docx
(because "ADMISSI_" is the first 7 characters of the folder name)

--------------------------------------------------------------
Usage
--------------------------------------------------------------
1. Save as: rename_with_prefix.py
2. Open PowerShell in the target root folder.
3. Run:
       python rename_with_prefix.py
   By default it only shows what it *would* do.
4. When satisfied, set DRY_RUN = False at the bottom of the file,
   and run it again to apply the renames.
"""

import os
from pathlib import Path

def rename_files(root: Path, dry_run: bool = True):
    """
    Crawl through 'root' recursively and rename files.

    Parameters
    ----------
    root : Path
        The top folder where the scan begins.
    dry_run : bool
        If True, nothing is renamed, only printed.
        If False, actual renaming happens.
    """

    # Walk through every folder and file inside 'root'
    for dirpath, dirnames, filenames in os.walk(root):
        parent = Path(dirpath)

        # Skip the root folder itself; only rename inside subfolders
        if parent == root:
            continue

        # Take the first 7 characters of the folder name
        prefix = parent.name[:7]

        # Process each file in this subfolder
        for fname in filenames:
            src = parent / fname   # the current file
            stem = src.stem        # filename without extension
            ext = src.suffix       # extension, like ".docx"

            # Build new filename: prefix + original filename + extension
            new_name = f"{prefix}{stem}{ext}"
            dst = parent / new_name

            # If a file with the new name already exists, skip it
            if dst.exists():
                print(f"SKIP (would overwrite): {src} -> {dst}")
                continue

            # Either preview or actually rename
            if dry_run:
                print(f"Would rename: {src} -> {dst}")
            else:
                print(f"Renaming: {src} -> {dst}")
                src.rename(dst)

if __name__ == "__main__":
    # Folder where script is run
    root = Path.cwd()
    print(f"Scanning: {root}")

    # Dry run first to preview
    DRY_RUN = False
    rename_files(root, dry_run=DRY_RUN)

    # To apply changes, set:
    # DRY_RUN = False
    # and run again
