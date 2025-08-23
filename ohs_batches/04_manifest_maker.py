#!/usr/bin/env python3
"""
One flat zip + one manifest with description.

Behavior
- Crawl all subdirectories of the current folder.
- Add every file to a single ZIP at the archive root (no folders).
- record_id = first 6 chars of the file's immediate parent folder name.
- description = filename stem with any leading folder markers removed:
    * first 6 chars of the parent folder
    * OR the full parent folder name
    * then strip leading separators (_ - space)
  If nothing remains, description is empty.
- Manifest columns: filename,record_id,description.
- If a name collision occurs in the ZIP root, append _1, _2, ...

Outputs (in the folder you run this from)
- combined_files.zip
- veracross-mapping-file.csv
"""

import os
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import csv

ZIP_NAME = "combined_files.zip"
MANIFEST_NAME = "veracross-mapping-file.csv"
SEP_SET = {"_", "-", " "}  # separators to strip if they follow the marker

def unique_name(name: str, used: set[str]) -> str:
    """Ensure 'name' is unique within 'used' by appending _N before the extension."""
    if name not in used:
        used.add(name)
        return name
    stem, dot, ext = name.partition(".")
    n = 1
    while True:
        candidate = f"{stem}_{n}.{ext}" if dot else f"{stem}_{n}"
        if candidate not in used:
            used.add(candidate)
            return candidate
        n += 1

def derive_description(filename_stem: str, parent_name: str, rid6: str) -> str:
    """
    Remove a leading folder marker from filename_stem.
    Marker preference:
      1) exact parent folder name
      2) first 6 chars (record id)
    After removing, strip leading separators. Return remaining text (may be '').
    """
    s = filename_stem

    # Try full parent name first
    if s.startswith(parent_name):
        s = s[len(parent_name):]
    # Else try the 6-char record id
    elif s.startswith(rid6):
        s = s[len(rid6):]

    # Strip leading separators repeatedly
    while s and s[0] in SEP_SET:
        s = s[1:]

    return s  # may be empty

def main():
    root = Path.cwd()

    # Collect all files under subdirectories
    workload: list[tuple[Path, str, str]] = []  # (file_path, record_id, description)
    for dirpath, dirnames, filenames in os.walk(root):
        parent = Path(dirpath)
        if parent == root:
            continue  # exclude files at the root
        rid6 = parent.name[:6]
        for fname in filenames:
            src = parent / fname
            if src.name in {ZIP_NAME, MANIFEST_NAME} and parent == root:
                continue
            stem = src.stem
            desc = derive_description(stem, parent.name, rid6)
            workload.append((src, rid6, desc))

    if not workload:
        print("No files found in subdirectories.")
        return

    used_names: set[str] = set()
    zip_path = root / ZIP_NAME
    manifest_path = root / MANIFEST_NAME

    with ZipFile(zip_path, mode="w", compression=ZIP_DEFLATED) as zf, \
         open(manifest_path, "w", newline="", encoding="utf-8") as mf:
        writer = csv.writer(mf)
        writer.writerow(["filename", "record_id", "description"])

        for src, rid, desc in workload:
            final_name = unique_name(src.name, used_names)
            zf.write(src, arcname=final_name)  # flat placement
            writer.writerow([final_name, rid, desc])

    print(f"Created: {zip_path.name}")
    print(f"Created: {manifest_path.name}")
    print(f"Files packed: {len(workload)}")

if __name__ == "__main__":
    main()
