#!/usr/bin/env python3
"""
One-zip + one-manifest packer.

Behavior
- Recursively crawl all subdirectories of the current folder.
- For each file:
    * record_id = first 6 chars of the file's immediate parent folder name
    * add the file to a single zip at the root of the archive (no folders)
    * write one CSV manifest with columns: filename,record_id
- If two files would have the same name inside the zip, append a numeric
  suffix before the extension to keep names unique, and use that final
  name in the manifest.

Outputs (created in the folder you run this from)
- combined_files.zip
- manifest.csv
"""

import os
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import csv

# Config
ZIP_NAME = "combined_files.zip"
MANIFEST_NAME = "manifest.csv"

def unique_name(name: str, used: set[str]) -> str:
    """
    Ensure 'name' is unique within 'used'. If taken, append _1, _2, ...
    before the extension.
    """
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

def main():
    root = Path.cwd()

    # Collect all candidate files from subdirectories only
    workload: list[tuple[Path, str]] = []  # (file_path, record_id)
    for dirpath, dirnames, filenames in os.walk(root):
        parent = Path(dirpath)
        if parent == root:
            # Only process items inside subfolders, not files in root
            continue
        record_id = parent.name[:6]  # first six characters (shorter if folder name is shorter)
        for fname in filenames:
            src = parent / fname
            # Skip our own outputs if the script is re-run
            if src.name in {ZIP_NAME, MANIFEST_NAME} and parent == root:
                continue
            workload.append((src, record_id))

    if not workload:
        print("No files found in subdirectories. Nothing to do.")
        return

    # Create zip and manifest
    used_names: set[str] = set()
    with ZipFile(root / ZIP_NAME, mode="w", compression=ZIP_DEFLATED) as zf, \
         open(root / MANIFEST_NAME, "w", newline="", encoding="utf-8") as mf:
        writer = csv.writer(mf)
        writer.writerow(["filename", "record_id"])  # header

        for src, rid in workload:
            # Flatten into the zip root: arcname = filename only, made unique if needed
            final_name = unique_name(src.name, used_names)
            zf.write(src, arcname=final_name)
            writer.writerow([final_name, rid])

    print(f"Created: {ZIP_NAME}")
    print(f"Created: {MANIFEST_NAME}")
    print(f"Files packed: {len(workload)}")

if __name__ == "__main__":
    main()
