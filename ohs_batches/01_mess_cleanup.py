#!/usr/bin/env python3
"""
==============================================================
Delete .ini and .db files from current folder and subfolders
==============================================================

What this script does
---------------------
1. Looks at the folder where you run the script.
2. Searches through every subfolder inside it.
3. Finds all files that end with ".ini" or ".db".
4. Counts how many such files exist.
5. Announces how many it will delete.
6. Deletes them one by one.
7. While deleting, shows progress in the form: [x/N] Deleted: filename
   where x is the number deleted so far and N is the total number.

Important notes
---------------
- Files will be permanently deleted. They will not go to the Recycle Bin.
- You can stop the script at any time by pressing CTRL+C.
- If a file cannot be deleted (for example, it is in use), the script will
  tell you that it failed for that file and then move on.
"""

# We import the modules (collections of ready-made code) we need.
import os       # lets us walk through folders and files
import sys      # lets us print progress on the same line
from pathlib import Path  # easier and safer way to handle file/folder paths

# Define which file types we want to remove.
# ".ini" and ".db" are the two file extensions.
TARGET_EXTS = {".ini", ".db"}

def find_targets(root: Path):
    """
    Look inside 'root' folder and all of its subfolders.
    Build a list of files that end with .ini or .db.

    Parameters
    ----------
    root : Path
        The folder where we begin our search.

    Returns
    -------
    targets : list of Path objects
        Each item is a file path that should be deleted.
    """
    targets = []  # an empty list to store all matches we find

    # os.walk goes through every folder and subfolder step by step
    for dirpath, dirnames, filenames in os.walk(root):
        for fname in filenames:
            fpath = Path(dirpath) / fname  # full path to this file
            # Check if the file’s extension (suffix) is either .ini or .db
            if fpath.suffix.lower() in TARGET_EXTS:
                targets.append(fpath)  # add this file to our list

    return targets  # send back the list of files to delete

def delete_files(files):
    """
    Delete each file in the 'files' list and show progress.

    Parameters
    ----------
    files : list of Path objects
        The list of files we want to delete.
    """
    total = len(files)  # how many files we need to delete

    if total == 0:
        # If there are no files to delete, announce that and stop here
        print("No .ini or .db files found.")
        return

    # Tell the user how many files we will delete
    print(f"Deleting {total} file(s)...")

    deleted = 0  # counter of how many we actually removed

    # Go through each file one by one
    # enumerate gives us both a counter (idx) and the file itself (fpath)
    for idx, fpath in enumerate(files, start=1):
        try:
            # Try to remove the file from disk
            fpath.unlink()
            deleted += 1
            # Print progress on the same line, overwriting as we go
            sys.stdout.write(f"\r[{idx}/{total}] Deleted: {fpath}")
            sys.stdout.flush()
        except Exception as e:
            # If deleting fails (file in use, permission denied, etc.)
            sys.stdout.write(f"\r[{idx}/{total}] Failed to delete {fpath}: {e}\n")
            sys.stdout.flush()

    # When the loop finishes, move to a new line and show summary
    print(f"\nCompleted. Deleted {deleted} file(s).")

if __name__ == "__main__":
    # Path.cwd() is "current working directory",
    # meaning the folder where you ran the script.
    root = Path.cwd()
    print(f"Scanning: {root}")

    # First find all target files
    targets = find_targets(root)

    # Then delete them while showing progress
    delete_files(targets)
