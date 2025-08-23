import os
import shutil
from tqdm import tqdm  # For progress bar

# === Explanation for Non-Coders ===
# This script does the following:
# 1. Walks through every folder in a given "root" directory.
# 2. If a folder named "Portfolio" exists, it takes all files inside it.
# 3. Moves those files into the parent folder (one level up).
# 4. Shows a progress bar with how many files are moved.

# === Configuration ===
# Change this to the top-level directory where all your folders are stored.
root_directory = r"U:\General Portfolio"

# === Step 1: Collect all portfolio files ===
portfolio_files = []  # Will store (source, destination) pairs

for current_path, folders, files in os.walk(root_directory):
    # Check if this is a "Portfolio" subdirectory
    if os.path.basename(current_path).lower() == "portfolio":
        parent_dir = os.path.dirname(current_path)  # Go one level up
        for file in files:
            source = os.path.join(current_path, file)       # Where the file is now
            destination = os.path.join(parent_dir, file)    # Where it should go
            portfolio_files.append((source, destination))   # Save for later moving

# === Step 2: Move files with progress bar ===
print(f"Found {len(portfolio_files)} files to move.\n")

for source, destination in tqdm(portfolio_files, desc="Moving files", unit="file"):
    try:
        # If a file with same name already exists in parent, rename to avoid overwrite
        if os.path.exists(destination):
            base, ext = os.path.splitext(destination)
            counter = 1
            new_destination = f"{base}_{counter}{ext}"
            while os.path.exists(new_destination):
                counter += 1
                new_destination = f"{base}_{counter}{ext}"
            destination = new_destination

        shutil.move(source, destination)  # Move the file
    except Exception as e:
        print(f"Could not move {source}: {e}")

print("\nFile moving complete.")
