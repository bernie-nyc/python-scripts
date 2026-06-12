import os
import shutil
from tqdm import tqdm  # for a progress bar

# === CONFIGURATION ===
# Change this to the top-level directory where your folders are stored
root_directory = r"U:\General Portfolio"

# === Step 1: Collect all Portfolio folders and files ===
portfolio_tasks = []   # list of (source_file, destination_file)
portfolio_folders = [] # list of portfolio folder paths

for current_path, folders, files in os.walk(root_directory):
    if os.path.basename(current_path).lower() == "portfolio":
        parent_dir = os.path.dirname(current_path)
        portfolio_folders.append(current_path)

        for file in files:
            source = os.path.join(current_path, file)
            destination = os.path.join(parent_dir, file)
            portfolio_tasks.append((source, destination))

# === Step 2: Move files ===
print(f"Found {len(portfolio_tasks)} file(s) to move from {len(portfolio_folders)} Portfolio folder(s).\n")

for source, destination in tqdm(portfolio_tasks, desc="Moving files", unit="file"):
    try:
        # Ensure no overwrite — if a duplicate exists, rename with _1, _2...
        if os.path.exists(destination):
            base, ext = os.path.splitext(destination)
            counter = 1
            new_destination = f"{base}_{counter}{ext}"
            while os.path.exists(new_destination):
                counter += 1
                new_destination = f"{base}_{counter}{ext}"
            destination = new_destination

        shutil.move(source, destination)
    except Exception as e:
        print(f"\nCould not move {source}: {e}")

# === Step 3: Remove empty Portfolio folders ===
for folder in portfolio_folders:
    try:
        if not os.listdir(folder):  # check if folder is empty
            os.rmdir(folder)
            print(f"Removed empty folder: {folder}")
        else:
            print(f"Skipped (not empty): {folder}")
    except Exception as e:
        print(f"Could not remove folder {folder}: {e}")

print("\nCleanup complete.")
