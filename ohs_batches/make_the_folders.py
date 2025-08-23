# Import the pandas library.
# Pandas is a Python tool that makes it easy to read and work with data from files such as CSVs.
import pandas as pd

# Import the os (operating system) library.
# This library lets Python interact with the computer's file system:
# - to find out the current working directory
# - to create new folders
import os

# Step 1: Read the CSV file into a "DataFrame" (a kind of table).
# Replace the file name below with the exact name of your CSV file.
# The file is expected to have a header row with the column "FileFolderName".
df = pd.read_csv("admission candidate - Copy of FullProspectList.csv")

# Step 2: Get the location where this script is running.
# All new folders will be created inside this directory.
base_dir = os.getcwd()

# Step 3: Loop through each row in the column called "FileFolderName".
# .dropna() ensures that if there are any blank entries, they are skipped.
for folder_name in df["FileFolderName"].dropna():
    # Make sure the folder name is a clean string (remove any extra spaces).
    folder_name = str(folder_name).strip()
    
    # Combine the base directory path with the folder name
    # to get the full location of where the folder should be created.
    folder_path = os.path.join(base_dir, folder_name)
    
    # Step 4: Create the folder.
    # os.makedirs will create the folder if it does not already exist.
    # exist_ok=True means: if the folder is already there, do nothing (no error).
    os.makedirs(folder_path, exist_ok=True)

# End of script.
# After running this, you will see one folder created for each row in the CSV file,
# inside the same folder where this script is located.
