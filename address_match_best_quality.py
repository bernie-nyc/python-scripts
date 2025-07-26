
import pandas as pd                      # Library for handling spreadsheets and tabular data
from fuzzywuzzy import process          # Library to find close (fuzzy) string matches
from tqdm import tqdm                   # Library to show progress bars
import time                             # Standard Python library to measure time elapsed

# Step 1: Load the CSV files exported from your system
# Ensure these CSV files are in the same folder as this script when you run it
dev_batch = pd.read_csv("Development Batch - Updated_Development_Batch_2025.csv")
person_query = pd.read_csv("Person Query (1124 records) 2025-07-26.csv")

# Step 2: Prepare and normalize address data
# This means: make everything lowercase, remove extra spaces, and join address + city
# This helps make fuzzy comparisons more consistent and accurate
dev_batch['address_key'] = (
    dev_batch['person_address_1'].fillna('').str.lower().str.strip() + ', ' +
    dev_batch['person_city'].fillna('').str.lower().str.strip()
)

person_query['address_key'] = (
    person_query['Address 1'].fillna('').str.lower().str.strip() + ', ' +
    person_query['City'].fillna('').str.lower().str.strip()
)

# Step 3: Function to find the best match based on address
# - It compares the current row's address against all addresses in the person query list
# - If a match is close enough (similarity score >= threshold), we accept it
# - We return the matched Person ID and set "new_person" flag to NO
# - If no match is good enough, we leave person_id blank and set "new_person" to YES
def best_address_match(row, threshold=90):
    target = row['address_key']  # The address we're trying to match
    result = process.extractOne(target, person_query['address_key'])  # Find best match
    if result and result[1] >= threshold:
        matched_row = person_query[person_query['address_key'] == result[0]]
        return matched_row['Person ID'].values[0], 'NO'
    return '', 'YES'

# Step 4: Loop through each row in the development batch and match addresses
# We'll track time, progress, and results
start_time = time.time()  # Start a stopwatch to measure how long this takes
person_ids = []           # List to store matched Person IDs
new_person_flags = []     # List to store YES/NO if person is new

print("Starting fuzzy address matching...")

# The tqdm() function shows a live progress bar in the terminal
for row in tqdm(dev_batch.itertuples(index=False), total=len(dev_batch), desc="Matching"):
    pid, new_flag = best_address_match(row)
    person_ids.append(pid)
    new_person_flags.append(new_flag)

# Step 5: Store the results back into the spreadsheet
dev_batch['person_id'] = person_ids
dev_batch['new_person'] = new_person_flags

# Step 6: Save the updated file with matches
# This creates a new CSV file in the same folder with the results
dev_batch.to_csv("Development_Batch_Matched.csv", index=False)

# Step 7: Print summary statistics
match_count = sum(flag == 'NO' for flag in new_person_flags)
no_match_count = len(dev_batch) - match_count
elapsed = time.time() - start_time

print("\n--- Matching Summary ---")
print(f"Total Records Processed: {len(dev_batch)}")
print(f"Matched Records: {match_count}")
print(f"Unmatched Records: {no_match_count}")
print(f"Total Time: {elapsed:.2f} seconds")
