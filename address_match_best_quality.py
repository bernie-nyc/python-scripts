# Importing necessary Python libraries

import pandas as pd                  # For reading and working with spreadsheet data (CSV files)
from fuzzywuzzy import process      # For matching similar-looking text using "fuzzy logic"
from tqdm import tqdm               # For showing a visual progress bar during processing
import time                         # For tracking how long the script takes to run
import os                           # For working with file names and extensions

# --- STEP 1: Ask the user to provide the file names ---

# This will prompt the user to type in the path to the input file (e.g., a list of addresses to match)
input_file = input("Enter the path to the input file (e.g., dev batch): ").strip()

# This will prompt the user to type in the path to the comparison file (e.g., a known address list to match against)
comparison_file = input("Enter the path to the comparison file (e.g., person query): ").strip()

# --- STEP 2: Load the CSV files into memory ---

# The input file (e.g., "Development Batch") is read into a table called `dev_batch`
dev_batch = pd.read_csv(input_file)

# The comparison file (e.g., "Person Query") is read into a table called `person_query`
person_query = pd.read_csv(comparison_file)

# --- STEP 3: Prepare addresses to make them easier to compare ---

# We create a new column called `address_key` for both files.
# This column combines street address and city, and makes everything lowercase with no extra spaces.
# This helps avoid mismatches due to capital letters, commas, or extra spaces.

# In the input file:
dev_batch['address_key'] = (
    dev_batch['person_address_1'].fillna('').str.lower().str.strip() + ', ' +
    dev_batch['person_city'].fillna('').str.lower().str.strip()
)

# In the comparison file:
person_query['address_key'] = (
    person_query['Address 1'].fillna('').str.lower().str.strip() + ', ' +
    person_query['City'].fillna('').str.lower().str.strip()
)

# --- STEP 4: Define the fuzzy matching logic ---

# This function will take one row from the input file, and try to find the best address match in the comparison file.
def best_address_match(row, threshold=90):
    target = row['address_key']  # Get the cleaned address from the input row

    # Use fuzzy logic to find the closest match in the comparison address list
    result = process.extractOne(target, person_query['address_key'])

    # If the best match has a high enough similarity score, we consider it a valid match
    if result and result[1] >= threshold:
        # Find the row in the comparison file that matched
        matched_row = person_query[person_query['address_key'] == result[0]]
        # Return the Person ID from that matched row, and mark as "NO" (not new)
        return matched_row['Person ID'].values[0], 'NO'

    # If no good match was found, return empty and mark as "YES" (new person)
    return '', 'YES'

# --- STEP 5: Loop through all input addresses and try to find matches ---

# Start a stopwatch to measure how long the whole process takes
start_time = time.time()

# Create empty lists to store results
person_ids = []         # List of matched Person IDs
new_person_flags = []   # List of "YES"/"NO" indicating whether it's a new person

# Print a message so the user knows the process has started
print("Starting fuzzy address matching...")

# Loop through each row in the input file, and show a progress bar
for _, row in tqdm(dev_batch.iterrows(), total=len(dev_batch), desc="Matching"):
    pid, new_flag = best_address_match(row)  # Call the function to get match results
    person_ids.append(pid)                   # Save the matched Person ID (or blank)
    new_person_flags.append(new_flag)        # Save the YES/NO flag

# --- STEP 6: Add the results to the original spreadsheet ---

# Add a new column for matched Person IDs
dev_batch['person_id'] = person_ids

# Add a column that indicates whether the person is new (i.e., no match found)
dev_batch['new_person'] = new_person_flags

# --- STEP 7: Create the name for the output file ---

# This part automatically creates a new filename with "_matched" added before the file extension.
# For example, if the input file was "addresses.csv", the output will be "addresses_matched.csv"
base, ext = os.path.splitext(input_file)
output_file = f"{base}_matched{ext}"

# Save the updated spreadsheet as a new CSV file
dev_batch.to_csv(output_file, index=False)

# --- STEP 8: Print out a summary of what happened ---

# Count how many matches we found
match_count = sum(flag == 'NO' for flag in new_person_flags)
no_match_count = len(dev_batch) - match_count
elapsed = time.time() - start_time

# Print results for the user
print("\n--- Matching Summary ---")
print(f"Input File: {input_file}")
print(f"Comparison File: {comparison_file}")
print(f"Output File: {output_file}")
print(f"Total Records Processed: {len(dev_batch)}")
print(f"Matched Records: {match_count}")
print(f"Unmatched Records: {no_match_count}")
print(f"Total Time: {elapsed:.2f} seconds")
