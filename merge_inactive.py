import pandas as pd

# Load the CSV file
file_path = 'Master Inactive Students - master_student_data_aligned (1).csv'
df = pd.read_csv(file_path)

# Define a function to merge duplicates based on the most complete row per 'UNIQUE ID'
def merge_duplicates(df, unique_id_col):
    merged_rows = []

    # Group by 'UNIQUE ID' and process each group
    for _, group in df.groupby(unique_id_col):
        # Find the row with the most non-null values (most complete data)
        most_complete_row = group.loc[group.notna().sum(axis=1).idxmax()]

        # Fill in missing values in the most complete row with values from other rows in the group
        for _, row in group.iterrows():
            most_complete_row = most_complete_row.combine_first(row)

        # Append the merged row to the result list
        merged_rows.append(most_complete_row)

    # Convert the list of merged rows to a DataFrame
    return pd.DataFrame(merged_rows)

# Apply the deduplication process
deduped_df = merge_duplicates(df, unique_id_col='UNIQUE ID')

# Save the deduplicated DataFrame to a new CSV file
output_path = 'master_student_data_merged_deduped.csv'
deduped_df.to_csv(output_path, index=False)

print(f"Deduplicated and merged file saved as {output_path}")
