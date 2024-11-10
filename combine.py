import os
import pandas as pd

# Define paths for each CSV file
csv_paths = [
'9900_Student.csv',
'9899_Student.csv',
'9798_Student.csv',
'9697_Student.csv',
'9596_Student.csv',
'9495_Student.csv',
'9394_Student.csv',
'1819_Student.csv',
'1718_Student.csv',
'1617_Student.csv',
'1516_Student.csv',
'1415_Student.csv',
'1314_Student.csv',
'1213_Student.csv',
'1112_Student.csv',
'1011_Student.csv',
'0910_Student.csv',
'0809_Student.csv',
'0708_Student.csv',
'0607_Student.csv',
'0506_Student.csv',
'0405_Student.csv',
'0304_Student.csv',
'0203_Student.csv',
'0102_Student.csv',
'0001_Student.csv',
'2425_Student.csv',
'2324_Student.csv',
'2223_Student.csv',
'2122_Student.csv',
'2021_Student.csv',
'1920_Student.csv'
]

# Load the '2122_Student.csv' file to use its headers as the master column set
master_columns_path = '2122_Student.csv'
master_df = pd.read_csv(master_columns_path)

# Get the master column headers from the 2122 CSV file
master_columns = master_df.columns

# Initialize a list to store DataFrames that match the master column headers
aligned_dataframes = []

# Process each CSV file: load, align columns to master, and append to list
for path in csv_paths:
    df = pd.read_csv(path)
    
    # Drop columns that are entirely empty
    df = df.dropna(axis=1, how='all')
    
    # Reindex the DataFrame to match the master columns, filling missing columns with NaN
    aligned_df = df.reindex(columns=master_columns)
    
    # Append the aligned DataFrame to the list
    aligned_dataframes.append(aligned_df)

# Concatenate all aligned DataFrames along rows, aligning columns by name
combined_df = pd.concat(aligned_dataframes, axis=0, ignore_index=True)

# Save the combined data to a single master CSV file
final_master_csv_path = 'master_student_data_aligned.csv'
combined_df.to_csv(final_master_csv_path, index=False)

print(f"Combined CSV file saved as {final_master_csv_path}")
